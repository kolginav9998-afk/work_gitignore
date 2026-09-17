Option Explicit

' PRIME_04_Posting
' Единственный posting engine (single_posting_engine=true). Все формы (Заказы, Выдачи,
' цеховые/офисные листы, Возвраты, Перемещения, Инвентаризация) строят PrimeDocPlan и вызывают
' PRIME_PostDocument - сами НИКОГДА не пишут в DB_PRIME_* напрямую
' (ui_forms_must_not_write_movements_directly).
'
' PRIME 2.1.0: добавлены контуры остатка (STOCK_CONTOUR), lot-lineage для TRANSFER/ADJUSTMENT,
' committed_only_everywhere для Returns/Allocations, allocation-level учёт частичных возвратов,
' резервирование остатка между строками одного документа (ISSUE/TRANSFER), отложенное создание
' нового товара (только после полной валидации плана), лимит строк документа поднят до 1000.

Type PrimeDocLine
    ProductCode As String        ' если пусто - будет создан новый товар (только когда это осознанно разрешено вызывающим)
    ProductName As String        ' для создания карточки товара, если ProductCode пуст
    IsNewProduct As Boolean      ' 2.1.0 (R11): True, если валидация решила, что это будет новый товар -
                                 ' фактическое создание отложено до записи (после полной валидации плана)
    QtyInput As Double
    UnitInput As String
    LocationFrom As String
    LocationTo As String
    Contour As String            ' 2.1.0: контур для RECEIPT/ISSUE/RETURN/ADJUSTMENT (см. PRIME_ContourForSheet)
    ContourFrom As String        ' 2.1.0: используется только TRANSFER
    ContourTo As String          ' 2.1.0: используется только TRANSFER
    DestinationProject As String ' "Назначение / проект" - обязательное поле, теряемое в 1.4.1
    Recipient As String
    Comment As String
    Price As Double
    OriginalDocLineId As String  ' для RETURN: ссылка на исходную строку выдачи
    OrderLineId As String        ' 2.1.0: для RECEIPT из "Заказы" - ссылка на _PRIME_LineID (полный snapshot)
    QtyBase As Double            ' заполняется на этапе валидации (после конвертации единиц)
    LotId As String              ' заполняется на этапе валидации (для однопартийных операций) либо пусто (FIFO по нескольким партиям)
End Type

Type PrimeDocPlan
    DocType As String
    DocDate As String
    SourceSheet As String
    SourceKey As String
    OrderId As String
    Lines(999) As PrimeDocLine
    LineCount As Long
End Type

Private gLastPostError As String
Private gLastPostDocId As String

Public Function PRIME_LastPostError() As String
    PRIME_LastPostError = gLastPostError
End Function

Public Sub PRIME_InitPlan(ByRef plan As PrimeDocPlan, ByVal docType As String, ByVal sourceSheet As String, ByVal sourceKey As String)
    plan.DocType = docType
    plan.DocDate = Format(Now, "YYYY-MM-DD")
    plan.SourceSheet = sourceSheet
    plan.SourceKey = sourceKey
    plan.OrderId = ""
    plan.LineCount = 0
End Sub

' 999, не UBound(plan.Lines): StarBasic не компилирует UBound() на поле-массиве внутри Type -
' граница держится в отдельной константе, синхронно с "Lines(999) As PrimeDocLine" в Type.
Public Sub PRIME_PlanAddLine(ByRef plan As PrimeDocPlan, ByRef docLine As PrimeDocLine)
    If plan.LineCount > PRIME_DOC_PLAN_MAX_LINE_INDEX Then
        Err.Raise 1010, "PRIME_Posting.PRIME_PlanAddLine", "Документ превышает лимит строк одной операции (1000). Разбейте на несколько проведений."
    End If
    If docLine.Contour = "" Then docLine.Contour = SC_GENERAL
    plan.Lines(plan.LineCount) = docLine
    plan.LineCount = plan.LineCount + 1
End Sub

' === Главная точка входа =====================================================================
' Возвращает DOC_ID успешно проведённого (или уже ранее проведённого - идемпотентность) документа,
' либо "" при ошибке (подробности - PRIME_LastPostError()).
Public Function PRIME_PostDocument(ByRef plan As PrimeDocPlan) As String
    gLastPostError = ""
    Dim opId As String
    Dim docId As String
    Dim txWritten As Boolean
    txWritten = False
    Dim committedKeyRegistered As Boolean
    committedKeyRegistered = False

    ' operation_lock (2.0.1): PRIME_TryEnter теперь реально возвращает False, если проведение
    ' уже идёт - раньше результат игнорировался, и повторный/двойной вызов проходил насквозь.
    ' Здесь - единственное место, которое имеет право писать в DB_PRIME_*, поэтому проверка
    ' именно тут защищает вообще любой источник повторного входа (любая кнопка, любой лист).
    If Not PRIME_TryEnter() Then
        gLastPostError = "Операция уже выполняется, дождитесь завершения текущего проведения."
        PRIME_AuditLog("", STAGE_ERROR, plan.SourceSheet, "REENTRANT_CALL_REJECTED:" & plan.SourceKey)
        PRIME_PostDocument = ""
        Exit Function
    End If
    PRIME_AuditLog("", STAGE_BUTTON_ENTER, plan.SourceSheet, plan.SourceKey)

    On Error Goto PostFailed

    ' 1. Идемпотентность - проверяем ДО любых изменений справочников (закрывает дефект 1.4.1,
    ' где EnsureProductCon выполнялся до проверки повтора движения).
    Dim existingDoc As String
    existingDoc = PRIME_FindCommittedBySourceKey(plan.SourceKey)
    If existingDoc <> "" Then
        PRIME_AuditLog("", STAGE_VALIDATION_OK, plan.SourceSheet, "already-posted:" & existingDoc)
        PRIME_Leave()
        PRIME_PostDocument = existingDoc
        Exit Function
    End If

    PRIME_AuditLog("", STAGE_VALIDATION_START, plan.SourceSheet, plan.SourceKey)

    ' 2. Проверка и построение полного плана в памяти (продукты/партии/FIFO/конвертации).
    ' transaction_protocol: эта функция и всё, что она вызывает, НЕ ИМЕЮТ ПРАВА писать в
    ' DB_PRIME_* - включая создание нового товара. Для новой строки (IsNewProduct=True)
    ' допускается только READ-ONLY "подглядывание" будущего ЕИ-кода через
    ' PRIME_PeekSequenceValue (см. PRIME_ValidateReceipt) - код УЖЕ решён и попадает в план,
    ' но счётчик SYS_PRIME_SEQ и сама строка DB_PRIME_PRODUCTS физически ещё НЕ существуют.
    Dim errMsg As String
    If Not PRIME_ValidateAndExpandPlan(plan, errMsg) Then
        gLastPostError = errMsg
        PRIME_AuditLog("", STAGE_ERROR, plan.SourceSheet, errMsg)
        PRIME_Leave()
        PRIME_PostDocument = ""
        Exit Function
    End If
    PRIME_AuditLog("", STAGE_VALIDATION_OK, plan.SourceSheet, plan.SourceKey)

    ' 3. OP_ID / DOC_ID, запись PREPARED. ДО этой точки ни одна физическая запись в DB_PRIME_*
    ' не допускается ни при каких обстоятельствах (transaction_protocol) - PREPARED должен
    ' существовать раньше любой строки, которую он потенциально описывает.
    opId = PRIME_NewOpId()
    docId = "DOC-" & Format(PRIME_SequenceNext("DOC_ID"), "00000000")
    PRIME_WriteTxRow(opId, plan.SourceKey, docId, TX_PREPARED, "")
    txWritten = True
    PRIME_AuditLog(opId, STAGE_TX_PREPARED, plan.SourceSheet, docId)

    ' 3b. Теперь, когда PREPARED уже физически записан, можно писать спланированные новые
    ' товары (IsNewProduct=True, код уже решён на шаге 2 через PRIME_PeekSequenceValue). Каждая
    ' такая строка получает OP_ID=opId и остаётся невидимой обычному поиску товара
    ' (PRIME_BuildProductIndex/PRIME_ProductExists/PRIME_GetProduct), пока opId не станет
    ' COMMITTED - см. PRIME_03_Catalog.PRIME_BuildProductIndex.
    PRIME_WriteNewProductsIfAny plan, opId
    PRIME_AuditLog(opId, STAGE_PRODUCTS_WRITTEN, plan.SourceSheet, docId)

    ' 4. Пакетная запись документа/строк/партий/движений - по типу документа.
    PRIME_WriteDocumentHeader(docId, opId, plan)
    PRIME_AuditLog(opId, STAGE_DOCS_WRITTEN, plan.SourceSheet, docId)
    PRIME_WriteDocLines(docId, plan)
    PRIME_AuditLog(opId, STAGE_LINES_WRITTEN, plan.SourceSheet, docId)

    Select Case plan.DocType
        Case DOC_RECEIPT
            PRIME_PostReceiptLines(docId, opId, plan)
        Case DOC_ISSUE
            PRIME_PostIssueLines(docId, opId, plan)
        Case DOC_RETURN
            PRIME_PostReturnLines(docId, opId, plan)
        Case DOC_TRANSFER
            PRIME_PostTransferLines(docId, opId, plan)
        Case DOC_ADJUSTMENT
            PRIME_PostAdjustmentLines(docId, opId, plan)
        Case Else
            Err.Raise 1011, "PRIME_Posting.PRIME_PostDocument", "Неизвестный тип документа: " & plan.DocType
    End Select
    PRIME_AuditLog(opId, STAGE_MOVEMENTS_WRITTEN, plan.SourceSheet, docId)

    ' 5. COMMITTED - последний логический шаг перед UI/store.
    PRIME_UpdateTxState(opId, TX_COMMITTED, "")
    PRIME_RegisterCommittedKey(plan.SourceKey, docId, opId)
    committedKeyRegistered = True
    PRIME_AuditLog(opId, STAGE_TX_COMMITTED, plan.SourceSheet, docId)

    ' 6. Один store() на весь документ (даже при 1000 строк - см. R09/R12).
    PRIME_AuditLog(opId, STAGE_STORE_START, plan.SourceSheet, docId)
    ThisComponent.store()
    PRIME_AuditLog(opId, STAGE_STORE_OK, plan.SourceSheet, docId)

    PRIME_AuditLog(opId, STAGE_UI_FINALIZED, plan.SourceSheet, docId)
    PRIME_Leave()
    gLastPostDocId = docId
    PRIME_PostDocument = docId
    Exit Function

PostFailed:
    Dim errDesc As String
    errDesc = Error$ & " (Erl=" & Erl & ")"
    gLastPostError = errDesc
    If txWritten Then
        PRIME_TryMarkTxFailed(opId, errDesc)
    End If
    ' transaction_cache_store_consistency (2.0.1/2.1.0): если COMMITTED успели выставить и
    ' зарегистрировать SOURCE_KEY/OP_ID/DOC_ID (например, ThisComponent.store() провалился
    ' ПОСЛЕ шага 5), откатываем кэши вместе с TX - иначе повторная попытка того же SOURCE_KEY
    ' получит "уже проведено", а Returns/Documents/Lines этой же операции продолжат ошибочно
    ' считаться COMMITTED.
    If committedKeyRegistered Then
        PRIME_UnregisterCommittedKey(plan.SourceKey, docId, opId)
    End If
    PRIME_AuditLog(opId, STAGE_ERROR, plan.SourceSheet, errDesc)
    PRIME_Leave()
    PRIME_PostDocument = ""
End Function

' === Валидация и построение плана по типу документа =========================================
Private Function PRIME_ValidateAndExpandPlan(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    If plan.LineCount = 0 Then
        errMsg = "Документ не содержит строк."
        PRIME_ValidateAndExpandPlan = False
        Exit Function
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        ' unique_ei_per_receipt (2.1.1): приход ВСЕГДА создаёт новую складскую позицию (ЕИ-код) -
        ' даже если "Код товара" был предзаполнен (например, кнопкой "Распознать по артикулам")
        ' для удобного автозаполнения наименования/категории. Этот код НЕ используется как
        ' идентичность физического остатка на приходе (см. PRIME_ValidateReceipt) - единственное
        ' реальное требование для строки прихода - наименование (само оно уже подставлено
        ' автозаполнением, если код был указан). Для остальных типов документа (Issue/Return/
        ' Transfer/Adjustment) код обязателен и должен ссылаться на существующую позицию.
        If plan.DocType = DOC_RECEIPT Then
            If plan.Lines(i).ProductName = "" Then
                errMsg = "Строка " & (i + 1) & ": не указано наименование товара."
                PRIME_ValidateAndExpandPlan = False
                Exit Function
            End If
        ElseIf plan.Lines(i).ProductCode = "" Then
            errMsg = "Строка " & (i + 1) & ": не указан код товара (ЕИ-код)."
            PRIME_ValidateAndExpandPlan = False
            Exit Function
        ElseIf Not PRIME_ProductExists(plan.Lines(i).ProductCode) Then
            errMsg = "Строка " & (i + 1) & ": товар с кодом " & plan.Lines(i).ProductCode & " не найден."
            PRIME_ValidateAndExpandPlan = False
            Exit Function
        End If

        ' ADJUSTMENT допускает отрицательный ввод (недостача) - разница инвентаризации.
        If plan.Lines(i).QtyInput = 0 Then
            errMsg = "Строка " & (i + 1) & ": количество (разница) не может быть нулевым."
            PRIME_ValidateAndExpandPlan = False
            Exit Function
        ElseIf plan.Lines(i).QtyInput < 0 And plan.DocType <> DOC_ADJUSTMENT Then
            errMsg = "Строка " & (i + 1) & ": количество должно быть больше нуля."
            PRIME_ValidateAndExpandPlan = False
            Exit Function
        End If
    Next i

    Select Case plan.DocType
        Case DOC_RECEIPT
            PRIME_ValidateAndExpandPlan = PRIME_ValidateReceipt(plan, errMsg)
        Case DOC_ISSUE
            PRIME_ValidateAndExpandPlan = PRIME_ValidateIssue(plan, errMsg)
        Case DOC_RETURN
            PRIME_ValidateAndExpandPlan = PRIME_ValidateReturn(plan, errMsg)
        Case DOC_TRANSFER
            PRIME_ValidateAndExpandPlan = PRIME_ValidateTransfer(plan, errMsg)
        Case DOC_ADJUSTMENT
            PRIME_ValidateAndExpandPlan = PRIME_ValidateAdjustment(plan, errMsg)
        Case Else
            errMsg = "Неизвестный тип документа: " & plan.DocType
            PRIME_ValidateAndExpandPlan = False
    End Select
End Function

' Разница инвентаризации: знак сохраняется через конвертацию (недостача - отрицательная QtyBase).
' no_empty_lot_fallback (2.1.1): недостача обязана списываться с конкретной реальной партии
' (EI_CODE) - если найденных по (товар, место, контур) партий не хватает на всю недостачу
' (снимок инвентаризации устарел - движения появились уже после него), весь batch отклоняется
' ЗДЕСЬ, на этапе валидации с нулевой записью, а не молча пишется движение с пустым LOT_ID на
' этапе PRIME_PostAdjustmentLines (см. её комментарий - там это теперь Err.Raise, а не тихий
' fallback). PRIME_HasMovementsSince уже отдельно ловит часть таких случаев как "конфликт" до
' построения плана, эта проверка - вторая, финальная линия защиты прямо перед записью.
Private Function PRIME_ValidateAdjustment(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim sign As Double
        sign = Sgn(plan.Lines(i).QtyInput)
        Dim baseQty As Variant
        baseQty = PRIME_ConvertQtyToBase(plan.Lines(i).ProductCode, plan.Lines(i).UnitInput, Abs(plan.Lines(i).QtyInput))
        If IsEmpty(baseQty) Then
            errMsg = "Строка " & (i + 1) & ": нет коэффициента пересчёта для единицы """ & plan.Lines(i).UnitInput & """."
            PRIME_ValidateAdjustment = False
            Exit Function
        End If
        plan.Lines(i).QtyBase = CDbl(baseQty) * sign

        If plan.Lines(i).QtyBase < 0 Then
            Dim available As Double
            available = PRIME_LocationContourBalance(plan.Lines(i).ProductCode, plan.Lines(i).LocationTo, plan.Lines(i).Contour)
            If available < -plan.Lines(i).QtyBase - 0.0000005 Then
                errMsg = "Строка " & (i + 1) & ": недостача " & (-plan.Lines(i).QtyBase) & " по " & plan.Lines(i).ProductCode & _
                    " превышает фактически доступный остаток этой позиции (" & available & "). Обновите остатки и повторите инвентаризацию."
                PRIME_ValidateAdjustment = False
                Exit Function
            End If
        End If
    Next i
    PRIME_ValidateAdjustment = True
End Function

' unique_ei_per_receipt (2.1.1): КАЖДАЯ строка прихода - это отдельная, независимая складская
' позиция (ЕИ-код) со своим собственным остатком, даже если это буквально тот же товар, тот же
' артикул и тот же поставщик, что и в предыдущем приходе. Одинаковое наименование НИКОГДА
' автоматически не суммируется в одну позицию (см. docs/REQUIREMENTS_MATRIX.md R32) - поэтому,
' в отличие от 2.1.0, здесь больше нет ветки "код указан -> считать существующим товаром": ЛЮБАЯ
' строка прихода получает НОВЫЙ код, а любой предзаполненный ProductCode (например, через
' "Распознать по артикулам") используется только для того, чтобы автозаполнить наименование до
' этого места - на саму идентичность физического остатка он не влияет.
'
' transaction_protocol: валидация обязана быть чистой (ноль записей в DB_PRIME_*). Будущий
' ЕИ-код нужен прямо сейчас, чтобы попасть в план (шаг "построить план в памяти, включая новые
' товары и зарезервированные ЕИ-коды") - поэтому он ТОЛЬКО ПОДГЛЯДЫВАЕТСЯ через read-only
' PRIME_PeekSequenceValue (не пишет ничего), а не выдаётся через PRIME_SequenceNext/
' PRIME_NextProductCode (которые физически увеличивают и пишут счётчик). Весь вызов
' PRIME_PostDocument сериализован одним мьютексом (PRIME_TryEnter), поэтому между этим peek и
' фактической записью счётчика на этапе PRIME_WriteNewProductsIfAny никто другой не может
' вклиниться и увидеть/забрать тот же код. Коэффициент пересчёта тривиален (введённая единица
' становится базовой единицей ЭТОЙ позиции, фактор=1 - другого выбора и не может быть, раз
' позиция физически создаётся впервые). Физическая запись карточки товара откладывается до
' PRIME_WriteNewProductsIfAny - вызывается из PRIME_PostDocument ПОСЛЕ того, как
' SYS_PRIME_TX.STATE=PREPARED уже физически записан.
Private Function PRIME_ValidateReceipt(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim productCodeBase As Long
    productCodeBase = PRIME_PeekSequenceValue("PRODUCT_CODE")

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        plan.Lines(i).IsNewProduct = True
        plan.Lines(i).ProductCode = PRIME_FormatProductCode(productCodeBase + i + 1)
        plan.Lines(i).QtyBase = PRIME_RoundQty(plan.Lines(i).QtyInput) ' новая позиция: введённая единица = её базовая, фактор 1
        plan.Lines(i).LotId = "LOT-" ' финальный номер присваивается на этапе записи (PRIME_PostReceiptLines)
    Next i
    PRIME_ValidateReceipt = True
End Function

' R10 (2.1.0): резервирование остатка между строками одного документа - если строка 1 уже
' "списала" часть остатка (product|location|contour) в памяти плана, строка 2 того же ключа
' обязана проверяться против УЖЕ УМЕНЬШЕННОГО остатка, иначе обе строки могут независимо
' пройти проверку против одного и того же физического остатка и в сумме увести его в минус.
' R02: доступность считается СТРОГО по месту+контуру (PRIME_FifoLotsForProduct), а не по
' общему остатку товара по всем местам/контурам сразу.
Private Function PRIME_ValidateIssue(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim reserved As New Collection
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim baseQty As Variant
        baseQty = PRIME_ConvertQtyToBase(plan.Lines(i).ProductCode, plan.Lines(i).UnitInput, plan.Lines(i).QtyInput)
        If IsEmpty(baseQty) Then
            errMsg = "Строка " & (i + 1) & ": нет коэффициента пересчёта для единицы """ & plan.Lines(i).UnitInput & """."
            PRIME_ValidateIssue = False
            Exit Function
        End If
        plan.Lines(i).QtyBase = CDbl(baseQty)

        Dim key As String
        key = plan.Lines(i).ProductCode & "|" & plan.Lines(i).LocationFrom & "|" & plan.Lines(i).Contour

        Dim available As Double
        available = PRIME_LocationContourBalance(plan.Lines(i).ProductCode, plan.Lines(i).LocationFrom, plan.Lines(i).Contour)

        Dim alreadyReservedInPlan As Variant
        If Not PRIME_CollectionTryGet(reserved, key, alreadyReservedInPlan) Then alreadyReservedInPlan = 0#

        Dim remaining As Double
        remaining = available - CDbl(alreadyReservedInPlan)
        If remaining < plan.Lines(i).QtyBase - 0.0000005 Then
            ' no_cross_ei_fifo (2.1.1): остаток проверяется строго по ЭТОМУ ЕИ-коду на этом
            ' месте/контуре - при нехватке документ ЦЕЛИКОМ отклоняется, а не списывается
            ' частично с другой (пусть даже одноимённой) позиции. Пользователь сам добавляет
            ' вторую позицию отдельной строкой, если нужно списать с двух ЕИ-кодов сразу.
            errMsg = "Строка " & (i + 1) & ": доступно по " & plan.Lines(i).ProductCode & " на месте """ & plan.Lines(i).LocationFrom & _
                """ (" & PRIME_ContourDisplayName(plan.Lines(i).Contour) & "): " & remaining & _
                " шт. Не хватает: " & (plan.Lines(i).QtyBase - remaining) & " шт. Добавьте другую позицию (ЕИ-код) отдельной строкой. Документ не проведён целиком."
            PRIME_ValidateIssue = False
            Exit Function
        End If

        On Error Resume Next
        reserved.Remove(key)
        On Error Goto 0
        reserved.Add(CDbl(alreadyReservedInPlan) + plan.Lines(i).QtyBase, key)
    Next i
    PRIME_ValidateIssue = True
End Function

' R07/R08 (2.1.0): issuedQty/returnedQty теперь считаются только по COMMITTED документам/
' возвратам (см. PRIME_DocLineQtyBase/PRIME_AlreadyReturnedQtyBase ниже).
Private Function PRIME_ValidateReturn(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim issuedQty As Double, returnedQty As Double
        issuedQty = PRIME_DocLineQtyBase(plan.Lines(i).OriginalDocLineId)
        returnedQty = PRIME_AlreadyReturnedQtyBase(plan.Lines(i).OriginalDocLineId)

        Dim baseQty As Variant
        baseQty = PRIME_ConvertQtyToBase(plan.Lines(i).ProductCode, plan.Lines(i).UnitInput, plan.Lines(i).QtyInput)
        If IsEmpty(baseQty) Then
            errMsg = "Строка " & (i + 1) & ": нет коэффициента пересчёта для единицы """ & plan.Lines(i).UnitInput & """."
            PRIME_ValidateReturn = False
            Exit Function
        End If
        plan.Lines(i).QtyBase = CDbl(baseQty)

        If plan.Lines(i).QtyBase > (issuedQty - returnedQty) + 0.0000005 Then
            errMsg = "Строка " & (i + 1) & ": возврат " & plan.Lines(i).QtyBase & " превышает остаток к возврату " & (issuedQty - returnedQty) & "."
            PRIME_ValidateReturn = False
            Exit Function
        End If
    Next i
    PRIME_ValidateReturn = True
End Function

' R02/R05/R10 (2.1.0): доступность на "Откуда" считается строго по месту+контуру, с
' резервированием между строками одного документа (как в Issue) - несколько строк перемещения
' одного товара с одного и того же места не могут в сумме увести остаток там в минус.
Private Function PRIME_ValidateTransfer(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim reserved As New Collection
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim baseQty As Variant
        baseQty = PRIME_ConvertQtyToBase(plan.Lines(i).ProductCode, plan.Lines(i).UnitInput, plan.Lines(i).QtyInput)
        If IsEmpty(baseQty) Then
            errMsg = "Строка " & (i + 1) & ": нет коэффициента пересчёта для единицы """ & plan.Lines(i).UnitInput & """."
            PRIME_ValidateTransfer = False
            Exit Function
        End If
        plan.Lines(i).QtyBase = CDbl(baseQty)

        If plan.Lines(i).ContourFrom = plan.Lines(i).ContourTo And plan.Lines(i).LocationFrom = plan.Lines(i).LocationTo Then
            errMsg = "Строка " & (i + 1) & ": место и контур назначения совпадают с исходными - перемещение не имеет смысла."
            PRIME_ValidateTransfer = False
            Exit Function
        End If

        Dim key As String
        key = plan.Lines(i).ProductCode & "|" & plan.Lines(i).LocationFrom & "|" & plan.Lines(i).ContourFrom

        Dim available As Double
        available = PRIME_LocationContourBalance(plan.Lines(i).ProductCode, plan.Lines(i).LocationFrom, plan.Lines(i).ContourFrom)

        Dim alreadyReservedInPlan As Variant
        If Not PRIME_CollectionTryGet(reserved, key, alreadyReservedInPlan) Then alreadyReservedInPlan = 0#

        Dim remaining As Double
        remaining = available - CDbl(alreadyReservedInPlan)
        If remaining < plan.Lines(i).QtyBase - 0.0000005 Then
            errMsg = "Строка " & (i + 1) & ": на месте """ & plan.Lines(i).LocationFrom & """ (" & _
                PRIME_ContourDisplayName(plan.Lines(i).ContourFrom) & ") недостаточно остатка для перемещения (доступно " & remaining & ")."
            PRIME_ValidateTransfer = False
            Exit Function
        End If

        On Error Resume Next
        reserved.Remove(key)
        On Error Goto 0
        reserved.Add(CDbl(alreadyReservedInPlan) + plan.Lines(i).QtyBase, key)
    Next i
    PRIME_ValidateTransfer = True
End Function

' transaction_protocol: единственное место на пути проведения, которое физически пишет новые
' строки в DB_PRIME_PRODUCTS - вызывается из PRIME_PostDocument ТОЛЬКО ПОСЛЕ того, как
' SYS_PRIME_TX.STATE=PREPARED уже записан (шаг 3 в PRIME_PostDocument), НИКОГДА раньше. Код
' каждой строки уже решён на этапе валидации (PRIME_ValidateReceipt, через read-only
' PRIME_PeekSequenceValue) - здесь он только персистируется вместе с OP_ID текущей транзакции.
' НЕ используем PRIME_CreateProduct (она сама вызывает PRIME_NextProductCode/PRIME_SequenceNext
' и выдала бы ВТОРОЙ, отличный от уже запланированного, код) - пишем batch-ом напрямую и
' фиксируем финальный сдвиг счётчика одним PRIME_AdvanceSequenceTo.
' Пока OP_ID этой транзакции не станет COMMITTED, такая строка не резолвится обычным поиском
' товара как активная запись каталога (см. PRIME_03_Catalog.PRIME_BuildProductIndex) - то есть
' до COMMITTED она физически существует, но не пригодна для использования бизнес-логикой.
Private Sub PRIME_WriteNewProductsIfAny(ByRef plan As PrimeDocPlan, ByVal opId As String)
    Dim n As Long
    n = 0
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        If plan.Lines(i).IsNewProduct Then n = n + 1
    Next i
    If n = 0 Then Exit Sub

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim colOpId As Long
    colOpId = PRIME_ColIndex(headers, "OP_ID")
    Dim now As String
    now = Format(Now, "YYYY-MM-DD HH:MM:SS")

    Dim rows(n - 1) As Variant
    Dim r As Long
    r = 0
    For i = 0 To plan.LineCount - 1
        ' Dim прямо на теле For (не глубже) - тот же паттерн, что уже используется в рабочем
        ' PRIME_WriteDocLines/PRIME_WriteDocumentHeader.
        Dim row(UBound(headers)) As Variant
        If plan.Lines(i).IsNewProduct Then
            row(PRIME_ColIndex(headers, "PRODUCT_CODE")) = plan.Lines(i).ProductCode
            row(PRIME_ColIndex(headers, "PRODUCT_NAME")) = plan.Lines(i).ProductName
            row(PRIME_ColIndex(headers, "BASE_UNIT")) = plan.Lines(i).UnitInput
            row(PRIME_ColIndex(headers, "DEFAULT_LOCATION")) = plan.Lines(i).LocationTo
            row(PRIME_ColIndex(headers, "CATEGORY")) = ""
            row(PRIME_ColIndex(headers, "SUBCATEGORY")) = ""
            row(PRIME_ColIndex(headers, "RETURNABLE")) = 0
            row(PRIME_ColIndex(headers, "ACCOUNT_TYPE")) = ""
            row(PRIME_ColIndex(headers, "ACTIVE")) = 1
            row(PRIME_ColIndex(headers, "CREATED_AT")) = now
            row(PRIME_ColIndex(headers, "UPDATED_AT")) = now
            If colOpId >= 0 Then row(colOpId) = opId
            rows(r) = row
            r = r + 1
        End If
    Next i
    PRIME_AppendRowsBatch(SH_DB_PRODUCTS, rows)

    ' Персистируем сдвиг счётчика РОВНО на n - код каждой новой строки был решён на этапе
    ' валидации как PRIME_PeekSequenceValue()+1..+n (см. PRIME_ValidateReceipt); весь вызов
    ' PRIME_PostDocument сериализован PRIME_TryEnter, поэтому peek здесь гарантированно вернёт
    ' то же значение, что видела валидация этого же вызова.
    Dim seqBaseVal As Long
    seqBaseVal = PRIME_PeekSequenceValue("PRODUCT_CODE")
    PRIME_AdvanceSequenceTo "PRODUCT_CODE", seqBaseVal + n

    PRIME_InvalidateProductIndex()
End Sub

' === Запись шапки/строк документа ============================================================
Private Sub PRIME_WriteDocumentHeader(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "DOC_ID")) = docId
    row(PRIME_ColIndex(headers, "DOC_TYPE")) = plan.DocType
    row(PRIME_ColIndex(headers, "DOC_DATE")) = plan.DocDate
    row(PRIME_ColIndex(headers, "SOURCE_SHEET")) = plan.SourceSheet
    row(PRIME_ColIndex(headers, "SOURCE_KEY")) = plan.SourceKey
    row(PRIME_ColIndex(headers, "ORDER_ID")) = plan.OrderId
    row(PRIME_ColIndex(headers, "STATUS")) = "PROVEDENO"
    Dim colOpId As Long
    colOpId = PRIME_ColIndex(headers, "OP_ID")
    If colOpId >= 0 Then row(colOpId) = opId

    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_DB_DOCUMENTS, rows)
End Sub

Private Sub PRIME_WriteDocLines(ByVal docId As String, ByRef plan As PrimeDocPlan)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim rows(plan.LineCount - 1) As Variant
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim row(UBound(headers)) As Variant
        Dim lineId As String
        lineId = docId & "-L" & (i + 1)
        row(PRIME_ColIndex(headers, "DOC_LINE_ID")) = lineId
        row(PRIME_ColIndex(headers, "DOC_ID")) = docId
        row(PRIME_ColIndex(headers, "PRODUCT_CODE")) = plan.Lines(i).ProductCode
        row(PRIME_ColIndex(headers, "QTY_BASE")) = plan.Lines(i).QtyBase
        row(PRIME_ColIndex(headers, "UNIT")) = plan.Lines(i).UnitInput
        row(PRIME_ColIndex(headers, "PRICE")) = plan.Lines(i).Price
        row(PRIME_ColIndex(headers, "LOCATION_FROM")) = plan.Lines(i).LocationFrom
        row(PRIME_ColIndex(headers, "LOCATION_TO")) = plan.Lines(i).LocationTo
        row(PRIME_ColIndex(headers, "DESTINATION_PROJECT")) = plan.Lines(i).DestinationProject
        row(PRIME_ColIndex(headers, "RECIPIENT")) = plan.Lines(i).Recipient
        row(PRIME_ColIndex(headers, "COMMENT")) = plan.Lines(i).Comment
        rows(i) = row
        ' Сгенерированный DOC_LINE_ID нужен дальше (партии/allocations/returns), но не является
        ' полем плана - используем отдельный рантайм-буфер (см. gLineIdBuffer/PRIME_LineIdFor).
        gLineIdBuffer(i) = lineId
    Next i
    PRIME_AppendRowsBatch(SH_DB_DOC_LINES, rows)
End Sub

' Буфер сгенерированных DOC_LINE_ID на время одного PostDocument (не персистентно, только рантайм).
Private gLineIdBuffer(999) As String

Private Function PRIME_LineIdFor(ByVal idx As Long) As String
    PRIME_LineIdFor = gLineIdBuffer(idx)
End Function

' === RECEIPT: создаёт партию на каждую строку и одно движение прихода =======================
Private Sub PRIME_PostReceiptLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim lotRows(plan.LineCount - 1) As Variant

    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim moveRows(plan.LineCount - 1) As Variant

    Dim snapHeaders As Variant
    Dim hasSnapshot As Boolean
    ' SourceSheet, а не plan.OrderId<>"" - в батче "Провести все готовые" (несколько заказов
    ' сразу) документ-уровневый plan.OrderId может быть пуст (строки принадлежат РАЗНЫМ заказам,
    ' см. PRIME_05_Orders.PRIME_Orders_ConductAllReadyButton), но каждая СТРОКА всё равно несёт
    ' свой OrderLineId и обязана попасть в snapshot.
    hasSnapshot = PRIME_SheetExists(SH_DB_ORDER_SNAPSHOT) And plan.SourceSheet = SH_ORDERS
    Dim snapRows() As Variant
    If hasSnapshot Then
        snapHeaders = PRIME_HeaderMap(SH_DB_ORDER_SNAPSHOT)
        ReDim snapRows(plan.LineCount - 1)
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        ' transaction_protocol: любой IsNewProduct=True уже физически записан в DB_PRIME_PRODUCTS
        ' в PRIME_PostDocument ДО этого места (см. PRIME_WriteNewProductsIfAny, вызывается сразу
        ' после записи TX=PREPARED) - здесь ProductCode уже заполнен и строка товара существует.
        Dim lotId As String
        lotId = "LOT-" & Format(PRIME_SequenceNext("LOT_ID"), "00000000")
        plan.Lines(i).LotId = lotId

        Dim lotRow(UBound(lotHeaders)) As Variant
        lotRow(PRIME_ColIndex(lotHeaders, "LOT_ID")) = lotId
        lotRow(PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")) = plan.Lines(i).ProductCode
        lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DOC_ID")) = docId
        lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_LINE_ID")) = PRIME_LineIdFor(i)
        lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")) = plan.DocDate
        lotRow(PRIME_ColIndex(lotHeaders, "LOCATION")) = plan.Lines(i).LocationTo
        lotRow(PRIME_ColIndex(lotHeaders, "ORIGINAL_QTY_BASE")) = plan.Lines(i).QtyBase
        lotRow(PRIME_ColIndex(lotHeaders, "BASE_UNIT")) = PRIME_GetProductField(plan.Lines(i).ProductCode, "BASE_UNIT")
        lotRow(PRIME_ColIndex(lotHeaders, "ORIGIN")) = plan.SourceSheet
        lotRow(PRIME_ColIndex(lotHeaders, "ORDER_ID")) = plan.OrderId
        Dim colContour As Long
        colContour = PRIME_ColIndex(lotHeaders, "STOCK_CONTOUR")
        If colContour >= 0 Then lotRow(colContour) = plan.Lines(i).Contour
        lotRows(i) = lotRow

        moveRows(i) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
            lotId, plan.Lines(i).QtyBase, plan.Lines(i).LocationTo, plan.DocDate, opId, plan.Lines(i).Contour)

        If hasSnapshot Then
            snapRows(i) = PRIME_BuildOrderSnapshotRow(snapHeaders, plan, i, docId, lotId)
        End If
    Next i

    PRIME_AppendRowsBatch(SH_DB_LOTS, lotRows)
    PRIME_AuditLog(opId, STAGE_LOTS_WRITTEN, plan.SourceSheet, docId)
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRows)
    If hasSnapshot Then PRIME_AppendRowsBatch(SH_DB_ORDER_SNAPSHOT, snapRows)
End Sub

' R13 (2.1.0): полный immutable snapshot - копирует ВСЕ значимые реквизиты заказа/поставки из
' текущей строки листа "Заказы" (найденной по OrderLineId), а не только количество/партию, как
' раньше. Источник правды на момент commit - сама строка "Заказы"; после commit её можно менять
' (новая цена/поставщик следующей поставки), а снимок этой конкретной поставки не изменится.
Private Function PRIME_BuildOrderSnapshotRow(ByVal headers As Variant, ByRef plan As PrimeDocPlan, ByVal idx As Long, ByVal docId As String, ByVal lotId As String) As Variant
    Dim row(UBound(headers)) As Variant
    Dim orderRow As Variant
    orderRow = PRIME_Orders_RowByLineId(plan.Lines(idx).OrderLineId)

    Dim c As Long
    Dim lineOrderId As String
    lineOrderId = plan.OrderId
    If Not IsEmpty(orderRow) Then
        Dim oHeadersForId As Variant
        oHeadersForId = PRIME_HeaderMap(SH_ORDERS)
        Dim idCol As Long
        idCol = PRIME_ColIndex(oHeadersForId, "_PRIME_OrderID")
        If idCol >= 0 Then lineOrderId = CStr(orderRow(idCol))
    End If
    c = PRIME_ColIndex(headers, "ORDER_ID") : If c >= 0 Then row(c) = lineOrderId
    c = PRIME_ColIndex(headers, "ORDER_LINE_ID") : If c >= 0 Then row(c) = plan.Lines(idx).OrderLineId
    c = PRIME_ColIndex(headers, "RECEIPT_DOC_ID") : If c >= 0 Then row(c) = docId
    c = PRIME_ColIndex(headers, "RECEIPT_LINE_ID") : If c >= 0 Then row(c) = PRIME_LineIdFor(idx)
    c = PRIME_ColIndex(headers, "LOT_ID") : If c >= 0 Then row(c) = lotId
    c = PRIME_ColIndex(headers, "PRODUCT_CODE") : If c >= 0 Then row(c) = plan.Lines(idx).ProductCode
    c = PRIME_ColIndex(headers, "RECEIPT_DATE") : If c >= 0 Then row(c) = plan.DocDate
    c = PRIME_ColIndex(headers, "RECEIVED_QTY") : If c >= 0 Then row(c) = plan.Lines(idx).QtyBase
    c = PRIME_ColIndex(headers, "UNIT") : If c >= 0 Then row(c) = plan.Lines(idx).UnitInput
    c = PRIME_ColIndex(headers, "PRICE") : If c >= 0 Then row(c) = plan.Lines(idx).Price
    c = PRIME_ColIndex(headers, "AMOUNT") : If c >= 0 Then row(c) = plan.Lines(idx).Price * plan.Lines(idx).QtyInput
    c = PRIME_ColIndex(headers, "LOCATION") : If c >= 0 Then row(c) = plan.Lines(idx).LocationTo
    c = PRIME_ColIndex(headers, "STOCK_CONTOUR") : If c >= 0 Then row(c) = plan.Lines(idx).Contour
    c = PRIME_ColIndex(headers, "DESTINATION_PROJECT") : If c >= 0 Then row(c) = plan.Lines(idx).DestinationProject
    c = PRIME_ColIndex(headers, "COMMENT") : If c >= 0 Then row(c) = plan.Lines(idx).Comment

    If Not IsEmpty(orderRow) Then
        Dim oHeaders As Variant
        oHeaders = PRIME_HeaderMap(SH_ORDERS)
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "PRODUCT_NAME", "Полное наименование товара"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "SUPPLIER_OR_PLATFORM", "От кого / площадка"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "SELLER", "Продавец"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "SUPPLIER_CODE", "Код поставщика"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "SUPPLIER_ARTICLE", "Артикул поставщика"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "INVOICE_NUMBER", "Номер счета"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "DOCUMENT_NUMBER", "№ документа"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "ORDER_DATE", "Дата заказа"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "EXPECTED_DATE", "Ожидаемая дата"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "DOCUMENT_DATE", "Дата документа"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "ORDERED_QTY", "Количество"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "BUYER", "Покупатель"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "CATEGORY", "Категория"
        PRIME_CopyOrderField row, headers, orderRow, oHeaders, "SUBCATEGORY", "Подкатегория"
    End If

    PRIME_BuildOrderSnapshotRow = row
End Function

Private Sub PRIME_CopyOrderField(ByRef targetRow As Variant, ByVal targetHeaders As Variant, ByVal sourceRow As Variant, _
        ByVal sourceHeaders As Variant, ByVal targetColName As String, ByVal sourceColName As String)
    Dim tc As Long, sc As Long
    tc = PRIME_ColIndex(targetHeaders, targetColName)
    sc = PRIME_ColIndex(sourceHeaders, sourceColName)
    If tc >= 0 And sc >= 0 Then targetRow(tc) = sourceRow(sc)
End Sub

' Линейный поиск строки "Заказы" по _PRIME_LineID - таблица заказов умеренного размера
' (та же оценка объёма, что и для DB_PRIME_PRODUCTS/ALIASES, см. PRIME_03_Catalog).
Private Function PRIME_Orders_RowByLineId(ByVal orderLineId As String) As Variant
    If orderLineId = "" Or Not PRIME_SheetExists(SH_ORDERS) Then
        PRIME_Orders_RowByLineId = Empty
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colLineId As Long
    colLineId = PRIME_ColIndex(headers, "_PRIME_LineID")
    If colLineId < 0 Then
        PRIME_Orders_RowByLineId = Empty
        Exit Function
    End If
    Dim table As Variant
    table = PRIME_ReadTable(SH_ORDERS)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colLineId, orderLineId)
    If idx = -1 Then
        PRIME_Orders_RowByLineId = Empty
    Else
        PRIME_Orders_RowByLineId = table(idx)
    End If
End Function

' === ISSUE: FIFO-разбиение по партиям (строго место+контур), allocations, отрицательные движения ==
Private Sub PRIME_PostIssueLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim allocHeaders As Variant
    allocHeaders = PRIME_HeaderMap(SH_DB_ALLOCATIONS)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)

    Dim allocRowsBuf() As Variant
    Dim moveRowsBuf() As Variant
    Dim allocCount As Long, moveCount As Long
    allocCount = 0 : moveCount = 0
    ReDim allocRowsBuf(4000)
    ReDim moveRowsBuf(4000)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim lots() As String
        Dim balances() As Double
        PRIME_FifoLotsForProduct(plan.Lines(i).ProductCode, plan.Lines(i).LocationFrom, plan.Lines(i).Contour, lots, balances)

        Dim remaining As Double
        remaining = plan.Lines(i).QtyBase
        Dim j As Long
        For j = LBound(lots) To UBound(lots)
            If remaining <= 0.0000005 Then Exit For
            If balances(j) > 0 Then
                Dim take As Double
                take = remaining
                If balances(j) < take Then take = balances(j)

                If allocCount > UBound(allocRowsBuf) Then ReDim Preserve allocRowsBuf(UBound(allocRowsBuf) + 4000)
                allocRowsBuf(allocCount) = PRIME_BuildAllocationRow(allocHeaders, PRIME_LineIdFor(i), lots(j), take)
                allocCount = allocCount + 1

                If moveCount > UBound(moveRowsBuf) Then ReDim Preserve moveRowsBuf(UBound(moveRowsBuf) + 4000)
                moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                    lots(j), -take, plan.Lines(i).LocationFrom, plan.DocDate, opId, plan.Lines(i).Contour)
                moveCount = moveCount + 1

                remaining = remaining - take
            End If
        Next j

        If remaining > 0.0000005 Then
            ' Не должно происходить после PRIME_ValidateIssue, но перестраховка важнее красоты кода:
            ' не допускаем частично проведённый документ.
            Err.Raise 1020, "PRIME_Posting.PRIME_PostIssueLines", "Недостаточно партий для списания строки " & (i + 1) & " после валидации - проведение отменено."
        End If
    Next i

    If allocCount > 0 Then
        ReDim Preserve allocRowsBuf(allocCount - 1)
        PRIME_AppendRowsBatch(SH_DB_ALLOCATIONS, allocRowsBuf)
    End If
    PRIME_AuditLog(opId, STAGE_LOTS_WRITTEN, plan.SourceSheet, docId)
    If moveCount > 0 Then
        ReDim Preserve moveRowsBuf(moveCount - 1)
        PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRowsBuf)
    End If
End Sub

Private Function PRIME_BuildAllocationRow(ByVal headers As Variant, ByVal docLineId As String, ByVal lotId As String, ByVal qty As Double) As Variant
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "ALLOC_ID")) = "ALC-" & docLineId & "-" & lotId
    row(PRIME_ColIndex(headers, "DOC_LINE_ID")) = docLineId
    row(PRIME_ColIndex(headers, "LOT_ID")) = lotId
    row(PRIME_ColIndex(headers, "QTY_BASE")) = qty
    PRIME_BuildAllocationRow = row
End Function

' === RETURN: возврат симметричен FIFO-списанию исходной выдачи, allocation-aware (R08) ========
' 2.1.0 fix: раньше повторный частичный возврат снова проходил allocations исходной выдачи по
' порядку FIFO БЕЗ учёта того, что часть каждой allocation уже была возвращена ранее - второй
' возврат мог повторно "попасть" в ту же самую (уже частично возвращённую) партию вместо
' следующей по очереди. Теперь для каждой allocation вычитается уже возвращённое по НЕЙ ЖЕ
' (DB_PRIME_RETURN_ALLOCATIONS, только COMMITTED) - остаток по каждой партии восстанавливается
' в правильном порядке, а не только правильной суммой.
Private Sub PRIME_PostReturnLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim retHeaders As Variant
    retHeaders = PRIME_HeaderMap(SH_DB_RETURNS)
    Dim hasRetAlloc As Boolean
    hasRetAlloc = PRIME_SheetExists(SH_DB_RETURN_ALLOCATIONS)
    Dim retAllocHeaders As Variant
    If hasRetAlloc Then retAllocHeaders = PRIME_HeaderMap(SH_DB_RETURN_ALLOCATIONS)

    Dim moveRowsBuf() As Variant
    Dim retRowsBuf(plan.LineCount - 1) As Variant
    Dim retAllocRowsBuf() As Variant
    Dim moveCount As Long, retAllocCount As Long
    moveCount = 0 : retAllocCount = 0
    ReDim moveRowsBuf(4000)
    ReDim retAllocRowsBuf(4000)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim returnId As String
        returnId = docId & "-R" & (i + 1)

        Dim lots() As String
        Dim allocatedQty() As Double
        PRIME_AllocationsForDocLine(plan.Lines(i).OriginalDocLineId, lots, allocatedQty)

        Dim remaining As Double
        remaining = plan.Lines(i).QtyBase
        Dim j As Long
        If UBound(lots) >= LBound(lots) Then
            For j = LBound(lots) To UBound(lots)
                If remaining <= 0.0000005 Then Exit For
                Dim allocId As String
                allocId = "ALC-" & plan.Lines(i).OriginalDocLineId & "-" & lots(j)
                Dim alreadyReturnedHere As Double
                alreadyReturnedHere = PRIME_AlreadyReturnedForAllocation(allocId)
                Dim availableInAlloc As Double
                availableInAlloc = allocatedQty(j) - alreadyReturnedHere
                If availableInAlloc > 0.0000005 Then
                    Dim take As Double
                    take = remaining
                    If availableInAlloc < take Then take = availableInAlloc
                    ' transaction_protocol/return_destination_fix (2.1.1): возврат обязан
                    ' зачисляться туда, ОТКУДА реально была выдача - в конкретную партию lots(j)
                    ' (место+контур), а НЕ в DEFAULT_LOCATION/GENERAL по умолчанию. Читаем
                    ' фактические LOCATION/STOCK_CONTOUR исходной партии, а не line-level поля
                    ' plan.Lines(i).LocationTo/.Contour (которые до этого фикса задавались как
                    ' общий дефолт для ВСЕЙ строки возврата и игнорировали, откуда именно был
                    ' списан каждый конкретный allocation).
                    moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                        lots(j), take, PRIME_GetLotField(lots(j), "LOCATION"), plan.DocDate, opId, PRIME_GetLotField(lots(j), "STOCK_CONTOUR"))
                    moveCount = moveCount + 1
                    If hasRetAlloc Then
                        If retAllocCount > UBound(retAllocRowsBuf) Then ReDim Preserve retAllocRowsBuf(UBound(retAllocRowsBuf) + 4000)
                        retAllocRowsBuf(retAllocCount) = PRIME_BuildReturnAllocationRow(retAllocHeaders, returnId, allocId, lots(j), take, opId)
                        retAllocCount = retAllocCount + 1
                    End If
                    remaining = remaining - take
                End If
            Next j
        End If
        If remaining > 0.0000005 Then
            ' Исходная выдача не найдена по allocations (например, легаси-данные без миграции
            ' allocations) - возврат всё равно проводим на условное "безлотовое" движение,
            ' чтобы не заблокировать документ, но это ухудшает трассируемость партии.
            moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                "", remaining, plan.Lines(i).LocationTo, plan.DocDate, opId, plan.Lines(i).Contour)
            moveCount = moveCount + 1
        End If

        Dim retRow(UBound(retHeaders)) As Variant
        retRow(PRIME_ColIndex(retHeaders, "RETURN_ID")) = returnId
        retRow(PRIME_ColIndex(retHeaders, "ORIGINAL_ISSUE_DOC_LINE_ID")) = plan.Lines(i).OriginalDocLineId
        retRow(PRIME_ColIndex(retHeaders, "RETURN_DOC_ID")) = docId
        retRow(PRIME_ColIndex(retHeaders, "QTY_BASE")) = plan.Lines(i).QtyBase
        retRow(PRIME_ColIndex(retHeaders, "RETURN_DATE")) = plan.DocDate
        Dim colOpIdRet As Long
        colOpIdRet = PRIME_ColIndex(retHeaders, "OP_ID")
        If colOpIdRet >= 0 Then retRow(colOpIdRet) = opId
        retRowsBuf(i) = retRow
    Next i

    If moveCount > 0 Then
        ReDim Preserve moveRowsBuf(moveCount - 1)
        PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRowsBuf)
    End If
    PRIME_AppendRowsBatch(SH_DB_RETURNS, retRowsBuf)
    If hasRetAlloc And retAllocCount > 0 Then
        ReDim Preserve retAllocRowsBuf(retAllocCount - 1)
        PRIME_AppendRowsBatch(SH_DB_RETURN_ALLOCATIONS, retAllocRowsBuf)
    End If
End Sub

Private Function PRIME_BuildReturnAllocationRow(ByVal headers As Variant, ByVal returnId As String, ByVal allocId As String, _
        ByVal lotId As String, ByVal qty As Double, ByVal opId As String) As Variant
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "RETURN_ID")) = returnId
    row(PRIME_ColIndex(headers, "ALLOC_ID")) = allocId
    row(PRIME_ColIndex(headers, "LOT_ID")) = lotId
    row(PRIME_ColIndex(headers, "QTY_BASE")) = qty
    row(PRIME_ColIndex(headers, "OP_ID")) = opId
    PRIME_BuildReturnAllocationRow = row
End Function

' R08: сколько уже COMMITTED-возвращено по конкретной allocation (не по всей строке выдачи).
Public Function PRIME_AlreadyReturnedForAllocation(ByVal allocId As String) As Double
    If allocId = "" Or Not PRIME_SheetExists(SH_DB_RETURN_ALLOCATIONS) Then
        PRIME_AlreadyReturnedForAllocation = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_RETURN_ALLOCATIONS)
    Dim colAlloc As Long, colQty As Long, colOpId As Long
    colAlloc = PRIME_ColIndex(headers, "ALLOC_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    colOpId = PRIME_ColIndex(headers, "OP_ID")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_RETURN_ALLOCATIONS)
    Dim total As Double
    total = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If CStr(table(i)(colAlloc)) = allocId Then
                If PRIME_IsOpIdCommitted(CStr(table(i)(colOpId))) Then
                    total = total + CDbl(table(i)(colQty))
                End If
            End If
        Next i
    End If
    PRIME_AlreadyReturnedForAllocation = total
End Function

' === TRANSFER: FIFO-разбиение по партиям исходного (место,контур), lot lineage (R05) ===========
' 2.1.0 полный редизайн: раньше писал два "безлотовых" движения (LOT_ID="") - партия физически
' не сохранялась, из-за чего дальнейший FIFO по партиям товара не видел перемещённый остаток
' как отдельную партию (R05/R06). Теперь для каждого куска, взятого из исходной партии по FIFO,
' создаётся НОВАЯ партия-назначение с PARENT_LOT_ID=исходная (полная трассируемость), и именно
' она увеличивает остаток в новом месте/контуре.
Private Sub PRIME_PostTransferLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)

    Dim moveRowsBuf() As Variant
    Dim lotRowsBuf() As Variant
    Dim moveCount As Long, lotCount As Long
    moveCount = 0 : lotCount = 0
    ReDim moveRowsBuf(4000)
    ReDim lotRowsBuf(4000)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim lots() As String
        Dim balances() As Double
        PRIME_FifoLotsForProduct(plan.Lines(i).ProductCode, plan.Lines(i).LocationFrom, plan.Lines(i).ContourFrom, lots, balances)

        Dim remaining As Double
        remaining = plan.Lines(i).QtyBase
        Dim j As Long
        For j = LBound(lots) To UBound(lots)
            If remaining <= 0.0000005 Then Exit For
            If balances(j) > 0 Then
                Dim take As Double
                take = remaining
                If balances(j) < take Then take = balances(j)

                If moveCount > UBound(moveRowsBuf) - 1 Then
                    ReDim Preserve moveRowsBuf(UBound(moveRowsBuf) + 4000)
                End If
                moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                    lots(j), -take, plan.Lines(i).LocationFrom, plan.DocDate, opId, plan.Lines(i).ContourFrom)
                moveCount = moveCount + 1

                Dim destLotId As String
                destLotId = "LOT-" & Format(PRIME_SequenceNext("LOT_ID"), "00000000")
                Dim srcReceiptDate As String
                srcReceiptDate = PRIME_GetLotField(lots(j), "RECEIPT_DATE")
                If srcReceiptDate = "" Then srcReceiptDate = plan.DocDate

                If lotCount > UBound(lotRowsBuf) Then ReDim Preserve lotRowsBuf(UBound(lotRowsBuf) + 4000)
                Dim lotRow(UBound(lotHeaders)) As Variant
                lotRow(PRIME_ColIndex(lotHeaders, "LOT_ID")) = destLotId
                lotRow(PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")) = plan.Lines(i).ProductCode
                lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DOC_ID")) = docId
                lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_LINE_ID")) = PRIME_LineIdFor(i)
                lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")) = srcReceiptDate ' сохраняем возраст партии для FIFO
                lotRow(PRIME_ColIndex(lotHeaders, "LOCATION")) = plan.Lines(i).LocationTo
                lotRow(PRIME_ColIndex(lotHeaders, "ORIGINAL_QTY_BASE")) = take
                lotRow(PRIME_ColIndex(lotHeaders, "BASE_UNIT")) = PRIME_GetProductField(plan.Lines(i).ProductCode, "BASE_UNIT")
                lotRow(PRIME_ColIndex(lotHeaders, "ORIGIN")) = "TRANSFER"
                lotRow(PRIME_ColIndex(lotHeaders, "ORDER_ID")) = ""
                Dim colContourTo As Long
                colContourTo = PRIME_ColIndex(lotHeaders, "STOCK_CONTOUR")
                If colContourTo >= 0 Then lotRow(colContourTo) = plan.Lines(i).ContourTo
                Dim colParent As Long
                colParent = PRIME_ColIndex(lotHeaders, "PARENT_LOT_ID")
                If colParent >= 0 Then lotRow(colParent) = lots(j)
                lotRowsBuf(lotCount) = lotRow
                lotCount = lotCount + 1

                moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                    destLotId, take, plan.Lines(i).LocationTo, plan.DocDate, opId, plan.Lines(i).ContourTo)
                moveCount = moveCount + 1

                remaining = remaining - take
            End If
        Next j

        If remaining > 0.0000005 Then
            Err.Raise 1021, "PRIME_Posting.PRIME_PostTransferLines", "Недостаточно партий для перемещения строки " & (i + 1) & " после валидации - проведение отменено."
        End If
    Next i

    If lotCount > 0 Then
        ReDim Preserve lotRowsBuf(lotCount - 1)
        PRIME_AppendRowsBatch(SH_DB_LOTS, lotRowsBuf)
    End If
    PRIME_AuditLog(opId, STAGE_LOTS_WRITTEN, plan.SourceSheet, docId)
    If moveCount > 0 Then
        ReDim Preserve moveRowsBuf(moveCount - 1)
        PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRowsBuf)
    End If
End Sub

' === ADJUSTMENT: инвентаризация, lot-consistent (R06) ==========================================
' 2.1.0 полный редизайн: раньше писал одно "безлотовое" движение (LOT_ID="") - расходилось с
' партийным FIFO (stock по месту менялся, а сумма балансов партий - нет). Теперь недостача
' списывается СО СПЕЦИФИЧНЫХ партий по FIFO (как обычное списание), а излишек создаёт НОВУЮ
' партию происхождения "ADJUSTMENT" - после проведения stock == sum(lot balances) гарантированно.
Private Sub PRIME_PostAdjustmentLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)

    Dim moveRowsBuf() As Variant
    Dim lotRowsBuf() As Variant
    Dim moveCount As Long, lotCount As Long
    moveCount = 0 : lotCount = 0
    ReDim moveRowsBuf(4000)
    ReDim lotRowsBuf(plan.LineCount)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        If plan.Lines(i).QtyBase > 0 Then
            ' Излишек - новая партия происхождения "ADJUSTMENT".
            Dim newLotId As String
            newLotId = "LOT-" & Format(PRIME_SequenceNext("LOT_ID"), "00000000")
            Dim lotRow(UBound(lotHeaders)) As Variant
            lotRow(PRIME_ColIndex(lotHeaders, "LOT_ID")) = newLotId
            lotRow(PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")) = plan.Lines(i).ProductCode
            lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DOC_ID")) = docId
            lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_LINE_ID")) = PRIME_LineIdFor(i)
            lotRow(PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")) = plan.DocDate
            lotRow(PRIME_ColIndex(lotHeaders, "LOCATION")) = plan.Lines(i).LocationTo
            lotRow(PRIME_ColIndex(lotHeaders, "ORIGINAL_QTY_BASE")) = plan.Lines(i).QtyBase
            lotRow(PRIME_ColIndex(lotHeaders, "BASE_UNIT")) = PRIME_GetProductField(plan.Lines(i).ProductCode, "BASE_UNIT")
            lotRow(PRIME_ColIndex(lotHeaders, "ORIGIN")) = "ADJUSTMENT"
            lotRow(PRIME_ColIndex(lotHeaders, "ORDER_ID")) = ""
            Dim colContourAdj As Long
            colContourAdj = PRIME_ColIndex(lotHeaders, "STOCK_CONTOUR")
            If colContourAdj >= 0 Then lotRow(colContourAdj) = plan.Lines(i).Contour
            lotRowsBuf(lotCount) = lotRow
            lotCount = lotCount + 1

            moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                newLotId, plan.Lines(i).QtyBase, plan.Lines(i).LocationTo, plan.DocDate, opId, plan.Lines(i).Contour)
            moveCount = moveCount + 1
        Else
            ' Недостача - списываем со специфичных партий по FIFO (место+контур), как ISSUE.
            Dim shortage As Double
            shortage = -plan.Lines(i).QtyBase
            Dim lots() As String
            Dim balances() As Double
            PRIME_FifoLotsForProduct(plan.Lines(i).ProductCode, plan.Lines(i).LocationTo, plan.Lines(i).Contour, lots, balances)
            Dim j As Long
            For j = LBound(lots) To UBound(lots)
                If shortage <= 0.0000005 Then Exit For
                If balances(j) > 0 Then
                    Dim take As Double
                    take = shortage
                    If balances(j) < take Then take = balances(j)
                    If moveCount > UBound(moveRowsBuf) Then ReDim Preserve moveRowsBuf(UBound(moveRowsBuf) + 4000)
                    moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                        lots(j), -take, plan.Lines(i).LocationTo, plan.DocDate, opId, plan.Lines(i).Contour)
                    moveCount = moveCount + 1
                    shortage = shortage - take
                End If
            Next j
            ' no_empty_lot_fallback (2.1.1): PRIME_ValidateAdjustment уже отклонила бы весь batch
            ' до этой точки, если найденных партий не хватает на всю недостачу - см. её
            ' комментарий. Если это условие всё же достигнуто (гонка между validation и записью
            ' внутри ОДНОГО сериализованного PRIME_TryEnter-вызова теоретически невозможна, но
            ' second line of defense), НЕ пишем движение с пустым LOT_ID (untraceable stock) -
            ' поднимаем ошибку, откатывая всю транзакцию, как и любую другую внутреннюю ошибку.
            If shortage > 0.0000005 Then
                Err.Raise 1012, "PRIME_Posting.PRIME_PostAdjustmentLines", _
                    "Внутренняя ошибка: недостача по " & plan.Lines(i).ProductCode & " превышает найденные партии после прохождения валидации."
            End If
        End If
    Next i

    If lotCount > 0 Then
        ReDim Preserve lotRowsBuf(lotCount - 1)
        PRIME_AppendRowsBatch(SH_DB_LOTS, lotRowsBuf)
    End If
    If moveCount > 0 Then
        ReDim Preserve moveRowsBuf(moveCount - 1)
        PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRowsBuf)
    End If
End Sub

Private Function PRIME_BuildMovementRow(ByVal headers As Variant, ByVal docId As String, ByVal docLineId As String, _
        ByVal productCode As String, ByVal lotId As String, ByVal qtyBase As Double, ByVal location As String, _
        ByVal moveDate As String, ByVal opId As String, ByVal stockContour As String) As Variant
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "MOVE_ID")) = "MOV-" & Format(PRIME_SequenceNext("MOVE_ID"), "00000000")
    row(PRIME_ColIndex(headers, "DOC_ID")) = docId
    row(PRIME_ColIndex(headers, "DOC_LINE_ID")) = docLineId
    row(PRIME_ColIndex(headers, "PRODUCT_CODE")) = productCode
    row(PRIME_ColIndex(headers, "LOT_ID")) = lotId
    row(PRIME_ColIndex(headers, "QTY_BASE")) = qtyBase
    row(PRIME_ColIndex(headers, "LOCATION")) = location
    row(PRIME_ColIndex(headers, "MOVE_DATE")) = moveDate
    row(PRIME_ColIndex(headers, "OP_ID")) = opId
    Dim colContour As Long
    colContour = PRIME_ColIndex(headers, "STOCK_CONTOUR")
    If colContour >= 0 Then row(colContour) = stockContour
    PRIME_BuildMovementRow = row
End Function

' === FIFO / остатки по партиям =================================================================
' R02 (2.1.0): партии товара строго на конкретном (месте, контуре), отсортированные по дате
' прихода и LOT_ID (fifo.sort_order), с текущим балансом (сумма COMMITTED-движений). Только
' партии с положительным балансом полезны для списания, но возвращаем все для прозрачности.
Public Sub PRIME_FifoLotsForProduct(ByVal productCode As String, ByVal location As String, ByVal contour As String, ByRef lots() As String, ByRef balances() As Double)
    If Not PRIME_SheetExists(SH_DB_LOTS) Then
        ReDim lots(-1) : ReDim balances(-1)
        Exit Sub
    End If
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim colCode As Long, colLotId As Long, colDate As Long, colLoc As Long, colContour As Long
    colCode = PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")
    colLotId = PRIME_ColIndex(lotHeaders, "LOT_ID")
    colDate = PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")
    colLoc = PRIME_ColIndex(lotHeaders, "LOCATION")
    colContour = PRIME_ColIndex(lotHeaders, "STOCK_CONTOUR")

    Dim lotTable As Variant
    lotTable = PRIME_ReadTable(SH_DB_LOTS)
    If UBound(lotTable) < 1 Then
        ReDim lots(-1) : ReDim balances(-1)
        Exit Sub
    End If

    ' Собираем список ID партий товара с датой, затем сортируем простой сортировкой вставками
    ' (число партий одного товара обычно невелико - десятки/сотни, не тысячи).
    Dim candLots() As String
    Dim candDates() As String
    Dim n As Long
    n = 0
    ReDim candLots(UBound(lotTable))
    ReDim candDates(UBound(lotTable))
    Dim i As Long
    For i = 1 To UBound(lotTable)
        If CStr(lotTable(i)(colCode)) = productCode Then
            If location = "" Or CStr(lotTable(i)(colLoc)) = location Then
                If contour = "" Or colContour < 0 Or CStr(lotTable(i)(colContour)) = contour Then
                    candLots(n) = CStr(lotTable(i)(colLotId))
                    candDates(n) = CStr(lotTable(i)(colDate))
                    n = n + 1
                End If
            End If
        End If
    Next i

    If n = 0 Then
        ReDim lots(-1) : ReDim balances(-1)
        Exit Sub
    End If

    Dim k As Long, m As Long
    For k = 1 To n - 1
        Dim keyDate As String, keyLot As String
        keyDate = candDates(k) : keyLot = candLots(k)
        m = k - 1
        Do While m >= 0 And (candDates(m) > keyDate Or (candDates(m) = keyDate And candLots(m) > keyLot))
            candDates(m + 1) = candDates(m)
            candLots(m + 1) = candLots(m)
            m = m - 1
        Loop
        candDates(m + 1) = keyDate
        candLots(m + 1) = keyLot
    Next k

    ReDim lots(n - 1)
    ReDim balances(n - 1)
    For i = 0 To n - 1
        lots(i) = candLots(i)
        balances(i) = PRIME_LotBalance(candLots(i))
    Next i
End Sub

' Тот же список партий, но БЕЗ фильтра по месту/контуру - только для информационных/read-only
' экранов (Комплекты - предварительная проверка достаточности, Поиск - "Партии товара"), где
' пользователь ещё не выбрал конкретное место списания. НЕ использовать при реальном проведении.
Public Sub PRIME_FifoLotsForProductAny(ByVal productCode As String, ByRef lots() As String, ByRef balances() As Double)
    PRIME_FifoLotsForProduct(productCode, "", "", lots, balances)
End Sub

' committed_only_stock (2.0.1): суммирует ТОЛЬКО движения, чей OP_ID зафиксирован как
' COMMITTED в SYS_PRIME_TX. Раньше суммировались ВСЕ строки DB_PRIME_MOVEMENTS независимо от
' состояния транзакции - движения, физически записанные на шаге 4 PRIME_PostDocument, но не
' доведённые до TX_COMMITTED (сбой/ошибка между записью движений и коммитом, либо между
' коммитом и успешным store()), ошибочно увеличивали остаток, хотя операция не завершилась
' или была помечена FAILED. См. golden invariant "Only COMMITTED movements affect stock.".
Public Function PRIME_LotBalance(ByVal lotId As String) As Double
    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        PRIME_LotBalance = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colLot As Long, colQty As Long, colOpId As Long
    colLot = PRIME_ColIndex(headers, "LOT_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    colOpId = PRIME_ColIndex(headers, "OP_ID")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    Dim total As Double
    total = 0
    Dim i As Long
    If UBound(table) >= 1 Then
        For i = 1 To UBound(table)
            If CStr(table(i)(colLot)) = lotId Then
                If PRIME_IsOpIdCommitted(CStr(table(i)(colOpId))) Then
                    total = total + CDbl(table(i)(colQty))
                End If
            End If
        Next i
    End If
    PRIME_LotBalance = total
End Function

' R02: остаток строго по (товар, место, контур) - сумма балансов партий, попадающих под эти
' три ключа. Единственный источник правды для проверки доступности при ISSUE/TRANSFER.
Public Function PRIME_LocationContourBalance(ByVal productCode As String, ByVal location As String, ByVal contour As String) As Double
    Dim lots() As String
    Dim balances() As Double
    PRIME_FifoLotsForProduct(productCode, location, contour, lots, balances)
    Dim total As Double
    total = 0
    If UBound(lots) >= LBound(lots) Then
        Dim i As Long
        For i = LBound(balances) To UBound(balances)
            total = total + balances(i)
        Next i
    End If
    PRIME_LocationContourBalance = total
End Function

' Общий остаток товара по ВСЕМ местам/контурам сразу - только для информационных сводок
' (Комплекты, "Наличие"), не для проверки доступности конкретной операции (см. R02 выше).
Public Function PRIME_TotalLotBalance(ByVal productCode As String) As Double
    Dim lots() As String
    Dim balances() As Double
    PRIME_FifoLotsForProductAny(productCode, lots, balances)
    Dim total As Double
    total = 0
    If UBound(lots) >= LBound(lots) Then
        Dim i As Long
        For i = LBound(balances) To UBound(balances)
            total = total + balances(i)
        Next i
    End If
    PRIME_TotalLotBalance = total
End Function

' Значение произвольного поля партии по LOT_ID (используется TRANSFER, чтобы унаследовать
' RECEIPT_DATE исходной партии в новую партию-назначение и не "омолаживать" её FIFO-возраст).
Public Function PRIME_GetLotField(ByVal lotId As String, ByVal fieldName As String) As String
    If lotId = "" Or Not PRIME_SheetExists(SH_DB_LOTS) Then
        PRIME_GetLotField = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_LOTS)
    Dim colLotId As Long, colField As Long
    colLotId = PRIME_ColIndex(headers, "LOT_ID")
    colField = PRIME_ColIndex(headers, fieldName)
    If colField < 0 Then
        PRIME_GetLotField = ""
        Exit Function
    End If
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_LOTS)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colLotId, lotId)
    If idx = -1 Then
        PRIME_GetLotField = ""
    Else
        PRIME_GetLotField = CStr(table(idx)(colField))
    End If
End Function

' Разбиение исходной строки выдачи по партиям (для симметричного возврата).
Public Sub PRIME_AllocationsForDocLine(ByVal docLineId As String, ByRef lots() As String, ByRef qtys() As Double)
    If docLineId = "" Or Not PRIME_SheetExists(SH_DB_ALLOCATIONS) Then
        ReDim lots(-1) : ReDim qtys(-1)
        Exit Sub
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_ALLOCATIONS)
    Dim colLine As Long, colLot As Long, colQty As Long
    colLine = PRIME_ColIndex(headers, "DOC_LINE_ID")
    colLot = PRIME_ColIndex(headers, "LOT_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_ALLOCATIONS)
    If UBound(table) < 1 Then
        ReDim lots(-1) : ReDim qtys(-1)
        Exit Sub
    End If

    Dim n As Long
    n = 0
    ReDim lots(UBound(table))
    ReDim qtys(UBound(table))
    Dim i As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(colLine)) = docLineId Then
            lots(n) = CStr(table(i)(colLot))
            qtys(n) = CDbl(table(i)(colQty))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim lots(-1) : ReDim qtys(-1)
    Else
        ReDim Preserve lots(n - 1)
        ReDim Preserve qtys(n - 1)
    End If
End Sub

' R07 (2.1.0): учитывает только строки COMMITTED документа - раньше читал DB_PRIME_DOC_LINES
' без проверки состояния владеющей транзакции, поэтому строка документа, физически записанная
' до отката на FAILED, ошибочно считалась бы существующей выдачей при проверке возврата.
Public Function PRIME_DocLineQtyBase(ByVal docLineId As String) As Double
    If docLineId = "" Then
        PRIME_DocLineQtyBase = 0
        Exit Function
    End If
    If Not PRIME_IsDocIdCommitted(PRIME_DocIdFromLineId(docLineId)) Then
        PRIME_DocLineQtyBase = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim colLine As Long, colQty As Long
    colLine = PRIME_ColIndex(headers, "DOC_LINE_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colLine, docLineId)
    If idx = -1 Then
        PRIME_DocLineQtyBase = 0
    Else
        PRIME_DocLineQtyBase = CDbl(table(idx)(colQty))
    End If
End Function

' R07 (2.1.0): суммирует только COMMITTED возвраты (см. PRIME_DbReturnsColumns - колонка OP_ID).
Public Function PRIME_AlreadyReturnedQtyBase(ByVal originalDocLineId As String) As Double
    If originalDocLineId = "" Or Not PRIME_SheetExists(SH_DB_RETURNS) Then
        PRIME_AlreadyReturnedQtyBase = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_RETURNS)
    Dim colOrig As Long, colQty As Long, colOpId As Long
    colOrig = PRIME_ColIndex(headers, "ORIGINAL_ISSUE_DOC_LINE_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    colOpId = PRIME_ColIndex(headers, "OP_ID")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_RETURNS)
    Dim total As Double
    total = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If CStr(table(i)(colOrig)) = originalDocLineId Then
                Dim countIt As Boolean
                countIt = True
                If colOpId >= 0 Then countIt = PRIME_IsOpIdCommitted(CStr(table(i)(colOpId)))
                If countIt Then total = total + CDbl(table(i)(colQty))
            End If
        Next i
    End If
    PRIME_AlreadyReturnedQtyBase = total
End Function

' === Журнал транзакций (SYS_PRIME_TX) и аудит-лог ============================================
Private Function PRIME_NewOpId() As String
    Randomize
    PRIME_NewOpId = "OP-" & Format(Now, "YYYYMMDDHHMMSS") & "-" & Format(Int(Rnd * 999999), "000000")
End Function

Private Sub PRIME_WriteTxRow(ByVal opId As String, ByVal sourceKey As String, ByVal docId As String, ByVal state As String, ByVal errText As String)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "OP_ID")) = opId
    row(PRIME_ColIndex(headers, "SOURCE_KEY")) = sourceKey
    row(PRIME_ColIndex(headers, "DOC_ID")) = docId
    row(PRIME_ColIndex(headers, "STATE")) = state
    row(PRIME_ColIndex(headers, "STARTED_AT")) = Format(Now, "YYYY-MM-DD HH:MM:SS")
    row(PRIME_ColIndex(headers, "COMMITTED_AT")) = ""
    row(PRIME_ColIndex(headers, "HASH")) = ""
    row(PRIME_ColIndex(headers, "ERROR")) = errText
    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_SYS_TX, rows)
End Sub

Private Sub PRIME_UpdateTxState(ByVal opId As String, ByVal newState As String, ByVal errText As String)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim colOp As Long, colState As Long, colCommitted As Long, colErr As Long
    colOp = PRIME_ColIndex(headers, "OP_ID")
    colState = PRIME_ColIndex(headers, "STATE")
    colCommitted = PRIME_ColIndex(headers, "COMMITTED_AT")
    colErr = PRIME_ColIndex(headers, "ERROR")

    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_TX)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colOp, opId)
    If idx = -1 Then Exit Sub

    Dim row(UBound(table(idx))) As Variant
    Dim c As Long
    For c = 0 To UBound(table(idx))
        row(c) = table(idx)(c)
    Next c
    row(colState) = newState
    If newState = TX_COMMITTED Then row(colCommitted) = Format(Now, "YYYY-MM-DD HH:MM:SS")
    If errText <> "" Then row(colErr) = errText

    Dim rows(0) As Variant
    rows(0) = row
    PRIME_UpdateRowsBatch(SH_SYS_TX, idx, rows)
End Sub

Private Sub PRIME_TryMarkTxFailed(ByVal opId As String, ByVal errText As String)
    On Error Resume Next
    PRIME_UpdateTxState(opId, TX_FAILED, errText)
End Sub

' Логирование НЕ должно уронить проведение при недоступности листа/файла (logging.must_not_break_operation_if_unavailable).
' Это единственное осознанное исключение из общего правила про On Error Resume Next -
' сама транзакционная запись (SYS_PRIME_TX) идёт отдельной функцией без подавления ошибок.
Public Sub PRIME_AuditLog(ByVal opId As String, ByVal stage As String, ByVal sheet As String, ByVal note As String)
    On Error Resume Next
    If Not PRIME_SheetExists(SH_DB_AUDIT) Then Exit Sub
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_AUDIT)
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "TS")) = Format(Now, "YYYY-MM-DD HH:MM:SS")
    row(PRIME_ColIndex(headers, "OP_ID")) = opId
    row(PRIME_ColIndex(headers, "STAGE")) = stage
    row(PRIME_ColIndex(headers, "SHEET")) = sheet
    row(PRIME_ColIndex(headers, "ERROR_TEXT")) = note
    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_DB_AUDIT, rows)
End Sub
