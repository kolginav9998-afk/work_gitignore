Option Explicit

' PRIME_04_Posting
' Единственный posting engine (single_posting_engine=true). Все формы (Заказы, Выдачи,
' цеховые/офисные листы, Возвраты, Инвентаризация) строят PrimeDocPlan и вызывают
' PRIME_PostDocument - сами НИКОГДА не пишут в DB_PRIME_* напрямую
' (ui_forms_must_not_write_movements_directly).

Type PrimeDocLine
    ProductCode As String        ' если пусто - будет создан новый товар (только когда это осознанно разрешено вызывающим)
    ProductName As String        ' для создания карточки товара, если ProductCode пуст
    QtyInput As Double
    UnitInput As String
    LocationFrom As String
    LocationTo As String
    DestinationProject As String ' "Назначение / проект" - обязательное поле, теряемое в 1.4.1
    Recipient As String
    Comment As String
    Price As Double
    OriginalDocLineId As String  ' для RETURN: ссылка на исходную строку выдачи
    QtyBase As Double            ' заполняется на этапе валидации (после конвертации единиц)
    LotId As String              ' заполняется на этапе валидации (для однопартийных операций) либо пусто (FIFO по нескольким партиям)
End Type

Type PrimeDocPlan
    DocType As String
    DocDate As String
    SourceSheet As String
    SourceKey As String
    OrderId As String
    Lines(99) As PrimeDocLine
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

Public Sub PRIME_PlanAddLine(ByRef plan As PrimeDocPlan, ByRef line As PrimeDocLine)
    If plan.LineCount > UBound(plan.Lines) Then
        Err.Raise(1010, "PRIME_Posting.PRIME_PlanAddLine", "Документ превышает лимит строк одной операции (100). Разбейте на несколько проведений.")
    End If
    plan.Lines(plan.LineCount) = line
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

    PRIME_TryEnter()
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
    Dim errMsg As String
    If Not PRIME_ValidateAndExpandPlan(plan, errMsg) Then
        gLastPostError = errMsg
        PRIME_AuditLog("", STAGE_ERROR, plan.SourceSheet, errMsg)
        PRIME_Leave()
        PRIME_PostDocument = ""
        Exit Function
    End If
    PRIME_AuditLog("", STAGE_VALIDATION_OK, plan.SourceSheet, plan.SourceKey)

    ' 3. OP_ID / DOC_ID, запись PREPARED.
    opId = PRIME_NewOpId()
    docId = "DOC-" & Format(PRIME_SequenceNext("DOC_ID"), "00000000")
    PRIME_WriteTxRow(opId, plan.SourceKey, docId, TX_PREPARED, "")
    txWritten = True
    PRIME_AuditLog(opId, STAGE_TX_PREPARED, plan.SourceSheet, docId)

    ' 4. Пакетная запись документа/строк/партий/движений - по типу документа.
    PRIME_WriteDocumentHeader(docId, plan)
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
            Err.Raise(1011, "PRIME_Posting.PRIME_PostDocument", "Неизвестный тип документа: " & plan.DocType)
    End Select
    PRIME_AuditLog(opId, STAGE_MOVEMENTS_WRITTEN, plan.SourceSheet, docId)

    ' 5. COMMITTED - последний логический шаг перед UI/store.
    PRIME_UpdateTxState(opId, TX_COMMITTED, "")
    PRIME_RegisterCommittedKey(plan.SourceKey, docId)
    PRIME_AuditLog(opId, STAGE_TX_COMMITTED, plan.SourceSheet, docId)

    ' 6. Один store() на весь документ.
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
        ' Идентичность товара: если код не указан, создаём новый ТОЛЬКО когда явно есть название
        ' (используется формами, где допустим ввод нового товара с нуля).
        If plan.Lines(i).ProductCode = "" Then
            If plan.Lines(i).ProductName = "" Then
                errMsg = "Строка " & (i + 1) & ": не указан ни код товара, ни наименование."
                PRIME_ValidateAndExpandPlan = False
                Exit Function
            End If
        ElseIf Not PRIME_ProductExists(plan.Lines(i).ProductCode) And plan.DocType <> DOC_RECEIPT Then
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
    Next i
    PRIME_ValidateAdjustment = True
End Function

Private Function PRIME_ValidateReceipt(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        If plan.Lines(i).ProductCode = "" Then
            ' Новый товар - forbidden запрещает создавать код, когда код УЖЕ указан; здесь код
            ' сознательно пуст, значит создание нового кода корректно.
            plan.Lines(i).ProductCode = PRIME_CreateProduct(plan.Lines(i).ProductName, plan.Lines(i).UnitInput, _
                plan.Lines(i).LocationTo, "", "", False, "")
        End If

        Dim baseQty As Variant
        baseQty = PRIME_ConvertQtyToBase(plan.Lines(i).ProductCode, plan.Lines(i).UnitInput, plan.Lines(i).QtyInput)
        If IsEmpty(baseQty) Then
            errMsg = "Строка " & (i + 1) & ": нет коэффициента пересчёта для единицы """ & plan.Lines(i).UnitInput & """."
            PRIME_ValidateReceipt = False
            Exit Function
        End If
        plan.Lines(i).QtyBase = CDbl(baseQty)
        plan.Lines(i).LotId = "LOT-" ' финальный номер присваивается на этапе записи (PRIME_PostReceiptLines)
    Next i
    PRIME_ValidateReceipt = True
End Function

Private Function PRIME_ValidateIssue(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
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

        ' Проверка достаточности остатка по FIFO (без физической записи) - shortage_behavior=Reject entire document.
        Dim available As Double
        available = PRIME_TotalLotBalance(plan.Lines(i).ProductCode)
        If available < plan.Lines(i).QtyBase Then
            errMsg = "Строка " & (i + 1) & ": недостаточно остатка (доступно " & available & ", требуется " & plan.Lines(i).QtyBase & "). Документ не проведён целиком."
            PRIME_ValidateIssue = False
            Exit Function
        End If
    Next i
    PRIME_ValidateIssue = True
End Function

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

Private Function PRIME_ValidateTransfer(ByRef plan As PrimeDocPlan, ByRef errMsg As String) As Boolean
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

        Dim locations() As String
        Dim quantities() As Double
        PRIME_StockByLocation(plan.Lines(i).ProductCode, locations, quantities)
        Dim avail As Double
        avail = 0
        Dim j As Long
        If UBound(locations) >= LBound(locations) Then
            For j = LBound(locations) To UBound(locations)
                If locations(j) = plan.Lines(i).LocationFrom Then avail = quantities(j)
            Next j
        End If
        If avail < plan.Lines(i).QtyBase Then
            errMsg = "Строка " & (i + 1) & ": на месте """ & plan.Lines(i).LocationFrom & """ недостаточно остатка для перемещения."
            PRIME_ValidateTransfer = False
            Exit Function
        End If
    Next i
    PRIME_ValidateTransfer = True
End Function

' === Запись шапки/строк документа ============================================================
Private Sub PRIME_WriteDocumentHeader(ByVal docId As String, ByRef plan As PrimeDocPlan)
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
Private gLineIdBuffer(99) As String

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
    hasSnapshot = PRIME_SheetExists(SH_DB_ORDER_SNAPSHOT) And plan.OrderId <> ""
    Dim snapRows() As Variant
    If hasSnapshot Then
        snapHeaders = PRIME_HeaderMap(SH_DB_ORDER_SNAPSHOT)
        ReDim snapRows(plan.LineCount - 1)
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
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
        lotRows(i) = lotRow

        moveRows(i) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
            lotId, plan.Lines(i).QtyBase, plan.Lines(i).LocationTo, plan.DocDate, opId)

        If hasSnapshot Then
            snapRows(i) = PRIME_BuildOrderSnapshotRow(snapHeaders, plan, i, docId, lotId)
        End If
    Next i

    PRIME_AppendRowsBatch(SH_DB_LOTS, lotRows)
    PRIME_AuditLog(opId, STAGE_LOTS_WRITTEN, plan.SourceSheet, docId)
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRows)
    If hasSnapshot Then PRIME_AppendRowsBatch(SH_DB_ORDER_SNAPSHOT, snapRows)
End Sub

Private Function PRIME_BuildOrderSnapshotRow(ByVal headers As Variant, ByRef plan As PrimeDocPlan, ByVal idx As Long, ByVal docId As String, ByVal lotId As String) As Variant
    Dim row(UBound(headers)) As Variant
    Dim c As Long
    c = PRIME_ColIndex(headers, "ORDER_ID") : If c >= 0 Then row(c) = plan.OrderId
    c = PRIME_ColIndex(headers, "RECEIPT_DOC_ID") : If c >= 0 Then row(c) = docId
    c = PRIME_ColIndex(headers, "LOT_ID") : If c >= 0 Then row(c) = lotId
    c = PRIME_ColIndex(headers, "DELIVERY_QTY") : If c >= 0 Then row(c) = plan.Lines(idx).QtyBase
    c = PRIME_ColIndex(headers, "PRODUCT_CODE") : If c >= 0 Then row(c) = plan.Lines(idx).ProductCode
    c = PRIME_ColIndex(headers, "DOC_DATE") : If c >= 0 Then row(c) = plan.DocDate
    PRIME_BuildOrderSnapshotRow = row
End Function

' === ISSUE: FIFO-разбиение по партиям, allocations, движения с отрицательным количеством =====
Private Sub PRIME_PostIssueLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim allocHeaders As Variant
    allocHeaders = PRIME_HeaderMap(SH_DB_ALLOCATIONS)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)

    Dim allocRowsBuf() As Variant
    Dim moveRowsBuf() As Variant
    Dim allocCount As Long, moveCount As Long
    allocCount = 0 : moveCount = 0
    ReDim allocRowsBuf(200)
    ReDim moveRowsBuf(200)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim lots() As String
        Dim balances() As Double
        PRIME_FifoLotsForProduct(plan.Lines(i).ProductCode, lots, balances)

        Dim remaining As Double
        remaining = plan.Lines(i).QtyBase
        Dim j As Long
        For j = LBound(lots) To UBound(lots)
            If remaining <= 0.0000005 Then Exit For
            If balances(j) > 0 Then
                Dim take As Double
                take = remaining
                If balances(j) < take Then take = balances(j)

                allocRowsBuf(allocCount) = PRIME_BuildAllocationRow(allocHeaders, PRIME_LineIdFor(i), lots(j), take)
                allocCount = allocCount + 1

                moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                    lots(j), -take, plan.Lines(i).LocationFrom, plan.DocDate, opId)
                moveCount = moveCount + 1

                remaining = remaining - take
            End If
        Next j

        If remaining > 0.0000005 Then
            ' Не должно происходить после PRIME_ValidateIssue, но перестраховка важнее красоты кода:
            ' не допускаем частично проведённый документ.
            Err.Raise(1020, "PRIME_Posting.PRIME_PostIssueLines", "Недостаточно партий для списания строки " & (i + 1) & " после валидации - проведение отменено.")
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

' === RETURN: возврат симметричен FIFO-списанию исходной выдачи (FIFO по allocations) =========
Private Sub PRIME_PostReturnLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim retHeaders As Variant
    retHeaders = PRIME_HeaderMap(SH_DB_RETURNS)

    Dim moveRowsBuf() As Variant
    Dim retRowsBuf(plan.LineCount - 1) As Variant
    Dim moveCount As Long
    moveCount = 0
    ReDim moveRowsBuf(200)

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        Dim lots() As String
        Dim allocatedQty() As Double
        PRIME_AllocationsForDocLine(plan.Lines(i).OriginalDocLineId, lots, allocatedQty)

        Dim remaining As Double
        remaining = plan.Lines(i).QtyBase
        Dim j As Long
        If UBound(lots) >= LBound(lots) Then
            For j = LBound(lots) To UBound(lots)
                If remaining <= 0.0000005 Then Exit For
                Dim take As Double
                take = remaining
                If allocatedQty(j) < take Then take = allocatedQty(j)
                If take > 0 Then
                    moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                        lots(j), take, plan.Lines(i).LocationTo, plan.DocDate, opId)
                    moveCount = moveCount + 1
                    remaining = remaining - take
                End If
            Next j
        End If
        If remaining > 0.0000005 Then
            ' Исходная выдача не найдена по allocations (например, легаси-данные без миграции
            ' allocations) - возврат всё равно проводим на условное "безлотовое" движение,
            ' чтобы не заблокировать документ, но это ухудшает трассируемость партии.
            moveRowsBuf(moveCount) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
                "", remaining, plan.Lines(i).LocationTo, plan.DocDate, opId)
            moveCount = moveCount + 1
        End If

        Dim retRow(UBound(retHeaders)) As Variant
        retRow(PRIME_ColIndex(retHeaders, "RETURN_ID")) = docId & "-R" & (i + 1)
        retRow(PRIME_ColIndex(retHeaders, "ORIGINAL_ISSUE_DOC_LINE_ID")) = plan.Lines(i).OriginalDocLineId
        retRow(PRIME_ColIndex(retHeaders, "RETURN_DOC_ID")) = docId
        retRow(PRIME_ColIndex(retHeaders, "QTY_BASE")) = plan.Lines(i).QtyBase
        retRow(PRIME_ColIndex(retHeaders, "RETURN_DATE")) = plan.DocDate
        retRowsBuf(i) = retRow
    Next i

    If moveCount > 0 Then
        ReDim Preserve moveRowsBuf(moveCount - 1)
        PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRowsBuf)
    End If
    PRIME_AppendRowsBatch(SH_DB_RETURNS, retRowsBuf)
End Sub

' === TRANSFER: OUT + IN, суммарный остаток товара не меняется ================================
Private Sub PRIME_PostTransferLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim rows(plan.LineCount * 2 - 1) As Variant
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        rows(i * 2) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
            "", -plan.Lines(i).QtyBase, plan.Lines(i).LocationFrom, plan.DocDate, opId)
        rows(i * 2 + 1) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
            "", plan.Lines(i).QtyBase, plan.Lines(i).LocationTo, plan.DocDate, opId)
    Next i
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, rows)
End Sub

' === ADJUSTMENT: разница инвентаризации, знак берётся из QtyBase (может быть отрицательным) ===
Private Sub PRIME_PostAdjustmentLines(ByVal docId As String, ByVal opId As String, ByRef plan As PrimeDocPlan)
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim rows(plan.LineCount - 1) As Variant
    Dim i As Long
    For i = 0 To plan.LineCount - 1
        rows(i) = PRIME_BuildMovementRow(moveHeaders, docId, PRIME_LineIdFor(i), plan.Lines(i).ProductCode, _
            "", plan.Lines(i).QtyBase, plan.Lines(i).LocationTo, plan.DocDate, opId)
    Next i
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, rows)
End Sub

Private Function PRIME_BuildMovementRow(ByVal headers As Variant, ByVal docId As String, ByVal docLineId As String, _
        ByVal productCode As String, ByVal lotId As String, ByVal qtyBase As Double, ByVal location As String, _
        ByVal moveDate As String, ByVal opId As String) As Variant
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
    PRIME_BuildMovementRow = row
End Function

' === FIFO / остатки по партиям =================================================================
' Партии товара, отсортированные по дате прихода и LOT_ID (fifo.sort_order), с текущим балансом
' (сумма COMMITTED-движений по каждой партии). Только партии с положительным балансом полезны
' для списания, но возвращаем все для прозрачности вызывающему.
Public Sub PRIME_FifoLotsForProduct(ByVal productCode As String, ByRef lots() As String, ByRef balances() As Double)
    If Not PRIME_SheetExists(SH_DB_LOTS) Then
        ReDim lots(-1) : ReDim balances(-1)
        Exit Sub
    End If
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim colCode As Long, colLotId As Long, colDate As Long
    colCode = PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")
    colLotId = PRIME_ColIndex(lotHeaders, "LOT_ID")
    colDate = PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")

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
            candLots(n) = CStr(lotTable(i)(colLotId))
            candDates(n) = CStr(lotTable(i)(colDate))
            n = n + 1
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

Public Function PRIME_LotBalance(ByVal lotId As String) As Double
    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        PRIME_LotBalance = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colLot As Long, colQty As Long
    colLot = PRIME_ColIndex(headers, "LOT_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    Dim total As Double
    total = 0
    Dim i As Long
    If UBound(table) >= 1 Then
        For i = 1 To UBound(table)
            If CStr(table(i)(colLot)) = lotId Then
                total = total + CDbl(table(i)(colQty))
            End If
        Next i
    End If
    PRIME_LotBalance = total
End Function

Public Function PRIME_TotalLotBalance(ByVal productCode As String) As Double
    Dim lots() As String
    Dim balances() As Double
    PRIME_FifoLotsForProduct(productCode, lots, balances)
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

Public Function PRIME_DocLineQtyBase(ByVal docLineId As String) As Double
    If docLineId = "" Then
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

Public Function PRIME_AlreadyReturnedQtyBase(ByVal originalDocLineId As String) As Double
    If originalDocLineId = "" Or Not PRIME_SheetExists(SH_DB_RETURNS) Then
        PRIME_AlreadyReturnedQtyBase = 0
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_RETURNS)
    Dim colOrig As Long, colQty As Long
    colOrig = PRIME_ColIndex(headers, "ORIGINAL_ISSUE_DOC_LINE_ID")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_RETURNS)
    Dim total As Double
    total = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If CStr(table(i)(colOrig)) = originalDocLineId Then
                total = total + CDbl(table(i)(colQty))
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
