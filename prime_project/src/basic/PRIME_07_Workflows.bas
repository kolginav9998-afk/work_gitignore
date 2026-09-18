Option Explicit

' Дублируем определение типов из PRIME_04_Posting - StarBasic не поддерживает совместное
' использование Type между модулями (Type виден только в модуле, где объявлен), поэтому
' идентичное объявление нужно в каждом модуле, который строит PrimeDocPlan/PrimeDocLine.
Type PrimeDocLine
    ProductCode As String
    ProductName As String
    IsNewProduct As Boolean
    QtyInput As Double
    UnitInput As String
    LocationFrom As String
    LocationTo As String
    Contour As String
    ContourFrom As String
    ContourTo As String
    DestinationProject As String
    Recipient As String
    Comment As String
    Price As Double
    OriginalDocLineId As String
    OrderLineId As String
    QtyBase As Double
    LotId As String
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


' PRIME_07_Workflows
' Единый механизм для активных receipt/issue-листов (Приход/Расход — Офис/Производство/Детали,
' плюс легаси "Приход/Расход — Цех" - скрыт из UI, но engine не тронут ради сохранности старых
' данных) - один и тот же posting engine, одна и та же лёгкая логика событий.
' single_physical_warehouse (FINAL): Офис/Производство/Детали/Заказы - МЕТКА источника/
' назначения движения, а не отдельный физический остаток - один склад, один общий остаток по
' (EI_CODE, место хранения), см. PRIME_04_Posting.PRIME_LocationContourBalance.

Private Function PRIME_Workflow_IsReceiptSheet(ByVal sheetName As String) As Boolean
    PRIME_Workflow_IsReceiptSheet = (sheetName = SH_RECEIPT_SHOP Or sheetName = SH_RECEIPT_OFFICE _
        Or sheetName = SH_RECEIPT_PRODUCTION Or sheetName = SH_RECEIPT_DETAILS)
End Function

Private Function PRIME_Workflow_IsIssueSheet(ByVal sheetName As String) As Boolean
    PRIME_Workflow_IsIssueSheet = (sheetName = SH_ISSUE_SHOP Or sheetName = SH_ISSUE_OFFICE _
        Or sheetName = SH_ISSUE_PRODUCTION Or sheetName = SH_ISSUE_DETAILS)
End Function

' Единственный обработчик события для всех 4 листов - LO передаёт индекс листа в oRangeAddr.Sheet,
' поэтому одна и та же лёгкая функция может быть привязана ко всем сразу (без дублирования кода
' по образцу трёх разных UI-движков 1.4.1).
Public Sub PRIME_OnContentChanged_Workflow(ByVal oRangeAddr As Variant)
    If Not PRIME_EventEnter() Then Exit Sub
    On Error Goto CleanExit

    If IsNull(oRangeAddr) Then GoTo CleanExit

    Dim oSheet As Object
    oSheet = ThisComponent.Sheets.getByIndex(oRangeAddr.Sheet)
    Dim sheetName As String
    sheetName = oSheet.Name

    If Not (PRIME_Workflow_IsReceiptSheet(sheetName) Or PRIME_Workflow_IsIssueSheet(sheetName)) Then GoTo CleanExit

    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    If colCode < 0 Then GoTo CleanExit

    Dim colQty As Long
    colQty = PRIME_ColIndex(headers, PRIME_Workflow_QtyColName(sheetName))

    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(sheetName)
    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= firstDataRow Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Workflow_AutofillByCode(oSheet, headers, sheetName, r)
                    PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, r)
                ElseIf c = colQty Then
                    PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, r)
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
End Sub

' Имя колонки "Кол-во"/"Место хранения" зависит от receipt/issue (см. новые unified-схемы в
' PRIME_00_Config.PRIME_WorkflowReceiptColumns/PRIME_WorkflowIssueColumns) - единая точка,
' чтобы не рассинхронизировать несколько мест, где нужно это имя.
Public Function PRIME_Workflow_QtyColName(ByVal sheetName As String) As String
    If PRIME_Workflow_IsReceiptSheet(sheetName) Then
        PRIME_Workflow_QtyColName = "Количество прихода"
    Else
        PRIME_Workflow_QtyColName = "Количество"
    End If
End Function

Public Function PRIME_Workflow_LocColName(ByVal sheetName As String) As String
    If PRIME_Workflow_IsReceiptSheet(sheetName) Then
        PRIME_Workflow_LocColName = "Место хранения на складе"
    Else
        PRIME_Workflow_LocColName = "Место хранения"
    End If
End Function

Private Sub PRIME_Workflow_AutofillByCode(ByVal oSheet As Object, ByVal headers As Variant, ByVal sheetName As String, ByVal row As Long)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    PRIME_SetCellIfEmpty(oSheet, headers, row, "Наименование", PRIME_GetProductField(code, "PRODUCT_NAME"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Ед. изм.", PRIME_GetProductField(code, "BASE_UNIT"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Категория", PRIME_GetProductField(code, "CATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Подкатегория", PRIME_GetProductField(code, "SUBCATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, PRIME_Workflow_LocColName(sheetName), PRIME_GetProductField(code, "DEFAULT_LOCATION"))
End Sub

' single_physical_warehouse (FINAL): раньше здесь были парные "Остаток .../Будет ..."-колонки,
' партиционированные по контуру листа (Остаток офиса/Будет в офисе и т.п. - forbidden_ui_terms
' в новом задании) - "Будет" вычислялась арифметикой (before +/- введённое количество) и после
' проведения, когда код позиции дописывается обратно в ячейку, событие срабатывало ПОВТОРНО и
' прибавляло/вычитало количество ЕЩЁ РАЗ поверх уже проведённого остатка (double-count bug).
' Заменено на единственную колонку "Остаток позиции на складе" - просто ТЕКУЩИЙ остаток именно
' этого EI_CODE (без учёта контура/источника - один физический склад), без какой-либо
' арифметики here - остаток всегда пересчитывается заново из COMMITTED-ledger, поэтому
' повторный вызов события идемпотентен.
Private Sub PRIME_Workflow_RefreshInlineStock(ByVal oSheet As Object, ByVal sheetName As String, ByVal headers As Variant, ByVal row As Long)
    Dim colBalance As Long, colCode As Long, colLoc As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    colLoc = PRIME_ColIndex(headers, PRIME_Workflow_LocColName(sheetName))
    colBalance = PRIME_ColIndex(headers, "Остаток позиции на складе")
    If colBalance < 0 Then Exit Sub

    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    Dim loc As String
    loc = ""
    If colLoc >= 0 Then loc = Trim(oSheet.getCellByPosition(colLoc, row).getString())
    oSheet.getCellByPosition(colBalance, row).setValue(PRIME_LocationContourBalance(code, loc, ""))
End Sub

' === Кнопки (одни и те же имена процедур используются на всех 4 листах; активный лист
' определяется через ThisComponent.CurrentController.ActiveSheet, а не жёстко зашит) ===========

Public Sub PRIME_Workflow_ConductRowButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < PRIME_FormSchemaFirstDataRow(oSheet.Name) Then Exit Sub
    PRIME_Workflow_ConductRow(oSheet, row)
End Sub

' R09 (2.1.0): все строки с заполненным кодом проводятся ОДНИМ документом (один DOC_ID/OP_ID/
' один store()) вместо цикла отдельных PostDocument-вызовов на каждую строку.
' batch_invalid_line fix (2.1.1): см. идентичный комментарий у PRIME_Issues_ConductAllButton -
' заполненная, но невалидная строка теперь останавливает построение всего батча, а не молча
' исключается из него.
Public Sub PRIME_Workflow_ConductAllButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim sheetName As String
    sheetName = oSheet.Name
    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim colName As Long
    colName = PRIME_ColIndex(headers, "Наименование")
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_WFState")
    Dim isReceiptSheet As Boolean
    isReceiptSheet = PRIME_Workflow_IsReceiptSheet(sheetName)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)

    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim batchKey As String
    batchKey = ""

    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(sheetName) To lastRow
        ' receipt_forms_still_require_existing_ei fix (2.1.2): на приходе строка считается
        ' заполненной, если есть код ИЛИ наименование (новый приход обычно вводится без кода -
        ' он появится только после проведения). На расходе, как и раньше, только код.
        Dim hasInput As Boolean
        hasInput = Trim(oSheet.getCellByPosition(colCode, r).getString()) <> ""
        If isReceiptSheet And colName >= 0 Then
            hasInput = hasInput Or Trim(oSheet.getCellByPosition(colName, r).getString()) <> ""
        End If
        If hasInput Then
            If colState < 0 Or Left(oSheet.getCellByPosition(colState, r).getString(), 5) <> "DONE:" Then
                Dim docLine As PrimeDocLine
                Dim rowKey As String
                If Not PRIME_Workflow_BuildLine(oSheet, sheetName, headers, r, docLine, rowKey) Then
                    MsgBox "Проведение не выполнено: строка " & (r + 1) & " заполнена, но невалидна. Исправьте её или очистите перед проведением всего батча."
                    Exit Sub
                End If
                If plan.LineCount = 0 Then
                    PRIME_InitPlan(plan, IIf(isReceiptSheet, DOC_RECEIPT, DOC_ISSUE), sheetName, "")
                End If
                batchKey = batchKey & rowKey & ","
                rowForLine(plan.LineCount) = r
                PRIME_PlanAddLine(plan, docLine)
            End If
        End If
    Next r

    If plan.LineCount = 0 Then Exit Sub
    plan.SourceKey = "BATCH:" & batchKey

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        ' EI_CODE новой позиции (решён внутри PRIME_PostDocument/PRIME_ValidateReceipt) обязан
        ' быть виден пользователю в строке прихода сразу после успешного проведения.
        If isReceiptSheet Then
            oSheet.getCellByPosition(colCode, rowForLine(i)).setString(plan.Lines(i).ProductCode)
        End If
        PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, rowForLine(i))
        PRIME_Workflow_MarkPostedRow(oSheet, headers, rowForLine(i), docId)
        If colState >= 0 Then
            oSheet.getCellByPosition(colState, rowForLine(i)).setString("DONE:" & _
                oSheet.getCellByPosition(colState, rowForLine(i)).getString())
        End If
    Next i
    MsgBox "Проведено строк: " & plan.LineCount & " (документ " & docId & ")"
End Sub

' receipt_registers (FINAL): "После проведения строка НЕ исчезает: в ней видны EI, количество
' прихода и текущий остаток этой позиции" + видимые DOC_ID/Статус - без MsgBox с номером
' документа как единственного способа его узнать. No-op на листах без этих колонок (Расход).
Private Sub PRIME_Workflow_MarkPostedRow(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, ByVal docId As String)
    Dim colDocId As Long, colStatus As Long
    colDocId = PRIME_ColIndex(headers, "DOC_ID")
    colStatus = PRIME_ColIndex(headers, "Статус")
    If colDocId >= 0 Then oSheet.getCellByPosition(colDocId, row).setString(docId)
    If colStatus >= 0 Then oSheet.getCellByPosition(colStatus, row).setString("Проведено")
End Sub

Public Sub PRIME_Workflow_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim sheetName As String
    sheetName = oSheet.Name
    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)

    Dim docLine As PrimeDocLine
    Dim draftId As String
    If Not PRIME_Workflow_BuildLine(oSheet, sheetName, headers, row, docLine, draftId) Then Exit Sub

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, IIf(PRIME_Workflow_IsReceiptSheet(sheetName), DOC_RECEIPT, DOC_ISSUE), sheetName, draftId)
    PRIME_PlanAddLine(plan, docLine)

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    If PRIME_Workflow_IsReceiptSheet(sheetName) Then
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), row).setString(plan.Lines(0).ProductCode)
    End If
    PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, row)
    PRIME_Workflow_MarkPostedRow(oSheet, headers, row, docId)
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_WFState")
    If colState >= 0 Then
        oSheet.getCellByPosition(colState, row).setString("DONE:" & oSheet.getCellByPosition(colState, row).getString())
    End If
    MsgBox "Проведено: " & docId
End Sub

' Читает строку Приход/Расход-Цех/Офис и строит PrimeDocLine - общий для одиночного и
' батч-проведения. Контур определяется листом (PRIME_ContourForSheet), не вводится пользователем.
' batch_duplicate_posting fix (2.1.1): см. идентичный комментарий у PRIME_06_Issues.
' PRIME_Issues_BuildLine - тот же паттерн "успешно проведённая строка никогда не помечалась и
' навсегда оставалась включаемой в следующий батч" был и здесь, на всех 4 workflow-листах.
Private Function PRIME_Workflow_BuildLine(ByVal oSheet As Object, ByVal sheetName As String, ByVal headers As Variant, ByVal row As Long, _
        ByRef docLine As PrimeDocLine, ByRef draftId As String) As Boolean
    Dim isReceipt As Boolean
    isReceipt = PRIME_Workflow_IsReceiptSheet(sheetName)

    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_WFState")
    If colState >= 0 Then draftId = oSheet.getCellByPosition(colState, row).getString()
    If Left(draftId, 5) = "DONE:" Then
        PRIME_Workflow_BuildLine = False
        Exit Function
    End If
    If draftId = "" Then
        draftId = "WF-" & Format(PRIME_SequenceNext("WF_DRAFT_ID"), "00000000")
        If colState >= 0 Then oSheet.getCellByPosition(colState, row).setString(draftId)
    End If

    Dim code As String
    code = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), row).getString())
    Dim itemName As String
    itemName = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Наименование"), row).getString())

    ' receipt_forms_still_require_existing_ei fix (2.1.2): приход НЕ обязан ссылаться на
    ' существующий код - по модели 2.1.1/2.1.2 каждый приход создаёт новую EI-позицию
    ' (см. PRIME_04_Posting.PRIME_ValidateReceipt), поэтому единственное реальное требование для
    ' строки прихода - наименование (код, если указан, всё равно игнорируется при создании новой
    ' позиции - оставлен только как возможное автозаполнение). Для расхода код (ЕИ-код конкретной
    ' позиции) остаётся строго обязательным - списывать без него нечего.
    If isReceipt Then
        If itemName = "" Then
            MsgBox "Строка " & (row + 1) & ": не указано наименование товара."
            PRIME_Workflow_BuildLine = False
            Exit Function
        End If
    ElseIf code = "" Then
        MsgBox "Строка " & (row + 1) & ": не указан внутренний код товара (ЕИ-код)."
        PRIME_Workflow_BuildLine = False
        Exit Function
    End If

    Dim qtyColName As String
    qtyColName = PRIME_Workflow_QtyColName(sheetName)
    Dim qtyStr As String
    qtyStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, qtyColName), row).getString())
    If qtyStr = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Строка " & (row + 1) & ": заполните корректное """ & qtyColName & """."
        PRIME_Workflow_BuildLine = False
        Exit Function
    End If

    docLine.ProductCode = code
    docLine.ProductName = itemName
    docLine.QtyInput = CDbl(qtyStr)
    docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    docLine.Contour = PRIME_ContourForSheet(sheetName)
    docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Документ"), row).getString() & _
        " " & oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), row).getString()

    Dim locColName As String
    locColName = PRIME_Workflow_LocColName(sheetName)
    If isReceipt Then
        docLine.LocationTo = oSheet.getCellByPosition(PRIME_ColIndex(headers, locColName), row).getString()
        docLine.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кто сдал"), row).getString()
    Else
        docLine.LocationFrom = oSheet.getCellByPosition(PRIME_ColIndex(headers, locColName), row).getString()
        docLine.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кто принял"), row).getString()
        Dim colDest As Long
        colDest = PRIME_ColIndex(headers, "Куда / назначение")
        If colDest >= 0 Then docLine.DestinationProject = oSheet.getCellByPosition(colDest, row).getString()
    End If
    PRIME_Workflow_BuildLine = True
End Function

Public Sub PRIME_Workflow_FillArticleButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim headers As Variant
    headers = PRIME_HeaderMap(oSheet.Name)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(oSheet.Name)
    Dim processed As Long
    processed = 0
    Dim r As Long
    For r = firstDataRow To lastRow
        PRIME_Workflow_AutofillByCode(oSheet, headers, oSheet.Name, r)
        processed = processed + 1
    Next r
    MsgBox "Автозаполнение выполнено для " & processed & " строк."
End Sub

' "Повторить значения" - копирует общие поля (Место хранения/Кто сдал/Кто принял/Документ) из
' предыдущей строки в новую, ускоряя ручной ввод серии однотипных операций.
Public Sub PRIME_Workflow_RepeatFieldsButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim headers As Variant
    headers = PRIME_HeaderMap(oSheet.Name)
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < PRIME_FormSchemaFirstDataRow(oSheet.Name) + 1 Then Exit Sub ' нужна предыдущая строка данных

    Dim repeatCols As Variant
    repeatCols = Array(PRIME_Workflow_LocColName(oSheet.Name), "Кто сдал", "Кто принял", "Документ")
    Dim i As Long
    For i = LBound(repeatCols) To UBound(repeatCols)
        Dim col As Long
        col = PRIME_ColIndex(headers, repeatCols(i))
        If col >= 0 Then
            If oSheet.getCellByPosition(col, row).getString() = "" Then
                oSheet.getCellByPosition(col, row).setString(oSheet.getCellByPosition(col, row - 1).getString())
            End If
        End If
    Next i
End Sub
