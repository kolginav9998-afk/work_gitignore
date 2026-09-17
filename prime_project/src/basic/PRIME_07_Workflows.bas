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
' Единый механизм для 4 активных цеховых/офисных листов (Приход/Расход - Цех/Офис) -
' один и тот же posting engine, одна и та же лёгкая логика событий. Легаси-формы
' "Производство"/"Детали" сюда не входят: они архив истории, не активные формы ввода
' (ARCHITECTURE §5, legacy_parallel_forms).

Private Function PRIME_Workflow_IsReceiptSheet(ByVal sheetName As String) As Boolean
    PRIME_Workflow_IsReceiptSheet = (sheetName = SH_RECEIPT_SHOP Or sheetName = SH_RECEIPT_OFFICE)
End Function

Private Function PRIME_Workflow_IsIssueSheet(ByVal sheetName As String) As Boolean
    PRIME_Workflow_IsIssueSheet = (sheetName = SH_ISSUE_SHOP Or sheetName = SH_ISSUE_OFFICE)
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
    colQty = PRIME_ColIndex(headers, "Кол-во")

    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(sheetName)
    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= firstDataRow Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Workflow_AutofillByCode(oSheet, headers, r)
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

Private Sub PRIME_Workflow_AutofillByCode(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    PRIME_SetCellIfEmpty(oSheet, headers, row, "Наименование", PRIME_GetProductField(code, "PRODUCT_NAME"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Ед. изм.", PRIME_GetProductField(code, "BASE_UNIT"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Категория", PRIME_GetProductField(code, "CATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Подкатегория", PRIME_GetProductField(code, "SUBCATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Место хранения", PRIME_GetProductField(code, "DEFAULT_LOCATION"))
End Sub

' R19/R20/R21 (2.1.0): "Остаток .../Приход или Расход/Будет ..." - контур берётся из имени листа
' (PRIME_ContourForSheet: Цех -> WORKSHOP_DETAILS, Офис -> OFFICE), место - из ячейки строки.
' Приход увеличивает "Будет", расход уменьшает - то же представление COMMITTED-ledger, что и
' лист "Наличие", просто с готовым фильтром по контуру этого листа.
Private Sub PRIME_Workflow_RefreshInlineStock(ByVal oSheet As Object, ByVal sheetName As String, ByVal headers As Variant, ByVal row As Long)
    Dim colBefore As Long, colAfter As Long, colCode As Long, colLoc As Long, colQty As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    colLoc = PRIME_ColIndex(headers, "Место хранения")
    colQty = PRIME_ColIndex(headers, "Кол-во")
    Dim isReceipt As Boolean
    isReceipt = PRIME_Workflow_IsReceiptSheet(sheetName)
    If sheetName = SH_RECEIPT_OFFICE Then
        colBefore = PRIME_ColIndex(headers, "Остаток офиса") : colAfter = PRIME_ColIndex(headers, "Будет в офисе")
    ElseIf sheetName = SH_ISSUE_OFFICE Then
        colBefore = PRIME_ColIndex(headers, "Остаток офиса") : colAfter = PRIME_ColIndex(headers, "Будет в офисе")
    Else
        colBefore = PRIME_ColIndex(headers, "Остаток деталей") : colAfter = PRIME_ColIndex(headers, "Будет деталей")
    End If
    If colBefore < 0 And colAfter < 0 Then Exit Sub

    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    Dim loc As String
    loc = ""
    If colLoc >= 0 Then loc = Trim(oSheet.getCellByPosition(colLoc, row).getString())
    Dim contour As String
    contour = PRIME_ContourForSheet(sheetName)
    Dim before As Double
    before = PRIME_LocationContourBalance(code, loc, contour)
    If colBefore >= 0 Then oSheet.getCellByPosition(colBefore, row).setValue(before)

    If colAfter >= 0 Then
        Dim qtyStr As String
        qtyStr = Trim(oSheet.getCellByPosition(colQty, row).getString())
        If qtyStr <> "" And IsNumeric(qtyStr) Then
            Dim delta As Double
            delta = CDbl(qtyStr)
            If Not isReceipt Then delta = -delta
            oSheet.getCellByPosition(colAfter, row).setValue(before + delta)
        End If
    End If
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
Public Sub PRIME_Workflow_ConductAllButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim sheetName As String
    sheetName = oSheet.Name
    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)

    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim batchKey As String
    batchKey = ""

    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(sheetName) To lastRow
        If Trim(oSheet.getCellByPosition(colCode, r).getString()) <> "" Then
            Dim docLine As PrimeDocLine
            Dim rowKey As String
            If PRIME_Workflow_BuildLine(oSheet, sheetName, headers, r, docLine, rowKey) Then
                If plan.LineCount = 0 Then
                    PRIME_InitPlan(plan, IIf(PRIME_Workflow_IsReceiptSheet(sheetName), DOC_RECEIPT, DOC_ISSUE), sheetName, "")
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
        PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, rowForLine(i))
    Next i
    MsgBox "Проведено строк: " & plan.LineCount & " (документ " & docId & ")"
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

    PRIME_Workflow_RefreshInlineStock(oSheet, sheetName, headers, row)
    MsgBox "Проведено: " & docId
End Sub

' Читает строку Приход/Расход-Цех/Офис и строит PrimeDocLine - общий для одиночного и
' батч-проведения. Контур определяется листом (PRIME_ContourForSheet), не вводится пользователем.
Private Function PRIME_Workflow_BuildLine(ByVal oSheet As Object, ByVal sheetName As String, ByVal headers As Variant, ByVal row As Long, _
        ByRef docLine As PrimeDocLine, ByRef draftId As String) As Boolean
    Dim isReceipt As Boolean
    isReceipt = PRIME_Workflow_IsReceiptSheet(sheetName)

    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_WFState")
    If colState >= 0 Then draftId = oSheet.getCellByPosition(colState, row).getString()
    If draftId = "" Then
        draftId = "WF-" & Format(PRIME_SequenceNext("WF_DRAFT_ID"), "00000000")
        If colState >= 0 Then oSheet.getCellByPosition(colState, row).setString(draftId)
    End If

    Dim code As String
    code = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), row).getString())
    If code = "" Then
        MsgBox "Строка " & (row + 1) & ": не указан внутренний код товара."
        PRIME_Workflow_BuildLine = False
        Exit Function
    End If

    Dim qtyStr As String
    qtyStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кол-во"), row).getString())
    If qtyStr = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Строка " & (row + 1) & ": заполните корректное ""Кол-во""."
        PRIME_Workflow_BuildLine = False
        Exit Function
    End If

    docLine.ProductCode = code
    docLine.QtyInput = CDbl(qtyStr)
    docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    docLine.Contour = PRIME_ContourForSheet(sheetName)
    docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Документ"), row).getString() & _
        " " & oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), row).getString()

    If isReceipt Then
        docLine.LocationTo = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
        docLine.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кто сдал"), row).getString()
    Else
        docLine.LocationFrom = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
        docLine.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кому"), row).getString()
        Dim colDest As Long
        colDest = PRIME_ColIndex(headers, "Назначение / проект")
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
        PRIME_Workflow_AutofillByCode(oSheet, headers, r)
        processed = processed + 1
    Next r
    MsgBox "Автозаполнение выполнено для " & processed & " строк."
End Sub

' "Повторить значения" - копирует общие поля (Место хранения/Кто сдал/Кому/Документ) из
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
    repeatCols = Array("Место хранения", "Кто сдал", "Кому", "Документ")
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
