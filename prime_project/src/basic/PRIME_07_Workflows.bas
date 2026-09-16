Option Explicit

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

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= 1 Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Workflow_AutofillByCode(oSheet, headers, r)
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

' === Кнопки (одни и те же имена процедур используются на всех 4 листах; активный лист
' определяется через ThisComponent.CurrentController.ActiveSheet, а не жёстко зашит) ===========

Public Sub PRIME_Workflow_ConductRowButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < 1 Then Exit Sub
    PRIME_Workflow_ConductRow(oSheet, row)
End Sub

Public Sub PRIME_Workflow_ConductAllButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim headers As Variant
    headers = PRIME_HeaderMap(oSheet.Name)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        If Trim(oSheet.getCellByPosition(colCode, r).getString()) <> "" Then
            PRIME_Workflow_ConductRow(oSheet, r)
        End If
    Next r
End Sub

Public Sub PRIME_Workflow_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim sheetName As String
    sheetName = oSheet.Name
    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)
    Dim isReceipt As Boolean
    isReceipt = PRIME_Workflow_IsReceiptSheet(sheetName)

    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_WFState")
    Dim draftId As String
    If colState >= 0 Then draftId = oSheet.getCellByPosition(colState, row).getString()
    If draftId = "" Then
        draftId = "WF-" & Format(PRIME_SequenceNext("WF_DRAFT_ID"), "00000000")
        If colState >= 0 Then oSheet.getCellByPosition(colState, row).setString(draftId)
    End If

    Dim code As String
    code = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), row).getString())
    If code = "" Then
        MsgBox "Строка " & (row + 1) & ": не указан внутренний код товара."
        Exit Sub
    End If

    Dim qtyStr As String
    qtyStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кол-во"), row).getString())
    If qtyStr = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Строка " & (row + 1) & ": заполните корректное ""Кол-во""."
        Exit Sub
    End If

    Dim plan As PrimeDocPlan
    Dim line As PrimeDocLine
    line.ProductCode = code
    line.QtyInput = CDbl(qtyStr)
    line.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    line.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Документ"), row).getString() & _
        " " & oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), row).getString()

    If isReceipt Then
        PRIME_InitPlan(plan, DOC_RECEIPT, sheetName, draftId)
        line.LocationTo = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
        line.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кто сдал"), row).getString()
    Else
        PRIME_InitPlan(plan, DOC_ISSUE, sheetName, draftId)
        line.LocationFrom = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
        line.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кому"), row).getString()
        Dim colDest As Long
        colDest = PRIME_ColIndex(headers, "Назначение / проект")
        If colDest >= 0 Then line.DestinationProject = oSheet.getCellByPosition(colDest, row).getString()
    End If
    PRIME_PlanAddLine(plan, line)

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    MsgBox "Проведено: " & docId
End Sub

Public Sub PRIME_Workflow_FillArticleButton()
    Dim oSheet As Object
    oSheet = ThisComponent.CurrentController.ActiveSheet
    Dim headers As Variant
    headers = PRIME_HeaderMap(oSheet.Name)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        PRIME_Workflow_AutofillByCode(oSheet, headers, r)
    Next r
    MsgBox "Автозаполнение выполнено для " & lastRow & " строк."
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
    If row < 2 Then Exit Sub

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
