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

' PRIME_15_Transfers (новый в 2.1.0, R05/R25)
' Лист "Перемещения" - TRANSFER между контурами/местами хранения с сохранением партийности
' (lot lineage, см. PRIME_04_Posting.PRIME_PostTransferLines). Раньше перемещение не имело
' собственного пользовательского листа вообще.

Public Sub PRIME_OnContentChanged_Transfers(ByVal oRangeAddr As Variant)
    If Not PRIME_EventEnter() Then Exit Sub
    On Error Goto CleanExit
    If IsNull(oRangeAddr) Then GoTo CleanExit

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_TRANSFERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_TRANSFERS)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= 1 Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Transfers_AutofillByCode(oSheet, headers, r)
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
End Sub

Private Sub PRIME_Transfers_AutofillByCode(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    PRIME_SetCellIfEmpty(oSheet, headers, row, "Наименование", PRIME_GetProductField(code, "PRODUCT_NAME"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Ед. изм.", PRIME_GetProductField(code, "BASE_UNIT"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Контур — откуда", "Склад")
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Место — откуда", PRIME_GetProductField(code, "DEFAULT_LOCATION"))
End Sub

Public Sub PRIME_Transfers_NewButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_TRANSFERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_TRANSFERS)
    Dim newRow As Long
    newRow = PRIME_FindLastRow(oSheet) + 1
    If newRow < 1 Then newRow = 1
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Дата"), newRow).setString(Format(Now, "YYYY-MM-DD"))
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_TransferState"), newRow).setString("")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    ThisComponent.CurrentController.select(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), newRow))
End Sub

Public Sub PRIME_Transfers_ConductSelectedButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_TRANSFERS)
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < 1 Then Exit Sub
    PRIME_Transfers_ConductRow(oSheet, row)
End Sub

' R09: несколько отмеченных строк перемещения проводятся ОДНИМ документом; R10: несколько строк
' одного товара с одного и того же (места, контура) резервируют остаток друг у друга (см.
' PRIME_04_Posting.PRIME_ValidateTransfer).
Public Sub PRIME_Transfers_ConductAllButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_TRANSFERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_TRANSFERS)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Внутренний код")
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)

    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim batchKey As String
    batchKey = ""

    Dim r As Long
    For r = 1 To lastRow
        If Trim(oSheet.getCellByPosition(colCode, r).getString()) <> "" Then
            Dim docLine As PrimeDocLine
            Dim rowKey As String
            If PRIME_Transfers_BuildLine(oSheet, headers, r, docLine, rowKey) Then
                If plan.LineCount = 0 Then PRIME_InitPlan(plan, DOC_TRANSFER, SH_TRANSFERS, "")
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
    MsgBox "Проведено перемещений: " & plan.LineCount & " (документ " & docId & ")"
End Sub

Public Sub PRIME_Transfers_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_TRANSFERS)

    Dim docLine As PrimeDocLine
    Dim draftId As String
    If Not PRIME_Transfers_BuildLine(oSheet, headers, row, docLine, draftId) Then Exit Sub

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_TRANSFER, SH_TRANSFERS, draftId)
    PRIME_PlanAddLine(plan, docLine)

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If
    MsgBox "Перемещение проведено: " & docId
End Sub

Private Function PRIME_Transfers_BuildLine(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, _
        ByRef docLine As PrimeDocLine, ByRef draftId As String) As Boolean
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_TransferState")
    If colState >= 0 Then draftId = oSheet.getCellByPosition(colState, row).getString()
    If draftId = "" Then
        draftId = "TRF-" & Format(PRIME_SequenceNext("TRANSFER_DRAFT_ID"), "00000000")
        If colState >= 0 Then oSheet.getCellByPosition(colState, row).setString(draftId)
    End If

    Dim code As String
    code = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Внутренний код"), row).getString())
    If code = "" Then
        MsgBox "Строка " & (row + 1) & ": не указан внутренний код товара."
        PRIME_Transfers_BuildLine = False
        Exit Function
    End If

    Dim qtyStr As String
    qtyStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кол-во"), row).getString())
    If qtyStr = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Строка " & (row + 1) & ": заполните корректное ""Кол-во""."
        PRIME_Transfers_BuildLine = False
        Exit Function
    End If

    docLine.ProductCode = code
    docLine.QtyInput = CDbl(qtyStr)
    docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    docLine.LocationFrom = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место — откуда"), row).getString()
    docLine.LocationTo = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место — куда"), row).getString()
    docLine.ContourFrom = PRIME_ContourFromDisplayName(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Контур — откуда"), row).getString())
    docLine.ContourTo = PRIME_ContourFromDisplayName(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Контур — куда"), row).getString())
    docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), row).getString()
    PRIME_Transfers_BuildLine = True
End Function
