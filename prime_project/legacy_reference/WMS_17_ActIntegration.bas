Option Explicit

Global Const WMSACTUI_VERSION = "0.3.0-PRODUCTION-ACTS"

Function WMSACTUI_SelectedRow(oSel As Object) As Long
    WMSACTUI_SelectedRow=0
    On Error GoTo TryRange

    WMSACTUI_SelectedRow=oSel.CellAddress.Row
    Exit Function

TryRange:
    Err=0
    On Error GoTo TryRanges
    WMSACTUI_SelectedRow=oSel.RangeAddress.StartRow
    Exit Function

TryRanges:
    Err=0
    On Error GoTo Fail
    WMSACTUI_SelectedRow=oSel.getRangeAddresses()(0).StartRow
    Exit Function

Fail:
    WMSACTUI_SelectedRow=0
End Function

Sub WMSACTUI_ReceiptActSelected()
    Dim oDoc As Object
    Dim oSh As Object
    Dim oSel As Object
    Dim r As Long
    Dim rowsText As String
    Dim partyFrom As String
    Dim operationID As String

    oDoc=ThisComponent
    oSh=oDoc.Sheets.getByName("Заказы")
    oSel=oDoc.CurrentSelection
    r=WMSACTUI_SelectedRow(oSel)

    If r<1 Then
        MsgBox "Выберите строку прихода на листе Заказы.",48,"WMS — Акт"
        Exit Sub
    End If

    If Not WMSACT_ShouldOfferReceipt(oSh,r) Then
        MsgBox "Акт предлагается для поступлений от Производства, Офиса и Деталей.",48,"WMS — Акт"
        Exit Sub
    End If

    operationID=WMSACT_ReceiptOperationID(oSh,r)
    partyFrom=WMSACT_ReceiptPartyFrom(oSh,r)
    rowsText=WMSACT_OrderRowsText(oSh,Trim(WMSDBO_TechText(oSh,r,"_WMS_ReceiptEventID")))

    If Trim(rowsText)="" Then
        rowsText=WMSACT_OrderRangeRowsText(oSh,r,r)
    End If

    WMSACT_AskAndCreate "IN",partyFrom,"Склад","Акт приёма-передачи ТМЦ",rowsText,operationID
End Sub

Sub WMSACTUI_IssueActSelected()
    Dim oDoc As Object
    Dim oSh As Object
    Dim oSel As Object
    Dim r As Long
    Dim toParty As String
    Dim rowsText As String
    Dim issueID As String

    oDoc=ThisComponent
    oSh=oDoc.Sheets.getByName("Выдачи")
    oSel=oDoc.CurrentSelection
    r=WMSACTUI_SelectedRow(oSel)

    If r<1 Then
        MsgBox "Выберите строку выдачи.",48,"WMS — Акт"
        Exit Sub
    End If

    toParty=Trim(oSh.getCellByPosition(5,r).String)
    If toParty="" Then
        MsgBox "Не указан получатель.",48,"WMS — Акт"
        Exit Sub
    End If

    rowsText=WMSACT_IssueRowText(oSh,r)
    issueID=Trim(Issues_CellText(oSh,r,"_WMS_SourceID"))
    If issueID="" Then
        issueID="ISSUE-" & Format(Now,"YYYYMMDD-HHMMSS")
    End If

    WMSACT_AskAndCreate "OUT","Склад",toParty,"Акт приёма-передачи ТМЦ",rowsText,issueID
End Sub

