Option Explicit

Global Const WMSUI_VERSION = "3.2.0-CLEAN-LIGHT-UI"

Sub WMSUI_ApplyAll()
    Dim oDoc As Object,names As Variant,i As Long
    oDoc=ThisComponent
    On Error GoTo EH
    oDoc.lockControllers()
    names=Array("Приход — Производство","Приход — Офис","Приход — Детали","Расход — Производство","Расход — Офис","Расход — Детали")
    For i=LBound(names) To UBound(names)
        If oDoc.Sheets.hasByName(CStr(names(i))) Then WMSUI_Workflow oDoc.Sheets.getByName(CStr(names(i)))
    Next i
    If oDoc.Sheets.hasByName("Остаток") Then WMSUI_Stock oDoc.Sheets.getByName("Остаток")
    If oDoc.Sheets.hasByName("База - Поиск") Then WMSUI_Search oDoc.Sheets.getByName("База - Поиск")
    If oDoc.Sheets.hasByName("Справочники") Then WMSUI_Reference oDoc.Sheets.getByName("Справочники")
Done:
    On Error Resume Next:oDoc.unlockControllers():On Error GoTo 0
    Exit Sub
EH:
    Resume Done
End Sub

Sub WMSUI_Workflow(oSh As Object)
    Dim lastCol As Long
    lastCol=WMSUI_LastHeaderCol(oSh,4):If lastCol<0 Then Exit Sub
    On Error Resume Next
    oSh.Rows.getByIndex(0).Height=900
    oSh.Rows.getByIndex(1).Height=650
    oSh.Rows.getByIndex(2).Height=650
    oSh.Rows.getByIndex(3).Height=950
    oSh.Rows.getByIndex(4).Height=850
    oSh.getCellRangeByPosition(0,0,lastCol,0).CharHeight=14
    oSh.getCellRangeByPosition(0,1,lastCol,2).CharHeight=10
    oSh.getCellRangeByPosition(0,4,lastCol,4).CharHeight=10
    oSh.getCellRangeByPosition(0,4,lastCol,204).VertJustify=2
    oSh.getCellRangeByPosition(0,5,lastCol,204).CharHeight=10
    oSh.getCellRangeByPosition(0,5,lastCol,204).CellBackColor=-1
    oSh.getCellRangeByPosition(0,5,lastCol,204).IsTextWrapped=True
    On Error GoTo 0
End Sub

Sub WMSUI_Stock(oSh As Object)
    On Error Resume Next
    oSh.Rows.getByIndex(0).Height=900:oSh.Rows.getByIndex(1).Height=650:oSh.Rows.getByIndex(2).Height=950:oSh.Rows.getByIndex(5).Height=850
    oSh.getCellRangeByPosition(0,0,8,0).CharHeight=14
    oSh.getCellRangeByPosition(0,5,8,5).CharHeight=10
    oSh.getCellRangeByPosition(0,6,8,1005).CharHeight=10
    On Error GoTo 0
End Sub

Sub WMSUI_Search(oSh As Object)
    On Error Resume Next
    oSh.Rows.getByIndex(0).Height=900:oSh.Rows.getByIndex(2).Height=800:oSh.Rows.getByIndex(4).Height=950:oSh.Rows.getByIndex(5).Height=850
    oSh.getCellRangeByPosition(0,0,9,0).CharHeight=14
    oSh.getCellRangeByPosition(0,5,9,5).CharHeight=10
    On Error GoTo 0
End Sub

Sub WMSUI_Reference(oSh As Object)
    On Error Resume Next
    oSh.getCellRangeByPosition(0,0,6,200).CharHeight=10
    On Error GoTo 0
End Sub

Function WMSUI_LastHeaderCol(oSh As Object,rowIndex As Long) As Long
    Dim i As Long
    WMSUI_LastHeaderCol=-1
    For i=0 To 40
        If Trim(oSh.getCellByPosition(i,rowIndex).String)<>"" Then WMSUI_LastHeaderCol=i
    Next i
End Function
