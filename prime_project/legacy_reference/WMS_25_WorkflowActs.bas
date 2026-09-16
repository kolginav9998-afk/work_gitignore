Option Explicit

Global Const WMSACTWF_VERSION = "3.2.0-INTERNAL-ACTS"

Sub WMSACTWF_CreateFromActiveSheet()
    Dim oDoc As Object,oSh As Object,sel As Object
    Dim firstR As Long,lastR As Long,r As Long,rowsText As String,firstDataRow As Long
    Dim senderName As String,receiverName As String,personRow As String
    Dim title As String,partyFrom As String,partyTo As String,directionName As String,opID As String
    oDoc=ThisComponent:oSh=oDoc.CurrentController.ActiveSheet
    If Not WMSWF_IsWorkflowSheet(oSh.Name) Then Exit Sub
    On Error GoTo EH

    sel=oDoc.CurrentSelection
    firstR=sel.RangeAddress.StartRow:lastR=sel.RangeAddress.EndRow
    If firstR<5 Then firstR=5
    If lastR<firstR Then lastR=firstR

    firstDataRow=-1
    For r=firstR To lastR
        If WMSWF_RowHasData(oSh,r) Then
            personRow=Trim(oSh.getCellByPosition(5,r).String)
            If firstDataRow=-1 Then
                firstDataRow=r
            ElseIf UCase(personRow)<>UCase(Trim(oSh.getCellByPosition(5,firstDataRow).String)) Then
                MsgBox "В выделенных строках разные передавшие/получатели. Создайте отдельные акты для каждого человека.",48,"WMS — Акт":Exit Sub
            End If
            rowsText=rowsText & WMSACTWF_RowText(oSh,r) & Chr(10)
        End If
    Next r
    If Trim(rowsText)="" Then
        MsgBox "Выделите одну или несколько заполненных строк таблицы.",48,"WMS — Акт"
        Exit Sub
    End If

    WMSACTWF_Context oSh,firstDataRow,directionName,title,partyFrom,partyTo
    opID="ACTWF-" & Format(Now,"YYYYMMDDHHMMSS") & "-" & CStr(Int(Timer*10))
    If directionName="IN" Then
        senderName=Trim(oSh.getCellByPosition(5,firstDataRow).String)
    Else
        receiverName=Trim(oSh.getCellByPosition(5,firstDataRow).String)
    End If
    WMSACT_AskAndCreate directionName,partyFrom,partyTo,title,rowsText,opID,senderName,receiverName
    Exit Sub
EH:
    MsgBox "Создание акта: " & CStr(Err) & " " & Error$,16,"WMS — Акт"
End Sub

Function WMSACTWF_RowText(oSh As Object,r As Long) As String
    Dim nm As String,art As String,qty As String,unitName As String
    nm=Trim(oSh.getCellByPosition(1,r).String)
    art=Trim(oSh.getCellByPosition(2,r).String)
    qty=Trim(oSh.getCellByPosition(3,r).String)
    If qty="" Then qty=CStr(oSh.getCellByPosition(3,r).Value)
    unitName=Trim(oSh.getCellByPosition(4,r).String)
    WMSACTWF_RowText=nm & Chr(9) & art & Chr(9) & qty & Chr(9) & unitName
End Function

Sub WMSACTWF_Context(oSh As Object,r As Long,ByRef directionName As String,ByRef title As String,ByRef partyFrom As String,ByRef partyTo As String)
    Dim person As String
    person=Trim(oSh.getCellByPosition(5,r).String)
    Select Case oSh.Name
        Case "Приход — Производство","Приход — Цех"
            directionName="IN":title="АКТ ПРИЕМА-ПЕРЕДАЧИ ТМЦ — ПРОИЗВОДСТВО → СКЛАД":partyFrom="Производство":partyTo="Склад"
        Case "Приход — Офис"
            directionName="IN":title="АКТ ПРИЕМА-ПЕРЕДАЧИ ТМЦ — ОФИС → СКЛАД":partyFrom="Офис":partyTo="Склад"
        Case "Приход — Детали"
            directionName="IN":title="АКТ ПРИЕМА-ПЕРЕДАЧИ ДЕТАЛЕЙ НА СКЛАД":partyFrom="Передающая сторона":partyTo="Склад"
        Case "Расход — Производство","Расход — Цех"
            directionName="OUT":title="АКТ ПЕРЕДАЧИ ТМЦ — СКЛАД → ПРОИЗВОДСТВО":partyFrom="Склад":partyTo="Производство"
        Case "Расход — Офис"
            directionName="OUT":title="АКТ ПЕРЕДАЧИ ТМЦ — СКЛАД → ОФИС":partyFrom="Склад":partyTo="Офис"
        Case "Расход — Детали"
            directionName="OUT":title="АКТ ВЫДАЧИ ДЕТАЛЕЙ СО СКЛАДА":partyFrom="Склад":partyTo="Получатель"
    End Select
End Sub
