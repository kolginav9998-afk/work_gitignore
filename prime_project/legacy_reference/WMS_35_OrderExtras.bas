Option Explicit

' Additive draft. Header lookup, no positional migration or database writes.
Global Const WMSOEX_VERSION = "0.1.0-DRAFT"

Sub WMSOEX_Install()
    Dim doc As Object,sh As Object,cur As Object,headers As Variant,h As Variant
    Dim c As Long,lastCol As Long,found As Long,lastRow As Long,fmt As Long
    Dim loc As New com.sun.star.lang.Locale
    doc=ThisComponent
    On Error GoTo EH
    sh=doc.Sheets.getByName("Заказы")
    cur=sh.createCursor():cur.gotoEndOfUsedArea(True)
    lastCol=cur.RangeAddress.EndColumn
    ' Do not write into a blank header above existing data.
    headers=Array("Назначение","Ожидаемая дата поступления","_WMS_DeliveryTiming","_WMS_ImportPayload")
    lastRow=WMSCore_LastContentRow(sh,lastCol):If lastRow<200 Then lastRow=200
    loc.Language="ru":loc.Country="RU"
    fmt=doc.NumberFormats.queryKey("DD.MM.YYYY",loc,True)
    If fmt=-1 Then fmt=doc.NumberFormats.addNew("DD.MM.YYYY",loc)
    For Each h In headers
        found=-1
        For c=0 To lastCol
            If Trim(sh.getCellByPosition(c,0).String)=CStr(h) Then
                If found>=0 Then MsgBox "Повторяется заголовок «" & CStr(h) & "». Уточните структуру перед обновлением.",48,"WMS":Exit Sub
                found=c
            End If
        Next c
        If found<0 Then
            lastCol=lastCol+1:found=lastCol
            sh.getCellByPosition(found,0).String=CStr(h)
        End If
        sh.Columns.getByIndex(found).IsVisible=(Left(CStr(h),5)<>"_WMS_")
        sh.Columns.getByIndex(found).Width=4900
        sh.getCellByPosition(found,0).CellBackColor=2240061
        sh.getCellByPosition(found,0).CharColor=16777215
        sh.getCellByPosition(found,0).CharWeight=150
        sh.getCellByPosition(found,0).IsTextWrapped=True
        If CStr(h)="Ожидаемая дата поступления" Then sh.getCellRangeByPosition(found,1,found,lastRow).NumberFormat=fmt
    Next h
    sh.Columns.getByIndex(4).IsVisible=False
    Orders_ResetHeaderCache
    WMSCore_InvalidateHeaderCache
    Exit Sub
EH:
    MsgBox "Дополнительные поля заказа: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSOEX_RefreshDeadlines()
    Dim sh As Object,lastRow As Long,lastCol As Long,cDate As Long,cStatus As Long,cQty As Long,cFact As Long,cGroup As Long
    Dim rows As Variant,item As Variant,deadlineRows() As Variant,keys() As String,dates() As Double,mixed() As Boolean,badDate() As Boolean
    Dim cap As Long,i As Long,slot As Long,found As Boolean,gid As String,d As Double,qty As Double,fact As Double,statusText As String
    On Error GoTo EH
    sh=ThisComponent.Sheets.getByName("Заказы")
    cDate=Orders_FindHeader(sh,"Ожидаемая дата поступления")
    cStatus=Orders_FindHeader(sh,"_WMS_DeliveryTiming")
    cQty=Orders_FindHeader(sh,"Количество"):cFact=Orders_FindHeader(sh,"Факт. количество")
    cGroup=Orders_FindHeader(sh,"_WMS_OrderGroupID")
    If cDate<0 Or cStatus<0 Or cQty<0 Or cFact<0 Or cGroup<0 Then
        MsgBox "Не найдены поля для проверки сроков. Сначала завершите обновление структуры заказов.",48,"WMS":Exit Sub
    End If
    lastCol=Orders_LastHeaderCol(sh)
    lastRow=WMSCore_LastContentRow(sh,lastCol)
    If lastRow<1 Then Exit Sub
    rows=sh.getCellRangeByPosition(0,1,lastCol,lastRow).getDataArray()
    ReDim deadlineRows(0 To UBound(rows)):ReDim badDate(0 To UBound(rows))
    cap=64:Do While cap<(UBound(rows)+1)*2:cap=cap*2:Loop
    ReDim keys(0 To cap-1):ReDim dates(0 To cap-1):ReDim mixed(0 To cap-1)
    ' Read explicit dates by stable group ID, independent of row adjacency.
    For i=0 To UBound(rows)
        item=rows(i):gid=Trim(CStr(item(cGroup))):d=0
        If VarType(item(cDate))=8 Then
            If Trim(CStr(item(cDate)))<>"" Then badDate(i)=True
        ElseIf IsNumeric(item(cDate)) Then
            d=CDbl(item(cDate))
            If d<0 Or d>2958465 Or d<>Int(d) Then badDate(i)=True:d=0
        End If
        If gid<>"" And d>0 Then
            slot=Orders_GroupHashSlot(keys(),cap,gid,found)
            If slot>=0 Then
                If Not found Then
                    keys(slot)=gid:dates(slot)=d
                ElseIf dates(slot)<>d Then
                    mixed(slot)=True
                End If
            End If
        End If
    Next i
    For i=0 To UBound(rows)
        item=rows(i):statusText="":d=0
        If Trim(CStr(item(1)))="" Then GoTo SaveStatus
        If badDate(i) Then statusText="Проверьте дату":GoTo SaveStatus
        If Not IsNumeric(item(cQty)) Then statusText="Проверьте количества":GoTo SaveStatus
        qty=CDbl(item(cQty)):fact=0
        If Trim(CStr(item(cFact)))<>"" Then
            If Not IsNumeric(item(cFact)) Then statusText="Проверьте количества":GoTo SaveStatus
            fact=CDbl(item(cFact))
        End If
        If qty<=0 Or fact<0 Then statusText="Проверьте количества":GoTo SaveStatus
        If fact>=qty Then statusText="Получено":GoTo SaveStatus
        If IsNumeric(item(cDate)) Then d=CDbl(item(cDate))
        gid=Trim(CStr(item(cGroup)))
        If d=0 And gid<>"" Then
            slot=Orders_GroupHashSlot(keys(),cap,gid,found)
            If slot>=0 And found Then
                If Not mixed(slot) Then d=dates(slot)
            End If
        End If
        If d=0 Then
            statusText="Срок не указан"
        ElseIf d<CDbl(Date) Then
            statusText="Просрочено на " & CStr(CLng(CDbl(Date)-d)) & " дн."
        ElseIf d=CDbl(Date) Then
            statusText="Ожидается сегодня"
        Else
            statusText="Ожидается " & Format(CDate(d),"DD.MM.YYYY")
        End If
SaveStatus:
        deadlineRows(i)=Array(statusText)
    Next i
    sh.getCellRangeByPosition(cStatus,1,cStatus,lastRow).setDataArray(deadlineRows())
    Exit Sub
EH:
    MsgBox "Сроки не обновлены: " & CStr(Err) & " " & Error$,48,"WMS"
End Sub
