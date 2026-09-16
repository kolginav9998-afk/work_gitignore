Option Explicit
Global Const WMSD_VERSION = "1.0.0"
Global Const WMSD_SHEET = "Дашборд"
Global Const WMSD_MARKER = "POKATAK_DASHBOARD_1"

Sub WMSD_Open()
    WMSD_Goto WMSD_SHEET
End Sub
Sub WMSD_OpenReport()
    WMSD_Goto "Отчет — ввод"
End Sub
Sub WMSD_OpenStock()
    WMSD_Goto "Остаток"
End Sub
Sub WMSD_OpenOrders()
    WMSD_Goto "Заказы"
End Sub
Sub WMSD_Goto(sheetName As String)
    On Error GoTo EH
    ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets.getByName(sheetName))
    Exit Sub
EH:
    MsgBox "Не удалось открыть лист «" & sheetName & "»: " & Error$,48,"ПОКАТАК"
End Sub

Sub WMSD_Refresh()
    Dim sh As Object,con As Object,st As Object,rs As Object,sErr As String
    Dim startDay As Double,endDay As Double,sql As String,periodSQL As String,filterSQL As String
    Dim moves As Variant,places As Variant,alerts As Variant,positiveCodes As Long,negativeGroups As Long,periodMoves As Long
    Dim truncatedMoves As Boolean,truncatedPlaces As Boolean,alertCount As Long,rowsScanned As Long
    Dim oldValues As Variant,haveBackup As Boolean,writing As Boolean,locked As Boolean
    If Not WMSDBX_TryEnter("DASHBOARD") Then MsgBox "Дождитесь завершения складской операции.",48,"ПОКАТАК":Exit Sub
    On Error GoTo EH
    sh=ThisComponent.Sheets.getByName(WMSD_SHEET)
    If sh.getCellByPosition(25,0).String<>WMSD_MARKER Then sErr="Лист дашборда не распознан.":GoTo Failed
    startDay=sh.getCellByPosition(1,2).Value:endDay=sh.getCellByPosition(3,2).Value
    If startDay<=0 Or endDay<startDay Or startDay<>Int(startDay) Or endDay<>Int(endDay) Or endDay>2958465 Then
        sErr="Укажите корректные даты периода: начало не позже окончания.":GoTo Failed
    End If
    filterSQL=" COALESCE(STATUS_NAME,'POSTED')='POSTED' "
    periodSQL=filterSQL & " AND MOVEMENT_DATE BETWEEN DATE '" & Format(CDate(startDay),"YYYY-MM-DD") & "' AND DATE '" & Format(CDate(endDay),"YYYY-MM-DD") & "' "
    con=WMSDB_GetConnectionEx(ThisComponent,sErr):If sErr<>"" Then GoTo Failed
    st=con.createStatement()
    sql="SELECT COUNT(DISTINCT PRODUCT_CODE) FROM (SELECT PRODUCT_CODE,UNIT_NAME,LOCATION_NAME,SUM(QTY) AS BAL FROM WMS_STOCK_MOVEMENTS WHERE " & filterSQL & " GROUP BY PRODUCT_CODE,UNIT_NAME,LOCATION_NAME HAVING SUM(QTY)>0) X"
    rs=st.executeQuery(sql):If rs.next() Then positiveCodes=rs.getLong(1)
    WMSDBX_CloseRS rs
    sql="SELECT COUNT(*) FROM (SELECT PRODUCT_CODE,UNIT_NAME,LOCATION_NAME FROM WMS_STOCK_MOVEMENTS WHERE " & filterSQL & " GROUP BY PRODUCT_CODE,UNIT_NAME,LOCATION_NAME HAVING SUM(QTY)<0) X"
    rs=st.executeQuery(sql):If rs.next() Then negativeGroups=rs.getLong(1)
    WMSDBX_CloseRS rs
    rs=st.executeQuery("SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE " & periodSQL)
    If rs.next() Then periodMoves=rs.getLong(1)
    WMSDBX_CloseRS rs
    sql="SELECT MOVEMENT_TYPE,UNIT_NAME,COUNT(*),COUNT(DISTINCT PRODUCT_CODE),CAST(SUM(QTY) AS DOUBLE PRECISION) FROM WMS_STOCK_MOVEMENTS WHERE " & periodSQL & " GROUP BY MOVEMENT_TYPE,UNIT_NAME ORDER BY MOVEMENT_TYPE,UNIT_NAME"
    moves=WMSD_Table(st,sql,8,truncatedMoves)
    sql="SELECT LOCATION_NAME,UNIT_NAME,COUNT(DISTINCT PRODUCT_CODE),CAST(SUM(QTY) AS DOUBLE PRECISION),COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE " & filterSQL & " GROUP BY LOCATION_NAME,UNIT_NAME ORDER BY LOCATION_NAME,UNIT_NAME"
    places=WMSD_Table(st,sql,10,truncatedPlaces)
    alerts=WMSD_Alerts(alertCount,rowsScanned,sErr)
    If sErr<>"" Then GoTo Failed
    ' All reads finished before replacing the previous dashboard snapshot.
    oldValues=sh.getCellRangeByPosition(0,5,5,67).getDataArray():haveBackup=True
    ThisComponent.lockControllers():locked=True:writing=True
    sh.getCellByPosition(0,8).Value=positiveCodes
    sh.getCellByPosition(2,8).Value=negativeGroups
    sh.getCellByPosition(4,8).Value=periodMoves
    sh.getCellRangeByPosition(0,12,4,19).setDataArray(moves)
    sh.getCellRangeByPosition(0,23,4,32).setDataArray(places)
    sh.getCellRangeByPosition(0,37,5,66).setDataArray(alerts)
    sh.getCellByPosition(0,20).String="Все группы периода показаны. OUT — отрицательное количество; единицы не складываются между собой."
    If truncatedMoves Then sh.getCellByPosition(0,20).String="Показаны первые 8 групп. Полный журнал — на листах операций."
    sh.getCellByPosition(0,33).String="Текущий остаток по местам и единицам; период сверху к остаткам не применяется."
    If truncatedPlaces Then sh.getCellByPosition(0,33).String="Показаны первые 10 групп мест/единиц. Полный список — кнопка «Остатки»."
    sh.getCellByPosition(0,67).String="Строк заказов просмотрено: " & CStr(rowsScanned) & ". Требуют внимания: " & CStr(alertCount) & ". В списке не более 30."
    sh.getCellByPosition(0,5).String="Обновлено " & Format(Now,"DD.MM.YYYY HH:MM:SS") & ". Движения — из базы; контроль заказов — из листа."
    sh.getCellByPosition(0,5).CellBackColor=14808549
    writing=False
Done:
    On Error Resume Next
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    If locked Then ThisComponent.unlockControllers()
    WMSDBX_Leave
    Exit Sub
Failed:
    On Error Resume Next
    If writing And haveBackup Then sh.getCellRangeByPosition(0,5,5,67).setDataArray(oldValues)
    sh.getCellByPosition(0,5).String="НЕ ОБНОВЛЕНО: " & sErr & ". Предыдущий снимок не актуализирован."
    sh.getCellByPosition(0,5).CellBackColor=16768991
    MsgBox sErr,48,"ПОКАТАК — Дашборд"
    GoTo Done
EH:
    sErr=CStr(Err) & " " & Error$
    Resume Failed
End Sub

Function WMSD_Table(st As Object,sql As String,limitRows As Long,ByRef hasMore As Boolean) As Variant
    Dim rs As Object,items() As Variant,i As Long,n As Long,errNo As Long,errText As String
    On Error GoTo EH
    ReDim items(0 To limitRows-1)
    For i=0 To limitRows-1:items(i)=Array("","","","",""):Next i
    rs=st.executeQuery(sql):n=0:hasMore=False
    Do While rs.next()
        If n>=limitRows Then hasMore=True:Exit Do
        items(n)=Array(rs.getString(1),rs.getString(2),rs.getLong(3),rs.getDouble(4),rs.getDouble(5))
        n=n+1
    Loop
    WMSDBX_CloseRS rs
    WMSD_Table=items()
    Exit Function
EH:
    errNo=Err:errText=Error$
    WMSDBX_CloseRS rs
    Error errNo
End Function

Function WMSD_Alerts(ByRef total As Long,ByRef scanned As Long,ByRef sErr As String) As Variant
    Dim sh As Object,lastR As Long,lastC As Long,cDate As Long,cDest As Long,r As Long,i As Long
    Dim data As Variant,item As Variant,items() As Variant,reason As String,statusText As String,expected As Double,dest As String
    sh=ThisComponent.Sheets.getByName("Заказы")
    lastC=Orders_LastHeaderCol(sh):lastR=WMSCore_LastContentRow(sh,lastC)
    sErr=""
    If lastR>20000 Then
        sErr="Дашборд не обновлён: в заказах более 20 000 строк. Предыдущий снимок сохранён; требуется обработка большего объёма."
        Exit Function
    End If
    cDate=Orders_FindHeader(sh,"Ожидаемая дата поступления")
    cDest=Orders_FindHeader(sh,"Назначение")
    ReDim items(0 To 29):For i=0 To 29:items(i)=Array("","","","","",""):Next i
    total=0:scanned=lastR
    If lastR<1 Then WMSD_Alerts=items():Exit Function
    data=sh.getCellRangeByPosition(0,1,lastC,lastR).getDataArray()
    For r=0 To UBound(data)
        item=data(r):reason="":expected=0:dest=""
        If Trim(CStr(item(1)))<>"" Then
            statusText=UCase(Trim(CStr(item(16))))
            If statusText<>"ОТМЕНЕНО" Then
                If Trim(CStr(item(11)))="" Then reason="Нет поставщика; "
                If VarType(item(6))<>8 Then
                    If Trim(CStr(item(2)))="" And CDbl(item(6))>0 Then reason=reason & "Факт без документа; "
                Else
                    If Trim(CStr(item(6)))<>"" Then reason=reason & "Факт записан текстом; "
                End If
                If cDate>=0 Then
                    If VarType(item(cDate))<>8 Then expected=CDbl(item(cDate))
                    If expected>0 And expected<CDbl(Date) And statusText<>"ПОЛУЧЕНО" And statusText<>"ОПРИХОДОВАНО" And statusText<>"ПОЛУЧЕНО БЕЗ ДОКУМЕНТОВ" Then reason=reason & "Срок прошёл; проверьте статус; "
                End If
                If cDest>=0 Then dest=CStr(item(cDest))
                If Right(reason,2)="; " Then reason=Left(reason,Len(reason)-2)
                If reason<>"" Then
                    If total<30 Then items(total)=Array(reason,CStr(item(1)),CStr(item(3)),CStr(item(11)),dest,"Строка " & CStr(r+2) & ": " & CStr(item(16)))
                    total=total+1
                End If
            End If
        End If
    Next r
    WMSD_Alerts=items()
End Function

Function WMSD_ReportDetails(con As Object,dayValue As Double) As String
    Dim st As Object,rs As Object,sql As String,body As String,n As Long,docNo As String,supplier As String,errNo As Long
    On Error GoTo EH
    sql="SELECT l.SUPPLIER_NAME,l.DOC_NO,m.UNIT_NAME,CAST(SUM(m.QTY) AS DOUBLE PRECISION),COUNT(DISTINCT m.PRODUCT_CODE) " & _
        "FROM WMS_ORDER_LINES l JOIN WMS_STOCK_MOVEMENTS m ON m.MOVEMENT_ID='IN-' || l.SOURCE_ID " & _
        "WHERE COALESCE(m.STATUS_NAME,'POSTED')='POSTED' AND m.MOVEMENT_TYPE='IN' AND m.MOVEMENT_DATE=DATE '" & Format(CDate(dayValue),"YYYY-MM-DD") & "' " & _
        "GROUP BY l.SUPPLIER_NAME,l.DOC_NO,m.UNIT_NAME ORDER BY l.SUPPLIER_NAME,l.DOC_NO,m.UNIT_NAME"
    st=con.createStatement():rs=st.executeQuery(sql):n=0:body=""
    Do While rs.next()
        If n>=20 Then body=body & "• Детализация ограничена первыми 20 группами поставщик/документ/единица; полный перечень — в заказах." & Chr(10):Exit Do
        supplier=Trim(rs.getString(1)):docNo=Trim(rs.getString(2))
        If supplier="" Then supplier="поставщик не указан"
        If docNo="" Then docNo="закрывающий документ не указан"
        body=body & "• Приход: " & supplier & "; " & docNo & "; " & CStr(rs.getDouble(4)) & " " & rs.getString(3) & _
            " в базовых единицах; кодов товаров: " & CStr(rs.getLong(5)) & "." & Chr(10)
        n=n+1
    Loop
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    WMSD_ReportDetails=body
    Exit Function
EH:
    errNo=Err
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Error errNo
End Function
