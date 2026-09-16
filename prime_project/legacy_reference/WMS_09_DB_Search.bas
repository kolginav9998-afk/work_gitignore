Option Explicit

' ============================================================================
' WMS_09_DB_Search.bas
' v0.1.0 — "База / Поиск" interface for portable Firebird WMS.
'
' Reads history from:
'   WMS_ORDER_LINES
'   WMS_ISSUES
'   WMS_ISSUE_RETURNS
'
' Calc is only the viewer. Historical records remain in Firebird.
' ============================================================================

Global Const WMSDBS_VERSION = "0.2.1-RECEIPT-EVENT-SEARCH"
Global Const WMSDBS_SHEET = "База - Поиск"
Global Const WMSDBS_FORM = "WMS_DB_SEARCH_PANEL"
Global Const WMSDBS_MAX_ROWS = 1000
Global gWMSDBS_Busy As Boolean

Sub WMSDBS_Install()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    If Not WMSDBS_GetOrCreateSheet(oDoc,oSh,sErr) Then
        MsgBox sErr,16,"WMS — База / Поиск"
        Exit Sub
    End If

    WMSDBS_BuildLayout oDoc,oSh
    If Not WMSDBS_InstallButtons(oDoc,oSh,sErr) Then
        MsgBox "Лист создан, но кнопки не установлены:" & Chr(10) & sErr,48,"WMS — База / Поиск"
        Exit Sub
    End If

    oDoc.CurrentController.setActiveSheet(oSh)
    MsgBox "Интерфейс 'База / Поиск' установлен." & Chr(10) & _
           "Версия: " & WMSDBS_VERSION,64,"WMS — База / Поиск"
    Exit Sub
EH:
    MsgBox "Установка 'База / Поиск': " & CStr(Err) & " " & Error$,16,"WMS — База / Поиск"
End Sub


Function WMSDBS_GetOrCreateSheet(oDoc As Object,ByRef oSh As Object,ByRef sErr As String) As Boolean
    WMSDBS_GetOrCreateSheet=False:sErr=""
    On Error GoTo EH

    If oDoc.Sheets.hasByName(WMSDBS_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSDBS_SHEET)
        WMSDBS_GetOrCreateSheet=True
        Exit Function
    End If

    If oDoc.Sheets.hasByName("База_Поиск") Then
        oSh=oDoc.Sheets.getByName("База_Поиск")
        On Error Resume Next
        oSh.Name=WMSDBS_SHEET
        On Error GoTo EH
        WMSDBS_GetOrCreateSheet=True
        Exit Function
    End If

    oDoc.Sheets.insertNewByName(WMSDBS_SHEET,oDoc.Sheets.Count)
    oSh=oDoc.Sheets.getByName(WMSDBS_SHEET)
    WMSDBS_GetOrCreateSheet=True
    Exit Function
EH:
    sErr="Не удалось создать лист '" & WMSDBS_SHEET & "': " & CStr(Err) & " " & Error$
End Function


Sub WMSDBS_Search()
    Dim oDoc As Object,oSh As Object,sErr As String,n As Long
    oDoc=ThisComponent

    If gWMSDBS_Busy Then Exit Sub
    If Not WMSDBS_GetSheet(oDoc,oSh) Then Exit Sub

    gWMSDBS_Busy=True
    On Error GoTo EH
    oDoc.lockControllers()

    WMSDBS_ClearResults oSh
    n=WMSDBS_RunSearch(oDoc,oSh,False,sErr)

    oDoc.unlockControllers()
    gWMSDBS_Busy=False

    If sErr<>"" Then
        MsgBox sErr,16,"WMS — База / Поиск"
    Else
        oSh.getCellByPosition(0,9).String="Найдено записей: " & CStr(n)
    End If
    Exit Sub

EH:
    On Error Resume Next
    oDoc.unlockControllers()
    WMSDB_Close
    gWMSDBS_Busy=False
    MsgBox "Поиск остановлен: " & CStr(Err) & " " & Error$,16,"WMS — База / Поиск"
End Sub

Sub WMSDBS_Today()
    Dim oSh As Object
    If Not WMSDBS_GetSheet(ThisComponent,oSh) Then Exit Sub
    oSh.getCellByPosition(1,4).Value=CDbl(Date)
    oSh.getCellByPosition(1,5).Value=CDbl(Date)
    WMSDBS_FormatDateCell ThisComponent,oSh.getCellByPosition(1,4)
    WMSDBS_FormatDateCell ThisComponent,oSh.getCellByPosition(1,5)
    WMSDBS_Search
End Sub

Sub WMSDBS_Last7Days()
    Dim oSh As Object
    If Not WMSDBS_GetSheet(ThisComponent,oSh) Then Exit Sub
    oSh.getCellByPosition(1,4).Value=CDbl(Date-6)
    oSh.getCellByPosition(1,5).Value=CDbl(Date)
    WMSDBS_FormatDateCell ThisComponent,oSh.getCellByPosition(1,4)
    WMSDBS_FormatDateCell ThisComponent,oSh.getCellByPosition(1,5)
    WMSDBS_Search
End Sub

Sub WMSDBS_OpenReturns()
    Dim oDoc As Object,oSh As Object,sErr As String,n As Long
    oDoc=ThisComponent
    If Not WMSDBS_GetSheet(oDoc,oSh) Then Exit Sub
    WMSDBS_ClearResults oSh
    n=WMSDBS_QueryOpenReturns(oDoc,oSh,11,sErr)
    If sErr<>"" Then
        MsgBox sErr,16,"WMS — Открытые возвраты"
    Else
        oSh.getCellByPosition(0,9).String="Открытых возвратных выдач: " & CStr(n)
    End If
End Sub

Sub WMSDBS_ProductHistory()
    Dim oSh As Object,s As String
    If Not WMSDBS_GetSheet(ThisComponent,oSh) Then Exit Sub
    s=Trim(oSh.getCellByPosition(1,6).String)
    If s="" Then
        s=InputBox("Введите название товара, артикул или код товара:","WMS — История товара")
        If Trim(s)="" Then Exit Sub
        oSh.getCellByPosition(1,6).String=s
    End If
    oSh.getCellByPosition(1,2).String="Все"
    WMSDBS_Search
End Sub

Sub WMSDBS_Clear()
    Dim oDoc As Object,oSh As Object
    oDoc=ThisComponent
    If gWMSDBS_Busy Then Exit Sub
    If Not WMSDBS_GetSheet(oDoc,oSh) Then Exit Sub

    gWMSDBS_Busy=True
    On Error GoTo EH
    oDoc.lockControllers()

    oSh.getCellByPosition(1,2).String="Все"
    oSh.getCellByPosition(1,3).String=""
    oSh.getCellByPosition(1,4).String=""
    oSh.getCellByPosition(1,5).String=""
    oSh.getCellByPosition(1,6).String=""
    oSh.getCellByPosition(1,7).String=""
    oSh.getCellByPosition(1,8).String=""
    WMSDBS_ClearResults oSh
    oSh.getCellByPosition(0,9).String="Фильтры очищены."

    oDoc.unlockControllers()
    gWMSDBS_Busy=False
    Exit Sub

EH:
    On Error Resume Next
    oDoc.unlockControllers()
    gWMSDBS_Busy=False
    MsgBox "Очистка остановлена: " & CStr(Err) & " " & Error$,16,"WMS — База / Поиск"
End Sub

Sub WMSDBS_Refresh()
    WMSDBS_Search
End Sub

Sub WMSDBS_SelfCheck()
    Dim oCon As Object,sErr As String,s As String
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — База / Поиск":Exit Sub

    s="Firebird: OK" & Chr(10) & _
      "WMS_ORDER_LINES: " & WMSDBS_TableState(oCon,"WMS_ORDER_LINES") & Chr(10) & _
      "WMS_ISSUES: " & WMSDBS_TableState(oCon,"WMS_ISSUES") & Chr(10) & _
      "WMS_ISSUE_RETURNS: " & WMSDBS_TableState(oCon,"WMS_ISSUE_RETURNS") & Chr(10) & _
      "Модуль: " & WMSDBS_VERSION
    WMSDB_Close
    MsgBox s,64,"WMS — База / Поиск"
End Sub

Function WMSDBS_RunSearch(oDoc As Object,oSh As Object,openOnly As Boolean,ByRef sErr As String) As Long
    Dim typ As String,rowOut As Long,n As Long,total As Long,oCon As Object
    WMSDBS_RunSearch=0
    typ=WMSDBS_CanonType(oSh.getCellByPosition(1,2).String)
    rowOut=11:total=0:sErr=""

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    On Error GoTo EH

    If typ="ALL" Or typ="ORDERS" Then
        n=WMSDBS_QueryOrdersCon(oCon,oSh,rowOut,sErr)
        If sErr<>"" Then GoTo Finish
        total=total+n:rowOut=rowOut+n
    End If

    If typ="ALL" Or typ="ISSUES" Then
        n=WMSDBS_QueryIssuesCon(oCon,oSh,rowOut,sErr)
        If sErr<>"" Then GoTo Finish
        total=total+n:rowOut=rowOut+n
    End If

    If typ="ALL" Or typ="RETURNS" Then
        n=WMSDBS_QueryReturnsCon(oCon,oSh,rowOut,sErr)
        If sErr<>"" Then GoTo Finish
        total=total+n
    End If

Finish:
    WMSDB_Close
    WMSDBS_RunSearch=total
    Exit Function

EH:
    sErr="Ошибка запроса Firebird: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
    WMSDBS_RunSearch=total
End Function


Function WMSDBS_QueryOrdersCon(oCon As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object,sql As String,whereSQL As String,q As String,p As String,s As String,d As String
    Dim hasV2 As Boolean
    WMSDBS_QueryOrdersCon=0:sErr=""
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ORDER_LINES",sErr) Then Exit Function

    hasV2=WMSDBS_ColumnExists(oCon,"WMS_ORDER_LINES","RECEIPT_EVENT_ID")
    q=Trim(oSh.getCellByPosition(1,3).String)
    p=Trim(oSh.getCellByPosition(1,6).String)
    s=Trim(oSh.getCellByPosition(1,7).String)
    d=Trim(oSh.getCellByPosition(1,8).String)

    whereSQL=" WHERE 1=1"
    If q<>"" Then
        whereSQL=whereSQL & " AND (o.PRODUCT_NAME CONTAINING " & WMSDBS_Q(q) & _
            " OR o.PRODUCT_CODE CONTAINING " & WMSDBS_Q(q) & _
            " OR o.SUPPLIER_NAME CONTAINING " & WMSDBS_Q(q) & _
            " OR o.DOC_NO CONTAINING " & WMSDBS_Q(q) & _
            " OR o.SUPPLIER_ARTICLE CONTAINING " & WMSDBS_Q(q)
        If hasV2 Then whereSQL=whereSQL & " OR o.RECEIPT_EVENT_ID CONTAINING " & WMSDBS_Q(q) & _
            " OR o.EXTERNAL_ORDER_NO CONTAINING " & WMSDBS_Q(q) & _
            " OR o.UPD_NO CONTAINING " & WMSDBS_Q(q)
        whereSQL=whereSQL & ")"
    End If
    If p<>"" Then
        whereSQL=whereSQL & " AND (o.PRODUCT_NAME CONTAINING " & WMSDBS_Q(p) & _
            " OR o.PRODUCT_CODE CONTAINING " & WMSDBS_Q(p) & _
            " OR o.SUPPLIER_ARTICLE CONTAINING " & WMSDBS_Q(p) & ")"
    End If
    If s<>"" Then whereSQL=whereSQL & " AND o.SUPPLIER_NAME CONTAINING " & WMSDBS_Q(s)
    If d<>"" Then
        whereSQL=whereSQL & " AND (o.DOC_NO CONTAINING " & WMSDBS_Q(d)
        If hasV2 Then whereSQL=whereSQL & " OR o.RECEIPT_EVENT_ID CONTAINING " & WMSDBS_Q(d) & _
            " OR o.EXTERNAL_ORDER_NO CONTAINING " & WMSDBS_Q(d) & _
            " OR o.UPD_NO CONTAINING " & WMSDBS_Q(d)
        whereSQL=whereSQL & ")"
    End If
    whereSQL=whereSQL & WMSDBS_DateWhere(oSh,"o.RECEIPT_DATE")

    If hasV2 Then
        sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
            " o.RECEIPT_DATE,o.PRODUCT_NAME,o.PRODUCT_CODE,o.FACT_QTY,o.UNIT_NAME,o.SUPPLIER_NAME,o.DOC_NO," & _
            "o.STATUS_NAME,o.SOURCE_ID," & _
            "COALESCE(o.SUPPLIER_ARTICLE,'') || CASE WHEN COALESCE(o.RECEIPT_EVENT_ID,'')='' THEN '' ELSE ' | Receipt=' || o.RECEIPT_EVENT_ID END || " & _
            "CASE WHEN COALESCE(o.EXTERNAL_ORDER_NO,'')='' THEN '' ELSE ' | Order=' || o.EXTERNAL_ORDER_NO END || " & _
            "CASE WHEN COALESCE(o.UPD_NO,'')='' THEN '' ELSE ' | UPD=' || o.UPD_NO END," & _
            "o.INVOICE_NO,o.LOCATION_NAME," & _
            "COALESCE(l.LOT_ID,''),COALESCE(l.DOC_QTY,0),COALESCE(l.DOC_UNIT,''),COALESCE(l.BASE_QTY,0),COALESCE(l.BASE_UNIT,'') " & _
            "FROM WMS_ORDER_LINES o LEFT JOIN WMS_STOCK_LOTS l ON l.SOURCE_ID=o.SOURCE_ID" & _
            whereSQL & " ORDER BY o.RECEIPT_DATE DESC,o.CREATED_AT DESC"
    Else
        sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
            " o.RECEIPT_DATE,o.PRODUCT_NAME,o.PRODUCT_CODE,o.FACT_QTY,o.UNIT_NAME,o.SUPPLIER_NAME,o.DOC_NO," & _
            "o.STATUS_NAME,o.SOURCE_ID,COALESCE(o.SUPPLIER_ARTICLE,''),o.INVOICE_NO,o.LOCATION_NAME," & _
            "COALESCE(l.LOT_ID,''),COALESCE(l.DOC_QTY,0),COALESCE(l.DOC_UNIT,''),COALESCE(l.BASE_QTY,0),COALESCE(l.BASE_UNIT,'') " & _
            "FROM WMS_ORDER_LINES o LEFT JOIN WMS_STOCK_LOTS l ON l.SOURCE_ID=o.SOURCE_ID" & _
            whereSQL & " ORDER BY o.RECEIPT_DATE DESC,o.CREATED_AT DESC"
    End If

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    WMSDBS_QueryOrdersCon=WMSDBS_WriteRSLot(oSh,rowOut,oRS,"Приход")
End Function

Function WMSDBS_QueryIssuesCon(oCon As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object,sql As String,whereSQL As String
    WMSDBS_QueryIssuesCon=0:sErr=""
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ISSUES",sErr) Then Exit Function
    whereSQL=WMSDBS_CommonWhere(oSh,"i.ISSUE_DATE","i.PRODUCT_NAME","i.PRODUCT_CODE","i.EMPLOYEE_NAME","i.SOURCE_ID")
    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " i.ISSUE_DATE,i.PRODUCT_NAME,i.PRODUCT_CODE,i.ISSUE_QTY,i.UNIT_NAME,i.EMPLOYEE_NAME,i.SOURCE_ID," & _
        "i.RETURN_STATE,i.SOURCE_ID,'',i.RETURNABLE,i.SOURCE_LOCATION," & _
        "COALESCE(i.LOT_ID,''),COALESCE(i.ENTRY_QTY,i.ISSUE_QTY),COALESCE(i.ENTRY_UNIT,i.UNIT_NAME)," & _
        "COALESCE(i.BASE_QTY,i.ISSUE_QTY),COALESCE(i.BASE_UNIT,i.UNIT_NAME) " & _
        "FROM WMS_ISSUES i" & whereSQL & " ORDER BY i.ISSUE_DATE DESC,i.CREATED_AT DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    WMSDBS_QueryIssuesCon=WMSDBS_WriteRSLot(oSh,rowOut,oRS,"Выдача")
End Function

Function WMSDBS_QueryReturnsCon(oCon As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object,sql As String,whereSQL As String
    WMSDBS_QueryReturnsCon=0:sErr=""
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ISSUE_RETURNS",sErr) Then Exit Function
    whereSQL=WMSDBS_ReturnWhere(oSh)
    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " r.RETURN_DATE,i.PRODUCT_NAME,i.PRODUCT_CODE,r.RETURN_QTY,i.UNIT_NAME,i.EMPLOYEE_NAME," & _
        "r.RETURN_ID,i.RETURN_STATE,i.SOURCE_ID,'',r.NOTE_TEXT,i.SOURCE_LOCATION," & _
        "COALESCE(r.LOT_ID,i.LOT_ID,''),r.RETURN_QTY,COALESCE(r.ENTRY_UNIT,i.UNIT_NAME)," & _
        "COALESCE(r.BASE_QTY,r.RETURN_QTY),COALESCE(r.BASE_UNIT,i.UNIT_NAME) " & _
        "FROM WMS_ISSUE_RETURNS r JOIN WMS_ISSUES i ON i.SOURCE_ID=r.SOURCE_ID" & _
        whereSQL & " ORDER BY r.RETURN_DATE DESC,r.CREATED_AT DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    WMSDBS_QueryReturnsCon=WMSDBS_WriteRSLot(oSh,rowOut,oRS,"Возврат")
End Function

Function WMSDBS_QueryOrders(oDoc As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String,whereSQL As String,n As Long
    WMSDBS_QueryOrders=0:sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ORDER_LINES",sErr) Then WMSDB_Close:Exit Function

    whereSQL=WMSDBS_CommonWhere(oSh,"RECEIPT_DATE", _
        "PRODUCT_NAME","PRODUCT_CODE","SUPPLIER_NAME","DOC_NO")
    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " RECEIPT_DATE,PRODUCT_NAME,PRODUCT_CODE,FACT_QTY,UNIT_NAME,SUPPLIER_NAME,DOC_NO," & _
        "STATUS_NAME,SOURCE_ID,SUPPLIER_ARTICLE,INVOICE_NO,LOCATION_NAME " & _
        "FROM WMS_ORDER_LINES" & whereSQL & " ORDER BY RECEIPT_DATE DESC,CREATED_AT DESC"

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    n=WMSDBS_WriteRS(oSh,rowOut,oRS,"Заказ")
    WMSDB_Close
    WMSDBS_QueryOrders=n
End Function

Function WMSDBS_QueryIssues(oDoc As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String,whereSQL As String,n As Long
    WMSDBS_QueryIssues=0:sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ISSUES",sErr) Then WMSDB_Close:Exit Function

    whereSQL=WMSDBS_CommonWhere(oSh,"ISSUE_DATE", _
        "PRODUCT_NAME","PRODUCT_CODE","EMPLOYEE_NAME","SOURCE_ID")
    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " ISSUE_DATE,PRODUCT_NAME,PRODUCT_CODE,ISSUE_QTY,UNIT_NAME,EMPLOYEE_NAME,SOURCE_ID," & _
        "RETURN_STATE,SOURCE_ID,'',RETURNABLE,SOURCE_LOCATION " & _
        "FROM WMS_ISSUES" & whereSQL & " ORDER BY ISSUE_DATE DESC,CREATED_AT DESC"

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    n=WMSDBS_WriteRS(oSh,rowOut,oRS,"Выдача")
    WMSDB_Close
    WMSDBS_QueryIssues=n
End Function

Function WMSDBS_QueryReturns(oDoc As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String,whereSQL As String,n As Long
    WMSDBS_QueryReturns=0:sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ISSUE_RETURNS",sErr) Then WMSDB_Close:Exit Function

    whereSQL=WMSDBS_ReturnWhere(oSh)
    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " r.RETURN_DATE,i.PRODUCT_NAME,i.PRODUCT_CODE,r.RETURN_QTY,i.UNIT_NAME,i.EMPLOYEE_NAME," & _
        "r.RETURN_ID,i.RETURN_STATE,i.SOURCE_ID,'',r.NOTE_TEXT,i.SOURCE_LOCATION " & _
        "FROM WMS_ISSUE_RETURNS r JOIN WMS_ISSUES i ON i.SOURCE_ID=r.SOURCE_ID" & _
        whereSQL & " ORDER BY r.RETURN_DATE DESC,r.CREATED_AT DESC"

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    n=WMSDBS_WriteRS(oSh,rowOut,oRS,"Возврат")
    WMSDB_Close
    WMSDBS_QueryReturns=n
End Function

Function WMSDBS_QueryOpenReturns(oDoc As Object,oSh As Object,rowOut As Long,ByRef sErr As String) As Long
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String,n As Long
    WMSDBS_QueryOpenReturns=0:sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    sql="SELECT FIRST " & CStr(WMSDBS_MAX_ROWS) & _
        " ISSUE_DATE,PRODUCT_NAME,PRODUCT_CODE,(ISSUE_QTY-COALESCE(RETURNED_QTY,0)),UNIT_NAME," & _
        "EMPLOYEE_NAME,SOURCE_ID,RETURN_STATE,SOURCE_ID,'',RETURNABLE,SOURCE_LOCATION " & _
        "FROM WMS_ISSUES WHERE " & _
        "UPPER(TRIM(COALESCE(RETURNABLE,''))) NOT IN ('','НЕТ','NO','FALSE','0') " & _
        "AND COALESCE(RETURNED_QTY,0)<ISSUE_QTY " & _
        "ORDER BY ISSUE_DATE,CREATED_AT"

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    n=WMSDBS_WriteRS(oSh,rowOut,oRS,"Открытый возврат")
    WMSDB_Close
    WMSDBS_QueryOpenReturns=n
End Function

Function WMSDBS_WriteRSLot(oSh As Object,rowOut As Long,oRS As Object,typeName As String) As Long
    Dim r As Long,n As Long
    r=rowOut:n=0
    Do While oRS.next()
        oSh.getCellByPosition(0,r).String=typeName
        oSh.getCellByPosition(1,r).String=oRS.getString(1)
        oSh.getCellByPosition(2,r).String=oRS.getString(2)
        oSh.getCellByPosition(3,r).String=oRS.getString(3)
        oSh.getCellByPosition(4,r).Value=oRS.getDouble(4)
        oSh.getCellByPosition(5,r).String=oRS.getString(5)
        oSh.getCellByPosition(6,r).String=oRS.getString(6)
        oSh.getCellByPosition(7,r).String=oRS.getString(7)
        oSh.getCellByPosition(8,r).String=oRS.getString(8)
        oSh.getCellByPosition(9,r).String=oRS.getString(9)
        oSh.getCellByPosition(10,r).String=oRS.getString(10)
        oSh.getCellByPosition(11,r).String=oRS.getString(11)
        oSh.getCellByPosition(12,r).String=oRS.getString(12)
        oSh.getCellByPosition(13,r).String=oRS.getString(13)
        oSh.getCellByPosition(14,r).Value=oRS.getDouble(14)
        oSh.getCellByPosition(15,r).String=oRS.getString(15)
        oSh.getCellByPosition(16,r).Value=oRS.getDouble(16)
        oSh.getCellByPosition(17,r).String=oRS.getString(17)
        r=r+1:n=n+1
    Loop
    WMSDBS_WriteRSLot=n
End Function

Function WMSDBS_WriteRS(oSh As Object,rowOut As Long,oRS As Object,typeName As String) As Long
    Dim r As Long,n As Long
    r=rowOut:n=0
    Do While oRS.next()
        oSh.getCellByPosition(0,r).String=typeName
        oSh.getCellByPosition(1,r).String=oRS.getString(1)
        oSh.getCellByPosition(2,r).String=oRS.getString(2)
        oSh.getCellByPosition(3,r).String=oRS.getString(3)
        oSh.getCellByPosition(4,r).Value=oRS.getDouble(4)
        oSh.getCellByPosition(5,r).String=oRS.getString(5)
        oSh.getCellByPosition(6,r).String=oRS.getString(6)
        oSh.getCellByPosition(7,r).String=oRS.getString(7)
        oSh.getCellByPosition(8,r).String=oRS.getString(8)
        oSh.getCellByPosition(9,r).String=oRS.getString(9)
        oSh.getCellByPosition(10,r).String=oRS.getString(10)
        oSh.getCellByPosition(11,r).String=oRS.getString(11)
        oSh.getCellByPosition(12,r).String=oRS.getString(12)
        r=r+1:n=n+1
    Loop
    WMSDBS_WriteRS=n
End Function

Function WMSDBS_ColumnExists(oCon As Object,t As String,c As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBS_ColumnExists=False
    On Error GoTo Done
    sql="SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE TRIM(RDB$RELATION_NAME)=" & WMSDBS_Q(UCase(t)) & _
        " AND TRIM(RDB$FIELD_NAME)=" & WMSDBS_Q(UCase(c))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBS_ColumnExists=(oRS.getInt(1)>0)
Done:
End Function

Function WMSDBS_CommonWhere(oSh As Object,dateField As String,productField As String,codeField As String,partyField As String,docField As String) As String
    Dim a As String,q As String,p As String,d As String,s As String
    a=" WHERE 1=1"
    q=Trim(oSh.getCellByPosition(1,3).String)
    p=Trim(oSh.getCellByPosition(1,6).String)
    s=Trim(oSh.getCellByPosition(1,7).String)
    d=Trim(oSh.getCellByPosition(1,8).String)

    If q<>"" Then a=a & " AND (" & productField & " CONTAINING " & WMSDBS_Q(q) & _
        " OR " & codeField & " CONTAINING " & WMSDBS_Q(q) & _
        " OR " & partyField & " CONTAINING " & WMSDBS_Q(q) & _
        " OR " & docField & " CONTAINING " & WMSDBS_Q(q) & ")"
    If p<>"" Then a=a & " AND (" & productField & " CONTAINING " & WMSDBS_Q(p) & _
        " OR " & codeField & " CONTAINING " & WMSDBS_Q(p) & ")"
    If s<>"" Then a=a & " AND " & partyField & " CONTAINING " & WMSDBS_Q(s)
    If d<>"" Then a=a & " AND " & docField & " CONTAINING " & WMSDBS_Q(d)

    a=a & WMSDBS_DateWhere(oSh,dateField)
    WMSDBS_CommonWhere=a
End Function

Function WMSDBS_ReturnWhere(oSh As Object) As String
    Dim a As String,q As String,p As String,s As String,d As String
    a=" WHERE 1=1"
    q=Trim(oSh.getCellByPosition(1,3).String)
    p=Trim(oSh.getCellByPosition(1,6).String)
    s=Trim(oSh.getCellByPosition(1,7).String)
    d=Trim(oSh.getCellByPosition(1,8).String)

    If q<>"" Then a=a & " AND (i.PRODUCT_NAME CONTAINING " & WMSDBS_Q(q) & _
        " OR i.PRODUCT_CODE CONTAINING " & WMSDBS_Q(q) & _
        " OR i.EMPLOYEE_NAME CONTAINING " & WMSDBS_Q(q) & _
        " OR r.RETURN_ID CONTAINING " & WMSDBS_Q(q) & ")"
    If p<>"" Then a=a & " AND (i.PRODUCT_NAME CONTAINING " & WMSDBS_Q(p) & _
        " OR i.PRODUCT_CODE CONTAINING " & WMSDBS_Q(p) & ")"
    If s<>"" Then a=a & " AND i.EMPLOYEE_NAME CONTAINING " & WMSDBS_Q(s)
    If d<>"" Then a=a & " AND r.RETURN_ID CONTAINING " & WMSDBS_Q(d)

    a=a & WMSDBS_DateWhere(oSh,"r.RETURN_DATE")
    WMSDBS_ReturnWhere=a
End Function

Function WMSDBS_DateWhere(oSh As Object,dateField As String) As String
    Dim s As String,v1 As Double,v2 As Double
    s="":v1=oSh.getCellByPosition(1,4).Value:v2=oSh.getCellByPosition(1,5).Value
    If v1>0 Then s=s & " AND " & dateField & ">=" & WMSDBS_DateSQL(v1)
    If v2>0 Then s=s & " AND " & dateField & "<=" & WMSDBS_DateSQL(v2)
    WMSDBS_DateWhere=s
End Function

Function WMSDBS_CanonType(s As String) As String
    Dim t As String
    t=UCase(Trim(s))
    If t="" Or t="ВСЕ" Or t="ALL" Then WMSDBS_CanonType="ALL":Exit Function
    If InStr(t,"ЗАКАЗ")>0 Then WMSDBS_CanonType="ORDERS":Exit Function
    If InStr(t,"ВЫДА")>0 Then WMSDBS_CanonType="ISSUES":Exit Function
    If InStr(t,"ВОЗВР")>0 Then WMSDBS_CanonType="RETURNS":Exit Function
    WMSDBS_CanonType="ALL"
End Function

Sub WMSDBS_BuildLayout(oDoc As Object,oSh As Object)
    Dim headers As Variant,i As Long
    oSh.getCellRangeByPosition(0,0,17,2000).clearContents(1023)

    oSh.getCellByPosition(0,0).String="WMS — БАЗА / ПОИСК"
    oSh.getCellByPosition(0,2).String="Тип"
    oSh.getCellByPosition(1,2).String="Все"
    oSh.getCellByPosition(0,3).String="Общий поиск"
    oSh.getCellByPosition(0,4).String="Дата с"
    oSh.getCellByPosition(0,5).String="Дата по"
    oSh.getCellByPosition(0,6).String="Товар / код / артикул"
    oSh.getCellByPosition(0,7).String="Поставщик / сотрудник"
    oSh.getCellByPosition(0,8).String="УПД / заказ / ReceiptID"
    oSh.getCellByPosition(0,9).String="Готово к поиску."

    headers=Array("Тип","Дата","Товар","Код товара","Кол-во","Ед.","Поставщик / получил", _
                  "Документ / событие","Статус","SourceID","Артикул / доп.","Комментарий","Место", _
                  "LOT","Док. кол-во","Док. ед.","Base кол-во","Base ед.")
    For i=0 To UBound(headers)
        oSh.getCellByPosition(i,10).String=headers(i)
    Next i

    oSh.Columns.getByIndex(0).Width=2600
    oSh.Columns.getByIndex(1).Width=2700
    oSh.Columns.getByIndex(2).Width=6500
    oSh.Columns.getByIndex(3).Width=3600
    oSh.Columns.getByIndex(4).Width=2200
    oSh.Columns.getByIndex(5).Width=1800
    oSh.Columns.getByIndex(6).Width=5200
    oSh.Columns.getByIndex(7).Width=4400
    oSh.Columns.getByIndex(8).Width=3000
    oSh.Columns.getByIndex(9).Width=5600
    oSh.Columns.getByIndex(10).Width=4200
    oSh.Columns.getByIndex(11).Width=5200
    oSh.Columns.getByIndex(12).Width=3600
    oSh.Columns.getByIndex(13).Width=4200
    oSh.Columns.getByIndex(14).Width=2600
    oSh.Columns.getByIndex(15).Width=2200
    oSh.Columns.getByIndex(16).Width=2600
    oSh.Columns.getByIndex(17).Width=2200

    oSh.getCellRangeByPosition(0,0,17,0).CharWeight=150
    oSh.getCellRangeByPosition(0,10,17,10).CharWeight=150
    oSh.getCellRangeByPosition(0,2,0,8).CharWeight=150

    WMSDBS_FormatDateCell oDoc,oSh.getCellByPosition(1,4)
    WMSDBS_FormatDateCell oDoc,oSh.getCellByPosition(1,5)

    oSh.Rows.getByIndex(0).Height=850
    oSh.Rows.getByIndex(10).Height=700
    On Error Resume Next
    oSh.getCellRangeByPosition(0,10,17,10).IsCellBackgroundTransparent=False
    oSh.getCellRangeByPosition(0,10,17,10).CellBackColor=12632256
    oSh.getCellRangeByPosition(0,0,17,0).IsCellBackgroundTransparent=False
    oSh.getCellRangeByPosition(0,0,17,0).CellBackColor=14737632
    On Error GoTo 0
End Sub

Sub WMSDBS_ClearResults(oSh As Object)
    Dim cur As Object,lastRow As Long
    On Error GoTo Fallback

    cur=oSh.createCursor()
    cur.gotoEndOfUsedArea(True)
    lastRow=cur.RangeAddress.EndRow

    If lastRow>=11 Then
        oSh.getCellRangeByPosition(0,11,17,lastRow).clearContents(1023)
    End If
    Exit Sub

Fallback:
    ' Safe fallback is intentionally modest; never repaint/clear thousands
    ' of empty rows on every search click.
    On Error Resume Next
    oSh.getCellRangeByPosition(0,11,17,200).clearContents(1023)
End Sub

Function WMSDBS_GetSheet(oDoc As Object,ByRef oSh As Object) As Boolean
    WMSDBS_GetSheet=False
    If Not oDoc.Sheets.hasByName(WMSDBS_SHEET) Then
        MsgBox "Лист 'База - Поиск' не установлен. Запустите WMSDBS_Install.",48,"WMS — База / Поиск"
        Exit Function
    End If
    oSh=oDoc.Sheets.getByName(WMSDBS_SHEET)
    WMSDBS_GetSheet=True
End Function

Function WMSDBS_InstallButtons(oDoc As Object,oSh As Object,ByRef sErr As String) As Boolean
    Dim oDP As Object,oForms As Object,oForm As Object
    Dim x As Long,y As Long,w As Long,h As Long,g As Long
    WMSDBS_InstallButtons=False:sErr=""
    On Error GoTo EH

    WMSDBS_RemoveButtons oSh
    oDP=oSh.DrawPage:oForms=oDP.Forms
    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name=WMSDBS_FORM:oForms.insertByName(WMSDBS_FORM,oForm)

    x=oSh.getCellByPosition(3,2).Position.X
    y=oSh.getCellByPosition(3,2).Position.Y
    w=3900:h=650:g=120

    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_FIND","Найти",x,y,w,h,"WMSDBS_Search"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_TODAY","Сегодня",x,y+(h+g),w,h,"WMSDBS_Today"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_7D","Последние 7 дней",x,y+2*(h+g),w,h,"WMSDBS_Last7Days"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_OPENRET","Открытые возвраты",x,y+3*(h+g),w,h,"WMSDBS_OpenReturns"

    x=x+w+250
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_PROD","История товара",x,y,w,h,"WMSDBS_ProductHistory"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_CLEAR","Очистить",x,y+(h+g),w,h,"WMSDBS_Clear"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_REFRESH","Обновить",x,y+2*(h+g),w,h,"WMSDBS_Refresh"
    WMSDBS_AddButton oDoc,oSh,oForm,"WMS_DBS_SELF","SelfCheck",x,y+3*(h+g),w,h,"WMSDBS_SelfCheck"

    WMSDBS_InstallButtons=True
    Exit Function
EH:
    sErr=CStr(Err) & " " & Error$
End Function

Sub WMSDBS_RemoveButtons(oSh As Object)
    Dim oDP As Object,oForms As Object,i As Long,oShape As Object,oCtl As Object,nm As String
    On Error Resume Next
    oDP=oSh.DrawPage
    For i=oDP.Count-1 To 0 Step -1
        oShape=oDP.getByIndex(i)
        nm="":oCtl=oShape.Control:nm=oCtl.Name
        If Left(nm,8)="WMS_DBS_" Then oDP.remove(oShape)
    Next i
    oForms=oDP.Forms
    If oForms.hasByName(WMSDBS_FORM) Then oForms.removeByName(WMSDBS_FORM)
    On Error GoTo 0
End Sub

Sub WMSDBS_AddButton(oDoc As Object,oSh As Object,oForm As Object,ctlName As String,labelText As String,x As Long,y As Long,w As Long,h As Long,macroName As String)
    Dim oCtl As Object,oShape As Object,ev As New com.sun.star.script.ScriptEventDescriptor
    Dim p As New com.sun.star.awt.Point,z As New com.sun.star.awt.Size,idx As Long
    oCtl=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    oCtl.Name=ctlName:oCtl.Label=labelText:oCtl.Tabstop=False
    oForm.insertByName(ctlName,oCtl):idx=oForm.Count-1

    oShape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    p.X=x:p.Y=y:z.Width=w:z.Height=h
    oShape.Position=p:oShape.Size=z:oShape.Control=oCtl
    oSh.DrawPage.add(oShape)

    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.AddListenerParam=""
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_09_DB_Search." & macroName & "?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
End Sub

Function WMSDBS_TableState(oCon As Object,t As String) As String
    Dim sErr As String
    If WMSDB_TableExistsSimple(oCon,t,sErr) Then
        WMSDBS_TableState="OK"
    ElseIf sErr<>"" Then
        WMSDBS_TableState="Ошибка: " & sErr
    Else
        WMSDBS_TableState="нет"
    End If
End Function

Function WMSDBS_Q(s As String) As String
    WMSDBS_Q="'" & Replace(s,"'","''") & "'"
End Function

Function WMSDBS_DateSQL(v As Double) As String
    Dim d As Date
    d=CDate(v)
    WMSDBS_DateSQL="DATE '" & Right("0000" & CStr(Year(d)),4) & "-" & _
                    Right("00" & CStr(Month(d)),2) & "-" & _
                    Right("00" & CStr(Day(d)),2) & "'"
End Function

Sub WMSDBS_FormatDateCell(oDoc As Object,oCell As Object)
    Dim nf As Object,k As Long,loc As New com.sun.star.lang.Locale
    On Error Resume Next
    nf=oDoc.NumberFormats
    loc.Language="ru":loc.Country="RU"
    k=nf.queryKey("DD.MM.YYYY",loc,True)
    If k=-1 Then k=nf.addNew("DD.MM.YYYY",loc)
    oCell.NumberFormat=k
    On Error GoTo 0
End Sub

