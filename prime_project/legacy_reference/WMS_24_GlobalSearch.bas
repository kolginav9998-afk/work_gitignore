Option Explicit

Global Const WMSSEARCH2_VERSION = "3.1.0-BOUNDED-GUARDED-SEARCH"
Global gWMSSEARCH2_Busy As Boolean
Global Const WMSSEARCH2_SHEET = "База - Поиск"
Global Const WMSSEARCH2_FORM = "WMS_SEARCH2_FORM"

Sub WMSSEARCH2_Install()
    Dim oDoc As Object,oSh As Object,h As Variant,i As Long
    oDoc=ThisComponent
    If oDoc.Sheets.hasByName(WMSSEARCH2_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSSEARCH2_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSSEARCH2_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSSEARCH2_SHEET)
    End If
    oSh.IsVisible=True

    oSh.getCellRangeByPosition(0,0,12,5).clearContents(1023)
    oSh.getCellByPosition(0,0).String="БАЗА — ПОИСК"
    oSh.getCellByPosition(0,1).String="Введите запрос в B3. Ищется по товарам, артикулам, движениям, заказам, выдачам и актам."
    oSh.getCellByPosition(0,2).String="Поиск:"
    oSh.getCellByPosition(1,2).String=""

    oSh.getCellRangeByPosition(0,0,9,0).CellBackColor=2500134
    oSh.getCellRangeByPosition(0,0,9,0).CharColor=16777215:oSh.getCellRangeByPosition(0,0,9,0).CharWeight=150
    oSh.getCellRangeByPosition(0,1,9,1).CellBackColor=15132390

    h=Array("Тип","Дата","ID","Код товара","Артикул","Наименование","Кол-во","Ед.","Место / контрагент","Подробности")
    For i=0 To UBound(h):oSh.getCellByPosition(i,5).String=CStr(h(i)):Next i
    oSh.getCellRangeByPosition(0,5,9,5).CellBackColor=4473924:oSh.getCellRangeByPosition(0,5,9,5).CharColor=16777215:oSh.getCellRangeByPosition(0,5,9,5).CharWeight=150
    oSh.Columns.getByIndex(0).Width=3400:oSh.Columns.getByIndex(1).Width=2800:oSh.Columns.getByIndex(2).Width=4600:oSh.Columns.getByIndex(3).Width=3500
    oSh.Columns.getByIndex(4).Width=3600:oSh.Columns.getByIndex(5).Width=6500:oSh.Columns.getByIndex(8).Width=5200:oSh.Columns.getByIndex(9).Width=8500

    WMSSEARCH2_InstallButtons oDoc,oSh
End Sub

Sub WMSSEARCH2_Run()
    Dim oSh As Object,q As String
    oSh=ThisComponent.Sheets.getByName(WMSSEARCH2_SHEET):q=Trim(oSh.getCellByPosition(1,2).String)
    WMSSEARCH2_Search q
End Sub

Sub WMSSEARCH2_Today()
    WMSSEARCH2_Search ""
End Sub

Sub WMSSEARCH2_Clear()
    Dim oSh As Object,c As Object,lastR As Long
    oSh=ThisComponent.Sheets.getByName(WMSSEARCH2_SHEET)
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow:If lastR<6 Then lastR=6
    oSh.getCellRangeByPosition(0,6,9,lastR).clearContents(1023)
    oSh.getCellByPosition(1,2).String=""
End Sub

Sub WMSSEARCH2_Search(q As String)
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String,r As Long,t0 As Double
    If Not WMSDBX_TryEnter("SEARCH") Then MsgBox "Другая операция WMS ещё выполняется. Дождитесь завершения.",48,"WMS — Поиск":Exit Sub
    If gWMSSEARCH2_Busy Then WMSDBX_Leave:Exit Sub
    gWMSSEARCH2_Busy=True
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSSEARCH2_SHEET)
    On Error GoTo EH
    t0=Timer:oDoc.lockControllers()
    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then GoTo Fail
    oSh.getCellRangeByPosition(0,6,9,305).clearContents(1023)
    r=6
    WMSSEARCH2_AddMovements oCon,oSh,q,r,sErr:If sErr<>"" Then GoTo Fail
    If r<206 Then WMSSEARCH2_AddOrders oCon,oSh,q,r,sErr:If sErr<>"" Then GoTo Fail
    If r<206 Then WMSSEARCH2_AddIssues oCon,oSh,q,r,sErr:If sErr<>"" Then GoTo Fail
    If r<206 Then WMSSEARCH2_AddActs oCon,oSh,q,r,sErr:If sErr<>"" Then GoTo Fail
    WMSDB_Close
    oSh.getCellByPosition(0,3).String="Найдено: " & CStr(r-6) & " | максимум 200 | " & CStr(CLng((Timer-t0)*1000)) & " мс"
Done:
    On Error Resume Next:oDoc.unlockControllers():On Error GoTo 0
    gWMSSEARCH2_Busy=False:WMSDBX_Leave
    Exit Sub
Fail:
    On Error Resume Next:WMSDB_Close:oDoc.unlockControllers():On Error GoTo 0
    gWMSSEARCH2_Busy=False:WMSDBX_Leave
    MsgBox sErr,16,"WMS — Поиск":Exit Sub
EH:
    sErr="Глобальный поиск: " & CStr(Err) & " " & Error$:Resume Fail
End Sub

Sub WMSSEARCH2_AddMovements(oCon As Object,oSh As Object,q As String,ByRef r As Long,ByRef sErr As String)
    Dim oStmt As Object,oRS As Object,sql As String,w As String
    sErr=""
    On Error GoTo EH
    w=WMSSEARCH2_Where(q,Array("m.PRODUCT_CODE","m.PRODUCT_NAME","COALESCE(p.SUPPLIER_ARTICLE,'')","m.LOCATION_NAME","COALESCE(m.NOTE_TEXT,'')","m.SOURCE_ID","COALESCE(m.LOT_ID,'')"))
    sql="SELECT FIRST 120 m.MOVEMENT_TYPE,m.MOVEMENT_DATE,m.MOVEMENT_ID,m.PRODUCT_CODE,COALESCE(p.SUPPLIER_ARTICLE,''),m.PRODUCT_NAME,m.QTY,m.UNIT_NAME,m.LOCATION_NAME," & _
        "COALESCE(m.FLOW_CHANNEL,'') || ' | LOT: ' || COALESCE(m.LOT_ID,'') || ' | ' || COALESCE(m.NOTE_TEXT,'') " & _
        "FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=m.PRODUCT_CODE" & w & " ORDER BY m.MOVEMENT_DATE DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        WMSSEARCH2_Write oSh,r,"Движение " & oRS.getString(1),oRS.getString(2),oRS.getString(3),oRS.getString(4),oRS.getString(5),oRS.getString(6),oRS.getDouble(7),oRS.getString(8),oRS.getString(9),oRS.getString(10)
        r=r+1:If r>205 Then Exit Do
    Loop
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo EH
    Exit Sub
EH:
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo 0
    sErr="Поиск движений: " & CStr(Err) & " " & Error$
End Sub

Sub WMSSEARCH2_AddOrders(oCon As Object,oSh As Object,q As String,ByRef r As Long,ByRef sErr As String)
    Dim oStmt As Object,oRS As Object,sql As String,w As String
    sErr="":On Error GoTo EH
    w=WMSSEARCH2_Where(q,Array("COALESCE(l.PRODUCT_CODE,'')","l.PRODUCT_NAME","COALESCE(l.SUPPLIER_ARTICLE,'')","COALESCE(l.SUPPLIER_NAME,'')","COALESCE(l.SELLER_NAME,'')","COALESCE(l.UPD_NO,'')","COALESCE(l.EXTERNAL_ORDER_NO,'')"))
    sql="SELECT FIRST 80 COALESCE(l.RECEIPT_DATE,l.ORDER_DATE),l.SOURCE_ID,COALESCE(l.PRODUCT_CODE,''),COALESCE(l.SUPPLIER_ARTICLE,''),l.PRODUCT_NAME,COALESCE(l.FACT_QTY,l.ORDER_QTY,0),COALESCE(l.UNIT_NAME,''),COALESCE(l.SUPPLIER_NAME,'')," & _
        "'Заказ: ' || COALESCE(l.EXTERNAL_ORDER_NO,'') || ' | УПД: ' || COALESCE(l.UPD_NO,'') || ' | Статус: ' || COALESCE(l.STATUS_NAME,'') " & _
        "FROM WMS_ORDER_LINES l" & w & " ORDER BY COALESCE(l.RECEIPT_DATE,l.ORDER_DATE) DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        WMSSEARCH2_Write oSh,r,"Заказ",oRS.getString(1),oRS.getString(2),oRS.getString(3),oRS.getString(4),oRS.getString(5),oRS.getDouble(6),oRS.getString(7),oRS.getString(8),oRS.getString(9)
        r=r+1:If r>205 Then Exit Do
    Loop
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo EH
    Exit Sub
EH:
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo 0
    sErr="Поиск заказов: " & CStr(Err) & " " & Error$
End Sub

Sub WMSSEARCH2_AddIssues(oCon As Object,oSh As Object,q As String,ByRef r As Long,ByRef sErr As String)
    Dim oStmt As Object,oRS As Object,sql As String,w As String
    sErr="":On Error GoTo EH
    w=WMSSEARCH2_Where(q,Array("COALESCE(i.PRODUCT_CODE,'')","i.PRODUCT_NAME","i.EMPLOYEE_NAME","COALESCE(i.SOURCE_LOCATION,'')","COALESCE(i.NOTE_TEXT,'')","i.SOURCE_ID"))
    sql="SELECT FIRST 80 i.ISSUE_DATE,i.SOURCE_ID,COALESCE(i.PRODUCT_CODE,''),COALESCE(p.SUPPLIER_ARTICLE,''),i.PRODUCT_NAME,i.ISSUE_QTY,i.UNIT_NAME,i.EMPLOYEE_NAME," & _
        "'Откуда: ' || COALESCE(i.SOURCE_LOCATION,'') || ' | LOT: ' || COALESCE(i.LOT_ID,'') || ' | Возвратный: ' || COALESCE(i.RETURNABLE,'') " & _
        "FROM WMS_ISSUES i LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=i.PRODUCT_CODE" & w & " ORDER BY i.ISSUE_DATE DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        WMSSEARCH2_Write oSh,r,"Выдача",oRS.getString(1),oRS.getString(2),oRS.getString(3),oRS.getString(4),oRS.getString(5),oRS.getDouble(6),oRS.getString(7),oRS.getString(8),oRS.getString(9)
        r=r+1:If r>205 Then Exit Do
    Loop
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo EH
    Exit Sub
EH:
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo 0
    sErr="Поиск выдач: " & CStr(Err) & " " & Error$
End Sub

Sub WMSSEARCH2_AddActs(oCon As Object,oSh As Object,q As String,ByRef r As Long,ByRef sErr As String)
    Dim oStmt As Object,oRS As Object,sql As String,w As String
    sErr="":On Error GoTo EH
    w=WMSSEARCH2_Where(q,Array("a.ACT_NO","a.ACT_TYPE","COALESCE(a.PARTY_FROM,'')","COALESCE(a.PARTY_TO,'')","COALESCE(l.PRODUCT_NAME,'')","COALESCE(l.ARTICLE_CODE,'')"))
    sql="SELECT FIRST 60 a.ACT_DATE,a.ACT_ID,COALESCE(l.PRODUCT_CODE,''),COALESCE(l.ARTICLE_CODE,''),COALESCE(l.PRODUCT_NAME,a.DOC_TITLE,''),COALESCE(l.QTY,0),COALESCE(l.UNIT_NAME,''),COALESCE(a.PARTY_TO,a.PARTY_FROM,'')," & _
        "'Акт: ' || a.ACT_NO || ' | ' || a.ACT_TYPE || ' | Статус: ' || a.STATUS_NAME " & _
        "FROM WMS_ACTS a LEFT JOIN WMS_ACT_LINES l ON l.ACT_ID=a.ACT_ID" & w & " ORDER BY a.ACT_DATE DESC"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        WMSSEARCH2_Write oSh,r,"Акт",oRS.getString(1),oRS.getString(2),oRS.getString(3),oRS.getString(4),oRS.getString(5),oRS.getDouble(6),oRS.getString(7),oRS.getString(8),oRS.getString(9)
        r=r+1:If r>205 Then Exit Do
    Loop
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo EH
    Exit Sub
EH:
    On Error Resume Next:oRS.close():oStmt.close():On Error GoTo 0
    sErr="Поиск актов: " & CStr(Err) & " " & Error$
End Sub

Function WMSSEARCH2_Where(q As String,fields As Variant) As String
    Dim i As Long,s As String,term As String
    If Trim(q)="" Then WMSSEARCH2_Where="":Exit Function
    term=Replace(UCase(Trim(q)),"'","''")
    s=" WHERE "
    For i=LBound(fields) To UBound(fields)
        If i>LBound(fields) Then s=s & " OR "
        s=s & "UPPER(" & CStr(fields(i)) & ") CONTAINING '" & term & "'"
    Next i
    WMSSEARCH2_Where=s
End Function

Sub WMSSEARCH2_Write(oSh As Object,r As Long,typ As String,dt As String,id As String,code As String,art As String,nm As String,qty As Double,unitName As String,place As String,detail As String)
    oSh.getCellByPosition(0,r).String=typ:oSh.getCellByPosition(1,r).String=dt:oSh.getCellByPosition(2,r).String=id:oSh.getCellByPosition(3,r).String=code
    oSh.getCellByPosition(4,r).String=art:oSh.getCellByPosition(5,r).String=nm:oSh.getCellByPosition(6,r).Value=qty:oSh.getCellByPosition(7,r).String=unitName
    oSh.getCellByPosition(8,r).String=place:oSh.getCellByPosition(9,r).String=detail
End Sub

Sub WMSSEARCH2_InstallButtons(oDoc As Object,oSh As Object)
    Dim forms As Object,form As Object
    WMSSEARCH2_RemoveButtons oSh
    forms=oSh.DrawPage.Forms
    On Error Resume Next:If forms.hasByName(WMSSEARCH2_FORM) Then forms.removeByName(WMSSEARCH2_FORM):On Error GoTo EH
    form=oDoc.createInstance("com.sun.star.form.component.Form"):form.Name=WMSSEARCH2_FORM:forms.insertByName(WMSSEARCH2_FORM,form)
    WMSSEARCH2_AddButton oDoc,oSh,form,"WMS_S2_FIND","Найти",0,3,2800,800,"WMSSEARCH2_Run"
    WMSSEARCH2_AddButton oDoc,oSh,form,"WMS_S2_ALL","Показать всё",1,3,3200,800,"WMSSEARCH2_Today"
    WMSSEARCH2_AddButton oDoc,oSh,form,"WMS_S2_CLEAR","Очистить",2,3,2800,800,"WMSSEARCH2_Clear"
    Exit Sub
EH:
    MsgBox "Кнопки поиска: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSSEARCH2_AddButton(oDoc As Object,oSh As Object,oForm As Object,sName As String,sLabel As String,nCol As Long,nRow As Long,nW As Long,nH As Long,sMacro As String)
    Dim model As Object,shape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    model=oDoc.createInstance("com.sun.star.form.component.CommandButton"):model.Name=sName:model.Label=sLabel:oForm.insertByName(sName,model):idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=model:a=oSh.getCellByPosition(nCol,nRow).Position:shape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=nW:sz.Height=nH:shape.Size=sz:oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed":ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_24_GlobalSearch." & sMacro & "?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSSEARCH2_RemoveButtons(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i):nm="":ctl=shape.Control:nm=ctl.Name
        If Left(nm,7)="WMS_S2_" Or Left(nm,8)="WMS_DBS_" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub
