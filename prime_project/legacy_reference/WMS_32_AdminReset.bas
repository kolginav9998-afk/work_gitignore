Option Explicit

Global Const WMSCLEAN_VERSION = "3.3.0-GENERAL-CLEANUP-CENTER"
Global Const WMSCLEAN_PASSWORD = "555"
Global Const WMSCLEAN_FORM = "WMS_CLEAN_FORM"

Sub WMSRESET_Install()
    WMSCLEAN_Install
End Sub

Sub WMSCLEAN_Install()
    Dim oDoc As Object,oSh As Object,forms As Object,form As Object
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName("Инфо") Then Exit Sub
    oSh=oDoc.Sheets.getByName("Инфо")
    On Error GoTo EH
    WMSCLEAN_RemoveButtons oSh
    oSh.getCellByPosition(0,22).String="ЦЕНТР БЕЗОПАСНОЙ ОЧИСТКИ"
    oSh.getCellByPosition(0,23).String="Только для исправления ошибочно внесённых данных. Каждая очистка: backup → пароль 555 → предварительный просмотр → текстовое подтверждение → контроль после операции."
    oSh.getCellRangeByPosition(0,22,10,22).CharWeight=150
    oSh.getCellRangeByPosition(0,23,10,23).IsTextWrapped=True
    forms=oSh.DrawPage.Forms
    On Error Resume Next
    If forms.hasByName(WMSCLEAN_FORM) Then forms.removeByName(WMSCLEAN_FORM)
    On Error GoTo EH
    form=oDoc.createInstance("com.sun.star.form.component.Form")
    form.Name=WMSCLEAN_FORM
    forms.insertByName(WMSCLEAN_FORM,form)
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_PRODUCT","Очистить товар",0,24,3000,"WMSCLEAN_Product"
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_ORDER","Очистить заказ",2,24,3000,"WMSCLEAN_Order"
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_DOC","Очистить документ / приход",4,24,3900,"WMSCLEAN_Document"
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_LOT","Очистить партию",7,24,3000,"WMSCLEAN_Lot"
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_OPER","Очистить все операции",0,26,3500,"WMSCLEAN_AllOperations"
    WMSCLEAN_AddButton oDoc,oSh,form,"WMS_CLEAN_ALL","ПОЛНАЯ ЧИСТКА WMS",3,26,4200,"WMSCLEAN_FullReset"
    Exit Sub
EH:
    MsgBox "Центр очистки: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSCLEAN_AddButton(oDoc As Object,oSh As Object,oForm As Object,nm As String,caption As String,col As Long,row As Long,w As Long,macroName As String)
    Dim model As Object,shape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    model=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    model.Name=nm
    model.Label=caption
    model.Tabstop=False
    oForm.insertByName(nm,model)
    idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    shape.Control=model
    a=oSh.getCellByPosition(col,row).Position
    shape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size")
    sz.Width=w
    sz.Height=850
    shape.Size=sz
    oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_32_AdminReset." & macroName & "?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSCLEAN_RemoveButtons(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i)
        nm=""
        ctl=shape.Control
        nm=ctl.Name
        If Left(nm,10)="WMS_CLEAN_" Or nm="WMS_RESET_LAST2" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub

Function WMSCLEAN_Authorize(titleText As String,confirmWord As String,preview As String,ByRef backupDir As String,ByRef sErr As String) As Boolean
    Dim pwd As String,t As String
    WMSCLEAN_Authorize=False
    pwd=InputBox(titleText & Chr(10) & Chr(10) & "Введите пароль:","WMS — защищённая очистка","")
    If pwd<>WMSCLEAN_PASSWORD Then
        If pwd<>"" Then MsgBox "Неверный пароль.",48,"WMS"
        Exit Function
    End If
    If Not WMSCLEAN_Backup(backupDir,sErr) Then Exit Function
    t=InputBox(preview & Chr(10) & Chr(10) & "Backup: " & backupDir & Chr(10) & Chr(10) & "Для подтверждения введите ровно: " & confirmWord,"WMS — последнее подтверждение","")
    If UCase(Trim(t))<>UCase(confirmWord) Then Exit Function
    WMSCLEAN_Authorize=True
End Function

Sub WMSCLEAN_Product()
    Dim key As String,code As String,nm As String,sErr As String,backupDir As String,preview As String
    Dim oCon As Object,n As Long,sql As String,afterN As Long
    If Not WMSCLEAN_Enter() Then Exit Sub
    On Error GoTo EH
    key=Trim(InputBox("Введите точный внутренний код товара или артикул.","WMS — очистить товар",""))
    If key="" Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    code=WMSCLEAN_ResolveProduct(oCon,key,nm,sErr)
    If sErr<>"" Or code="" Then GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code),sErr)
    If sErr<>"" Then GoTo Fail
    preview="Будет полностью очищен товар:" & Chr(10) & nm & Chr(10) & "Код: " & code & Chr(10) & "Движений: " & CStr(n) & Chr(10) & "Также удалятся его партии, выдачи, строки документов, строки заказов и сопоставления артикула."
    WMSDB_Close
    If Not WMSCLEAN_Authorize("ОЧИСТКА ОДНОГО ТОВАРА","УДАЛИТЬ ТОВАР",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql="EXECUTE BLOCK AS BEGIN " & _
        "DELETE FROM WMS_ISSUE_RETURNS WHERE SOURCE_ID IN (SELECT SOURCE_ID FROM WMS_ISSUES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "); " & _
        "DELETE FROM WMS_ISSUES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_ACT_LINES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_LOT_UNITS WHERE LOT_ID IN (SELECT LOT_ID FROM WMS_STOCK_LOTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "); " & _
        "DELETE FROM WMS_STOCK_MOVEMENTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_DOCUMENT_LINES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_STOCK_LOTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_ORDER_LINES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_PRODUCT_ALIASES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_PRODUCTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & "; " & _
        "DELETE FROM WMS_ACTS a WHERE NOT EXISTS (SELECT 1 FROM WMS_ACT_LINES l WHERE l.ACT_ID=a.ACT_ID); " & _
        "DELETE FROM WMS_DOCUMENTS d WHERE NOT EXISTS (SELECT 1 FROM WMS_DOCUMENT_LINES l WHERE l.DOC_ID=d.DOC_ID); " & _
        "DELETE FROM WMS_ORDERS o WHERE NOT EXISTS (SELECT 1 FROM WMS_ORDER_LINES l WHERE l.ORDER_GROUP_ID=o.ORDER_GROUP_ID); END"
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    afterN=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_PRODUCTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(code),sErr)
    If sErr<>"" Or afterN<>0 Then sErr="Контроль после очистки товара не пройден.":GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"Товар очищен полностью."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Очистка товара: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSCLEAN_Order()
    Dim key As String,gid As String,sErr As String,backupDir As String,preview As String,oCon As Object,n As Long,sql As String
    If Not WMSCLEAN_Enter() Then Exit Sub
    On Error GoTo EH
    key=Trim(InputBox("Введите номер/ID заказа. Поиск идёт по ORDER_GROUP_ID и номеру накладной.","WMS — очистить заказ",""))
    If key="" Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    gid=WMSCLEAN_ResolveOrder(oCon,key,sErr)
    If sErr<>"" Or gid="" Then GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_ORDER_LINES WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid),sErr)
    If sErr<>"" Then GoTo Fail
    preview="Заказ: " & gid & Chr(10) & "Строк заказа: " & CStr(n) & Chr(10) & "Удаляются только данные заказа из WMS_ORDERS/WMS_ORDER_LINES. Уже проведённые складские движения не удаляются автоматически."
    WMSDB_Close
    If Not WMSCLEAN_Authorize("ОЧИСТКА ОДНОГО ЗАКАЗА","УДАЛИТЬ ЗАКАЗ",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql="EXECUTE BLOCK AS BEGIN DELETE FROM WMS_ORDER_LINES WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid) & "; DELETE FROM WMS_ORDERS WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid) & "; END"
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"Заказ очищен. Складские движения не затрагивались."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Очистка заказа: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSCLEAN_Document()
    Dim key As String,docID As String,sErr As String,backupDir As String,preview As String,oCon As Object,n As Long,sql As String
    If Not WMSCLEAN_Enter() Then Exit Sub
    On Error GoTo EH
    key=Trim(InputBox("Введите DOC_ID или номер документа прихода.","WMS — очистить документ",""))
    If key="" Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    docID=WMSCLEAN_ResolveDocument(oCon,key,sErr)
    If sErr<>"" Or docID="" Then GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE SOURCE_ID=" & WMSDB_SQLText(docID),sErr)
    If sErr<>"" Then GoTo Fail
    preview="Документ: " & docID & Chr(10) & "Связанных движений: " & CStr(n) & Chr(10) & "Если по его партиям уже были расходы, очистка будет заблокирована."
    If WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID IN (SELECT LOT_ID FROM WMS_STOCK_LOTS WHERE SOURCE_ID=" & WMSDB_SQLText(docID) & ") AND SOURCE_ID<>" & WMSDB_SQLText(docID),sErr)>0 Then
        sErr="По партиям документа уже есть последующие движения. Сначала очищайте конкретные зависимые операции/товар, либо используйте полную очистку."
        GoTo Fail
    End If
    WMSDB_Close
    If Not WMSCLEAN_Authorize("ОЧИСТКА ДОКУМЕНТА / ПРИХОДА","УДАЛИТЬ ДОКУМЕНТ",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql="EXECUTE BLOCK AS BEGIN " & _
        "DELETE FROM WMS_LOT_UNITS WHERE LOT_ID IN (SELECT LOT_ID FROM WMS_STOCK_LOTS WHERE SOURCE_ID=" & WMSDB_SQLText(docID) & "); " & _
        "DELETE FROM WMS_STOCK_MOVEMENTS WHERE SOURCE_ID=" & WMSDB_SQLText(docID) & "; " & _
        "DELETE FROM WMS_STOCK_LOTS WHERE SOURCE_ID=" & WMSDB_SQLText(docID) & "; " & _
        "DELETE FROM WMS_DOCUMENT_LINES WHERE DOC_ID=" & WMSDB_SQLText(docID) & "; " & _
        "DELETE FROM WMS_DOCUMENTS WHERE DOC_ID=" & WMSDB_SQLText(docID) & "; END"
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"Документ/приход очищен."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Очистка документа: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSCLEAN_Lot()
    Dim lotID As String,sErr As String,backupDir As String,preview As String,oCon As Object,n As Long,sql As String
    If Not WMSCLEAN_Enter() Then Exit Sub
    On Error GoTo EH
    lotID=Trim(InputBox("Введите точный LOT_ID партии.","WMS — очистить партию",""))
    If lotID="" Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_LOTS WHERE LOT_ID=" & WMSDB_SQLText(lotID),sErr)
    If sErr<>"" Then GoTo Fail
    If n<>1 Then sErr="Партия не найдена или неоднозначна.":GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID=" & WMSDB_SQLText(lotID),sErr)
    preview="Партия: " & lotID & Chr(10) & "Движений партии: " & CStr(n) & Chr(10) & "Будут удалены все движения именно этой партии и связанные выдачи."
    WMSDB_Close
    If Not WMSCLEAN_Authorize("ОЧИСТКА ОДНОЙ ПАРТИИ","УДАЛИТЬ ПАРТИЮ",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql="EXECUTE BLOCK AS BEGIN " & _
        "DELETE FROM WMS_ISSUE_RETURNS WHERE SOURCE_ID IN (SELECT SOURCE_ID FROM WMS_ISSUES WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "); " & _
        "DELETE FROM WMS_ISSUES WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "; " & _
        "DELETE FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "; " & _
        "DELETE FROM WMS_DOCUMENT_LINES WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "; " & _
        "DELETE FROM WMS_LOT_UNITS WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "; " & _
        "DELETE FROM WMS_STOCK_LOTS WHERE LOT_ID=" & WMSDB_SQLText(lotID) & "; END"
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"Партия очищена."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Очистка партии: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSCLEAN_AllOperations()
    Dim sErr As String,backupDir As String,preview As String,oCon As Object,sql As String,n As Long
    If Not WMSCLEAN_Enter() Then Exit Sub
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS",sErr)
    preview="Будут удалены ВСЕ складские операции: движения, партии, выдачи/возвраты, документы операций, инвентаризации и акты." & Chr(10) & "Движений сейчас: " & CStr(n) & Chr(10) & "Заказы, номенклатура и справочники сохраняются."
    WMSDB_Close
    If Not WMSCLEAN_Authorize("ОЧИСТКА ВСЕХ СКЛАДСКИХ ОПЕРАЦИЙ","ОЧИСТИТЬ ОПЕРАЦИИ",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql=WMSCLEAN_SQL_AllOperations(False)
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"Все складские операции очищены. Заказы, товары и справочники сохранены."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Очистка операций: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSCLEAN_FullReset()
    Dim sErr As String,backupDir As String,preview As String,oCon As Object,sql As String
    If Not WMSCLEAN_Enter() Then Exit Sub
    preview="ПОЛНАЯ ЧИСТКА удалит рабочие данные WMS: складские операции, заказы, номенклатуру и артикулы." & Chr(10) & "Структура Firebird, настройки, WMS_META и справочники WMS_REFERENCE сохраняются."
    If Not WMSCLEAN_Authorize("ПОЛНАЯ ЧИСТКА WMS","ПОЛНАЯ ЧИСТКА",preview,backupDir,sErr) Then GoTo Cancelled
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql=WMSCLEAN_SQL_AllOperations(True)
    If Not WMSCLEAN_ExecAtomic(oCon,sql,sErr) Then GoTo Fail
    WMSCLEAN_Finish oCon,backupDir,"WMS очищена до пустого рабочего состояния. Структура и справочники сохранены."
    Exit Sub
Fail:
    WMSCLEAN_Fail oCon,backupDir,sErr
    Exit Sub
Cancelled:
    WMSCLEAN_Cancel
    Exit Sub
EH:
    sErr="Полная чистка: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSCLEAN_SQL_AllOperations(includeMaster As Boolean) As String
    Dim s As String
    s="EXECUTE BLOCK AS BEGIN "
    s=s & "DELETE FROM WMS_ISSUE_RETURNS; DELETE FROM WMS_ISSUES; "
    s=s & "DELETE FROM WMS_ACT_LINES; DELETE FROM WMS_ACTS; "
    s=s & "DELETE FROM WMS_INVENTORY_LINES; DELETE FROM WMS_INVENTORY; "
    s=s & "DELETE FROM WMS_LOT_UNITS; DELETE FROM WMS_STOCK_MOVEMENTS; DELETE FROM WMS_STOCK_LOTS; "
    s=s & "DELETE FROM WMS_DOCUMENT_LINES; DELETE FROM WMS_DOCUMENTS; "
    s=s & "DELETE FROM WMS_CORRECTIONS; DELETE FROM WMS_OPERATION_GUARD; "
    If includeMaster Then
        s=s & "DELETE FROM WMS_ORDER_LINES; DELETE FROM WMS_ORDERS; DELETE FROM WMS_PRODUCT_ALIASES; DELETE FROM WMS_PRODUCTS; "
    End If
    s=s & "END"
    WMSCLEAN_SQL_AllOperations=s
End Function

Function WMSCLEAN_ResolveProduct(oCon As Object,key As String,ByRef nm As String,ByRef sErr As String) As String
    Dim ps As Object,rs As Object,n As Long,code As String
    WMSCLEAN_ResolveProduct=""
    nm=""
    sErr=""
    On Error GoTo EH
    ps=oCon.prepareStatement("SELECT DISTINCT p.PRODUCT_CODE,p.PRODUCT_NAME FROM WMS_PRODUCTS p LEFT JOIN WMS_PRODUCT_ALIASES a ON a.PRODUCT_CODE=p.PRODUCT_CODE WHERE UPPER(TRIM(p.PRODUCT_CODE))=UPPER(TRIM(?)) OR UPPER(TRIM(COALESCE(p.SUPPLIER_ARTICLE,'')))=UPPER(TRIM(?)) OR UPPER(TRIM(COALESCE(a.ARTICLE_CODE,'')))=UPPER(TRIM(?))")
    ps.setString(1,key)
    ps.setString(2,key)
    ps.setString(3,key)
    rs=ps.executeQuery()
    Do While rs.next()
        n=n+1
        If n=1 Then code=rs.getString(1):nm=rs.getString(2)
        If n>1 Then Exit Do
    Loop
    If n=0 Then sErr="Товар не найден."
    If n>1 Then sErr="Найдено несколько товаров. Используйте внутренний код товара."
    If n=1 Then WMSCLEAN_ResolveProduct=code
Done:
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="Поиск товара: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSCLEAN_ResolveOrder(oCon As Object,key As String,ByRef sErr As String) As String
    Dim ps As Object,rs As Object,n As Long,gid As String
    WMSCLEAN_ResolveOrder=""
    sErr=""
    On Error GoTo EH
    ps=oCon.prepareStatement("SELECT ORDER_GROUP_ID FROM WMS_ORDERS WHERE UPPER(TRIM(ORDER_GROUP_ID))=UPPER(TRIM(?)) OR UPPER(TRIM(COALESCE(INVOICE_NO,'')))=UPPER(TRIM(?))")
    ps.setString(1,key)
    ps.setString(2,key)
    rs=ps.executeQuery()
    Do While rs.next()
        n=n+1
        If n=1 Then gid=rs.getString(1)
        If n>1 Then Exit Do
    Loop
    If n=0 Then sErr="Заказ не найден."
    If n>1 Then sErr="Найдено несколько заказов. Используйте точный ORDER_GROUP_ID."
    If n=1 Then WMSCLEAN_ResolveOrder=gid
Done:
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="Поиск заказа: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSCLEAN_ResolveDocument(oCon As Object,key As String,ByRef sErr As String) As String
    Dim ps As Object,rs As Object,n As Long,id As String
    WMSCLEAN_ResolveDocument=""
    sErr=""
    On Error GoTo EH
    ps=oCon.prepareStatement("SELECT DOC_ID FROM WMS_DOCUMENTS WHERE UPPER(TRIM(DOC_ID))=UPPER(TRIM(?)) OR UPPER(TRIM(COALESCE(DOC_NO,'')))=UPPER(TRIM(?))")
    ps.setString(1,key)
    ps.setString(2,key)
    rs=ps.executeQuery()
    Do While rs.next()
        n=n+1
        If n=1 Then id=rs.getString(1)
        If n>1 Then Exit Do
    Loop
    If n=0 Then sErr="Документ не найден."
    If n>1 Then sErr="Найдено несколько документов с таким номером. Используйте DOC_ID."
    If n=1 Then WMSCLEAN_ResolveDocument=id
Done:
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="Поиск документа: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSCLEAN_ExecAtomic(oCon As Object,sql As String,ByRef sErr As String) As Boolean
    Dim st As Object
    WMSCLEAN_ExecAtomic=False
    sErr=""
    On Error GoTo EH
    st=oCon.createStatement()
    st.execute(sql)
    WMSDBX_CloseStmt st
    oCon.commit()
    WMSCLEAN_ExecAtomic=True
    Exit Function
EH:
    sErr="Очистка Firebird: " & CStr(Err) & " " & Error$
    WMSDBX_CloseStmt st
End Function

Function WMSCLEAN_Enter() As Boolean
    WMSCLEAN_Enter=False
    If Not WMSDBX_TryEnter("CLEANUP") Then
        MsgBox "Другая операция WMS ещё выполняется.",48,"WMS"
        Exit Function
    End If
    WMSCLEAN_Enter=True
End Function

Sub WMSCLEAN_Finish(oCon As Object,backupDir As String,msg As String)
    Dim e As String
    On Error Resume Next
    oCon.commit()
    WMSDBX_SaveODB e
    WMSDB_Close
    WMSDBST_RefreshStock
    WMSREF_RefreshCache
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox msg & Chr(10) & Chr(10) & "Резервная копия: " & backupDir,64,"WMS — Очистка"
End Sub

Sub WMSCLEAN_Fail(oCon As Object,backupDir As String,sErr As String)
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox "Очистка НЕ завершена." & Chr(10) & sErr & Chr(10) & Chr(10) & "Резервная копия: " & backupDir,16,"WMS — Очистка"
End Sub

Sub WMSCLEAN_Cancel()
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    WMSDBX_Leave
End Sub

Function WMSCLEAN_Backup(ByRef backupDir As String,ByRef sErr As String) As Boolean
    Dim oSFA As Object,baseFolder As String,root As String,sep As String,odbURL As String
    WMSCLEAN_Backup=False
    backupDir=""
    sErr=""
    On Error GoTo EH
    ThisComponent.store()
    On Error Resume Next
    WMSDB_Close
    On Error GoTo EH
    sep=WMSACT_PathSep()
    baseFolder=WMSACT_BaseFolder()
    If baseFolder="" Then sErr="WMS должна быть сохранена как ODS.":Exit Function
    root=baseFolder & "WMS_Cleanup_Backup" & sep
    WMSACT_EnsureFolder root
    backupDir=root & Format(Now,"YYYYMMDD_HHMMSS") & sep
    WMSACT_EnsureFolder backupDir
    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    oSFA.copy(ThisComponent.URL,ConvertToURL(backupDir & WMSCLEAN_FileName(ThisComponent.URL)))
    odbURL=ConvertToURL(baseFolder & "WMS_DATA_PORTABLE.odb")
    If Not oSFA.exists(odbURL) Then sErr="WMS_DATA_PORTABLE.odb рядом с ODS не найден.":Exit Function
    oSFA.copy(odbURL,ConvertToURL(backupDir & "WMS_DATA_PORTABLE.odb"))
    WMSCLEAN_Backup=True
    Exit Function
EH:
    sErr="Backup перед очисткой: " & CStr(Err) & " " & Error$
End Function

Function WMSCLEAN_FileName(fileURL As String) As String
    Dim p As String,sep As String,n As Long
    p=ConvertFromURL(fileURL)
    sep=WMSACT_PathSep()
    n=WMSACT_LastSepPos(p,sep)
    If n>0 Then p=Mid(p,n+1)
    WMSCLEAN_FileName=p
End Function

Sub WMSRESET_ResetLastTwoReceiptSessions()
    MsgBox "Функция 'два последних прихода' заменена Центром безопасной очистки на листе Инфо.",64,"WMS 3.3"
End Sub

Sub WMSCLEAN_SyntaxProbe()
End Sub
