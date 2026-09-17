Option Explicit

Global Const WMSX_VERSION = "3.3.0-ORDERS-WITH-STOCK"
Global Const WMSX_FORM = "WMS_EXPORT_FORM"

Sub WMSX_InstallUI()
    Dim oDoc As Object,oSh As Object,forms As Object,form As Object
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName("Инфо") Then Exit Sub
    oSh=oDoc.Sheets.getByName("Инфо")
    On Error GoTo EH
    WMSX_RemoveButton oSh
    oSh.getCellByPosition(0,18).String="ТОВАРЫ ИЗ ЗАКАЗОВ"
    oSh.getCellByPosition(0,19).String="Создаёт один локальный файл Exports/orders_goods.tsv: все строки заказов + текущий остаток товара. Используется отдельным WMS_ТОВАРЫ.ods."
    oSh.getCellRangeByPosition(0,18,9,18).CharWeight=150
    oSh.getCellRangeByPosition(0,19,9,19).IsTextWrapped=True
    forms=oSh.DrawPage.Forms
    On Error Resume Next
    If forms.hasByName(WMSX_FORM) Then forms.removeByName(WMSX_FORM)
    On Error GoTo EH
    form=oDoc.createInstance("com.sun.star.form.component.Form")
    form.Name=WMSX_FORM
    forms.insertByName(WMSX_FORM,form)
    WMSX_AddButton oDoc,oSh,form
    Exit Sub
EH:
    MsgBox "Кнопка экспорта товаров: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSX_AddButton(oDoc As Object,oSh As Object,oForm As Object)
    Dim model As Object,shape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    model=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    model.Name="WMS_EXPORT_REPORT_DATA"
    model.Label="ОБНОВИТЬ ФАЙЛ ТОВАРОВ"
    model.Tabstop=False
    oForm.insertByName(model.Name,model)
    idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    shape.Control=model
    a=oSh.getCellByPosition(0,20).Position
    shape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size")
    sz.Width=6200
    sz.Height=850
    shape.Size=sz
    oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_33_OfflineExports.WMSX_ExportAll?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSX_RemoveButton(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i)
        nm=""
        ctl=shape.Control
        nm=ctl.Name
        If nm="WMS_EXPORT_REPORT_DATA" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub

Sub WMSX_ExportAll()
    Dim oCon As Object,sErr As String,baseDir As String,expDir As String,filePath As String,sql As String
    If Not WMSDBX_TryEnter("REPORT_EXPORT") Then MsgBox "Другая операция WMS ещё выполняется.",48,"WMS":Exit Sub
    On Error GoTo EH
    baseDir=WMSACT_BaseFolder()
    If baseDir="" Then sErr="Не определена папка WMS.":GoTo Fail
    expDir=baseDir & "Exports" & WMSACT_PathSep()
    WMSACT_EnsureFolder expDir
    filePath=expDir & "orders_goods.tsv"
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    sql="SELECT l.ORDER_GROUP_ID AS ORDER_GROUP_ID,l.POSITION_NO AS POSITION_NO,l.PRODUCT_NAME AS PRODUCT_NAME,l.SUPPLIER_ARTICLE AS SUPPLIER_ARTICLE," & _
        "CAST(l.ORDER_QTY AS DOUBLE PRECISION) AS ORDER_QTY,CAST(COALESCE(l.FACT_QTY,0) AS DOUBLE PRECISION) AS FACT_QTY,l.UNIT_NAME AS UNIT_NAME," & _
        "CAST(COALESCE(l.PRICE,0) AS DOUBLE PRECISION) AS PRICE,CAST(COALESCE(l.TOTAL_AMOUNT,0) AS DOUBLE PRECISION) AS TOTAL_AMOUNT," & _
        "l.SUPPLIER_NAME AS SUPPLIER_NAME,l.SELLER_NAME AS SELLER_NAME,l.ORDER_DATE AS ORDER_DATE,l.STATUS_NAME AS STATUS_NAME," & _
        "l.CATEGORY_NAME AS CATEGORY_NAME,l.LOCATION_NAME AS LOCATION_NAME,l.PRODUCT_CODE AS PRODUCT_CODE," & _
        "CAST(COALESCE(s.STOCK_QTY,0) AS DOUBLE PRECISION) AS STOCK_QTY,l.DESTINATION_NAME AS DESTINATION_NAME,l.EXPECTED_RECEIPT_DATE AS EXPECTED_RECEIPT_DATE " & _
        "FROM WMS_ORDER_LINES l LEFT JOIN (SELECT PRODUCT_CODE,SUM(QTY) STOCK_QTY FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' GROUP BY PRODUCT_CODE) s ON s.PRODUCT_CODE=l.PRODUCT_CODE " & _
        "ORDER BY l.ORDER_DATE,l.ORDER_GROUP_ID,l.POSITION_NO"
    If Not WMSX_ExportQuery(oCon,filePath,sql,sErr) Then GoTo Fail
    WMSDB_Close
    WMSDBX_Leave
    MsgBox "Готово." & Chr(10) & filePath & Chr(10) & Chr(10) & "В WMS_ТОВАРЫ.ods нажмите «Обновить». Все данные остаются локально.",64,"WMS — Товары"
    Exit Sub
Fail:
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox "Экспорт остановлен: " & sErr,16,"WMS — Товары"
    Exit Sub
EH:
    sErr="Экспорт: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSX_ExportQuery(oCon As Object,filePath As String,sql As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,md As Object,nCols As Long,i As Long
    Dim line As String,sfa As Object,outSt As Object,txt As Object
    Dim targetURL As String,pendingURL As String,previousURL As String
    Dim pendingCreated As Boolean,previousMoved As Boolean
    WMSX_ExportQuery=False
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    md=oRS.getMetaData()
    nCols=md.getColumnCount()
    sfa=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    targetURL=ConvertToURL(filePath)
    pendingURL=targetURL & ".pending"
    previousURL=targetURL & ".previous"
    If sfa.exists(pendingURL) Then
        sErr="Найден незавершённый экспорт: " & filePath & ".pending. Сохраните его отдельно и освободите это имя перед повтором."
        GoTo FailedExport
    End If
    If Not sfa.exists(targetURL) And sfa.exists(previousURL) Then
        sErr="Найдена предыдущая выгрузка: " & filePath & ".previous. Восстановите её под исходным именем перед повтором."
        GoTo FailedExport
    End If
    pendingCreated=True
    outSt=sfa.openFileWrite(pendingURL)
    txt=CreateUnoService("com.sun.star.io.TextOutputStream")
    txt.setOutputStream(outSt)
    txt.setEncoding("UTF-8")
    line=""
    For i=1 To nCols
        If i>1 Then line=line & Chr(9)
        line=line & WMSX_Clean(CStr(md.getColumnLabel(i)))
    Next i
    txt.writeString(line & Chr(10))
    Do While oRS.next()
        line=""
        For i=1 To nCols
            If i>1 Then line=line & Chr(9)
            line=line & WMSX_Clean(oRS.getString(i))
        Next i
        txt.writeString(line & Chr(10))
    Loop
    txt.closeOutput()
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    ' Only publish after complete write and successful close.
    ' Keep one recoverable previous export; this is not a cross-file transaction.
    If sfa.exists(targetURL) Then
        If sfa.exists(previousURL) Then sfa.kill(previousURL)
        sfa.move(targetURL,previousURL)
        previousMoved=True
    End If
    sfa.move(pendingURL,targetURL)
    pendingCreated=False
    WMSX_ExportQuery=True
    Exit Function
EH:
    sErr=CStr(Err) & " " & Error$
FailedExport:
    On Error Resume Next
    If previousMoved Then
        If Not sfa.exists(targetURL) Then sfa.move(previousURL,targetURL)
        sErr=sErr & " | При необходимости восстановите файл .previous; незавершённый .pending не используйте."
    End If
    If Not IsEmpty(txt) Then txt.closeOutput()
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    On Error GoTo 0
End Function

Function WMSX_Clean(v As Variant) As String
    Dim s As String
    s=CStr(v)
    s=Replace(s,Chr(9)," ")
    s=Replace(s,Chr(13)," ")
    s=Replace(s,Chr(10)," ")
    WMSX_Clean=s
End Function

Sub WMSX_SyntaxProbe()
End Sub

Function WMSX_SaveOrderSnapshot(con As Object,sh As Object,r As Long,sid As String,ByRef sErr As String) As Boolean
    Dim ps As Object,c As Long,lastCol As Long,h As String,v As String,n As Long
    WMSX_SaveOrderSnapshot=False:sErr=""
    On Error GoTo EH
    n=WMSDBX_ScalarLong(con,"SELECT COUNT(*) FROM WMS_ORDER_SNAPSHOT WHERE SOURCE_ID=" & WMSDB_SQLText(sid),sErr)
    If sErr<>"" Then Exit Function
    ' Caller only invokes this for a source with no posted movement.
    ' Replace an orphaned snapshot from an earlier failed attempt.
    If n>0 Then
        ps=con.prepareStatement("DELETE FROM WMS_ORDER_SNAPSHOT WHERE SOURCE_ID=?")
        ps.setString(1,sid):ps.executeUpdate():WMSDBX_CloseStmt ps
    End If
    lastCol=WMSDBO_LastHeaderCol(sh)
    ps=con.prepareStatement("INSERT INTO WMS_ORDER_SNAPSHOT (SOURCE_ID,COL_INDEX,HEADER_NAME,VALUE_TEXT) VALUES (?,?,?,?)")
    For c=0 To lastCol
        h=Trim(sh.getCellByPosition(c,0).String)
        If Left(h,5)<>"_WMS_" Then
            If h="" Then h="Столбец " & CStr(c+1)
            v=sh.getCellByPosition(c,r).String
            If Len(v)>8000 Or Len(h)>255 Then
                sErr="Поле заказа превышает размер снимка (8000 символов). Проведение остановлено без усечения текста."
                GoTo Done
            End If
            ps.setString(1,sid):ps.setInt(2,c):ps.setString(3,h):ps.setString(4,v)
            ps.executeUpdate()
        End If
    Next c
    WMSX_SaveOrderSnapshot=True
Done:
    WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="Снимок исходного заказа: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSX_OrderStockInstall()
    Dim doc As Object,sh As Object,form As Object,model As Object,shape As Object,i As Long
    Dim labels As Variant,actions As Variant
    Dim p As New com.sun.star.awt.Point,sz As New com.sun.star.awt.Size
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    doc=ThisComponent
    If Not doc.Sheets.hasByName("Остаток — Заказы") Then doc.Sheets.insertNewByName("Остаток — Заказы",doc.Sheets.Count)
    sh=doc.Sheets.getByName("Остаток — Заказы")
    sh.getCellByPosition(0,0).String="ПРОВЕДЁННЫЕ ЗАКАЗЫ И ОСТАТОК ИХ ПАРТИЙ"
    sh.getCellByPosition(0,1).String="Реквизиты из базы + сохранённый снимок ввода. Нулевые остатки включены. Новые снимки сохраняются начиная с 1.4.0."
    sh.getCellRangeByPosition(0,0,10,0).CellBackColor=2240061
    sh.getCellRangeByPosition(0,0,10,0).CharColor=16777215
    sh.Rows.getByIndex(0).Height=900:sh.Rows.getByIndex(1).Height=1000:sh.Rows.getByIndex(2).Height=1000
    sh.getCellRangeByPosition(0,1,10,1).IsTextWrapped=True
    For i=sh.DrawPage.Count-1 To 0 Step -1
        shape=sh.DrawPage.getByIndex(i)
        If shape.supportsService("com.sun.star.drawing.ControlShape") Then
            If Left(shape.Control.Name,8)="WMS_OXS_" Then sh.DrawPage.remove(shape)
        End If
    Next i
    If sh.DrawPage.Forms.hasByName("WMS_OXS") Then sh.DrawPage.Forms.removeByName("WMS_OXS")
    form=doc.createInstance("com.sun.star.form.component.Form"):form.Name="WMS_OXS"
    sh.DrawPage.Forms.insertByName("WMS_OXS",form)
    labels=Array("Обновить из базы","Выгрузить в Calc")
    actions=Array("WMSX_OrderStockRefresh","WMSX_OrderStockExport")
    For i=0 To 1
        model=doc.createInstance("com.sun.star.form.component.CommandButton")
        model.Name="WMS_OXS_" & CStr(i):model.Label=labels(i):model.BackgroundColor=2240061:model.TextColor=16777215
        form.insertByName(model.Name,model)
        shape=doc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=model
        sh.DrawPage.add(shape):shape.Anchor=sh
        p.X=i*6500:p.Y=2000:sz.Width=6000:sz.Height=900:shape.Position=p:shape.Size=sz
        ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed":ev.ScriptType="Script"
        ev.ScriptCode="vnd.sun.star.script:Standard.WMS_33_OfflineExports." & actions(i) & "?language=Basic&location=document"
        form.registerScriptEvent i,ev
    Next i
End Sub

Sub WMSX_OrderStockRefresh()
    Dim con As Object,st As Object,rs As Object,ps As Object,sr As Object,sh As Object
    Dim sErr As String,sql As String,fields As Variant,headers As Variant,nums As Variant
    Dim data() As Variant,line() As Variant,n As Long,i As Long,j As Long,k As Long,c As Long,sid As String
    Dim extraHeaders() As String,extraCols() As Long,extraN As Long,allHeaders() As Variant,oldLast As Long
    If Not WMSDBX_TryEnter("ORDER_STOCK") Then MsgBox "Дождитесь завершения операции.",48,"WMS":Exit Sub
    On Error GoTo EH
    con=WMSDB_GetConnectionEx(ThisComponent,sErr):If sErr<>"" Then GoTo Done
    If Not WMSDB_TableExistsSimple(con,"WMS_ORDER_SNAPSHOT",sErr) Then
        sErr="Сначала установите обновление WMS 1.4.0 кнопкой «Обновить WMS»."
        GoTo Done
    End If
    fields=Split("SOURCE_ID|POSITION_NO|PRODUCT_NAME|DOC_NO|INVOICE_NO|SUPPLIER_ARTICLE|FACT_QTY|ORDER_QTY|UNIT_NAME|PRICE|TOTAL_AMOUNT|SUPPLIER_NAME|RECEIPT_DATE|DOC_DATE|ORDER_DATE|BUYER|STATUS_NAME|CATEGORY_NAME|CONTROL_TEXT|LOCATION_NAME|COMMENT_TEXT|PRODUCT_CODE|SELLER_NAME|DESTINATION_NAME|EXPECTED_RECEIPT_DATE|EXTERNAL_ORDER_NO|RECEIVED_FROM|RECEIPT_SOURCE","|")
    headers=Split("ID строки|№ позиции|Наименование|Закрывающий документ|Счёт|Артикул|Факт при проведении|Заказано|Ед. документа|Цена|Сумма|Поставщик|Дата поступления|Дата документа|Дата заказа|Кто заказывал|Статус при проведении|Категория|Контроль|Место при поступлении|Комментарий|Внутренний код|Продавец|Назначение|Ожидаемая дата|Внешний номер заказа|От кого принято|Источник","|")
    st=con.createStatement()
    rs=st.executeQuery("SELECT DISTINCT COL_INDEX,HEADER_NAME FROM WMS_ORDER_SNAPSHOT ORDER BY COL_INDEX,HEADER_NAME")
    extraN=0
    Do While rs.next()
        ReDim Preserve extraHeaders(0 To extraN):ReDim Preserve extraCols(0 To extraN)
        extraCols(extraN)=rs.getLong(1):extraHeaders(extraN)=rs.getString(2):extraN=extraN+1
        If extraN>256 Then sErr="Более 256 разных полей снимков. Выгрузка остановлена без усечения.":GoTo Done
    Loop
    WMSDBX_CloseRS rs
    ReDim allHeaders(0 To UBound(headers)+3+extraN)
    For i=0 To UBound(headers):allHeaders(i)=headers(i):Next i
    k=UBound(headers)+1
    allHeaders(k)="Остаток партии":allHeaders(k+1)="Базовая единица":allHeaders(k+2)="Снимок ввода"
    For i=0 To extraN-1:allHeaders(k+3+i)="Ввод: " & extraHeaders(i):Next i
    sql="SELECT "
    For i=0 To UBound(fields)
        If i>0 Then sql=sql & ","
        If i=6 Or i=7 Or i=9 Or i=10 Then
            sql=sql & "CAST(l." & fields(i) & " AS DOUBLE PRECISION)"
        Else
            sql=sql & "l." & fields(i)
        End If
    Next i
    sql=sql & ",CAST(COALESCE(b.QTY,0) AS DOUBLE PRECISION),m.UNIT_NAME FROM WMS_ORDER_LINES l JOIN WMS_STOCK_MOVEMENTS m ON m.MOVEMENT_ID='IN-' || l.SOURCE_ID AND COALESCE(m.STATUS_NAME,'POSTED')='POSTED' LEFT JOIN (SELECT LOT_ID,SUM(QTY) QTY FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' GROUP BY LOT_ID) b ON b.LOT_ID=m.LOT_ID ORDER BY l.ORDER_DATE,l.ORDER_GROUP_ID,l.POSITION_NO,l.SOURCE_ID"
    rs=st.executeQuery(sql)
    ps=con.prepareStatement("SELECT COL_INDEX,HEADER_NAME,VALUE_TEXT FROM WMS_ORDER_SNAPSHOT WHERE SOURCE_ID=? ORDER BY COL_INDEX")
    n=0
    Do While rs.next()
        If n>=20000 Then sErr="Более 20000 строк. Выгрузка остановлена без публикации неполного списка.":GoTo Done
        ReDim line(0 To UBound(allHeaders))
        For i=0 To UBound(line):line(i)="":Next i
        For i=0 To UBound(fields):line(i)=rs.getString(i+1):Next i
        line(k)=rs.getDouble(k+1):line(k+1)=rs.getString(k+2)
        line(k+2)="Нет исторического снимка; реквизиты из базы"
        sid=rs.getString(1):ps.setString(1,sid):sr=ps.executeQuery()
        Do While sr.next()
            line(k+2)="Сохранён при проведении"
            For j=0 To extraN-1
                If extraCols(j)=sr.getLong(1) And extraHeaders(j)=sr.getString(2) Then line(k+3+j)=sr.getString(3):Exit For
            Next j
        Loop
        WMSDBX_CloseRS sr
        ReDim Preserve data(0 To n):data(n)=line():n=n+1
    Loop
    sh=ThisComponent.Sheets.getByName("Остаток — Заказы")
    oldLast=WMSCore_LastContentRow(sh,UBound(allHeaders)):If oldLast<4 Then oldLast=4
    sh.getCellRangeByPosition(0,4,UBound(allHeaders),oldLast).clearContents(1023)
    sh.getCellRangeByPosition(0,4,UBound(allHeaders),4).setDataArray(Array(allHeaders()))
    If n>0 Then sh.getCellRangeByPosition(0,5,UBound(allHeaders),n+4).setDataArray(data())
    sh.getCellRangeByPosition(0,4,UBound(allHeaders),4).CellBackColor=2240061
    sh.getCellRangeByPosition(0,4,UBound(allHeaders),4).CharColor=16777215
    sh.getCellRangeByPosition(0,4,UBound(allHeaders),4).IsTextWrapped=True
    sh.Rows.getByIndex(4).Height=1600
    For i=0 To UBound(allHeaders):sh.Columns.getByIndex(i).Width=4200:Next i
    sh.Columns.getByIndex(2).Width=7500
    sh.getCellByPosition(0,3).String="Обновлено " & Format(Now,"DD.MM.YYYY HH:MM:SS") & "; проведённых строк: " & CStr(n)
    ThisComponent.CurrentController.setActiveSheet(sh)
Done:
    WMSDBX_CloseRS sr:WMSDBX_CloseRS rs:WMSDBX_CloseStmt ps:WMSDBX_CloseStmt st
    WMSDBX_Leave
    If sErr<>"" Then MsgBox sErr,48,"Остаток — Заказы"
    Exit Sub
EH:
    sErr="Выгрузка проведённых заказов: " & CStr(Err) & " " & Error$
    Resume Done
End Sub

Sub WMSX_OrderStockExport()
    Dim src As Object,dest As Object,sh As Object,cur As Object,args(0) As New com.sun.star.beans.PropertyValue
    Dim path As String,i As Long
    On Error GoTo EH
    src=ThisComponent:sh=src.Sheets.getByName("Остаток — Заказы")
    If Left(sh.getCellByPosition(0,3).String,9)<>"Обновлено" Then MsgBox "Сначала нажмите «Обновить из базы».",48,"WMS":Exit Sub
    path=WMSACT_BaseFolder() & "Exports" & WMSACT_PathSep()
    WMSACT_EnsureFolder path
    path=path & "orders_stock_" & Format(Now,"YYYYMMDD_HHMMSS") & ".ods"
    If WMSDB_FileExists(ConvertToURL(path)) Then MsgBox "Файл с таким временем уже существует. Повторите через секунду.",48,"WMS":Exit Sub
    dest=StarDesktop.loadComponentFromURL("private:factory/scalc","_blank",0,Array())
    cur=sh.createCursor():cur.gotoEndOfUsedArea(True)
    dest.Sheets.getByIndex(0).getCellRangeByPosition(0,0,cur.RangeAddress.EndColumn,cur.RangeAddress.EndRow).setDataArray(sh.getCellRangeByPosition(0,0,cur.RangeAddress.EndColumn,cur.RangeAddress.EndRow).getDataArray())
    For i=0 To cur.RangeAddress.EndColumn
        dest.Sheets.getByIndex(0).Columns.getByIndex(i).Width=sh.Columns.getByIndex(i).Width
    Next i
    dest.Sheets.getByIndex(0).getCellRangeByPosition(0,4,cur.RangeAddress.EndColumn,4).CellBackColor=2240061
    dest.Sheets.getByIndex(0).getCellRangeByPosition(0,4,cur.RangeAddress.EndColumn,4).CharColor=16777215
    dest.Sheets.getByIndex(0).getCellRangeByPosition(0,4,cur.RangeAddress.EndColumn,4).IsTextWrapped=True
    dest.Sheets.getByIndex(0).Rows.getByIndex(4).Height=1600
    args(0).Name="FilterName":args(0).Value="calc8"
    dest.storeAsURL(ConvertToURL(path),args())
    MsgBox "Выгрузка сохранена и открыта в Calc:" & Chr(10) & path,64,"WMS"
    Exit Sub
EH:
    MsgBox "Файл не выгружен: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub
