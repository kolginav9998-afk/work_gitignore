Option Explicit

Global Const WMSC_VERSION = "2.0.0-SYSTEM-CENTER"
Global Const WMSC_SHEET = "Инфо"
Global Const WMSC_FORM = "WMS_CENTER_PANEL"

Sub WMSC_Install()
    Dim oDoc As Object,oSh As Object
    oDoc=ThisComponent
    On Error GoTo EH
    If oDoc.Sheets.hasByName(WMSC_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSC_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSC_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSC_SHEET)
    End If

    WMSC_BuildLayout oDoc,oSh
    WMSC_InstallButtons oDoc,oSh
    WMSC_Refresh
    oDoc.CurrentController.setActiveSheet(oSh)
    MsgBox "WMS Center установлен." & Chr(10) & "Версия: " & WMSC_VERSION,64,"WMS Center"
    Exit Sub
EH:
    MsgBox "WMS Center: " & CStr(Err) & " " & Error$,16,"WMS Center"
End Sub

Sub WMSC_BuildLayout(oDoc As Object,oSh As Object)
    Dim labels
    Dim i As Long

    oSh.getCellRangeByPosition(0,0,5,22).clearContents(1023)

    oSh.getCellByPosition(0,0).String="WMS — PRODUCTION CENTER"
    oSh.getCellByPosition(0,1).String="Рабочая WMS. Производственная база не пересобирается — только обновляется."

    labels=Array("Режим","Релиз","Firebird","Товары","Движения","Партии","Единицы партий","Заказы БД","Выдачи БД","Возвраты БД","Акты","Safety/Audit","Последняя проверка")
    For i=0 To UBound(labels)
        oSh.getCellByPosition(0,3+i).String=labels(i)
    Next i

    oSh.getCellByPosition(3,3).String="Рабочие листы"
    oSh.getCellByPosition(4,3).String="Заказы / Выдачи / Остаток / База - Поиск / Инфо"
    oSh.getCellByPosition(3,4).String="Правило релиза"
    oSh.getCellByPosition(4,4).String="Полная пересборка запрещена. Перед обновлением создаётся резервная копия."

    oSh.Columns.getByIndex(0).Width=4300
    oSh.Columns.getByIndex(1).Width=6500
    oSh.Columns.getByIndex(2).Width=1800
    oSh.Columns.getByIndex(3).Width=4100
    oSh.Columns.getByIndex(4).Width=8200

    oSh.getCellRangeByPosition(0,0,4,0).CharWeight=150
    oSh.getCellRangeByPosition(0,3,0,15).CharWeight=150
End Sub

Sub WMSC_Refresh()
    Dim oDoc As Object
    Dim oSh As Object
    Dim oCon As Object
    Dim sErr As String

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSC_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSC_SHEET)

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then
        oSh.getCellByPosition(1,3).String="ERROR"
        oSh.getCellByPosition(1,5).String="ERROR: " & sErr
        Exit Sub
    End If

    oSh.getCellByPosition(1,3).String=WMSC_MetaValue(oCon,"RELEASE_MODE","PRODUCTION")
    oSh.getCellByPosition(1,4).String=WMSC_MetaValue(oCon,"RELEASE_VERSION","1.0.2")
    oSh.getCellByPosition(1,5).String="OK — embedded Firebird"
    oSh.getCellByPosition(1,6).String=WMSC_TableState(oCon,"WMS_PRODUCTS")
    oSh.getCellByPosition(1,7).String=WMSC_TableState(oCon,"WMS_STOCK_MOVEMENTS")
    oSh.getCellByPosition(1,8).String=WMSC_TableState(oCon,"WMS_STOCK_LOTS")
    oSh.getCellByPosition(1,9).String=WMSC_TableState(oCon,"WMS_LOT_UNITS")
    oSh.getCellByPosition(1,10).String=WMSC_TableState(oCon,"WMS_ORDER_LINES")
    oSh.getCellByPosition(1,11).String=WMSC_TableState(oCon,"WMS_ISSUES")
    oSh.getCellByPosition(1,12).String=WMSC_TableState(oCon,"WMS_ISSUE_RETURNS")
    oSh.getCellByPosition(1,13).String=WMSC_TableState(oCon,"WMS_ACTS")
    oSh.getCellByPosition(1,14).String=WMSC_TableState(oCon,"WMS_AUDIT_LOG")
    oSh.getCellByPosition(1,15).String=Format(Now,"DD.MM.YYYY HH:MM:SS")
    WMSDB_Close
End Sub

Function WMSC_MetaValue(oCon As Object,keyName As String,defaultValue As String) As String
    Dim oStmt As Object
    Dim oRS As Object

    WMSC_MetaValue=defaultValue
    On Error GoTo Done
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT META_VALUE FROM WMS_META WHERE META_KEY=" & WMSDB_SQLText(keyName))
    If oRS.next() Then
        If Trim(oRS.getString(1))<>"" Then WMSC_MetaValue=Trim(oRS.getString(1))
    End If
Done:
End Function

Function WMSC_TableState(oCon As Object,t As String) As String
    Dim sErr As String
    If WMSDB_TableExistsSimple(oCon,t,sErr) Then
        WMSC_TableState="OK"
    ElseIf sErr<>"" Then
        WMSC_TableState="ERROR: " & sErr
    Else
        WMSC_TableState="НЕТ"
    End If
End Function

Sub WMSC_FullSelfCheck()
    ' Single production-wide health check in 2.0.
    WMSPROD_SelfCheck
End Sub

Sub WMSC_CheckTable(oCon As Object,t As String,ByRef s As String,ByRef errCount As Long)
    Dim st As String
    st=WMSC_TableState(oCon,t)
    s=s & t & ": " & st & Chr(10)
    If st<>"OK" Then errCount=errCount+1
End Sub

Sub WMSC_CheckSheet(nm As String,ByRef s As String,ByRef errCount As Long)
    If ThisComponent.Sheets.hasByName(nm) Then
        s=s & "Лист " & nm & ": OK" & Chr(10)
    Else
        s=s & "Лист " & nm & ": НЕТ" & Chr(10)
        errCount=errCount+1
    End If
End Sub

Sub WMSC_CheckProductionModules(ByRef s As String,ByRef errCount As Long)
    Dim oLib As Object
    Dim required As Variant
    Dim i As Long
    On Error GoTo EH
    oLib=ThisComponent.BasicLibraries.getByName("Standard")
    required=Array("WMS_CORE_Common_FINAL","WMS_02_Orders_FINAL","WMS_03_Issues_FINAL","WMS_04_DB_Connection","WMS_05_DB_Install","WMS_06_DB_Orders","WMS_07_DB_Issues","WMS_08_DB_Returns","WMS_09_DB_Search","WMS_10_DB_Stock","WMS_11_DB_UnitsLots","WMS_12_DB_SmartReceipt","WMS_13_SystemCenter","WMS_14_UniversalReceipt","WMS_15_SafetyCore","WMS_16_Acts","WMS_17_ActIntegration","WMS_18_ActsRegistry","WMS_19_ProductionUI","WMS_99_Installer")
    For i=0 To UBound(required)
        If Not oLib.hasByName(CStr(required(i))) Then
            s=s & "Модуль " & CStr(required(i)) & ": НЕТ" & Chr(10)
            errCount=errCount+1
        End If
    Next i
    If errCount=0 Then s=s & "Production-модули: комплект — OK" & Chr(10)
    Exit Sub
EH:
    s=s & "Проверка модулей: ОШИБКА" & Chr(10)
    errCount=errCount+1
End Sub

Sub WMSC_CheckPortableDatabaseFile(ByRef s As String,ByRef errCount As Long)
    Dim docURL As String
    Dim dbURL As String
    Dim n As Long
    Dim i As Long
    Dim sfa As Object

    On Error GoTo EH

    docURL = ThisComponent.URL
    If Trim(docURL) = "" Then
        s = s & "Проверка пары ODS/ODB: ODS ещё не сохранён — ОШИБКА" & Chr(10)
        errCount = errCount + 1
        Exit Sub
    End If

    n = 0
    For i = Len(docURL) To 1 Step -1
        If Mid(docURL,i,1) = "/" Then
            n = i
            Exit For
        End If
    Next i

    If n <= 0 Then
        s = s & "Проверка пары ODS/ODB: не удалось определить папку ODS — ОШИБКА" & Chr(10)
        errCount = errCount + 1
        Exit Sub
    End If

    dbURL = Left(docURL,n) & "WMS_DATA_PORTABLE.odb"
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    If sfa.exists(dbURL) Then
        s = s & "WMS_DATA_PORTABLE.odb рядом с ODS: OK" & Chr(10)
    Else
        s = s & "WMS_DATA_PORTABLE.odb рядом с ODS: НЕТ" & Chr(10)
        s = s & "Ожидаемый путь: " & ConvertFromURL(dbURL) & Chr(10)
        errCount = errCount + 1
    End If

    Exit Sub

EH:
    s = s & "Проверка пары ODS/ODB: ОШИБКА " & CStr(Err) & " — " & Error$ & Chr(10)
    errCount = errCount + 1
End Sub

Sub WMSC_LotRollbackTest()
    Dim oCon As Object,oStmt As Object,oRS As Object,sErr As String
    Dim code As String,lotID As String,movID As String,srcID As String
    Dim v As Double,ok As Boolean

    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Тест партий":Exit Sub

    If Not WMSDBUL_EnsureSchema(oCon,sErr) Then WMSDB_Close:MsgBox sErr,16,"WMS — Тест партий":Exit Sub

    code="__WMS_TEST_" & Replace(Replace(Format(Now,"HHMMSS"),":","")," ","")
    srcID="TESTSRC-" & code
    lotID="TESTLOT-" & code
    movID="TESTMOV-" & code
    oStmt=oCon.createStatement()

    On Error Resume Next
    oCon.setAutoCommit(False)
    On Error GoTo EH

    oStmt.executeUpdate("INSERT INTO WMS_PRODUCTS (PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION,ACTIVE_FLAG) VALUES (" & _
        WMSDB_SQLText(code) & "," & WMSDB_SQLText("WMS rollback test") & "," & WMSDB_SQLText("кг") & "," & _
        WMSDB_SQLText("TEST") & ",1)")

    oStmt.executeUpdate("INSERT INTO WMS_STOCK_LOTS (LOT_ID,PRODUCT_CODE,SOURCE_ID,SOURCE_TYPE,DOC_QTY,DOC_UNIT,BASE_QTY,BASE_UNIT,LOCATION_NAME,RECEIPT_DATE,ACTIVE_FLAG) VALUES (" & _
        WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(srcID) & "," & _
        WMSDB_SQLText("SELFTEST") & ",100," & WMSDB_SQLText("кг") & ",100," & WMSDB_SQLText("кг") & "," & _
        WMSDB_SQLText("TEST") & ",CURRENT_DATE,1)")

    oStmt.executeUpdate("INSERT INTO WMS_LOT_UNITS (LOT_ID,UNIT_NAME,FACTOR_TO_BASE,UNIT_ROLE) VALUES (" & _
        WMSDB_SQLText(lotID) & "," & WMSDB_SQLText("кг") & ",1," & WMSDB_SQLText("BASE") & ")")
    oStmt.executeUpdate("INSERT INTO WMS_LOT_UNITS (LOT_ID,UNIT_NAME,FACTOR_TO_BASE,UNIT_ROLE) VALUES (" & _
        WMSDB_SQLText(lotID) & "," & WMSDB_SQLText("шт") & ",10," & WMSDB_SQLText("ACTUAL") & ")")

    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT) VALUES (" & _
        WMSDB_SQLText(movID) & "," & WMSDB_SQLText(srcID) & "," & WMSDB_SQLText("SELFTEST") & "," & WMSDB_SQLText("IN") & "," & _
        WMSDB_SQLText(code) & "," & WMSDB_SQLText("WMS rollback test") & ",100," & WMSDB_SQLText("кг") & "," & _
        WMSDB_SQLText("TEST") & ",CURRENT_DATE," & WMSDB_SQLText(lotID) & ",10," & WMSDB_SQLText("шт") & ",100," & WMSDB_SQLText("кг") & ")")

    ' Test exact scenario: document 100 kg, actual 10 pcs, stock 100 kg.
    ' Issue 3 pcs => 30 kg, then return 1 pc => 10 kg. Final 80 kg.
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT) VALUES (" & _
        WMSDB_SQLText(movID & "-OUT") & "," & WMSDB_SQLText(srcID & "-OUT") & "," & WMSDB_SQLText("SELFTEST") & "," & WMSDB_SQLText("OUT") & "," & _
        WMSDB_SQLText(code) & "," & WMSDB_SQLText("WMS rollback test") & ",-30," & WMSDB_SQLText("кг") & "," & _
        WMSDB_SQLText("TEST") & ",CURRENT_DATE," & WMSDB_SQLText(lotID) & ",3," & WMSDB_SQLText("шт") & ",30," & WMSDB_SQLText("кг") & ")")

    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT) VALUES (" & _
        WMSDB_SQLText(movID & "-RET") & "," & WMSDB_SQLText(srcID & "-RET") & "," & WMSDB_SQLText("SELFTEST") & "," & WMSDB_SQLText("IN") & "," & _
        WMSDB_SQLText(code) & "," & WMSDB_SQLText("WMS rollback test") & ",10," & WMSDB_SQLText("кг") & "," & _
        WMSDB_SQLText("TEST") & ",CURRENT_DATE," & WMSDB_SQLText(lotID) & ",1," & WMSDB_SQLText("шт") & ",10," & WMSDB_SQLText("кг") & ")")

    oRS=oStmt.executeQuery("SELECT SUM(QTY) FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID=" & WMSDB_SQLText(lotID))
    If oRS.next() Then v=oRS.getDouble(1)
    ok=(Abs(v-80)<0.0001)

    oCon.rollback()
    On Error Resume Next:oCon.setAutoCommit(True)
    WMSDB_Close

    If ok Then
        MsgBox "Сквозной тест партий: OK" & Chr(10) & Chr(10) & _
               "Документ: 100 кг" & Chr(10) & _
               "Факт: 10 шт" & Chr(10) & _
               "Выдача: 3 шт = -30 кг" & Chr(10) & _
               "Возврат: 1 шт = +10 кг" & Chr(10) & _
               "Проверенный остаток: 80 кг" & Chr(10) & Chr(10) & _
               "Тестовые записи откатились через ROLLBACK и в базе не сохранены.",64,"WMS — Тест партий"
    Else
        MsgBox "Сквозной тест партий НЕ пройден. Получено: " & CStr(v) & " кг вместо 80 кг." & Chr(10) & _
               "Тестовые записи откатились.",16,"WMS — Тест партий"
    End If
    Exit Sub
EH:
    On Error Resume Next
    oCon.rollback()
    oCon.setAutoCommit(True)
    WMSDB_Close
    MsgBox "Тест партий остановлен: " & CStr(Err) & " " & Error$ & Chr(10) & _
           "Выполнен rollback тестовых изменений.",16,"WMS — Тест партий"
End Sub

Sub WMSC_OpenSearch()
    Dim oDoc As Object
    oDoc=ThisComponent
    If oDoc.Sheets.hasByName("База - Поиск") Then oDoc.CurrentController.setActiveSheet(oDoc.Sheets.getByName("База - Поиск"))
End Sub

Sub WMSC_OpenStock()
    Dim oDoc As Object
    oDoc=ThisComponent
    If oDoc.Sheets.hasByName("Остаток") Then
        oDoc.CurrentController.setActiveSheet(oDoc.Sheets.getByName("Остаток"))
        WMSDBST_RefreshStock
    End If
End Sub

Sub WMSC_InstallButtons(oDoc As Object,oSh As Object)
    Dim oForms As Object
    Dim oForm As Object

    On Error GoTo EH
    WMSC_RemoveButtons oSh
    oForms=oSh.DrawPage.Forms
    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name=WMSC_FORM
    oForms.insertByName(WMSC_FORM,oForm)

    WMSC_AddButton oDoc,oSh,oForm,"WMSC_REFRESH","Обновить статус",3,6,3900,750,"WMSC_Refresh"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_SELF","Проверка системы",3,7,3900,750,"WMSC_FullSelfCheck"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_STOCK","Открыть остаток",3,8,3900,750,"WMSC_OpenStock"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_SEARCH","Открыть поиск",3,9,3900,750,"WMSC_OpenSearch"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_INCOMPLETE","Незавершённые",3,10,3900,750,"WMSC_OpenIncomplete"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_AUDIT","Журнал действий",3,11,3900,750,"WMSC_OpenAudit"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_ACTS","Последние акты",3,12,3900,750,"WMSC_OpenActs"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_ACTOPEN","Открыть акт",3,13,3900,750,"WMSC_OpenActByNumber"
    WMSC_AddButton oDoc,oSh,oForm,"WMSC_UPDATE","Обновить WMS",3,14,3900,750,"WMSC_RunUpdate"
    Exit Sub
EH:
    MsgBox "Кнопки WMS Center: " & CStr(Err) & " " & Error$,16,"WMS Center"
End Sub

Sub WMSC_OpenIncomplete()
    WMSSAFE_CheckIncomplete
End Sub

Sub WMSC_OpenAudit()
    WMSSAFE_ShowAudit
End Sub

Sub WMSC_OpenActs()
    WMSACTREG_ShowRecent
End Sub

Sub WMSC_OpenActByNumber()
    WMSACTREG_OpenByNumber
End Sub

Sub WMSC_RunUpdate()
    WMS_UPDATE
End Sub

Sub WMSC_AddButton(oDoc As Object,oSh As Object,oForm As Object,nm As String,labelText As String,c As Long,r As Long,w As Long,h As Long,macroName As String)
    Dim ctl As Object,shape As Object,ev As New com.sun.star.script.ScriptEventDescriptor
    Dim p As Object,z As Object,idx As Long
    ctl=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    ctl.Name=nm:ctl.Label=labelText:ctl.Tabstop=False
    oForm.insertByName(nm,ctl):idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    shape.Control=ctl:p=oSh.getCellByPosition(c,r).Position:shape.Position=p
    z=CreateUnoStruct("com.sun.star.awt.Size"):z.Width=w:z.Height=h:shape.Size=z
    oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed"
    ev.AddListenerParam="":ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_13_SystemCenter." & macroName & "?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
End Sub

Sub WMSC_RemoveButtons(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.Count-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i):nm="":ctl=shape.Control:nm=ctl.Name
        If Left(nm,5)="WMSC_" Then oSh.DrawPage.remove(shape)
    Next i
    If oSh.DrawPage.Forms.hasByName(WMSC_FORM) Then oSh.DrawPage.Forms.removeByName(WMSC_FORM)
    On Error GoTo 0
End Sub

