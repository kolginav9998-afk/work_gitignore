Option Explicit

' ============================================================================
' WMS_06_DB_Orders.bas
' Stage 03: working Orders DB layer: order headers + lines + single/group/bulk sync.
'
' PURPOSE
' - Keeps the current Calc sheet as the working UI.
' - Writes one selected position, the current numbered order, or all ready positions.
' - Never deletes the Calc row.
' - Separates WMS_ORDERS (header) and WMS_ORDER_LINES (positions).
' - Upserts by _WMS_SourceID / _WMS_OrderGroupID, so repeated clicks do not duplicate data.
' - Does NOT mark the order as "Проведено" and does NOT create stock movement.
' - Uses separate DB-sync metadata, so it does not hijack the existing DIRTY/
'   movement workflow of WMS_02_Orders.
'
' DEPENDENCIES
' - WMS_02_Orders_FINAL 1.0.10-UPD-FIX (current project module)
' - WMS_04_DB_Connection 0.3.0-PORTABLE
' - WMS_05_DB_Install 0.3.0-PORTABLE (WMS_DATA_PORTABLE.odb installed)
' ============================================================================

Global Const WMSDBO_VERSION = "0.6.3-GUARDED-CONDUCT"
Global Const WMSDBO_MODULE = "DB_Orders"
Global Const WMSDBO_SHEET = "Заказы"
Global Const WMSDBO_TABLE = "WMS_ORDER_LINES"
Global Const WMSDBO_ORDERS_TABLE = "WMS_ORDERS"

Global Const WMSDBO_C_DBSTATE = "_WMS_DBState"
Global Const WMSDBO_C_DBHASH = "_WMS_DBHash"
Global Const WMSDBO_C_DBLASTSYNC = "_WMS_DBLastSync"
Global Const WMSDBO_C_DBERROR = "_WMS_DBError"

Global gWMSDBO_LastError As String
Global gWMSDBO_LastReport As String

' ---------------------------------------------------------------------------
' PUBLIC ENTRY POINTS
' ---------------------------------------------------------------------------

Sub WMSDBO_Install()
    Dim sReport As String
    If WMSDBO_InstallCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — БД Заказы"
    Else
        MsgBox sReport, 16, "WMS — БД Заказы"
    End If
End Sub

Sub WMSDBO_SyncSelectedPosition()
    Dim sReport As String
    If WMSDBO_SyncSelectedPositionCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — позиция записана в БД"
    Else
        MsgBox sReport, 16, "WMS — позиция НЕ записана"
    End If
End Sub

Sub WMSDBO_SyncCurrentOrder()
    Dim sReport As String
    If WMSDBO_SyncCurrentOrderCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — заказ синхронизирован"
    Else
        MsgBox sReport, 16, "WMS — заказ синхронизирован не полностью"
    End If
End Sub

Sub WMSDBO_SyncAllReadyPositions()
    Dim sReport As String
    If WMSDBO_SyncAllReadyPositionsCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — готовые позиции синхронизированы"
    Else
        MsgBox sReport, 48, "WMS — синхронизация завершена с замечаниями"
    End If
End Sub

Sub WMSDBO_ViewSelectedPositionDB()
    Dim sReport As String
    If WMSDBO_ViewSelectedPositionDBCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — запись Firebird"
    Else
        MsgBox sReport, 48, "WMS — запись Firebird"
    End If
End Sub

Sub WMSDBO_SelectedPositionStatus()
    Dim oDoc As Object, oSh As Object
    Dim r As Long, sid As String, sState As String, sWhen As String, sErr As String

    oDoc = ThisComponent
    If Not WMSDBO_GetSelectedOrderRow(oDoc, oSh, r, sErr) Then
        MsgBox sErr, 48, "WMS — статус БД"
        Exit Sub
    End If

    sid = WMSDBO_TechText(oSh, r, "_WMS_SourceID")
    sState = WMSDBO_TechText(oSh, r, WMSDBO_C_DBSTATE)
    sWhen = WMSDBO_TechText(oSh, r, WMSDBO_C_DBLASTSYNC)
    sErr = WMSDBO_TechText(oSh, r, WMSDBO_C_DBERROR)

    If sState = "" Then sState = "Ещё не синхронизирована"

    MsgBox "Строка Calc: " & CStr(r + 1) & Chr(10) & _
           "SourceID: " & sid & Chr(10) & _
           "DBState: " & sState & Chr(10) & _
           "Последняя синхронизация: " & sWhen & Chr(10) & _
           IIf(sErr <> "", "Ошибка: " & sErr, ""), 64, "WMS — статус БД"
End Sub

Sub WMSDBO_SelfCheck()
    Dim oCon As Object, sErr As String, sReport As String
    Dim bTable As Boolean, bOrders As Boolean

    On Error GoTo EH
    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox "Соединение: ERROR" & Chr(10) & sErr, 16, "WMS — БД Заказы SelfCheck"
        Exit Sub
    End If

    bTable = WMSDBO_TableExists(oCon, WMSDBO_TABLE, sErr)
    If sErr <> "" Then
        MsgBox "Соединение: OK" & Chr(10) & _
               "Проверка таблицы позиций: ERROR" & Chr(10) & sErr, 16, "WMS — БД Заказы SelfCheck"
        Exit Sub
    End If

    bOrders = WMSDBO_TableExists(oCon, WMSDBO_ORDERS_TABLE, sErr)
    If sErr <> "" Then
        MsgBox "Соединение: OK" & Chr(10) & _
               "Проверка таблицы заказов: ERROR" & Chr(10) & sErr, 16, "WMS — БД Заказы SelfCheck"
        Exit Sub
    End If

    sReport = "Соединение: OK" & Chr(10) & _
              "Firebird: OK" & Chr(10) & _
              WMSDBO_ORDERS_TABLE & ": " & IIf(bOrders, "OK", "НЕТ") & Chr(10) & _
              WMSDBO_TABLE & ": " & IIf(bTable, "OK", "НЕТ") & Chr(10) & _
              "Модуль: " & WMSDBO_VERSION

    If bTable And bOrders Then
        MsgBox sReport, 64, "WMS — БД Заказы SelfCheck"
    Else
        MsgBox sReport & Chr(10) & "Запустите WMSDBO_Install.", 48, "WMS — БД Заказы SelfCheck"
    End If
    Exit Sub
EH:
    MsgBox "SelfCheck: " & CStr(Err) & " " & Error$, 16, "WMS — БД Заказы SelfCheck"
End Sub

Function WMSDBO_CompileProbe() As String
    WMSDBO_CompileProbe = WMSDBO_VERSION & "|" & WMSDBO_TABLE
End Function

Function WMSDBO_LastErrorText() As String
    WMSDBO_LastErrorText = gWMSDBO_LastError
End Function

Function WMSDBO_LastReportText() As String
    WMSDBO_LastReportText = gWMSDBO_LastReport
End Function

' ---------------------------------------------------------------------------
' INSTALL: DB TABLE + 4 HIDDEN SYNC COLUMNS IN "ЗАКАЗЫ"
' ---------------------------------------------------------------------------

Function WMSDBO_InstallCore(oDoc As Object, ByRef sReport As String) As Boolean
    Dim oCon As Object, oSh As Object
    Dim sErr As String

    WMSDBO_InstallCore = False
    sReport = ""
    gWMSDBO_LastError = ""

    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBO_SHEET) Then
        sReport = "В книге нет листа 'Заказы'."
        Exit Function
    End If
    oSh = oDoc.Sheets.getByName(WMSDBO_SHEET)

    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then
        sReport = sErr
        Exit Function
    End If

    If Not WMSDBO_EnsureTable(oCon, sErr) Then
        sReport = "Не удалось создать/проверить таблицу " & WMSDBO_TABLE & ":" & Chr(10) & sErr
        Exit Function
    End If

    If Not WMSDBO_EnsureOrdersTable(oCon, sErr) Then
        sReport = "Не удалось создать/проверить таблицу " & WMSDBO_ORDERS_TABLE & ":" & Chr(10) & sErr
        Exit Function
    End If

    ' Portable embedded Firebird: COMMIT + Base document store.
    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        sReport = "Не удалось зафиксировать CREATE TABLE в переносимой Base:" & Chr(10) & sErr
        Exit Function
    End If

    ' Hard persistence test: fully close SDBC, reconnect and verify table again.
    WMSDB_Close
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then
        sReport = "Таблица создана, но повторное открытие переносимой Base не прошло:" & Chr(10) & sErr
        Exit Function
    End If
    If Not WMSDBO_TableExists(oCon, WMSDBO_TABLE, sErr) Then
        If sErr = "" Then sErr = "После полного переподключения таблица позиций не найдена."
        sReport = "Firebird не подтвердил постоянное хранение " & WMSDBO_TABLE & " после reconnect:" & Chr(10) & sErr
        Exit Function
    End If
    If Not WMSDBO_TableExists(oCon, WMSDBO_ORDERS_TABLE, sErr) Then
        If sErr = "" Then sErr = "После полного переподключения таблица заказов не найдена."
        sReport = "Firebird не подтвердил постоянное хранение " & WMSDBO_ORDERS_TABLE & " после reconnect:" & Chr(10) & sErr
        Exit Function
    End If

    If Not WMSDBO_EnsureSyncColumns(oSh, sErr) Then
        sReport = "База готова, но не удалось добавить служебные DB-поля на лист 'Заказы':" & Chr(10) & sErr
        Exit Function
    End If

    sReport = "DB-модуль Заказов 0.3.1-PORTABLE установлен." & Chr(10) & _
              "Firebird: " & WMSDBO_ORDERS_TABLE & " + " & WMSDBO_TABLE & Chr(10) & _
              "Один заказ хранится отдельно от его позиций." & Chr(10) & _
              "Каждую позицию по-прежнему можно отправлять отдельно." & Chr(10) & _
              "Данные Calc не удалялись и автоматически массово не переносились."

    gWMSDBO_LastReport = sReport
    WMSDBO_InstallCore = True
    Exit Function

EH:
    sReport = "WMSDBO_Install: " & CStr(Err) & " " & Error$
    gWMSDBO_LastError = sReport
End Function

Function WMSDBO_EnsureTable(oCon As Object, ByRef sErr As String) As Boolean
    Dim sSQL As String, sVerifyErr As String

    WMSDBO_EnsureTable = False
    sErr = ""
    On Error GoTo EH

    If WMSDBO_TableExists(oCon, WMSDBO_TABLE, sVerifyErr) Then
        If Not WMSDBO_EnsureReceiptV2Columns(oCon,sErr) Then Exit Function
        WMSDBO_EnsureTable = True
        Exit Function
    End If
    If sVerifyErr <> "" Then
        sErr = "Не удалось проверить метаданные Firebird: " & sVerifyErr
        Exit Function
    End If

    ' Keep VARCHAR sizes deliberately conservative. Firebird stores VARCHAR
    ' with the database character set and an oversized declared row can make
    ' CREATE TABLE fail even when the real strings are short.
    sSQL = "CREATE TABLE " & WMSDBO_TABLE & " ("
    sSQL = sSQL & "SOURCE_ID VARCHAR(96) NOT NULL PRIMARY KEY,"
    sSQL = sSQL & "ORDER_GROUP_ID VARCHAR(96),"
    sSQL = sSQL & "POSITION_NO INTEGER,"
    sSQL = sSQL & "PRODUCT_NAME VARCHAR(255) NOT NULL,"
    sSQL = sSQL & "DOC_NO VARCHAR(96),"
    sSQL = sSQL & "INVOICE_NO VARCHAR(96),"
    sSQL = sSQL & "SUPPLIER_CODE VARCHAR(96),"
    sSQL = sSQL & "SUPPLIER_ARTICLE VARCHAR(120),"
    sSQL = sSQL & "FACT_QTY DECIMAL(18,6),"
    sSQL = sSQL & "ORDER_QTY DECIMAL(18,6),"
    sSQL = sSQL & "UNIT_NAME VARCHAR(48),"
    sSQL = sSQL & "PRICE DECIMAL(18,4),"
    sSQL = sSQL & "TOTAL_AMOUNT DECIMAL(18,2),"
    sSQL = sSQL & "SUPPLIER_NAME VARCHAR(180),"
    sSQL = sSQL & "RECEIPT_DATE DATE,"
    sSQL = sSQL & "DOC_DATE DATE,"
    sSQL = sSQL & "ORDER_DATE DATE,"
    sSQL = sSQL & "BUYER VARCHAR(120),"
    sSQL = sSQL & "STATUS_NAME VARCHAR(80),"
    sSQL = sSQL & "CATEGORY_NAME VARCHAR(120),"
    sSQL = sSQL & "CONTROL_TEXT VARCHAR(320),"
    sSQL = sSQL & "LOCATION_NAME VARCHAR(120),"
    sSQL = sSQL & "COMMENT_TEXT VARCHAR(640),"
    sSQL = sSQL & "PRODUCT_CODE VARCHAR(96),"
    sSQL = sSQL & "SELLER_NAME VARCHAR(180),"
    sSQL = sSQL & "ROW_STATE VARCHAR(48),"
    sSQL = sSQL & "ROW_MODE VARCHAR(48),"
    sSQL = sSQL & "BUSINESS_HASH VARCHAR(255),"
    sSQL = sSQL & "LAST_CALC_ROW INTEGER,"
    sSQL = sSQL & "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,"
    sSQL = sSQL & "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)"

    If Not WMSDBO_ExecDDL(oCon, sSQL, sErr) Then Exit Function

    On Error Resume Next
    oCon.commit()
    On Error GoTo EH

    ' Mandatory post-condition: never report install success unless Firebird
    ' itself exposes the table through metadata after COMMIT.
    sVerifyErr = ""
    If Not WMSDBO_TableExists(oCon, WMSDBO_TABLE, sVerifyErr) Then
        If sVerifyErr <> "" Then
            sErr = "CREATE TABLE выполнен, но контроль метаданных завершился ошибкой: " & sVerifyErr
        Else
            sErr = "CREATE TABLE не подтверждён: Firebird не видит " & WMSDBO_TABLE & " после COMMIT."
        End If
        Exit Function
    End If

    If Not WMSDBO_EnsureReceiptV2Columns(oCon,sErr) Then Exit Function
    WMSDBO_EnsureTable = True
    Exit Function

EH:
    sErr = CStr(Err) & " " & Error$
    On Error Resume Next
    oCon.rollback()
End Function

Function WMSDBO_EnsureReceiptV2Columns(oCon As Object,ByRef sErr As String) As Boolean
    Dim aCols,aTypes,i As Long
    WMSDBO_EnsureReceiptV2Columns=False:sErr=""
    aCols=Array("RECEIPT_EVENT_ID","RECEIPT_MODE","RECEIPT_SOURCE","RECEIVED_FROM","EXTERNAL_ORDER_NO","UPD_NO","DESTINATION_NAME","EXPECTED_RECEIPT_DATE")
    aTypes=Array("VARCHAR(128)","VARCHAR(48)","VARCHAR(120)","VARCHAR(180)","VARCHAR(120)","VARCHAR(120)","VARCHAR(180)","DATE")
    For i=0 To UBound(aCols)
        If Not WMSDBO_ColumnExistsV2(oCon,WMSDBO_TABLE,CStr(aCols(i)),sErr) Then
            If sErr<>"" Then Exit Function
            If Not WMSDBO_ExecDDL(oCon,"ALTER TABLE " & WMSDBO_TABLE & " ADD " & CStr(aCols(i)) & " " & CStr(aTypes(i)),sErr) Then Exit Function
        End If
    Next i
    On Error Resume Next:oCon.commit():On Error GoTo 0
    If Not WMSDB_TableExistsSimple(oCon,"WMS_ORDER_SNAPSHOT",sErr) Then
        If sErr<>"" Then Exit Function
        If Not WMSDBO_ExecDDL(oCon,"CREATE TABLE WMS_ORDER_SNAPSHOT (SOURCE_ID VARCHAR(96) NOT NULL,COL_INDEX INTEGER NOT NULL,HEADER_NAME VARCHAR(255),VALUE_TEXT VARCHAR(8000),PRIMARY KEY(SOURCE_ID,COL_INDEX))",sErr) Then Exit Function
    End If
    WMSDBO_EnsureReceiptV2Columns=True
End Function

Function WMSDBO_ColumnExistsV2(oCon As Object,t As String,c As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBO_ColumnExistsV2=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE TRIM(RDB$RELATION_NAME)=" & WMSDB_SQLText(UCase(t)) & _
        " AND TRIM(RDB$FIELD_NAME)=" & WMSDB_SQLText(UCase(c))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBO_ColumnExistsV2=(oRS.getInt(1)>0)
    Exit Function
EH:
    sErr="Проверка поля " & c & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBO_EnsureOrdersTable(oCon As Object, ByRef sErr As String) As Boolean
    Dim sSQL As String, sVerifyErr As String
    WMSDBO_EnsureOrdersTable = False
    sErr = ""
    On Error GoTo EH

    If WMSDBO_TableExists(oCon, WMSDBO_ORDERS_TABLE, sVerifyErr) Then
        WMSDBO_EnsureOrdersTable = True
        Exit Function
    End If
    If sVerifyErr <> "" Then
        sErr = sVerifyErr
        Exit Function
    End If

    sSQL = "CREATE TABLE " & WMSDBO_ORDERS_TABLE & " ("
    sSQL = sSQL & "ORDER_GROUP_ID VARCHAR(96) NOT NULL PRIMARY KEY,"
    sSQL = sSQL & "INVOICE_NO VARCHAR(96),"
    sSQL = sSQL & "SUPPLIER_NAME VARCHAR(180),"
    sSQL = sSQL & "ORDER_DATE DATE,"
    sSQL = sSQL & "BUYER VARCHAR(120),"
    sSQL = sSQL & "STATUS_NAME VARCHAR(80),"
    sSQL = sSQL & "FIRST_CALC_ROW INTEGER,"
    sSQL = sSQL & "LAST_CALC_ROW INTEGER,"
    sSQL = sSQL & "CALC_LINE_COUNT INTEGER,"
    sSQL = sSQL & "DB_LINE_COUNT INTEGER,"
    sSQL = sSQL & "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,"
    sSQL = sSQL & "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)"

    If Not WMSDBO_ExecDDL(oCon, sSQL, sErr) Then Exit Function
    On Error Resume Next
    oCon.commit()
    On Error GoTo EH

    If Not WMSDBO_TableExists(oCon, WMSDBO_ORDERS_TABLE, sVerifyErr) Then
        If sVerifyErr = "" Then sVerifyErr = "Firebird не видит таблицу после CREATE TABLE."
        sErr = sVerifyErr
        Exit Function
    End If

    WMSDBO_EnsureOrdersTable = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDBO_ExecDDL(oCon As Object, sSQL As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, bResult As Boolean
    WMSDBO_ExecDDL = False
    sErr = ""
    On Error GoTo EH

    oStmt = oCon.createStatement()
    ' XStatement.execute is appropriate for DDL. Its Boolean return only says
    ' whether a ResultSet exists; CREATE TABLE normally returns False.
    bResult = oStmt.execute(sSQL)
    WMSDBO_ExecDDL = True
    Exit Function
EH:
    sErr = Error$
End Function

Function WMSDBO_TableExists(oCon As Object, sTable As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sSQL As String
    WMSDBO_TableExists = False
    sErr = ""
    On Error GoTo EH

    ' Do not use COUNT(*) here. Some LibreOffice/Firebird SDBC builds have
    ' shown unreliable scalar fetch behaviour. Fetch the relation name itself.
    sSQL = "SELECT RDB$RELATION_NAME FROM RDB$RELATIONS WHERE " & _
           "RDB$RELATION_NAME=" & WMSDB_SQLText(UCase(Trim(sTable))) & _
           " AND COALESCE(RDB$SYSTEM_FLAG,0)=0"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If oRS.next() Then WMSDBO_TableExists = (Trim(oRS.getString(1)) <> "")
    Exit Function
EH:
    sErr = Error$
End Function

Function WMSDBO_EnsureSyncColumns(oSh As Object, ByRef sErr As String) As Boolean
    Dim a As Variant, i As Long, c As Long
    Dim oCols As Object

    WMSDBO_EnsureSyncColumns = False
    sErr = ""
    On Error GoTo EH

    a = Array(WMSDBO_C_DBSTATE, WMSDBO_C_DBHASH, WMSDBO_C_DBLASTSYNC, WMSDBO_C_DBERROR)
    oCols = oSh.Columns

    For i = LBound(a) To UBound(a)
        c = WMSDBO_FindHeader(oSh, CStr(a(i)))
        If c < 0 Then
            c = WMSDBO_LastHeaderCol(oSh) + 1
            oSh.getCellByPosition(c, 0).String = CStr(a(i))
        End If
        On Error Resume Next
        oCols.getByIndex(c).IsVisible = False
        On Error GoTo EH
    Next i

    WMSDBO_EnsureSyncColumns = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

' ---------------------------------------------------------------------------
' SYNC ONE SELECTED POSITION
' ---------------------------------------------------------------------------

Function WMSDBO_SyncSelectedPositionCore(oDoc As Object, ByRef sReport As String) As Boolean
    Dim oSh As Object, oCon As Object
    Dim r As Long, sErr As String, sid As String, sSeverity As String
    Dim sHash As String, nExists As Long, nChanged As Long

    WMSDBO_SyncSelectedPositionCore = False
    sReport = ""
    gWMSDBO_LastError = ""

    On Error GoTo EH

    If Not WMSDBO_GetSelectedOrderRow(oDoc, oSh, r, sErr) Then
        sReport = sErr
        Exit Function
    End If

    If Not WMSDBO_RowIsRealOrderPosition(oSh, r, sErr) Then
        sReport = sErr
        Exit Function
    End If

    ' Reuse the current Orders controller to normalize the row and create/heal
    ' SourceID + group metadata, but do not enqueue or mark it dirty here.
    Orders_ProcessRow oDoc, oSh, r, False, False, True, True

    sid = WMSDBO_TechText(oSh, r, "_WMS_SourceID")
    If sid = "" Then
        sReport = "Не удалось получить _WMS_SourceID для выбранной позиции."
        Exit Function
    End If

    sSeverity = UCase(WMSDBO_TechText(oSh, r, "_WMS_Severity"))
    If sSeverity = "CRITICAL" Then
        sReport = "Позиция не отправлена: в ней есть критическая ошибка." & Chr(10) & _
                  WMSDBO_TechText(oSh, r, "Контроль")
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sReport
        Exit Function
    End If

    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then
        sReport = sErr
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        Exit Function
    End If

    If Not WMSDBO_EnsureTable(oCon, sErr) Then
        sReport = sErr
        Exit Function
    End If
    If Not WMSDBO_EnsureOrdersTable(oCon, sErr) Then
        sReport = sErr
        Exit Function
    End If

    If WMSDBO_FindHeader(oSh, WMSDBO_C_DBSTATE) < 0 Then
        If Not WMSDBO_EnsureSyncColumns(oSh, sErr) Then
            sReport = sErr
            Exit Function
        End If
    End If

    sHash = WMSDBO_ExtendedFingerprint(oSh, r)
    nExists = WMSDBO_RecordExists(oCon, sid, sErr)
    If sErr <> "" Then
        sReport = "Не удалось проверить SourceID в БД: " & sErr
        Exit Function
    End If

    On Error Resume Next
    oCon.setAutoCommit(False)
    On Error GoTo EH

    If nExists > 0 Then
        nChanged = WMSDBO_UpdateRow(oCon, oSh, r, sid, sHash, sErr)
    Else
        nChanged = WMSDBO_InsertRow(oCon, oSh, r, sid, sHash, sErr)
    End If

    If sErr <> "" Or nChanged < 0 Then
        On Error Resume Next
        oCon.rollback()
        On Error GoTo EH
        If sErr = "" Then sErr = "Firebird не подтвердил запись."
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "Позиция не записана в БД:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If

    oCon.commit()
    On Error Resume Next
    oCon.setAutoCommit(True)
    On Error GoTo EH

    ' Mandatory read-back. No success message unless the committed record is
    ' physically readable from Firebird by the same SOURCE_ID.
    sErr = ""
    If Not WMSDBO_VerifyRecord(oCon, sid, sErr) Then
        If sErr = "" Then sErr = "После COMMIT запись не найдена по SOURCE_ID=" & sid
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "Firebird не подтвердил сохранение позиции:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If

    ' External FDB must survive a complete disconnect/reconnect before success.
    sErr = ""
    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "Не удалось сохранить изменения в переносимой Base:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If

    WMSDB_Close
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "После сохранения не удалось повторно открыть переносимую Base:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If
    If Not WMSDBO_VerifyRecord(oCon, sid, sErr) Then
        If sErr = "" Then sErr = "После полного переподключения запись не найдена по SOURCE_ID=" & sid
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "Firebird не подтвердил постоянное хранение позиции после reconnect:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If

    If Not WMSDBO_UpsertOrderHeader(oCon, oSh, r, sErr) Then
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        sReport = "Позиция сохранена, но заголовок заказа не обновлён:" & Chr(10) & sErr
        gWMSDBO_LastError = sReport
        Exit Function
    End If

    WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "SYNCED"
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBHASH, sHash
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBLASTSYNC, WMSDBO_Stamp(Now)
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, ""

    If nExists > 0 Then
        sReport = "Позиция обновлена в Base/Firebird." & Chr(10)
    Else
        sReport = "Позиция впервые записана в Base/Firebird." & Chr(10)
    End If
    sReport = sReport & _
              "Строка Calc: " & CStr(r + 1) & Chr(10) & _
              "SourceID: " & sid & Chr(10) & _
              "Calc-данные остались на месте."

    gWMSDBO_LastReport = sReport
    WMSDBO_SyncSelectedPositionCore = True
    Exit Function

EH:
    On Error Resume Next
    If Not (IsNull(oCon) Or IsEmpty(oCon)) Then
        oCon.rollback()
        oCon.setAutoCommit(True)
    End If
    On Error GoTo 0
    sReport = "WMSDBO_SyncSelectedPosition: " & CStr(Err) & " " & Error$
    gWMSDBO_LastError = sReport
    If Not (IsNull(oSh) Or IsEmpty(oSh)) And r > 0 Then
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sReport
    End If
End Function

Function WMSDBO_RecordExists(oCon As Object, sid As String, ByRef sErr As String) As Long
    Dim oStmt As Object, oRS As Object, sSQL As String
    WMSDBO_RecordExists = 0
    sErr = ""
    On Error GoTo EH

    ' Fetch SOURCE_ID directly instead of COUNT(*) for maximum SDBC stability.
    sSQL = "SELECT SOURCE_ID FROM " & WMSDBO_TABLE & _
           " WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If oRS.next() Then
        If Trim(oRS.getString(1)) = Trim(sid) Then WMSDBO_RecordExists = 1
    End If
    Exit Function
EH:
    sErr = Error$
    WMSDBO_RecordExists = -1
End Function

Function WMSDBO_VerifyRecord(oCon As Object, sid As String, ByRef sErr As String) As Boolean
    WMSDBO_VerifyRecord = (WMSDBO_RecordExists(oCon, sid, sErr) = 1)
End Function

Function WMSDBO_InsertRow(oCon As Object, oSh As Object, r As Long, sid As String, sHash As String, ByRef sErr As String) As Long
    Dim sSQL As String
    sErr = ""

    sSQL = "INSERT INTO " & WMSDBO_TABLE & " (" & _
           "SOURCE_ID,ORDER_GROUP_ID,POSITION_NO,PRODUCT_NAME,DOC_NO,INVOICE_NO," & _
           "SUPPLIER_CODE,SUPPLIER_ARTICLE,FACT_QTY,ORDER_QTY,UNIT_NAME,PRICE,TOTAL_AMOUNT," & _
           "SUPPLIER_NAME,RECEIPT_DATE,DOC_DATE,ORDER_DATE,BUYER,STATUS_NAME,CATEGORY_NAME," & _
           "CONTROL_TEXT,LOCATION_NAME,COMMENT_TEXT,PRODUCT_CODE,SELLER_NAME,ROW_STATE,ROW_MODE," & _
           "BUSINESS_HASH,LAST_CALC_ROW,RECEIPT_EVENT_ID,RECEIPT_MODE,RECEIPT_SOURCE,RECEIVED_FROM,EXTERNAL_ORDER_NO,UPD_NO,DESTINATION_NAME,EXPECTED_RECEIPT_DATE,CREATED_AT,UPDATED_AT) VALUES (" & _
           WMSDBO_RowValuesSQL(oSh, r, sid, sHash) & ",CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)"

    WMSDBO_InsertRow = WMSDB_ExecuteUpdate(oCon, sSQL, sErr)
End Function

Function WMSDBO_UpdateRow(oCon As Object, oSh As Object, r As Long, sid As String, sHash As String, ByRef sErr As String) As Long
    Dim sSQL As String
    sErr = ""

    sSQL = "UPDATE " & WMSDBO_TABLE & " SET " & _
           WMSDBO_RowSetSQL(oSh, r, sHash) & _
           ", UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sid)

    WMSDBO_UpdateRow = WMSDB_ExecuteUpdate(oCon, sSQL, sErr)
End Function

Function WMSDBO_RowValuesSQL(oSh As Object, r As Long, sid As String, sHash As String) As String
    Dim s As String
    s = WMSDB_SQLText(sid)
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_OrderGroupID"))
    s = s & "," & WMSDBO_DBPositionSQL(oSh, r)
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(1, r).String))
    s = s & "," & WMSDB_SQLText(Orders_EffectiveDocNo(oSh, r))
    s = s & "," & WMSDB_SQLText(Orders_EffectiveInvoice(oSh, r))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(4, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(5, r).String))
    s = s & "," & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(6, r))
    s = s & "," & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(7, r))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(8, r).String))
    s = s & "," & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(9, r))
    s = s & "," & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(10, r))
    s = s & "," & WMSDB_SQLText(Orders_EffectiveSupplier(oSh, r))
    s = s & "," & WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(12, r), True))
    s = s & "," & WMSDBO_SQLDate(Orders_EffectiveDocDate(oSh, r))
    s = s & "," & WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(14, r), True))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(15, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(16, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(17, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(18, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(19, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(20, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(21, r).String))
    s = s & "," & WMSDB_SQLText(Trim(oSh.getCellByPosition(22, r).String))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_State"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_Mode"))
    s = s & "," & WMSDB_SQLText(sHash)
    s = s & "," & CStr(r + 1)
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptEventID"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptMode"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptSource"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceivedFrom"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ExternalOrderNo"))
    s = s & "," & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_UPDNo"))
    s=s & "," & WMSDBO_ExtraSQL(oSh,r,"Назначение",False)
    s=s & "," & WMSDBO_ExtraSQL(oSh,r,"Ожидаемая дата поступления",True)
    WMSDBO_RowValuesSQL = s
End Function

Function WMSDBO_RowSetSQL(oSh As Object, r As Long, sHash As String) As String
    Dim s As String
    s = "ORDER_GROUP_ID=" & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_OrderGroupID"))
    s = s & ", POSITION_NO=" & WMSDBO_DBPositionSQL(oSh, r)
    s = s & ", PRODUCT_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(1, r).String))
    s = s & ", DOC_NO=" & WMSDB_SQLText(Orders_EffectiveDocNo(oSh, r))
    s = s & ", INVOICE_NO=" & WMSDB_SQLText(Orders_EffectiveInvoice(oSh, r))
    s = s & ", SUPPLIER_CODE=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(4, r).String))
    s = s & ", SUPPLIER_ARTICLE=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(5, r).String))
    s = s & ", FACT_QTY=" & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(6, r))
    s = s & ", ORDER_QTY=" & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(7, r))
    s = s & ", UNIT_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(8, r).String))
    s = s & ", PRICE=" & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(9, r))
    s = s & ", TOTAL_AMOUNT=" & WMSDBO_SQLNumberFromCell(oSh.getCellByPosition(10, r))
    s = s & ", SUPPLIER_NAME=" & WMSDB_SQLText(Orders_EffectiveSupplier(oSh, r))
    s = s & ", RECEIPT_DATE=" & WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(12, r), True))
    s = s & ", DOC_DATE=" & WMSDBO_SQLDate(Orders_EffectiveDocDate(oSh, r))
    s = s & ", ORDER_DATE=" & WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(14, r), True))
    s = s & ", BUYER=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(15, r).String))
    s = s & ", STATUS_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(16, r).String))
    s = s & ", CATEGORY_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(17, r).String))
    s = s & ", CONTROL_TEXT=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(18, r).String))
    s = s & ", LOCATION_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(19, r).String))
    s = s & ", COMMENT_TEXT=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(20, r).String))
    s = s & ", PRODUCT_CODE=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(21, r).String))
    s = s & ", SELLER_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(22, r).String))
    s = s & ", ROW_STATE=" & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_State"))
    s = s & ", ROW_MODE=" & WMSDB_SQLText(WMSDBO_TechText(oSh, r, "_WMS_Mode"))
    s = s & ", BUSINESS_HASH=" & WMSDB_SQLText(sHash)
    s = s & ", LAST_CALC_ROW=" & CStr(r + 1)
    s = s & ", RECEIPT_EVENT_ID=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptEventID"))
    s = s & ", RECEIPT_MODE=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptMode"))
    s = s & ", RECEIPT_SOURCE=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceiptSource"))
    s = s & ", RECEIVED_FROM=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ReceivedFrom"))
    s = s & ", EXTERNAL_ORDER_NO=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_ExternalOrderNo"))
    s = s & ", UPD_NO=" & WMSDB_SQLText(WMSDBO_TechText(oSh,r,"_WMS_UPDNo"))
    ' Old workbooks without the new headers must not erase stored values.
    If WMSDBO_FindHeader(oSh,"Назначение")>=0 Then s=s & ", DESTINATION_NAME=" & WMSDBO_ExtraSQL(oSh,r,"Назначение",False)
    If WMSDBO_FindHeader(oSh,"Ожидаемая дата поступления")>=0 Then s=s & ", EXPECTED_RECEIPT_DATE=" & WMSDBO_ExtraSQL(oSh,r,"Ожидаемая дата поступления",True)
    WMSDBO_RowSetSQL = s
End Function

Function WMSDBO_UpsertOrderHeader(oCon As Object, oSh As Object, r As Long, ByRef sErr As String) As Boolean
    Dim firstRow As Long, lastRow As Long, i As Long, nCalc As Long, nDB As Long
    Dim gid As String, sSQL As String, nExists As Long
    WMSDBO_UpsertOrderHeader = False
    sErr = ""
    On Error GoTo EH

    gid = WMSDBO_TechText(oSh, r, "_WMS_OrderGroupID")
    If gid = "" Then
        Orders_EnsureRowGroupContext oSh, r
        gid = WMSDBO_TechText(oSh, r, "_WMS_OrderGroupID")
    End If
    If gid = "" Then
        sErr = "Не удалось определить _WMS_OrderGroupID."
        Exit Function
    End If

    If Not Orders_GetNumberedGroupBounds(oSh, r, firstRow, lastRow) Then
        firstRow = r
        lastRow = r
    End If

    For i = firstRow To lastRow
        If Trim(oSh.getCellByPosition(1, i).String) <> "" Then nCalc = nCalc + 1
    Next i
    nDB = WMSDBO_DBLineCountForGroup(oCon, gid, sErr)
    If sErr <> "" Then Exit Function

    nExists = WMSDBO_OrderHeaderExists(oCon, gid, sErr)
    If sErr <> "" Then Exit Function

    If nExists = 0 Then
        sSQL = "INSERT INTO " & WMSDBO_ORDERS_TABLE & " (ORDER_GROUP_ID,INVOICE_NO,SUPPLIER_NAME,ORDER_DATE,BUYER,STATUS_NAME,FIRST_CALC_ROW,LAST_CALC_ROW,CALC_LINE_COUNT,DB_LINE_COUNT,CREATED_AT,UPDATED_AT) VALUES (" & _
               WMSDB_SQLText(gid) & "," & _
               WMSDB_SQLText(Orders_EffectiveInvoice(oSh, r)) & "," & _
               WMSDB_SQLText(Orders_EffectiveSupplier(oSh, r)) & "," & _
               WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(14, r), True)) & "," & _
               WMSDB_SQLText(Trim(oSh.getCellByPosition(15, r).String)) & "," & _
               WMSDB_SQLText(Trim(oSh.getCellByPosition(16, r).String)) & "," & _
               CStr(firstRow + 1) & "," & CStr(lastRow + 1) & "," & CStr(nCalc) & "," & CStr(nDB) & ",CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)"
    Else
        sSQL = "UPDATE " & WMSDBO_ORDERS_TABLE & " SET " & _
               "INVOICE_NO=" & WMSDB_SQLText(Orders_EffectiveInvoice(oSh, r)) & _
               ",SUPPLIER_NAME=" & WMSDB_SQLText(Orders_EffectiveSupplier(oSh, r)) & _
               ",ORDER_DATE=" & WMSDBO_SQLDate(Orders_CellDate(oSh.getCellByPosition(14, r), True)) & _
               ",BUYER=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(15, r).String)) & _
               ",STATUS_NAME=" & WMSDB_SQLText(Trim(oSh.getCellByPosition(16, r).String)) & _
               ",FIRST_CALC_ROW=" & CStr(firstRow + 1) & _
               ",LAST_CALC_ROW=" & CStr(lastRow + 1) & _
               ",CALC_LINE_COUNT=" & CStr(nCalc) & _
               ",DB_LINE_COUNT=" & CStr(nDB) & _
               ",UPDATED_AT=CURRENT_TIMESTAMP WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid)
    End If

    If WMSDB_ExecuteUpdate(oCon, sSQL, sErr) < 0 Or sErr <> "" Then Exit Function
    On Error Resume Next
    oCon.commit()
    On Error GoTo EH
    WMSDBO_UpsertOrderHeader = (WMSDBO_OrderHeaderExists(oCon, gid, sErr) = 1)
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDBO_OrderHeaderExists(oCon As Object, gid As String, ByRef sErr As String) As Long
    Dim oStmt As Object, oRS As Object, sSQL As String
    WMSDBO_OrderHeaderExists = 0
    sErr = ""
    On Error GoTo EH
    sSQL = "SELECT ORDER_GROUP_ID FROM " & WMSDBO_ORDERS_TABLE & " WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid)
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If oRS.next() Then WMSDBO_OrderHeaderExists = 1
    Exit Function
EH:
    sErr = Error$
    WMSDBO_OrderHeaderExists = -1
End Function

Function WMSDBO_DBLineCountForGroup(oCon As Object, gid As String, ByRef sErr As String) As Long
    Dim oStmt As Object, oRS As Object, sSQL As String, n As Long
    WMSDBO_DBLineCountForGroup = 0
    sErr = ""
    On Error GoTo EH
    sSQL = "SELECT SOURCE_ID FROM " & WMSDBO_TABLE & " WHERE ORDER_GROUP_ID=" & WMSDB_SQLText(gid)
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    Do While oRS.next()
        n = n + 1
    Loop
    WMSDBO_DBLineCountForGroup = n
    Exit Function
EH:
    sErr = Error$
    WMSDBO_DBLineCountForGroup = -1
End Function

Function WMSDBO_SyncSpecificRow(oDoc As Object, oSh As Object, r As Long, oCon As Object, ByRef sAction As String, ByRef sErr As String) As Boolean
    Dim sid As String, sHash As String, sSeverity As String, nExists As Long, nChanged As Long
    WMSDBO_SyncSpecificRow = False
    sAction = ""
    sErr = ""
    On Error GoTo EH

    If Not WMSDBO_RowIsRealOrderPosition(oSh, r, sErr) Then Exit Function
    Orders_ProcessRow oDoc, oSh, r, False, False, True, True
    sid = WMSDBO_TechText(oSh, r, "_WMS_SourceID")
    If sid = "" Then
        sErr = "Нет _WMS_SourceID."
        Exit Function
    End If
    sSeverity = UCase(WMSDBO_TechText(oSh, r, "_WMS_Severity"))
    If sSeverity = "CRITICAL" Then
        sErr = "Критическая ошибка: " & WMSDBO_TechText(oSh, r, "Контроль")
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "ERROR"
        WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, sErr
        Exit Function
    End If

    sHash = WMSDBO_ExtendedFingerprint(oSh, r)
    nExists = WMSDBO_RecordExists(oCon, sid, sErr)
    If sErr <> "" Then Exit Function

    If nExists > 0 And WMSDBO_TechText(oSh, r, WMSDBO_C_DBHASH) = sHash And UCase(WMSDBO_TechText(oSh, r, WMSDBO_C_DBSTATE)) = "SYNCED" Then
        sAction = "SKIP"
    ElseIf nExists > 0 Then
        nChanged = WMSDBO_UpdateRow(oCon, oSh, r, sid, sHash, sErr)
        If sErr <> "" Or nChanged < 0 Then Exit Function
        sAction = "UPDATE"
    Else
        nChanged = WMSDBO_InsertRow(oCon, oSh, r, sid, sHash, sErr)
        If sErr <> "" Or nChanged < 0 Then Exit Function
        sAction = "INSERT"
    End If

    On Error Resume Next
    oCon.commit()
    On Error GoTo EH
    If Not WMSDBO_VerifyRecord(oCon, sid, sErr) Then Exit Function
    If Not WMSDBO_UpsertOrderHeader(oCon, oSh, r, sErr) Then Exit Function

    WMSDBO_SetTech oSh, r, WMSDBO_C_DBSTATE, "SYNCED"
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBHASH, sHash
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBLASTSYNC, WMSDBO_Stamp(Now)
    WMSDBO_SetTech oSh, r, WMSDBO_C_DBERROR, ""
    WMSDBO_SyncSpecificRow = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDBO_SyncCurrentOrderCore(oDoc As Object, ByRef sReport As String) As Boolean
    Dim oSh As Object, oCon As Object, r As Long, firstRow As Long, lastRow As Long, i As Long
    Dim sErr As String, sAction As String, nOK As Long, nIns As Long, nUpd As Long, nSkip As Long, nErr As Long
    WMSDBO_SyncCurrentOrderCore = False
    If Not WMSDBO_GetSelectedOrderRow(oDoc, oSh, r, sErr) Then sReport = sErr: Exit Function
    Orders_ProcessRow oDoc, oSh, r, False, False, True, True
    If Not Orders_GetNumberedGroupBounds(oSh, r, firstRow, lastRow) Then firstRow = r: lastRow = r
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then sReport = sErr: Exit Function
    If Not WMSDBO_EnsureTable(oCon, sErr) Or Not WMSDBO_EnsureOrdersTable(oCon, sErr) Then sReport = sErr: Exit Function

    For i = firstRow To lastRow
        sErr = "": sAction = ""
        If WMSDBO_SyncSpecificRow(oDoc, oSh, i, oCon, sAction, sErr) Then
            nOK = nOK + 1
            If sAction = "INSERT" Then nIns = nIns + 1
            If sAction = "UPDATE" Then nUpd = nUpd + 1
            If sAction = "SKIP" Then nSkip = nSkip + 1
        Else
            nErr = nErr + 1
        End If
    Next i

    WMSDB_SaveDatabaseDocument sErr
    WMSDB_Close
    sReport = "Заказ: строки Calc " & CStr(firstRow + 1) & "–" & CStr(lastRow + 1) & Chr(10) & _
              "Успешно: " & CStr(nOK) & Chr(10) & _
              "Новых: " & CStr(nIns) & " | Обновлено: " & CStr(nUpd) & " | Без изменений: " & CStr(nSkip) & Chr(10) & _
              "Ошибок/пропущено: " & CStr(nErr) & Chr(10) & _
              "Calc-строки не удалялись."
    WMSDBO_SyncCurrentOrderCore = (nOK > 0 And nErr = 0)
End Function

Function WMSDBO_SyncAllReadyPositionsCore(oDoc As Object, ByRef sReport As String) As Boolean
    Dim oSh As Object, oCon As Object, oCur As Object, lastRow As Long, i As Long
    Dim sErr As String, sAction As String, nOK As Long, nIns As Long, nUpd As Long, nSkip As Long, nErr As Long, nEmpty As Long
    WMSDBO_SyncAllReadyPositionsCore = False
    If Not oDoc.Sheets.hasByName(WMSDBO_SHEET) Then sReport = "Нет листа 'Заказы'.": Exit Function
    oSh = oDoc.Sheets.getByName(WMSDBO_SHEET)
    oCur = oSh.createCursorByRange(oSh.getCellByPosition(0,0))
    oCur.gotoEndOfUsedArea(True)
    lastRow = oCur.RangeAddress.EndRow
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then sReport = sErr: Exit Function
    If Not WMSDBO_EnsureTable(oCon, sErr) Or Not WMSDBO_EnsureOrdersTable(oCon, sErr) Then sReport = sErr: Exit Function

    For i = 1 To lastRow
        sErr = "": sAction = ""
        If Trim(oSh.getCellByPosition(1, i).String) = "" Then
            nEmpty = nEmpty + 1
        ElseIf WMSDBO_SyncSpecificRow(oDoc, oSh, i, oCon, sAction, sErr) Then
            nOK = nOK + 1
            If sAction = "INSERT" Then nIns = nIns + 1
            If sAction = "UPDATE" Then nUpd = nUpd + 1
            If sAction = "SKIP" Then nSkip = nSkip + 1
        Else
            nErr = nErr + 1
        End If
    Next i
    WMSDB_SaveDatabaseDocument sErr
    WMSDB_Close
    sReport = "Массовая синхронизация 'Заказы' завершена." & Chr(10) & _
              "Успешно: " & CStr(nOK) & Chr(10) & _
              "Новых: " & CStr(nIns) & " | Обновлено: " & CStr(nUpd) & " | Без изменений: " & CStr(nSkip) & Chr(10) & _
              "Ошибок/неготовых: " & CStr(nErr) & Chr(10) & _
              "Пустых строк проигнорировано: " & CStr(nEmpty)
    WMSDBO_SyncAllReadyPositionsCore = (nOK > 0 And nErr = 0)
End Function

Function WMSDBO_ViewSelectedPositionDBCore(oDoc As Object, ByRef sReport As String) As Boolean
    Dim oSh As Object, oCon As Object, r As Long, sid As String, sErr As String, sSQL As String
    Dim oStmt As Object, oRS As Object
    WMSDBO_ViewSelectedPositionDBCore = False
    If Not WMSDBO_GetSelectedOrderRow(oDoc, oSh, r, sErr) Then sReport = sErr: Exit Function
    sid = WMSDBO_TechText(oSh, r, "_WMS_SourceID")
    If sid = "" Then sReport = "У строки пока нет SourceID.": Exit Function
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then sReport = sErr: Exit Function
    sSQL = "SELECT SOURCE_ID,ORDER_GROUP_ID,POSITION_NO,PRODUCT_NAME,ORDER_QTY,UNIT_NAME,STATUS_NAME,UPDATED_AT FROM " & WMSDBO_TABLE & " WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If Not oRS.next() Then sReport = "В Firebird записи с SourceID " & sid & " нет.": Exit Function
    sReport = "SourceID: " & oRS.getString(1) & Chr(10) & _
              "OrderGroupID: " & oRS.getString(2) & Chr(10) & _
              "Позиция: " & oRS.getString(3) & Chr(10) & _
              "Товар: " & oRS.getString(4) & Chr(10) & _
              "Количество: " & oRS.getString(5) & " " & oRS.getString(6) & Chr(10) & _
              "Статус: " & oRS.getString(7) & Chr(10) & _
              "Обновлено в БД: " & oRS.getString(8)
    WMSDBO_ViewSelectedPositionDBCore = True
    Exit Function
EH:
    sReport = "Ошибка чтения Firebird: " & CStr(Err) & " " & Error$
End Function

' ---------------------------------------------------------------------------
' SELECTION + ROW VALIDATION
' ---------------------------------------------------------------------------

Function WMSDBO_GetSelectedOrderRow(oDoc As Object, ByRef oSh As Object, ByRef r As Long, ByRef sErr As String) As Boolean
    Dim oSel As Object, addr As Variant
    WMSDBO_GetSelectedOrderRow = False
    sErr = ""
    r = -1

    On Error GoTo EH

    oSh = oDoc.CurrentController.ActiveSheet
    If oSh.Name <> WMSDBO_SHEET Then
        sErr = "Перейдите на лист 'Заказы' и выделите ячейку нужной позиции."
        Exit Function
    End If

    oSel = oDoc.CurrentController.getSelection()
    On Error Resume Next
    addr = oSel.CellAddress
    If Err = 0 Then
        r = addr.Row
    Else
        Err = 0
        addr = oSel.RangeAddress
        If Err = 0 Then r = addr.StartRow
    End If
    On Error GoTo EH

    If r < 1 Then
        sErr = "Выберите строку с позицией заказа, а не заголовок."
        Exit Function
    End If

    WMSDBO_GetSelectedOrderRow = True
    Exit Function
EH:
    sErr = "Не удалось определить выбранную строку: " & CStr(Err) & " " & Error$
End Function

Function WMSDBO_RowIsRealOrderPosition(oSh As Object, r As Long, ByRef sErr As String) As Boolean
    Dim sName As String, sQty As String, dQty As Double
    WMSDBO_RowIsRealOrderPosition = False
    sErr = ""

    sName = Trim(oSh.getCellByPosition(1, r).String)
    sQty = Trim(oSh.getCellByPosition(7, r).String)
    dQty = oSh.getCellByPosition(7, r).Value

    ' Deliberately do not treat automatic Status/Control/Category/Code as proof
    ' that a row is real. This prevents the old "empty row became alive" defect.
    If sName = "" And sQty = "" And Abs(dQty) < 0.0000001 Then
        sErr = "Выбранная строка не похожа на реальную позицию заказа: нет наименования и количества."
        Exit Function
    End If

    If sName = "" Then
        sErr = "Не заполнено наименование товара."
        Exit Function
    End If

    If sQty = "" And Abs(dQty) < 0.0000001 Then
        sErr = "Не заполнено заказанное количество."
        Exit Function
    End If

    WMSDBO_RowIsRealOrderPosition = True
End Function

' ---------------------------------------------------------------------------
' SQL VALUE HELPERS
' ---------------------------------------------------------------------------

Function WMSDBO_SQLNumberFromCell(oCell As Object) As String
    Dim s As String, d As Double
    s = Trim(oCell.String)
    d = oCell.Value

    If s = "" And Abs(d) < 0.0000000001 Then
        WMSDBO_SQLNumberFromCell = "NULL"
    Else
        WMSDBO_SQLNumberFromCell = Replace(Trim(Str(d)), ",", ".")
    End If
End Function

Function WMSDBO_SQLIntegerFromCell(oCell As Object) As String
    Dim s As String, d As Double
    s = Trim(oCell.String)
    d = oCell.Value

    If s = "" And Abs(d) < 0.0000000001 Then
        WMSDBO_SQLIntegerFromCell = "NULL"
    Else
        WMSDBO_SQLIntegerFromCell = CStr(CLng(d))
    End If
End Function

Function WMSDBO_SQLDate(d As Double) As String
    Dim dt As Date
    If d <= 0 Then
        WMSDBO_SQLDate = "NULL"
        Exit Function
    End If

    dt = CDate(d)
    WMSDBO_SQLDate = "CAST('" & _
        Right("0000" & CStr(Year(dt)), 4) & "-" & _
        Right("00" & CStr(Month(dt)), 2) & "-" & _
        Right("00" & CStr(Day(dt)), 2) & "' AS DATE)"
End Function

Function WMSDBO_PostStockForRange(oDoc As Object,oSh As Object,firstRow As Long,lastRow As Long,ByRef sErr As String) As Boolean
    Dim oCon As Object,r As Long
    WMSDBO_PostStockForRange=False:sErr=""
    On Error GoTo EH

    ' One shared connection for the entire order. Do not reopen embedded Base
    ' once per row.
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    For r=firstRow To lastRow
        If WMSDBO_RowHasUserData(oSh,r) Then
            If Not WMSDBO_PostStockForRowCon(oCon,oSh,r,sErr) Then
                sErr="Строка " & CStr(r+1) & ": " & sErr
                Exit Function
            End If
        End If
    Next r

    If Not WMSDB_SaveDatabaseDocument(sErr) Then Exit Function
    WMSDBO_PostStockForRange=True
    Exit Function
EH:
    sErr="Ошибка записи остатка заказа: " & CStr(Err) & " " & Error$
End Function

Function WMSDBO_PostStockForRow(oDoc As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim oCon As Object
    WMSDBO_PostStockForRow=False:sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    If Not WMSDBO_PostStockForRowCon(oCon,oSh,r,sErr) Then Exit Function
    If Not WMSDB_SaveDatabaseDocument(sErr) Then Exit Function

    WMSDBO_PostStockForRow=True
    Exit Function
EH:
    sErr="Ошибка прихода в остаток: " & CStr(Err) & " " & Error$
End Function

Function WMSDBO_PostStockForRowCon(oCon As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim sid As String,movementID As String,src As String,cat As String,subcat As String
    Dim code As String,nm As String,unitText As String,loc As String,noteText As String
    Dim qty As Double,dDate As Double,sourceType As String

    WMSDBO_PostStockForRowCon=False:sErr=""

    sid=WMSDBO_TechText(oSh,r,"_WMS_SourceID")
    code=Trim(oSh.getCellByPosition(21,r).String)
    nm=Trim(oSh.getCellByPosition(1,r).String)
    qty=oSh.getCellByPosition(6,r).Value
    unitText=Trim(oSh.getCellByPosition(8,r).String)
    loc=Trim(oSh.getCellByPosition(19,r).String)
    src=Trim(oSh.getCellByPosition(23,r).String)
    cat=Trim(oSh.getCellByPosition(17,r).String)
    subcat=Trim(oSh.getCellByPosition(24,r).String)
    dDate=Orders_CellDate(oSh.getCellByPosition(12,r),True)
    noteText=Trim(oSh.getCellByPosition(20,r).String)

    If sid="" Then sErr="Нет SourceID.":Exit Function
    If code="" Then sErr="Нет внутреннего кода товара.":Exit Function
    If qty<=0 Then sErr="Фактическое количество должно быть больше 0.":Exit Function
    If unitText="" Then sErr="Не заполнена единица измерения.":Exit Function
    If loc="" Then sErr="Не заполнено место хранения.":Exit Function
    If src="" Then sErr="Не заполнен источник прихода.":Exit Function
    If dDate<=0 Then sErr="Не заполнена дата поступления.":Exit Function

    sourceType=WMSDBO_StockSourceType(src)
    movementID="IN-" & sid

    ' 0.5.0: WMS_12 owns document/base conversion and lot creation.
    ' WMS_10 still writes the immutable physical stock movement.
    On Error GoTo Missing
    If Not WMSDBST_MovementExists(oCon,movementID,sErr) Then
        If sErr<>"" Then Exit Function
        If Not WMSX_SaveOrderSnapshot(oCon,oSh,r,sid,sErr) Then Exit Function
    Else
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBSR_PostReceiptCon(oCon,oSh,r,movementID,sid,sourceType, _
        code,nm,qty,unitText,loc,src,cat,subcat,dDate,noteText,sErr) Then
        If sErr="" Then sErr="WMS_12 не подтвердил умный приход."
        Exit Function
    End If

    WMSDBO_PostStockForRowCon=True
    Exit Function
Missing:
    sErr="Нужны WMS_10 Stock 0.2.3+, WMS_11 UnitsLots 0.1.0+ и WMS_12 SmartReceipt 0.1.0+: " & _
         CStr(Err) & " " & Error$
End Function

Function WMSDBO_StockSourceType(src As String) As String
    Dim s As String
    s=LCase(Trim(src))
    If InStr(1,s,"производ",1)>0 Then
        WMSDBO_StockSourceType="PRODUCTION_IN"
    ElseIf InStr(1,s,"офис",1)>0 Then
        WMSDBO_StockSourceType="OFFICE_IN"
    ElseIf InStr(1,s,"прям",1)>0 Then
        WMSDBO_StockSourceType="DIRECT_RECEIPT"
    ElseIf InStr(1,s,"поставщик",1)>0 Or InStr(1,s,"заказ",1)>0 Then
        WMSDBO_StockSourceType="ORDER_RECEIPT"
    Else
        WMSDBO_StockSourceType="OTHER_IN"
    End If
End Function

Function WMSDBO_VerifyConductedAfterReconnect(oDoc As Object,oSh As Object,firstRow As Long,lastRow As Long,ByRef sErr As String) As Boolean
    Dim oCon As Object,r As Long,sid As String,movementID As String
    WMSDBO_VerifyConductedAfterReconnect=False:sErr=""
    On Error GoTo EH

    ' Exactly one close/reopen cycle for the final durability check.
    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    For r=firstRow To lastRow
        If WMSDBO_RowHasUserData(oSh,r) Then
            sid=WMSDBO_TechText(oSh,r,"_WMS_SourceID")
            If sid="" Then sErr="Строка " & CStr(r+1) & ": нет SourceID.":GoTo Fail
            If Not WMSDBO_VerifyRecord(oCon,sid,sErr) Then
                If sErr="" Then sErr="Строка " & CStr(r+1) & ": заказ не найден после reconnect."
                GoTo Fail
            End If
            movementID="IN-" & sid
            If Not WMSDBST_VerifyMovementCon(oCon,movementID,sErr) Then
                If sErr="" Then sErr="Строка " & CStr(r+1) & ": движение не найдено после reconnect."
                GoTo Fail
            End If
            If Not WMSDBSR_VerifyReceiptCon(oCon,movementID,sid,sErr) Then
                If sErr="" Then sErr="Строка " & CStr(r+1) & ": партия/единицы не подтверждены после reconnect."
                GoTo Fail
            End If
        End If
    Next r

    WMSDB_Close
    WMSDBO_VerifyConductedAfterReconnect=True
    Exit Function
Fail:
    WMSDB_Close
    Exit Function
EH:
    sErr="Финальная проверка Firebird: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

' ---------------------------------------------------------------------------
' TECH COLUMN HELPERS
' ---------------------------------------------------------------------------

Function WMSDBO_FindHeader(oSh As Object, sHeader As String) As Long
    Dim c As Long, last As Long
    WMSDBO_FindHeader = -1
    last = WMSDBO_LastHeaderCol(oSh)
    For c = 0 To last
        If Trim(oSh.getCellByPosition(c, 0).String) = Trim(sHeader) Then
            WMSDBO_FindHeader = c
            Exit Function
        End If
    Next c
End Function

Function WMSDBO_LastHeaderCol(oSh As Object) As Long
    Dim oCursor As Object
    WMSDBO_LastHeaderCol = 0
    On Error GoTo Done
    oCursor = oSh.createCursorByRange(oSh.getCellByPosition(0, 0))
    oCursor.gotoEndOfUsedArea(True)
    WMSDBO_LastHeaderCol = oCursor.RangeAddress.EndColumn
Done:
End Function

Function WMSDBO_TechText(oSh As Object, r As Long, sHeader As String) As String
    Dim c As Long
    WMSDBO_TechText = ""
    c = WMSDBO_FindHeader(oSh, sHeader)
    If c >= 0 Then WMSDBO_TechText = Trim(oSh.getCellByPosition(c, r).String)
End Function

Sub WMSDBO_SetTech(oSh As Object, r As Long, sHeader As String, sValue As String)
    Dim c As Long
    c = WMSDBO_FindHeader(oSh, sHeader)
    If c >= 0 Then oSh.getCellByPosition(c, r).String = sValue
End Sub

Function WMSDBO_Stamp(v As Variant) As String
    WMSDBO_Stamp = Right("0000" & CStr(Year(v)), 4) & "-" & _
                   Right("00" & CStr(Month(v)), 2) & "-" & _
                   Right("00" & CStr(Day(v)), 2) & " " & _
                   Right("00" & CStr(Hour(v)), 2) & ":" & _
                   Right("00" & CStr(Minute(v)), 2) & ":" & _
                   Right("00" & CStr(Second(v)), 2)
End Function


' ============================================================================
' SAFE CONDUCT 0.4.0
' ============================================================================


' ============================================================================
' SAFE MASS CONDUCT 0.6.0
'
' - Processes all ready receipt rows.
' - Uses one open Firebird connection for sync + lot/stock posting.
' - Does NOT delete rows during database writes.
' - Reconnects once at the end and verifies each prepared SOURCE_ID + movement + lot.
' - Deletes only verified rows, bottom-to-top.
' - Error rows remain visible in Calc.
' ============================================================================

Sub WMSDBO_ConductAllReady()
    Dim oDoc As Object,oSh As Object,oCon As Object
    Dim lastRow As Long,r As Long,nBusiness As Long,nPrepared As Long,nDeleted As Long,nErr As Long
    Dim sErr As String,sAction As String,sReport As String
    Dim rows() As Long,sids() As String,movs() As String,ok() As Boolean
    Dim answer As Integer,i As Long

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSDBO_SHEET) Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Провести все приходы"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSDBO_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    lastRow=WMSDBO_LastOrderRow(oSh)
    For r=1 To lastRow
        If WMSDBO_RowHasUserData(oSh,r) Then nBusiness=nBusiness+1
    Next r
    If nBusiness=0 Then
        MsgBox "Нет строк для проведения.",48,"WMS — Провести все приходы"
        Exit Sub
    End If

    answer=MsgBox("Провести все готовые приходы?" & Chr(10) & Chr(10) & _
                  "Строк с данными: " & CStr(nBusiness) & Chr(10) & _
                  "Ошибочные строки останутся на листе." & Chr(10) & _
                  "Подтверждённые Firebird строки будут удалены из Calc.", _
                  36,"WMS — Провести все приходы")
    If answer<>6 Then Exit Sub

    ReDim rows(1 To nBusiness)
    ReDim sids(1 To nBusiness)
    ReDim movs(1 To nBusiness)

    gWMSORD_Busy=True
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail
    If Not WMSDBO_EnsureTable(oCon,sErr) Or Not WMSDBO_EnsureOrdersTable(oCon,sErr) Then GoTo Fail

    ' Phase 1 — persist without deleting.
    For r=1 To lastRow
        If WMSDBO_RowHasUserData(oSh,r) Then
            sErr="":sAction=""

            ' Preserve original line numbers for order groups before any future deletion.
            Dim gid As String
            gid=WMSDBO_TechText(oSh,r,"_WMS_OrderGroupID")
            If gid<>"" Then WMSDBO_FreezeOriginalPositions oSh,gid

            If WMSDBO_SyncSpecificRow(oDoc,oSh,r,oCon,sAction,sErr) Then
                Dim guardSid As String
                guardSid=WMSDBO_TechText(oSh,r,"_WMS_SourceID")
                If guardSid="" Then
                    nErr=nErr+1
                    If Len(sReport)<7000 Then sReport=sReport & "Строка " & CStr(r+1) & ": отсутствует SourceID." & Chr(10)
                ElseIf Not WMSSAFE_GuardBeginCon(oCon,"RECEIPT",guardSid,"CONDUCT",sErr) Then
                    nErr=nErr+1
                    If Len(sReport)<7000 Then sReport=sReport & "Строка " & CStr(r+1) & ": " & sErr & Chr(10)
                ElseIf WMSDBO_PostStockForRowCon(oCon,oSh,r,sErr) Then
                    nPrepared=nPrepared+1
                    rows(nPrepared)=r
                    sids(nPrepared)=guardSid
                    movs(nPrepared)="IN-" & sids(nPrepared)
                Else
                    WMSSAFE_GuardFinishCon oCon,"RECEIPT",guardSid,"CONDUCT",False,sErr
                    nErr=nErr+1
                    If Len(sReport)<7000 Then sReport=sReport & "Строка " & CStr(r+1) & ": " & sErr & Chr(10)
                End If
            Else
                nErr=nErr+1
                If Len(sReport)<7000 Then sReport=sReport & "Строка " & CStr(r+1) & ": " & sErr & Chr(10)
            End If
        End If
    Next r

    If nPrepared=0 Then
        WMSDB_Close
        gWMSORD_Busy=False
        MsgBox "Ни одна строка не проведена." & IIf(sReport="","",Chr(10) & Chr(10) & sReport),48,"WMS — Провести все приходы"
        Exit Sub
    End If

    If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail
    WMSDB_Close

    ' Phase 2 — one durability reconnect.
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo FailKeepRows
    ReDim ok(1 To nPrepared)

    For i=1 To nPrepared
        sErr=""
        If WMSDBO_VerifyRecord(oCon,sids(i),sErr) Then
            If WMSDBST_VerifyMovementCon(oCon,movs(i),sErr) Then
                If WMSDBSR_VerifyReceiptCon(oCon,movs(i),sids(i),sErr) Then ok(i)=True
            End If
        End If
        If ok(i) Then
            WMSSAFE_GuardFinishCon oCon,"RECEIPT",sids(i),"CONDUCT",True,"Bulk verified"
        Else
            WMSSAFE_GuardFinishCon oCon,"RECEIPT",sids(i),"CONDUCT",False,"Bulk durability verification failed"
            nErr=nErr+1
            If Len(sReport)<7000 Then
                sReport=sReport & "Строка " & CStr(rows(i)+1) & ": финальная проверка не пройдена. " & sErr & Chr(10)
            End If
        End If
    Next i
    WMSDB_Close

    ' Phase 3 — delete only verified rows from bottom to top.
    For i=nPrepared To 1 Step -1
        If ok(i) Then
            oSh.Rows.removeByIndex(rows(i),1)
            nDeleted=nDeleted+1
        End If
    Next i

    Orders_ResetDupCache
    Orders_ResetHistoryCache
    Orders_RefreshFilter oDoc,oSh
    gWMSORD_Busy=False
    WMSSAFE_Audit "RECEIPT_BULK_CONDUCTED","BATCH","ORDERS-BULK-" & Format(Now,"YYYYMMDD-HHMMSS"), _
        IIf(nErr=0,"DONE","PARTIAL"),"Проведено: " & CStr(nDeleted) & "; ошибок: " & CStr(nErr)

    MsgBox "Массовое проведение завершено." & Chr(10) & _
           "Проведено и удалено: " & CStr(nDeleted) & Chr(10) & _
           "Осталось проверить: " & CStr(nBusiness-nDeleted) & _
           IIf(sReport="","",Chr(10) & Chr(10) & "Причины:" & Chr(10) & sReport), _
           IIf(nErr=0,64,48),"WMS — Провести все приходы"
    Exit Sub

FailKeepRows:
    WMSDB_Close
    gWMSORD_Busy=False
    MsgBox "Firebird получил часть данных, но финальная проверка не выполнена." & Chr(10) & _
           "Строки из Calc НЕ удалены." & Chr(10) & Chr(10) & sErr,16,"WMS — Провести все приходы"
    Exit Sub

Fail:
    WMSDB_Close
    gWMSORD_Busy=False
    MsgBox "Массовое проведение остановлено. Строки НЕ удалены." & Chr(10) & Chr(10) & sErr,16,"WMS — Провести все приходы"
    Exit Sub

EH:
    On Error Resume Next
    WMSDB_Close
    gWMSORD_Busy=False
    MsgBox "Массовое проведение: " & CStr(Err) & " " & Error$ & Chr(10) & _
           "Неподтверждённые строки автоматически не удалялись.",16,"WMS — Провести все приходы"
End Sub

Sub WMSDBO_ConductCurrentOrder()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long
    Dim firstRow As Long,lastRow As Long,grp As String,answer As Integer
    Dim sErr As String,sSyncReport As String,nRows As Long
    Dim actRowsText As String,actPartyFrom As String,actOperationID As String
    Dim offerAct As Boolean
    Dim guardID As String,guardStarted As Boolean

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName("Заказы") Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Проведение"
        Exit Sub
    End If
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>"Заказы" Then
        MsgBox "Откройте лист 'Заказы' и выберите строку нужного заказа.",48,"WMS — Проведение"
        Exit Sub
    End If

    On Error GoTo EH
    oSel=oDoc.CurrentSelection
    r=oSel.CellAddress.Row
    If r<1 Then GoTo EH

    If Not WMSDBO_FindCurrentOrderBounds(oSh,r,firstRow,lastRow,grp,sErr) Then
        MsgBox sErr,48,"WMS — Проведение"
        Exit Sub
    End If

    If Not WMSDBO_ValidateOrderForConduct(oSh,firstRow,lastRow,sErr) Then
        MsgBox "Заказ/поступление не готово к проведению:" & Chr(10) & Chr(10) & sErr,48,"WMS — Проведение"
        Exit Sub
    End If

    nRows=lastRow-firstRow+1
    answer=MsgBox("Готово к переносу в Firebird." & Chr(10) & _
                  "Позиций: " & CStr(nRows) & Chr(10) & _
                  "После контрольной проверки строки будут удалены из Calc." & Chr(10) & Chr(10) & _
                  "Продолжить?",36,"WMS — Проведение")
    If answer<>6 Then Exit Sub

    If Not WMSDBO_SyncCurrentOrderCore(oDoc,sSyncReport) Then
        MsgBox "Проведение остановлено. Данные в Calc НЕ удалены." & Chr(10) & Chr(10) & sSyncReport,16,"WMS — Проведение"
        Exit Sub
    End If

    guardID=WMSDBO_TechText(oSh,firstRow,"_WMS_SourceID")
    If guardID="" Then
        MsgBox "После записи отсутствует SourceID. Строки НЕ удалены.",16,"WMS — Проведение"
        Exit Sub
    End If
    If Not WMSSAFE_GuardBegin("RECEIPT_GROUP",guardID,"CONDUCT",sErr) Then
        MsgBox sErr,48,"WMS — Защита операции"
        Exit Sub
    End If
    guardStarted=True

    ' Write all stock movements through one connection.
    If Not WMSDBO_PostStockForRange(oDoc,oSh,firstRow,lastRow,sErr) Then
        WMSSAFE_GuardFinish "RECEIPT_GROUP",guardID,"CONDUCT",False,sErr
        guardStarted=False
        MsgBox "Заказ сохранён в БД заказов, но остаток НЕ подтверждён." & Chr(10) & _
               "Строки в Calc НЕ удалены." & Chr(10) & Chr(10) & sErr,16,"WMS — Проведение"
        Exit Sub
    End If

    ' One final reconnect verifies both order rows and all stock movements.
    If Not WMSDBO_VerifyConductedAfterReconnect(oDoc,oSh,firstRow,lastRow,sErr) Then
        WMSSAFE_GuardFinish "RECEIPT_GROUP",guardID,"CONDUCT",False,sErr
        guardStarted=False
        MsgBox "Финальная проверка Firebird не пройдена." & Chr(10) & _
               "Строки в Calc НЕ удалены." & Chr(10) & Chr(10) & sErr,16,"WMS — Проведение"
        Exit Sub
    End If

    WMSSAFE_GuardFinish "RECEIPT_GROUP",guardID,"CONDUCT",True,"Позиций: " & CStr(nRows)
    guardStarted=False

    ' Capture printable act data BEFORE Calc rows are removed.
    offerAct=WMSACT_ShouldOfferReceipt(oSh,firstRow)
    If offerAct Then
        actRowsText=WMSACT_OrderRangeRowsText(oSh,firstRow,lastRow)
        actPartyFrom=WMSACT_ReceiptPartyFrom(oSh,firstRow)
        actOperationID=WMSACT_ReceiptOperationID(oSh,firstRow)
    End If

    WMSDBO_DeleteRowsAfterConduct oSh,firstRow,lastRow
    WMSSAFE_Audit "RECEIPT_CONDUCTED","RECEIPT",grp,"DONE","Позиций: " & CStr(nRows)

    MsgBox "Проведение завершено." & Chr(10) & _
           "Позиций перенесено: " & CStr(nRows) & Chr(10) & _
           "Строки удалены из рабочего листа 'Заказы'.",64,"WMS — Проведение"

    If offerAct And Trim(actRowsText)<>"" Then
        WMSACT_AskAndCreate "IN",actPartyFrom,"Склад","Акт приёма-передачи ТМЦ",actRowsText,actOperationID
    End If
    Exit Sub
EH:
    If guardStarted And guardID<>"" Then WMSSAFE_GuardFinish "RECEIPT_GROUP",guardID,"CONDUCT",False,CStr(Err) & " " & Error$
    MsgBox "Не удалось провести выбранный заказ/поступление: " & CStr(Err) & " " & Error$,16,"WMS — Проведение"
End Sub

Function WMSDBO_FindCurrentOrderBounds(oSh As Object,r As Long,ByRef firstRow As Long,ByRef lastRow As Long,ByRef grp As String,ByRef sErr As String) As Boolean
    Dim rr As Long,lastContent As Long
    WMSDBO_FindCurrentOrderBounds=False:sErr=""
    If Trim(oSh.getCellByPosition(0,r).String)="" And oSh.getCellByPosition(0,r).Value=0 Then
        sErr="В выбранной строке нет номера позиции в колонке A."
        Exit Function
    End If

    firstRow=r
    Do While firstRow>1
        If oSh.getCellByPosition(0,firstRow).Value=1 Then Exit Do
        firstRow=firstRow-1
    Loop
    If oSh.getCellByPosition(0,firstRow).Value<>1 Then
        sErr="Не удалось найти начало заказа (строка с A=1)."
        Exit Function
    End If

    lastContent=WMSDBO_LastOrderRow(oSh)
    lastRow=firstRow
    For rr=firstRow+1 To lastContent
        If oSh.getCellByPosition(0,rr).Value=1 Then Exit For
        If WMSDBO_RowHasUserData(oSh,rr) Then lastRow=rr
    Next rr
    grp="ROW-" & CStr(firstRow+1)
    WMSDBO_FindCurrentOrderBounds=True
End Function


Function WMSDBO_QuoteSimple(s As String) As String
    WMSDBO_QuoteSimple = Replace(CStr(s), "'", "''")
End Function

Function WMSDBO_FillNameByCodeCon(oCon As Object, oSh As Object, r As Long, ByRef sErr As String) As Boolean
    Dim code As String, oStmt As Object, oRS As Object, sql As String, nm As String
    WMSDBO_FillNameByCodeCon = False
    sErr = ""
    On Error GoTo EH

    code = Trim(oSh.getCellByPosition(21,r).String)
    If code = "" Then Exit Function

    sql = "SELECT PRODUCT_NAME FROM WMS_PRODUCTS WHERE PRODUCT_CODE='" & WMSDBO_QuoteSimple(code) & "'"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)
    If oRS.next() Then
        nm = Trim(oRS.getString(1))
        If nm <> "" Then
            oSh.getCellByPosition(1,r).String = nm
            WMSDBO_FillNameByCodeCon = True
        End If
    End If
    Exit Function
EH:
    sErr = "Поиск товара по коду: " & CStr(Err) & " " & Error$
End Function

Function WMSDBO_ValidateOrderForConduct(oSh As Object,firstRow As Long,lastRow As Long,ByRef sErr As String) As Boolean
    Dim r As Long,miss As String,nm As String,qty As Double,unitText As String,receiptType As String
    Dim oCon As Object, dbErr As String, code As String
    sErr=""
    On Error GoTo EH

    ' Один read-only lookup connection на всю проверяемую группу.
    dbErr=""
    oCon=WMSDB_GetConnectionEx(ThisComponent,dbErr)

    For r=firstRow To lastRow
        miss=""
        code=Trim(oSh.getCellByPosition(21,r).String)

        ' Если введён код известного товара, имя восстанавливаем из Firebird.
        ' Поэтому после проведения старой позиции пользователю не надо помнить
        ' полное наименование для следующего прихода.
        If code<>"" And dbErr="" Then
            Dim lookupErr As String
            lookupErr=""
            Call WMSDBO_FillNameByCodeCon(oCon,oSh,r,lookupErr)
        End If

        nm=Trim(oSh.getCellByPosition(1,r).String)
        qty=oSh.getCellByPosition(6,r).Value
        unitText=Trim(oSh.getCellByPosition(8,r).String)

        If nm="" Then
            If code="" Then
                miss=miss & "наименование/код товара; "
            Else
                miss=miss & "код товара не найден в БД и наименование пустое; "
            End If
        End If
        If qty<=0 Then miss=miss & "фактическое количество; "
        If unitText="" Then miss=miss & "ед. изм.; "
        If code="" Then miss=miss & "код товара; "
        If Trim(oSh.getCellByPosition(19,r).String)="" Then miss=miss & "место хранения; "
        If Trim(oSh.getCellByPosition(23,r).String)="" Then miss=miss & "источник прихода; "

        ' Категория (R) и Подкатегория (Y) теперь обычные свойства товара.
        ' Они НЕ являются "категорией назначения" и не блокируют проведение.

        ' Supplier is mandatory only for supplier-origin receipts.
        If InStr(1,LCase(Trim(oSh.getCellByPosition(23,r).String)),"поставщик",1)>0 Then
            If Trim(oSh.getCellByPosition(11,r).String)="" Then miss=miss & "поставщик; "
        End If
        If Trim(oSh.getCellByPosition(12,r).String)="" And oSh.getCellByPosition(12,r).Value=0 Then miss=miss & "дата поступления; "

        If miss<>"" Then
            sErr=sErr & "Строка " & CStr(r+1) & ": " & miss & Chr(10)
        End If
    Next r

    WMSDB_Close
    WMSDBO_ValidateOrderForConduct=(sErr="")
    Exit Function
EH:
    On Error Resume Next
    WMSDB_Close
    sErr="Проверка заказа: " & CStr(Err) & " " & Error$
    WMSDBO_ValidateOrderForConduct=False
End Function

Function WMSDBO_VerifyOrderPersisted(oSh As Object,firstRow As Long,lastRow As Long,ByRef sErr As String) As Boolean
    Dim r As Long,sourceID As String,oCon As Object
    WMSDBO_VerifyOrderPersisted=False:sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function

    For r=firstRow To lastRow
        sourceID=WMSDBO_TechText(oSh,r,"_WMS_SourceID")
        If sourceID="" Then
            sErr="Строка " & CStr(r+1) & ": отсутствует SourceID после синхронизации."
            WMSDB_Close
            Exit Function
        End If
        If Not WMSDBO_VerifyRecord(oCon,sourceID,sErr) Then
            If sErr="" Then sErr="Строка " & CStr(r+1) & ": Firebird не подтвердил SOURCE_ID=" & sourceID
            WMSDB_Close
            Exit Function
        End If
    Next r

    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        WMSDB_Close
        Exit Function
    End If
    WMSDB_Close
    WMSDBO_VerifyOrderPersisted=True
    Exit Function
EH:
    sErr="Ошибка контрольного чтения Firebird: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

Sub WMSDBO_DeleteRowsAfterConduct(oSh As Object,firstRow As Long,lastRow As Long)
    Dim r As Long
    On Error GoTo EH
    For r=lastRow To firstRow Step -1
        oSh.Rows.removeByIndex(r,1)
    Next r
    Exit Sub
EH:
    Err.Raise Err,,"Не удалось удалить проведённые строки из Calc: " & Error$
End Sub

Function WMSDBO_LastOrderRow(oSh As Object) As Long
    Dim cur As Object
    cur=oSh.createCursor()
    cur.gotoEndOfUsedArea(True)
    WMSDBO_LastOrderRow=cur.RangeAddress.EndRow
End Function

Function WMSDBO_RowHasUserData(oSh As Object,r As Long) As Boolean
    Dim c As Long
    For c=1 To 25
        If c<>16 And c<>18 Then
            If Trim(oSh.getCellByPosition(c,r).Formula)<>"" Then WMSDBO_RowHasUserData=True:Exit Function
        End If
    Next c
End Function


' ============================================================================
' CONDUCT ONE POSITION 0.4.1
' ============================================================================
Sub WMSDBO_ConductSelectedPosition()
    Dim oDoc As Object,oSh As Object,r As Long,sErr As String,sReport As String
    Dim sid As String,gid As String,productName As String,answer As Integer
    Dim originalPos As Long
    Dim actRowsText As String,actPartyFrom As String,actOperationID As String
    Dim offerAct As Boolean,guardStarted As Boolean

    oDoc=ThisComponent
    If Not WMSDBO_GetSelectedOrderRow(oDoc,oSh,r,sErr) Then
        MsgBox sErr,48,"WMS — Провести позицию"
        Exit Sub
    End If

    If Not WMSDBO_RowIsRealOrderPosition(oSh,r,sErr) Then
        MsgBox sErr,48,"WMS — Провести позицию"
        Exit Sub
    End If

    ' Normalize/validate first. If anything mandatory is missing, the Orders
    ' controller writes the exact reason into "Контроль".
    Orders_ProcessRow oDoc,oSh,r,False,False,True,True
    If UCase(WMSDBO_TechText(oSh,r,"_WMS_Severity"))="CRITICAL" Then
        MsgBox "Позиция не готова к проведению." & Chr(10) & Chr(10) & _
               Trim(oSh.getCellByPosition(18,r).String),48,"WMS — Провести позицию"
        Exit Sub
    End If

    gid=WMSDBO_TechText(oSh,r,"_WMS_OrderGroupID")
    If gid="" Then
        MsgBox "Не удалось определить группу заказа.",16,"WMS — Провести позицию"
        Exit Sub
    End If

    originalPos=CLng(oSh.getCellByPosition(0,r).Value)
    productName=Trim(oSh.getCellByPosition(1,r).String)

    ' Freeze original DB position numbers for ALL still visible lines of this
    ' order before any row is removed. Calc can then be renumbered 1..N while
    ' Firebird keeps the original line numbers and never gets duplicates.
    WMSDBO_FreezeOriginalPositions oSh,gid

    answer=MsgBox("Провести только выбранную позицию?" & Chr(10) & Chr(10) & _
                  "Позиция: " & CStr(originalPos) & Chr(10) & _
                  "Товар: " & productName & Chr(10) & _
                  "После подтверждения Firebird строка будет удалена из Calc.", _
                  36,"WMS — Провести позицию")
    If answer<>6 Then Exit Sub

    If Not WMSDBO_SyncSelectedPositionCore(oDoc,sReport) Then
        MsgBox "Проведение остановлено. Строка в Calc НЕ удалена." & Chr(10) & Chr(10) & _
               sReport,16,"WMS — Провести позицию"
        Exit Sub
    End If

    sid=WMSDBO_TechText(oSh,r,"_WMS_SourceID")
    If sid="" Then
        MsgBox "После записи отсутствует SourceID. Строка в Calc НЕ удалена.",16,"WMS — Провести позицию"
        Exit Sub
    End If
    If Not WMSSAFE_GuardBegin("RECEIPT",sid,"CONDUCT",sErr) Then
        MsgBox sErr,48,"WMS — Защита операции"
        Exit Sub
    End If
    guardStarted=True

    ' SyncSelectedPositionCore already performed its own reconnect verification.
    ' Now write stock using the SAME open embedded-Base connection.
    If Not WMSDBO_PostStockForRow(oDoc,oSh,r,sErr) Then
        WMSSAFE_GuardFinish "RECEIPT",sid,"CONDUCT",False,sErr
        guardStarted=False
        MsgBox "Позиция сохранена в БД заказов, но приход в остаток НЕ подтверждён." & Chr(10) & _
               "Строка в Calc НЕ удалена." & Chr(10) & Chr(10) & sErr,16,"WMS — Провести позицию"
        Exit Sub
    End If

    ' One final reconnect verifies BOTH the order record and the stock movement.
    If Not WMSDBO_VerifyConductedAfterReconnect(oDoc,oSh,r,r,sErr) Then
        WMSSAFE_GuardFinish "RECEIPT",sid,"CONDUCT",False,sErr
        guardStarted=False
        MsgBox "Финальная проверка Firebird не пройдена." & Chr(10) & _
               "Строка в Calc НЕ удалена." & Chr(10) & Chr(10) & sErr,16,"WMS — Провести позицию"
        Exit Sub
    End If

    WMSSAFE_GuardFinish "RECEIPT",sid,"CONDUCT",True,productName
    guardStarted=False

    offerAct=WMSACT_ShouldOfferReceipt(oSh,r)
    If offerAct Then
        actRowsText=WMSACT_OrderRangeRowsText(oSh,r,r)
        actPartyFrom=WMSACT_ReceiptPartyFrom(oSh,r)
        actOperationID=WMSACT_ReceiptOperationID(oSh,r)
    End If

    If Not WMSDBO_RemoveConductedPosition(oSh,r,gid,sErr) Then
        MsgBox "Позиция уже сохранена в Firebird, но удалить её из Calc не удалось." & Chr(10) & _
               sErr,48,"WMS — Провести позицию"
        Exit Sub
    End If

    WMSSAFE_Audit "RECEIPT_POSITION_CONDUCTED","RECEIPT",sid,"DONE",productName

    MsgBox "Позиция проведена." & Chr(10) & _
           "Товар: " & productName & Chr(10) & _
           "Исходная позиция заказа: " & CStr(originalPos) & Chr(10) & _
           "Строка удалена из рабочего листа.",64,"WMS — Провести позицию"

    If offerAct And Trim(actRowsText)<>"" Then
        WMSACT_AskAndCreate "IN",actPartyFrom,"Склад","Акт приёма-передачи ТМЦ",actRowsText,actOperationID
    End If
End Sub

Sub WMSDBO_FreezeOriginalPositions(oSh As Object,gid As String)
    Dim cPos As Long,cGroup As Long,lastRow As Long,r As Long
    cPos=WMSDBO_GetOrCreateExtraTechColumn(oSh,"_WMS_DBPositionNo")
    cGroup=WMSDBO_FindHeader(oSh,"_WMS_OrderGroupID")
    If cGroup<0 Then Exit Sub

    lastRow=WMSDBO_LastOrderRow(oSh)
    For r=1 To lastRow
        If Trim(oSh.getCellByPosition(cGroup,r).String)=gid Then
            If Trim(oSh.getCellByPosition(cPos,r).String)="" Then
                oSh.getCellByPosition(cPos,r).Value=oSh.getCellByPosition(0,r).Value
            End If
        End If
    Next r
    On Error Resume Next
    oSh.Columns.getByIndex(cPos).IsVisible=False
    On Error GoTo 0
End Sub

Function WMSDBO_DBPositionSQL(oSh As Object,r As Long) As String
    Dim c As Long,v As Double
    c=WMSDBO_FindHeader(oSh,"_WMS_DBPositionNo")
    If c>=0 Then
        v=oSh.getCellByPosition(c,r).Value
        If v>0 Then
            WMSDBO_DBPositionSQL=CStr(CLng(v))
            Exit Function
        End If
    End If
    WMSDBO_DBPositionSQL=WMSDBO_SQLIntegerFromCell(oSh.getCellByPosition(0,r))
End Function

Function WMSDBO_GetOrCreateExtraTechColumn(oSh As Object,headerName As String) As Long
    Dim c As Long,lastCol As Long,cur As Object
    c=WMSDBO_FindHeader(oSh,headerName)
    If c>=0 Then
        WMSDBO_GetOrCreateExtraTechColumn=c
        Exit Function
    End If

    cur=oSh.createCursor()
    cur.gotoEndOfUsedArea(True)
    lastCol=cur.RangeAddress.EndColumn
    c=lastCol+1
    oSh.getCellByPosition(c,0).String=headerName
    On Error Resume Next
    oSh.Columns.getByIndex(c).IsVisible=False
    On Error GoTo 0
    WMSDBO_GetOrCreateExtraTechColumn=c
End Function

Function WMSDBO_VerifySourceAfterReconnect(oDoc As Object,sid As String,ByRef sErr As String) As Boolean
    Dim oCon As Object
    WMSDBO_VerifySourceAfterReconnect=False:sErr=""
    On Error GoTo EH

    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    If Not WMSDBO_VerifyRecord(oCon,sid,sErr) Then
        If sErr="" Then sErr="SOURCE_ID не найден после повторного подключения: " & sid
        WMSDB_Close
        Exit Function
    End If

    WMSDB_Close
    WMSDBO_VerifySourceAfterReconnect=True
    Exit Function
EH:
    sErr=CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

Function WMSDBO_RemoveConductedPosition(oSh As Object,r As Long,gid As String,ByRef sErr As String) As Boolean
    WMSDBO_RemoveConductedPosition=False:sErr=""
    On Error GoTo EH

    ' Suppress the normal sheet-change controller while deleting and visually
    ' renumbering the remaining rows. Their immutable GroupID/SourceID stay intact.
    gWMSORD_Busy=True
    oSh.Rows.removeByIndex(r,1)
    WMSDBO_RenumberVisibleGroup oSh,gid
    gWMSORD_Busy=False

    Orders_ResetDupCache
    Orders_ResetHistoryCache
    Orders_RefreshFilter ThisComponent,oSh
    WMSDBO_RemoveConductedPosition=True
    Exit Function
EH:
    gWMSORD_Busy=False
    sErr=CStr(Err) & " " & Error$
End Function

Sub WMSDBO_RenumberVisibleGroup(oSh As Object,gid As String)
    Dim cGroup As Long,lastRow As Long,r As Long,n As Long
    cGroup=WMSDBO_FindHeader(oSh,"_WMS_OrderGroupID")
    If cGroup<0 Then Exit Sub
    lastRow=WMSDBO_LastOrderRow(oSh)
    n=0
    For r=1 To lastRow
        If Trim(oSh.getCellByPosition(cGroup,r).String)=gid Then
            If WMSDBO_RowHasUserData(oSh,r) Then
                n=n+1
                oSh.getCellByPosition(0,r).Value=n
                Orders_ApplyRowVisual oSh,r
            End If
        End If
    Next r
End Sub



Function WMSDBO_ExtraSQL(oSh As Object,r As Long,headerName As String,isDateField As Boolean) As String
    Dim col As Long,cell As Object,textValue As String,dateValue As Double
    WMSDBO_ExtraSQL="NULL"
    col=WMSDBO_FindHeader(oSh,headerName)
    If col<0 Then Exit Function
    cell=oSh.getCellByPosition(col,r)
    textValue=Trim(cell.String)
    If textValue="" Then Exit Function
    If isDateField Then
        dateValue=Orders_CellDate(cell,True)
        If dateValue<=0 Or dateValue>2958465 Then Error 5
        WMSDBO_ExtraSQL=WMSDBO_SQLDate(dateValue)
    Else
        If Len(textValue)>180 Then Error 5
        WMSDBO_ExtraSQL=WMSDB_SQLText(textValue)
    End If
End Function

Function WMSDBO_ExtendedFingerprint(oSh As Object,r As Long) As String
    Dim extra As String
    extra=WMSDBO_ExtraSQL(oSh,r,"Назначение",False) & "|" & WMSDBO_ExtraSQL(oSh,r,"Ожидаемая дата поступления",True)
    WMSDBO_ExtendedFingerprint=Orders_BusinessFingerprint(oSh,r) & "-E1-" & CStr(WMSCore_SimpleHash(extra)) & "-" & CStr(Len(extra))
End Function
