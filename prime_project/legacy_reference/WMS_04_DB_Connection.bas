Option Explicit

' ============================================================================
' WMS_04_DB_Connection.bas
' v0.3.0 — PORTABLE embedded Firebird connection layer.
'
' Architecture:
'   current WMS.ods -> WMS_DATA_PORTABLE.odb -> embedded Firebird
'
' Why this mode:
' - no Firebird server configuration is required;
' - no DatabaseAccess / aliases / server folders;
' - the same Base file can travel with the ODS on Windows and Linux;
' - Calc stays UI; data lives outside Calc in Base/Firebird.
'
' IMPORTANT: WMS_DATA_PORTABLE.odb is deliberately a new filename so older
' experimental WMS_DATA.odb / WMS_DATA_EXTERNAL.odb are never overwritten.
' ============================================================================

Global Const WMSDB_VERSION = "0.3.0-PORTABLE"
Global Const WMSDB_MODULE = "DB_Connection"
Global Const WMSDB_DEFAULT_ODB = "WMS_DATA_PORTABLE.odb"
Global Const WMSDB_SETTING_ODB_URL = "WMS.DB.Portable.ODB.URL"

Global gWMSDB_BaseDoc As Object
Global gWMSDB_Connection As Object
Global gWMSDB_LastError As String
Global gWMSDB_LastReport As String

Sub WMSDB_TestConnection()
    Dim sReport As String
    If WMSDB_TestConnectionCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — переносимая база"
    Else
        MsgBox sReport, 16, "WMS — переносимая база"
    End If
End Sub

Sub WMSDB_Disconnect()
    WMSDB_Close
    MsgBox "Соединение WMS с WMS_DATA_PORTABLE.odb закрыто и Base-файл сохранён.", 64, "WMS — база данных"
End Sub

Sub WMSDB_SaveNow()
    Dim sErr As String
    If WMSDB_SaveDatabaseDocument(sErr) Then
        MsgBox "COMMIT и сохранение WMS_DATA_PORTABLE.odb выполнены.", 64, "WMS — база данных"
    Else
        MsgBox sErr, 16, "WMS — база данных"
    End If
End Sub

Function WMSDB_CompileProbe() As String
    WMSDB_CompileProbe = WMSDB_VERSION & "|" & WMSDB_DefaultBaseURL(ThisComponent)
End Function

Function WMSDB_LastErrorText() As String
    WMSDB_LastErrorText = gWMSDB_LastError
End Function

Function WMSDB_LastReportText() As String
    WMSDB_LastReportText = gWMSDB_LastReport
End Function

Function WMSDB_IsConnected() As Boolean
    WMSDB_IsConnected = False
    On Error GoTo Done
    If IsNull(gWMSDB_Connection) Or IsEmpty(gWMSDB_Connection) Then Exit Function
    WMSDB_IsConnected = Not gWMSDB_Connection.isClosed()
Done:
End Function

Function WMSDB_GetConnection() As Object
    Dim sErr As String
    WMSDB_GetConnection = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then gWMSDB_LastError = sErr
End Function

Function WMSDB_GetConnectionEx(oCalcDoc As Object, ByRef sErr As String) As Object
    Dim sODBURL As String
    Dim aLoad(0) As New com.sun.star.beans.PropertyValue

    WMSDB_GetConnectionEx = Nothing
    sErr = ""
    gWMSDB_LastError = ""
    If gWMSDBX_UnsafeConnection Then
        sErr=gWMSDBX_UnsafeReason:gWMSDB_LastError=sErr
        Exit Function
    End If
    On Error GoTo EH

    If WMSDB_IsConnected() Then
        WMSDB_GetConnectionEx = gWMSDB_Connection
        Exit Function
    End If

    sODBURL = WMSDB_ResolveBaseURL(oCalcDoc)
    If sODBURL = "" Then
        sErr = "Не удалось определить путь к WMS_DATA_PORTABLE.odb. Сначала сохраните ODS."
        Exit Function
    End If

    If Not WMSDB_FileExists(sODBURL) Then
        sErr = "Переносимая база ещё не создана: " & WMSDB_URLForDisplay(sODBURL) & Chr(10) & _
               "Запустите WMSDBI_InstallPortableDatabase из WMS_05_DB_Install."
        Exit Function
    End If

    aLoad(0).Name = "Hidden"
    aLoad(0).Value = True
    gWMSDB_BaseDoc = StarDesktop.loadComponentFromURL(sODBURL, "_blank", 0, aLoad())
    If IsNull(gWMSDB_BaseDoc) Or IsEmpty(gWMSDB_BaseDoc) Then
        sErr = "LibreOffice Base не смог открыть WMS_DATA_PORTABLE.odb."
        Exit Function
    End If

    gWMSDB_Connection = gWMSDB_BaseDoc.DataSource.getIsolatedConnection("", "")
    If IsNull(gWMSDB_Connection) Or IsEmpty(gWMSDB_Connection) Then
        sErr = "Base открылся, но embedded Firebird не вернул SDBC-соединение."
        WMSDB_Close
        Exit Function
    End If

    WMSDB_GetConnectionEx = gWMSDB_Connection
    Exit Function
EH:
    sErr = "Ошибка подключения к WMS_DATA_PORTABLE.odb: " & CStr(Err) & " " & Error$
    gWMSDB_LastError = sErr
    WMSDB_Close
End Function

Function WMSDB_TestConnectionCore(oCalcDoc As Object, ByRef sReport As String) As Boolean
    Dim oCon As Object
    Dim sErr As String, sODBURL As String, sProbe As String
    Dim sEngine As String, sTables As String, sOrderLines As String

    WMSDB_TestConnectionCore = False
    sReport = ""
    On Error GoTo EH

    sODBURL = WMSDB_ResolveBaseURL(oCalcDoc)
    oCon = WMSDB_GetConnectionEx(oCalcDoc, sErr)
    If sErr <> "" Then
        sReport = sErr
        gWMSDB_LastError = sErr
        gWMSDB_LastReport = sReport
        Exit Function
    End If

    sProbe = WMSDB_ScalarString(oCon, _
        "SELECT CAST('WMS_OK' AS VARCHAR(16)) FROM RDB$DATABASE", sErr)
    If sErr <> "" Or Trim(sProbe) <> "WMS_OK" Then
        If sErr = "" Then sErr = "Неожиданный результат контрольного SELECT: " & sProbe
        sReport = "Firebird открылся, но SQL execute/fetch не прошёл:" & Chr(10) & sErr
        gWMSDB_LastError = sReport
        gWMSDB_LastReport = sReport
        Exit Function
    End If

    sEngine = "Firebird"
    On Error Resume Next
    sEngine = oCon.MetaData.getDatabaseProductName()
    On Error GoTo EH

    sTables = WMSDB_UserTablesReport(oCon, sErr)
    If sErr <> "" Then sTables = "(ошибка чтения списка: " & sErr & ")"

    sErr = ""
    If WMSDB_TableExistsSimple(oCon, "WMS_ORDER_LINES", sErr) Then
        sOrderLines = "есть"
    ElseIf sErr <> "" Then
        sOrderLines = "ошибка проверки: " & sErr
    Else
        sOrderLines = "НЕТ"
    End If

    sReport = "Соединение: OK" & Chr(10) & _
              "Режим: PORTABLE EMBEDDED FIREBIRD" & Chr(10) & _
              "SQL execute/fetch: OK" & Chr(10) & _
              "Base: " & WMSDB_URLForDisplay(sODBURL) & Chr(10) & _
              "СУБД: " & sEngine & Chr(10) & _
              "WMS_ORDER_LINES: " & sOrderLines & Chr(10) & _
              "Пользовательские таблицы: " & sTables & Chr(10) & _
              "Модуль: " & WMSDB_VERSION

    gWMSDB_LastReport = sReport
    WMSDB_TestConnectionCore = True
    Exit Function
EH:
    sReport = "Тест переносимой БД завершился ошибкой: " & CStr(Err) & " " & Error$
    gWMSDB_LastError = sReport
    gWMSDB_LastReport = sReport
End Function

Function WMSDB_SaveDatabaseDocument(ByRef sErr As String) As Boolean
    WMSDB_SaveDatabaseDocument = False
    sErr = ""
    If gWMSDBX_UnsafeConnection Then sErr=gWMSDBX_UnsafeReason:Exit Function
    On Error GoTo EH

    If WMSDB_IsConnected() Then gWMSDB_Connection.commit()

    ' Absence of a Base document is a failure, never a successful save.
    gWMSDB_BaseDoc.store()

    WMSDB_SaveDatabaseDocument = True
    Exit Function
EH:
    sErr = "Не удалось зафиксировать/сохранить WMS_DATA_PORTABLE.odb: " & CStr(Err) & " " & Error$
    gWMSDB_LastError = sErr
    WMSDBX_BlockConnection sErr
End Function

Sub WMSDB_Close()
    Dim hasBaseDoc As Boolean,stage As String,detail As String
    If gWMSDBX_UnsafeConnection Then
        gWMSDB_LastError=gWMSDBX_UnsafeReason
        Exit Sub
    End If
    ' Only probe whether a document exists under Resume Next. No writes here.
    hasBaseDoc=False
    On Error Resume Next
    hasBaseDoc=gWMSDB_BaseDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument")
    On Error GoTo EH
    stage="Проверка соединения"
    If WMSDB_IsConnected() Then
        stage="COMMIT перед закрытием"
        gWMSDB_Connection.commit()
    End If
    ' Embedded Firebird must export its storage while the connection is alive.
    ' isModified is not proof that all committed rows are in the ODB file.
    If hasBaseDoc Then
        stage="Сохранение файла базы до закрытия соединения"
        gWMSDB_BaseDoc.store()
    End If
    If WMSDB_IsConnected() Then
        stage="Закрытие соединения"
        gWMSDB_Connection.close()
        gWMSDB_Connection=Nothing
    End If
    If hasBaseDoc Then
        stage="Сохранение контейнера после закрытия соединения"
        gWMSDB_BaseDoc.store()
        stage="Закрытие документа базы"
        gWMSDB_BaseDoc.close(True)
        gWMSDB_BaseDoc=Nothing
    End If
    Exit Sub
EH:
    detail=stage & ": " & CStr(Err) & " " & Error$
    gWMSDB_LastError=detail
    WMSDBX_BlockConnection detail
    MsgBox "Завершение работы с базой не подтверждено." & Chr(10) & detail & Chr(10) & _
        "Не повторяйте складскую операцию до проверки её записи и сохранения базы.",16,"ПОКАТАК — Сохранение базы"
End Sub

Function WMSDB_ResolveBaseURL(oCalcDoc As Object) As String
    Dim sURL As String
    WMSDB_ResolveBaseURL = ""
    On Error Resume Next
    sURL = WMSCore_GetSetting(oCalcDoc, WMSDB_SETTING_ODB_URL, "")
    On Error GoTo 0
    If Trim(sURL) <> "" Then
        WMSDB_ResolveBaseURL = WMSDB_NormalizeFileURL(sURL)
    Else
        WMSDB_ResolveBaseURL = WMSDB_DefaultBaseURL(oCalcDoc)
    End If
End Function

' Compatibility aliases for older DB modules.
Function WMSDB_ResolveFDBURL(oCalcDoc As Object) As String
    WMSDB_ResolveFDBURL = ""
End Function

Function WMSDB_DefaultBaseURL(oCalcDoc As Object) As String
    WMSDB_DefaultBaseURL = WMSDB_SiblingURL(oCalcDoc, WMSDB_DEFAULT_ODB)
End Function

Function WMSDB_DefaultFDBURL(oCalcDoc As Object) As String
    WMSDB_DefaultFDBURL = ""
End Function

Function WMSDB_SiblingURL(oCalcDoc As Object, sName As String) As String
    Dim sDocURL As String, p As Long
    WMSDB_SiblingURL = ""
    On Error GoTo Done
    sDocURL = oCalcDoc.URL
    If Trim(sDocURL) = "" Then Exit Function
    p = WMSDB_LastSlash(sDocURL)
    If p <= 0 Then Exit Function
    WMSDB_SiblingURL = Left(sDocURL, p) & sName
Done:
End Function

Function WMSDB_NormalizeFileURL(sPathOrURL As String) As String
    Dim s As String
    s = Trim(sPathOrURL)
    If Left(LCase(s), 7) = "file://" Then
        WMSDB_NormalizeFileURL = s
    Else
        WMSDB_NormalizeFileURL = ConvertToURL(s)
    End If
End Function

' Kept only for source compatibility. Portable mode does not call this.
Function WMSDB_FirebirdSDBCURL(sFDBURL As String) As String
    WMSDB_FirebirdSDBCURL = "sdbc:firebird:" & WMSDB_NormalizeFileURL(sFDBURL)
End Function

Function WMSDB_LastSlash(s As String) As Long
    Dim i As Long
    WMSDB_LastSlash = 0
    For i = Len(s) To 1 Step -1
        If Mid(s, i, 1) = "/" Then
            WMSDB_LastSlash = i
            Exit Function
        End If
    Next i
End Function

Function WMSDB_FileExists(sURL As String) As Boolean
    Dim oSFA As Object
    WMSDB_FileExists = False
    On Error GoTo Done
    oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    WMSDB_FileExists = oSFA.exists(WMSDB_NormalizeFileURL(sURL))
Done:
End Function

Function WMSDB_URLForDisplay(sURL As String) As String
    WMSDB_URLForDisplay = sURL
    On Error Resume Next
    WMSDB_URLForDisplay = ConvertFromURL(sURL)
    On Error GoTo 0
End Function

Function WMSDB_ExecuteUpdate(oCon As Object, sSQL As String, ByRef sErr As String) As Long
    Dim oStmt As Object
    WMSDB_ExecuteUpdate = -1
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    WMSDB_ExecuteUpdate = oStmt.executeUpdate(sSQL)
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDB_ScalarString(oCon As Object, sSQL As String, ByRef sErr As String) As String
    Dim oStmt As Object, oRS As Object
    WMSDB_ScalarString = ""
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If oRS.next() Then WMSDB_ScalarString = oRS.getString(1)
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDB_SQLText(v As Variant) As String
    Dim s As String
    If IsNull(v) Or IsEmpty(v) Then
        WMSDB_SQLText = "NULL"
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "'", "''")
    WMSDB_SQLText = "'" & s & "'"
End Function

Function WMSDB_TableExistsSimple(oCon As Object, sTable As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sSQL As String
    WMSDB_TableExistsSimple = False
    sErr = ""
    On Error GoTo EH
    sSQL = "SELECT RDB$RELATION_NAME FROM RDB$RELATIONS WHERE RDB$RELATION_NAME=" & _
           WMSDB_SQLText(UCase(Trim(sTable))) & " AND COALESCE(RDB$SYSTEM_FLAG,0)=0"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sSQL)
    If oRS.next() Then WMSDB_TableExistsSimple = (Trim(oRS.getString(1)) <> "")
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDB_UserTablesReport(oCon As Object, ByRef sErr As String) As String
    Dim oStmt As Object, oRS As Object, s As String, n As Long
    WMSDB_UserTablesReport = ""
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT RDB$RELATION_NAME FROM RDB$RELATIONS " & _
        "WHERE COALESCE(RDB$SYSTEM_FLAG,0)=0 AND RDB$VIEW_BLR IS NULL ORDER BY RDB$RELATION_NAME")
    Do While oRS.next()
        If Trim(oRS.getString(1)) <> "" Then
            If s <> "" Then s = s & ", "
            s = s & Trim(oRS.getString(1))
            n = n + 1
            If n >= 40 Then
                s = s & ", ..."
                Exit Do
            End If
        End If
    Loop
    If s = "" Then s = "(нет)"
    WMSDB_UserTablesReport = s
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

