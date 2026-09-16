Option Explicit

' ============================================================================
' WMS_05_DB_Install.bas
' v0.3.0 — creates WMS_DATA_PORTABLE.odb with EMBEDDED Firebird.
'
' This version intentionally avoids creation of an external .fdb. On Windows
' some Firebird configurations reject CREATE DATABASE outside allowed folders
' (DatabaseAccess restriction). Embedded Firebird is managed by LibreOffice
' Base itself and does not require Firebird server reconfiguration.
' ============================================================================

Global Const WMSDBI_VERSION = "0.3.0-PORTABLE"
Global Const WMSDBI_SCHEMA = "1"
Global Const WMSDBI_MODULE = "DB_Install"
Global Const WMSDBI_APP_ID = "WMS-LITE"
Global Const WMSDBI_STORAGE_MODE = "EMBEDDED_FIREBIRD_PORTABLE"

Global gWMSDBI_LastError As String
Global gWMSDBI_LastReport As String

Sub WMSDBI_InstallPortableDatabase()
    Dim sReport As String
    If WMSDBI_InstallPortableCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — переносимая Firebird"
    Else
        MsgBox sReport, 16, "WMS — переносимая Firebird"
    End If
End Sub

' Compatibility entry points: old menu/procedure names still work.
Sub WMSDBI_InstallExternalDatabase()
    WMSDBI_InstallPortableDatabase
End Sub

Sub WMSDBI_InstallDatabase()
    WMSDBI_InstallPortableDatabase
End Sub

Sub WMSDBI_SelfCheck()
    Dim sReport As String
    If WMSDBI_SelfCheckCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — DB SelfCheck"
    Else
        MsgBox sReport, 16, "WMS — DB SelfCheck"
    End If
End Sub

Sub WMSDBI_RoundTripTest()
    Dim sReport As String
    If WMSDBI_RoundTripCore(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — DB RoundTrip"
    Else
        MsgBox sReport, 16, "WMS — DB RoundTrip"
    End If
End Sub

Function WMSDBI_CompileProbe() As String
    WMSDBI_CompileProbe = WMSDBI_VERSION & "|PORTABLE|" & WMSDBI_SCHEMA
End Function

Function WMSDBI_InstallPortableCore(oCalcDoc As Object, ByRef sReport As String) As Boolean
    Dim sODBURL As String, sErr As String
    Dim oCon As Object
    Dim bCreated As Boolean

    WMSDBI_InstallPortableCore = False
    sReport = ""
    On Error GoTo EH

    If Trim(oCalcDoc.URL) = "" Then
        sReport = "Сначала сохраните WMS как .ods. База создаётся рядом с файлом."
        Exit Function
    End If

    sODBURL = WMSDB_DefaultBaseURL(oCalcDoc)
    WMSDB_Close

    If Not WMSDB_FileExists(sODBURL) Then
        If Not WMSDBI_CreateEmbeddedODB(sODBURL, sErr) Then
            sReport = "Не удалось создать WMS_DATA_PORTABLE.odb:" & Chr(10) & sErr
            Exit Function
        End If
        bCreated = True
    End If

    On Error Resume Next
    WMSCore_SetSetting oCalcDoc, WMSDB_SETTING_ODB_URL, sODBURL, "Переносимая Base/Firebird БД", True
    On Error GoTo EH

    oCon = WMSDB_GetConnectionEx(oCalcDoc, sErr)
    If sErr <> "" Then
        sReport = "Base-файл создан, но embedded Firebird не открылся:" & Chr(10) & sErr
        Exit Function
    End If

    If Not WMSDBI_EnsureSchema(oCon, sErr) Then
        sReport = "Не удалось создать служебную схему:" & Chr(10) & sErr
        Exit Function
    End If

    If Not WMSDBI_RoundTripWithConnection(oCon, sErr) Then
        sReport = "Схема создана, но RoundTrip не прошёл:" & Chr(10) & sErr
        Exit Function
    End If

    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        sReport = "Данные записались, но Base-файл не сохранился:" & Chr(10) & sErr
        Exit Function
    End If

    WMSDB_Close

    If Not WMSDBI_ReconnectPersistenceCheck(oCalcDoc, sErr) Then
        sReport = "База создана, но проверка после полного закрытия/открытия не прошла:" & Chr(10) & sErr
        Exit Function
    End If

    WMSDB_Close

    On Error Resume Next
    WMSCore_SetSetting oCalcDoc, "WMS.DB.Schema.Version", WMSDBI_SCHEMA, "Версия схемы Firebird", True
    WMSCore_SetSetting oCalcDoc, "WMS.DB.Status", "PORTABLE_READY", "Статус переносимой БД", True
    WMSCore_SetSetting oCalcDoc, "WMS.DB.LastCheck", WMSDBI_Stamp(Now), "Последняя успешная проверка БД", True
    On Error GoTo EH

    sReport = "Переносимая база WMS готова." & Chr(10) & _
              "Base: " & WMSDB_URLForDisplay(sODBURL) & Chr(10) & _
              "Firebird embedded: OK" & Chr(10) & _
              "RoundTrip: OK" & Chr(10) & _
              "Полное закрытие/повторное открытие: OK" & Chr(10) & _
              "Схема: " & WMSDBI_SCHEMA & Chr(10) & _
              "Настройки Firebird Server не требуются."

    gWMSDBI_LastReport = sReport
    WMSDBI_InstallPortableCore = True
    Exit Function
EH:
    sReport = "Критическая ошибка установки переносимой БД: " & CStr(Err) & " " & Error$
    gWMSDBI_LastError = sReport
    On Error Resume Next
    WMSDB_Close
End Function

Function WMSDBI_CreateEmbeddedODB(sODBURL As String, ByRef sErr As String) As Boolean
    Dim oBase As Object
    Dim aLoad(0) As New com.sun.star.beans.PropertyValue
    Dim aStore() As New com.sun.star.beans.PropertyValue

    WMSDBI_CreateEmbeddedODB = False
    sErr = ""
    On Error GoTo EH

    If WMSDB_FileExists(sODBURL) Then
        WMSDBI_CreateEmbeddedODB = True
        Exit Function
    End If

    aLoad(0).Name = "Hidden"
    aLoad(0).Value = True
    oBase = StarDesktop.loadComponentFromURL("private:factory/sdatabase", "_blank", 0, aLoad())
    If IsNull(oBase) Or IsEmpty(oBase) Then
        sErr = "LibreOffice Base не смог создать новый документ базы данных."
        Exit Function
    End If

    oBase.DataSource.URL = "sdbc:embedded:firebird"
    oBase.storeAsURL(sODBURL, aStore())
    oBase.close(True)

    If Not WMSDB_FileExists(sODBURL) Then
        sErr = "LibreOffice сообщил о сохранении, но WMS_DATA_PORTABLE.odb не найден на диске."
        Exit Function
    End If

    WMSDBI_CreateEmbeddedODB = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
    On Error Resume Next
    If Not (IsNull(oBase) Or IsEmpty(oBase)) Then oBase.close(True)
End Function

Function WMSDBI_EnsureSchema(oCon As Object, ByRef sErr As String) As Boolean
    WMSDBI_EnsureSchema = False
    sErr = ""
    On Error GoTo EH

    If Not WMSDBI_TableExists(oCon, "WMS_META") Then
        If Not WMSDBI_Exec(oCon, _
            "CREATE TABLE WMS_META (" & _
            "META_KEY VARCHAR(80) NOT NULL PRIMARY KEY," & _
            "META_VALUE VARCHAR(255)," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)", sErr) Then Exit Function
    End If

    If Not WMSDBI_TableExists(oCon, "WMS_SYNC_TEST") Then
        If Not WMSDBI_Exec(oCon, _
            "CREATE TABLE WMS_SYNC_TEST (" & _
            "SOURCE_ID VARCHAR(96) NOT NULL PRIMARY KEY," & _
            "PAYLOAD VARCHAR(255)," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)", sErr) Then Exit Function
    End If

    WMSDBI_PutMeta oCon, "APP_ID", WMSDBI_APP_ID
    WMSDBI_PutMeta oCon, "SCHEMA_VERSION", WMSDBI_SCHEMA
    WMSDBI_PutMeta oCon, "INSTALLER_VERSION", WMSDBI_VERSION
    WMSDBI_PutMeta oCon, "CONNECTION_VERSION", WMSDB_VERSION
    WMSDBI_PutMeta oCon, "STORAGE_MODE", WMSDBI_STORAGE_MODE
    oCon.commit()

    WMSDBI_EnsureSchema = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
    On Error Resume Next
    oCon.rollback()
End Function

Function WMSDBI_TableExists(oCon As Object, sTable As String) As Boolean
    Dim sErr As String
    WMSDBI_TableExists = WMSDB_TableExistsSimple(oCon, sTable, sErr)
End Function

Function WMSDBI_Exec(oCon As Object, sSQL As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSDBI_Exec = False
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oStmt.execute(sSQL)
    WMSDBI_Exec = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Sub WMSDBI_PutMeta(oCon As Object, sKey As String, sValue As String)
    Dim oStmt As Object, n As Long
    On Error GoTo EH
    oStmt = oCon.createStatement()
    n = oStmt.executeUpdate("UPDATE WMS_META SET META_VALUE=" & WMSDB_SQLText(sValue) & _
        ", UPDATED_AT=CURRENT_TIMESTAMP WHERE META_KEY=" & WMSDB_SQLText(sKey))
    If n = 0 Then
        oStmt.executeUpdate("INSERT INTO WMS_META (META_KEY,META_VALUE,UPDATED_AT) VALUES (" & _
            WMSDB_SQLText(sKey) & "," & WMSDB_SQLText(sValue) & ",CURRENT_TIMESTAMP)")
    End If
    Exit Sub
EH:
    Err = 0
End Sub

Function WMSDBI_SelfCheckCore(oCalcDoc As Object, ByRef sReport As String) As Boolean
    Dim oCon As Object, sErr As String, sSchema As String, sApp As String, sMode As String
    WMSDBI_SelfCheckCore = False
    sReport = ""
    On Error GoTo EH

    oCon = WMSDB_GetConnectionEx(oCalcDoc, sErr)
    If sErr <> "" Then sReport = sErr : Exit Function

    If Not WMSDBI_TableExists(oCon, "WMS_META") Then sReport = "WMS_META отсутствует." : Exit Function
    If Not WMSDBI_TableExists(oCon, "WMS_SYNC_TEST") Then sReport = "WMS_SYNC_TEST отсутствует." : Exit Function

    sApp = WMSDB_ScalarString(oCon, "SELECT META_VALUE FROM WMS_META WHERE META_KEY='APP_ID'", sErr)
    If sErr <> "" Or sApp <> WMSDBI_APP_ID Then sReport = "APP_ID не соответствует WMS: " & sApp : Exit Function

    sMode = WMSDB_ScalarString(oCon, "SELECT META_VALUE FROM WMS_META WHERE META_KEY='STORAGE_MODE'", sErr)
    If sErr <> "" Or sMode <> WMSDBI_STORAGE_MODE Then sReport = "Режим хранения не подтверждён: " & sMode : Exit Function

    sSchema = WMSDB_ScalarString(oCon, "SELECT META_VALUE FROM WMS_META WHERE META_KEY='SCHEMA_VERSION'", sErr)
    If sErr <> "" Then sReport = sErr : Exit Function

    If Not WMSDBI_RoundTripWithConnection(oCon, sErr) Then sReport = "RoundTrip FAIL: " & sErr : Exit Function
    If Not WMSDB_SaveDatabaseDocument(sErr) Then sReport = "Save FAIL: " & sErr : Exit Function
    WMSDB_Close

    If Not WMSDBI_ReconnectPersistenceCheck(oCalcDoc, sErr) Then sReport = "Reconnect FAIL: " & sErr : Exit Function
    WMSDB_Close

    sReport = "DB SelfCheck: OK" & Chr(10) & _
              "Storage: " & WMSDBI_STORAGE_MODE & Chr(10) & _
              "Schema: " & sSchema & Chr(10) & _
              "RoundTrip: OK" & Chr(10) & _
              "Close/reopen persistence: OK" & Chr(10) & _
              "Base: " & WMSDB_URLForDisplay(WMSDB_ResolveBaseURL(oCalcDoc))
    WMSDBI_SelfCheckCore = True
    Exit Function
EH:
    sReport = "DB SelfCheck: " & CStr(Err) & " " & Error$
End Function

Function WMSDBI_RoundTripCore(oCalcDoc As Object, ByRef sReport As String) As Boolean
    Dim oCon As Object, sErr As String
    WMSDBI_RoundTripCore = False
    oCon = WMSDB_GetConnectionEx(oCalcDoc, sErr)
    If sErr <> "" Then sReport = sErr : Exit Function
    If WMSDBI_RoundTripWithConnection(oCon, sErr) Then
        If Not WMSDB_SaveDatabaseDocument(sErr) Then
            sReport = "RoundTrip прошёл, но сохранение Base FAIL: " & sErr
            Exit Function
        End If
        sReport = "RoundTrip portable embedded Firebird: OK."
        WMSDBI_RoundTripCore = True
    Else
        sReport = "RoundTrip: FAIL — " & sErr
    End If
End Function

Function WMSDBI_RoundTripWithConnection(oCon As Object, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object
    Dim sID As String, sPayload As String, sRead As String
    WMSDBI_RoundTripWithConnection = False
    sErr = ""
    On Error GoTo EH

    sID = "SELFTEST-ROUNDTRIP"
    sPayload = "WMS portable test " & WMSDBI_Stamp(Now)
    oStmt = oCon.createStatement()

    On Error Resume Next
    oStmt.executeUpdate("DELETE FROM WMS_SYNC_TEST WHERE SOURCE_ID=" & WMSDB_SQLText(sID))
    oCon.commit()
    Err = 0
    On Error GoTo EH

    oStmt.executeUpdate("INSERT INTO WMS_SYNC_TEST (SOURCE_ID,PAYLOAD,UPDATED_AT) VALUES (" & _
        WMSDB_SQLText(sID) & "," & WMSDB_SQLText(sPayload) & ",CURRENT_TIMESTAMP)")
    oCon.commit()

    oRS = oStmt.executeQuery("SELECT PAYLOAD FROM WMS_SYNC_TEST WHERE SOURCE_ID=" & WMSDB_SQLText(sID))
    If Not oRS.next() Then sErr = "INSERT выполнен, но запись не прочиталась." : GoTo CleanupFail
    sRead = oRS.getString(1)
    If sRead <> sPayload Then sErr = "Прочитанное значение отличается." : GoTo CleanupFail

    oStmt.executeUpdate("DELETE FROM WMS_SYNC_TEST WHERE SOURCE_ID=" & WMSDB_SQLText(sID))
    oCon.commit()
    WMSDBI_RoundTripWithConnection = True
    Exit Function

CleanupFail:
    On Error Resume Next
    oStmt.executeUpdate("DELETE FROM WMS_SYNC_TEST WHERE SOURCE_ID=" & WMSDB_SQLText(sID))
    oCon.commit()
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
    On Error Resume Next
    oCon.rollback()
End Function

Function WMSDBI_ReconnectPersistenceCheck(oCalcDoc As Object, ByRef sErr As String) As Boolean
    Dim oCon As Object, sProbe As String
    WMSDBI_ReconnectPersistenceCheck = False
    sErr = ""
    On Error GoTo EH

    WMSDB_Close
    oCon = WMSDB_GetConnectionEx(oCalcDoc, sErr)
    If sErr <> "" Then Exit Function

    If Not WMSDBI_TableExists(oCon, "WMS_META") Then
        sErr = "После полного повторного открытия WMS_META отсутствует."
        Exit Function
    End If

    sProbe = WMSDB_ScalarString(oCon, "SELECT META_VALUE FROM WMS_META WHERE META_KEY='STORAGE_MODE'", sErr)
    If sErr <> "" Then Exit Function
    If sProbe <> WMSDBI_STORAGE_MODE Then
        sErr = "После повторного открытия STORAGE_MODE=" & sProbe
        Exit Function
    End If

    WMSDBI_ReconnectPersistenceCheck = True
    Exit Function
EH:
    sErr = CStr(Err) & " " & Error$
End Function

Function WMSDBI_Stamp(v As Variant) As String
    WMSDBI_Stamp = Right("0" & CStr(Day(v)),2) & "." & Right("0" & CStr(Month(v)),2) & "." & CStr(Year(v)) & " " & _
                   Right("0" & CStr(Hour(v)),2) & ":" & Right("0" & CStr(Minute(v)),2) & ":" & Right("0" & CStr(Second(v)),2)
End Function

