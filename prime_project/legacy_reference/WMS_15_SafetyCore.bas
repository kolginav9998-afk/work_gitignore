Option Explicit

Global Const WMSSAFE_VERSION = "0.3.1-OPERATION-GUARD-1.0.2"

Sub WMSSAFE_Install()
    Dim oCon As Object,sErr As String
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS Safety":Exit Sub
    If Not WMSSAFE_EnsureSchema(oCon,sErr) Then WMSDB_Close:MsgBox sErr,16,"WMS Safety":Exit Sub
    On Error Resume Next:oCon.commit():On Error GoTo 0
    If Not WMSDB_SaveDatabaseDocument(sErr) Then WMSDB_Close:MsgBox sErr,16,"WMS Safety":Exit Sub
    WMSDB_Close
    MsgBox "Safety Core установлен." & Chr(10) & _
           "Защита повторов, аудит, корректировки и диагностика готовы.",64,"WMS Safety"
End Sub

Function WMSSAFE_EnsureSchema(oCon As Object,ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSSAFE_EnsureSchema=False:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()

    If Not WMSSAFE_TableExists(oCon,"WMS_AUDIT_LOG") Then
        oStmt.execute("CREATE TABLE WMS_AUDIT_LOG (" & _
            "AUDIT_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "EVENT_TIME TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "ACTION_NAME VARCHAR(80) NOT NULL," & _
            "ENTITY_TYPE VARCHAR(40)," & _
            "ENTITY_ID VARCHAR(128)," & _
            "RESULT_STATE VARCHAR(40)," & _
            "DETAIL_TEXT VARCHAR(1000))")
    End If

    If Not WMSSAFE_TableExists(oCon,"WMS_OPERATION_GUARD") Then
        oStmt.execute("CREATE TABLE WMS_OPERATION_GUARD (" & _
            "GUARD_KEY VARCHAR(220) NOT NULL PRIMARY KEY," & _
            "ENTITY_TYPE VARCHAR(40)," & _
            "ENTITY_ID VARCHAR(128)," & _
            "OPERATION_NAME VARCHAR(80)," & _
            "STATE_NAME VARCHAR(40)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "UPDATED_AT TIMESTAMP)")
    End If

    If Not WMSSAFE_TableExists(oCon,"WMS_CORRECTIONS") Then
        oStmt.execute("CREATE TABLE WMS_CORRECTIONS (" & _
            "CORRECTION_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "SOURCE_ENTITY_TYPE VARCHAR(40) NOT NULL," & _
            "SOURCE_ENTITY_ID VARCHAR(128) NOT NULL," & _
            "CORRECTION_TYPE VARCHAR(40) NOT NULL," & _
            "REASON_TEXT VARCHAR(500) NOT NULL," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)")
    End If

    On Error Resume Next
    If WMSSAFE_TableExists(oCon,"WMS_META") Then
        If oStmt.executeUpdate("UPDATE WMS_META SET META_VALUE='1.0.2' WHERE META_KEY='RELEASE_VERSION'")=0 Then
            oStmt.executeUpdate("INSERT INTO WMS_META (META_KEY,META_VALUE) VALUES ('RELEASE_VERSION','1.0.2')")
        End If
        If oStmt.executeUpdate("UPDATE WMS_META SET META_VALUE='OPERATION_GUARD_V1' WHERE META_KEY='HARDENING_LEVEL'")=0 Then
            oStmt.executeUpdate("INSERT INTO WMS_META (META_KEY,META_VALUE) VALUES ('HARDENING_LEVEL','OPERATION_GUARD_V1')")
        End If
    End If
    On Error GoTo EH

    WMSSAFE_EnsureSchema=True
    Exit Function
EH:
    sErr="Safety schema: " & CStr(Err) & " " & Error$
End Function

Function WMSSAFE_TableExists(oCon As Object,t As String) As Boolean
    Dim oStmt As Object,oRS As Object
    WMSSAFE_TableExists=False
    On Error GoTo Done
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COUNT(*) FROM RDB$RELATIONS WHERE TRIM(RDB$RELATION_NAME)=" & WMSDB_SQLText(UCase(t)) & _
        " AND COALESCE(RDB$SYSTEM_FLAG,0)=0")
    If oRS.next() Then WMSSAFE_TableExists=(oRS.getInt(1)>0)
Done:
End Function


Function WMSSAFE_GuardBegin(entityType As String,entityID As String,operationName As String,ByRef sErr As String) As Boolean
    Dim oCon As Object
    WMSSAFE_GuardBegin=False
    sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSSAFE_GuardBeginCon(oCon,entityType,entityID,operationName,sErr) Then
        On Error Resume Next
        oCon.commit()
        WMSDB_Close
        Exit Function
    End If
    oCon.commit()
    WMSDB_Close
    WMSSAFE_GuardBegin=True
    Exit Function
EH:
    sErr="GuardBegin: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

Sub WMSSAFE_GuardFinish(entityType As String,entityID As String,operationName As String,ok As Boolean,detailText As String)
    Dim oCon As Object
    Dim sErr As String
    On Error GoTo Done
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Sub
    WMSSAFE_GuardFinishCon oCon,entityType,entityID,operationName,ok,detailText
    oCon.commit()
Done:
    On Error Resume Next
    WMSDB_Close
End Sub

Function WMSSAFE_GuardBeginCon(oCon As Object,entityType As String,entityID As String,operationName As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,k As String,stateName As String
    WMSSAFE_GuardBeginCon=False:sErr=""
    On Error GoTo EH
    If Not WMSSAFE_EnsureSchema(oCon,sErr) Then Exit Function
    k=UCase(Trim(entityType)) & "|" & Trim(entityID) & "|" & UCase(Trim(operationName))
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT STATE_NAME FROM WMS_OPERATION_GUARD WHERE GUARD_KEY=" & WMSDB_SQLText(k))
    If oRS.next() Then
        stateName=UCase(Trim(oRS.getString(1)))
        If stateName="DONE" Or stateName="RUNNING" Then
            sErr="Операция уже " & IIf(stateName="DONE","выполнена","выполняется") & ". Повтор заблокирован."
            WMSSAFE_AuditCon oCon,"DUPLICATE_BLOCKED",entityType,entityID,"BLOCKED",operationName
            Exit Function
        End If
        oStmt.executeUpdate("UPDATE WMS_OPERATION_GUARD SET STATE_NAME='RUNNING',UPDATED_AT=CURRENT_TIMESTAMP WHERE GUARD_KEY=" & WMSDB_SQLText(k))
    Else
        oStmt.executeUpdate("INSERT INTO WMS_OPERATION_GUARD (GUARD_KEY,ENTITY_TYPE,ENTITY_ID,OPERATION_NAME,STATE_NAME) VALUES (" & _
            WMSDB_SQLText(k) & "," & WMSDB_SQLText(entityType) & "," & WMSDB_SQLText(entityID) & "," & _
            WMSDB_SQLText(operationName) & ",'RUNNING')")
    End If
    WMSSAFE_GuardBeginCon=True
    Exit Function
EH:
    sErr="GuardBegin: " & CStr(Err) & " " & Error$
End Function

Sub WMSSAFE_GuardFinishCon(oCon As Object,entityType As String,entityID As String,operationName As String,ok As Boolean,detailText As String)
    Dim oStmt As Object,k As String,st As String
    On Error Resume Next
    k=UCase(Trim(entityType)) & "|" & Trim(entityID) & "|" & UCase(Trim(operationName))
    st=IIf(ok,"DONE","FAILED")
    oStmt=oCon.createStatement()
    oStmt.executeUpdate("UPDATE WMS_OPERATION_GUARD SET STATE_NAME=" & WMSDB_SQLText(st) & _
        ",UPDATED_AT=CURRENT_TIMESTAMP WHERE GUARD_KEY=" & WMSDB_SQLText(k))
    WMSSAFE_AuditCon oCon,operationName,entityType,entityID,st,detailText
End Sub

Sub WMSSAFE_AuditCon(oCon As Object,actionName As String,entityType As String,entityID As String,resultState As String,detailText As String)
    Dim oStmt As Object,id As String
    On Error Resume Next
    id=WMSSAFE_NewID("AUD")
    oStmt=oCon.createStatement()
    oStmt.executeUpdate("INSERT INTO WMS_AUDIT_LOG (AUDIT_ID,ACTION_NAME,ENTITY_TYPE,ENTITY_ID,RESULT_STATE,DETAIL_TEXT) VALUES (" & _
        WMSDB_SQLText(id) & "," & WMSDB_SQLText(actionName) & "," & WMSDB_SQLText(entityType) & "," & _
        WMSDB_SQLText(entityID) & "," & WMSDB_SQLText(resultState) & "," & WMSDB_SQLText(Left(detailText,1000)) & ")")
End Sub

Function WMSSAFE_ConfirmOverReceipt(orderNo As String,orderedQty As Double,receivedBefore As Double,nowQty As Double) As Boolean
    Dim afterQty As Double,msg As String
    afterQty=receivedBefore+nowQty
    If orderedQty<=0 Or afterQty<=orderedQty+0.0000001 Then
        WMSSAFE_ConfirmOverReceipt=True
        Exit Function
    End If
    msg="ВНИМАНИЕ: перепоставка." & Chr(10) & Chr(10) & _
        "Заказ: " & orderNo & Chr(10) & _
        "Заказано: " & CStr(orderedQty) & Chr(10) & _
        "Получено ранее: " & CStr(receivedBefore) & Chr(10) & _
        "Сейчас: " & CStr(nowQty) & Chr(10) & _
        "После проведения: " & CStr(afterQty) & Chr(10) & _
        "Превышение: " & CStr(afterQty-orderedQty) & Chr(10) & Chr(10) & _
        "Разрешить реальную перепоставку?"
    WMSSAFE_ConfirmOverReceipt=(MsgBox(msg,36,"WMS — Контроль количества")=6)
End Function

Sub WMSSAFE_CheckIncomplete()
    Dim oCon As Object,oStmt As Object,oRS As Object,sErr As String,s As String,n As Long
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Диагностика":Exit Sub
    If Not WMSSAFE_EnsureSchema(oCon,sErr) Then WMSDB_Close:MsgBox sErr,16,"WMS — Диагностика":Exit Sub

    s="Незавершённые операции:" & Chr(10) & Chr(10)
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT GUARD_KEY,STATE_NAME,UPDATED_AT FROM WMS_OPERATION_GUARD WHERE STATE_NAME IN ('RUNNING','FAILED') ORDER BY UPDATED_AT DESC")
    Do While oRS.next()
        n=n+1
        If n<=30 Then s=s & oRS.getString(1) & " | " & oRS.getString(2) & " | " & oRS.getString(3) & Chr(10)
    Loop
    WMSDB_Close
    If n=0 Then s=s & "Не найдено."
    MsgBox s,64,"WMS — Незавершённые операции"
End Sub

Sub WMSSAFE_ShowAudit()
    Dim oCon As Object,oStmt As Object,oRS As Object,sErr As String,s As String,n As Long
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Аудит":Exit Sub
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT FIRST 30 EVENT_TIME,ACTION_NAME,ENTITY_TYPE,ENTITY_ID,RESULT_STATE,DETAIL_TEXT FROM WMS_AUDIT_LOG ORDER BY EVENT_TIME DESC")
    Do While oRS.next()
        n=n+1
        s=s & oRS.getString(1) & " | " & oRS.getString(2) & " | " & oRS.getString(3) & Chr(10) & _
            oRS.getString(4) & " | " & oRS.getString(5) & Chr(10) & oRS.getString(6) & Chr(10) & Chr(10)
    Loop
    WMSDB_Close
    If n=0 Then s="Журнал пока пуст."
    MsgBox s,64,"WMS — Последние действия"
End Sub

Function WMSSAFE_NewID(prefix As String) As String
    Randomize
    WMSSAFE_NewID=prefix & "-" & Format(Now,"YYYYMMDD-HHMMSS") & "-" & Right("000000" & CStr(CLng(Rnd()*999999)),6)
End Function



Sub WMSSAFE_Audit(actionName As String,entityType As String,entityID As String,resultState As String,detailText As String)
    Dim oCon As Object
    Dim sErr As String

    On Error GoTo Done
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Sub
    If Not WMSSAFE_EnsureSchema(oCon,sErr) Then
        WMSDB_Close
        Exit Sub
    End If

    WMSSAFE_AuditCon oCon,actionName,entityType,entityID,resultState,detailText
    On Error Resume Next
    oCon.commit()
    On Error GoTo 0
    WMSDB_SaveDatabaseDocument sErr
    WMSDB_Close
Done:
End Sub
