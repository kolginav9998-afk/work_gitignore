Option Explicit

Global Const WMSDBX_VERSION = "3.3.0-SHARED-CONNECTION-SAFE"
Global gWMSDBX_Busy As Boolean
Global gWMSDBX_OpName As String
Global gWMSDBX_UnsafeConnection As Boolean
Global gWMSDBX_UnsafeReason As String

Function WMSDBX_TryEnter(opName As String) As Boolean
    If gWMSDBX_Busy Then
        WMSDBX_TryEnter=False
        Exit Function
    End If
    gWMSDBX_Busy=True
    gWMSDBX_OpName=opName
    WMSDBX_TryEnter=True
End Function

Sub WMSDBX_Leave()
    gWMSDBX_Busy=False
    gWMSDBX_OpName=""
End Sub

Sub WMSDBX_ResetBusy()
    WMSDBX_Leave
End Sub


Function WMSDBX_Begin(oCon As Object,ByRef sErr As String) As Boolean
    WMSDBX_Begin=False:sErr=""
    If gWMSDBX_UnsafeConnection Then sErr=gWMSDBX_UnsafeReason:Exit Function
    On Error GoTo EH
    If oCon.getAutoCommit() Then oCon.setAutoCommit(False)
    If oCon.getAutoCommit() Then
        sErr="Не подтверждён режим транзакции. Проведение остановлено до записи данных."
        Exit Function
    End If
    WMSDBX_Begin=True
    Exit Function
EH:
    sErr="Не удалось включить транзакцию: " & CStr(Err) & " " & Error$ & ". Проведение остановлено до записи данных."
End Function

Function WMSDBX_Commit(oCon As Object,ByRef sErr As String) As Boolean
    WMSDBX_Commit=False:sErr=""
    If gWMSDBX_UnsafeConnection Then sErr=gWMSDBX_UnsafeReason:Exit Function
    On Error GoTo EH
    oCon.commit()
    WMSDBX_Commit=True
    Exit Function
EH:
    sErr="Результат COMMIT не подтверждён: " & CStr(Err) & " " & Error$ & ". Не повторяйте операцию до проверки базы."
    WMSDBX_BlockConnection sErr
End Function

Function WMSDBX_SaveODB(ByRef sErr As String) As Boolean
    WMSDBX_SaveODB=WMSDB_SaveDatabaseDocument(sErr)
End Function

Function WMSDBX_Prepare(oCon As Object,sql As String,ByRef sErr As String) As Object
    sErr=""
    On Error GoTo EH
    WMSDBX_Prepare=oCon.prepareStatement(sql)
    Exit Function
EH:
    sErr="PrepareStatement: " & CStr(Err) & " " & Error$
End Function

Sub WMSDBX_CloseRS(oRS As Variant)
    On Error Resume Next
    oRS.close()
    On Error GoTo 0
End Sub

Sub WMSDBX_CloseStmt(oStmt As Variant)
    On Error Resume Next
    oStmt.close()
    On Error GoTo 0
End Sub

Sub WMSDBX_RollbackQuiet(oCon As Variant)
    Dim detail As String
    If IsNull(oCon) Or IsEmpty(oCon) Then Exit Sub
    If Not IsObject(oCon) Then Exit Sub
    On Error GoTo EH
    If oCon.isClosed() Then Exit Sub
    oCon.rollback()
    Exit Sub
EH:
    detail="Откат не подтверждён: " & CStr(Err) & " " & Error$ & ". Не повторяйте операцию до проверки базы."
    WMSDBX_BlockConnection detail
End Sub

Function WMSDBX_ScalarLong(oCon As Object,sql As String,ByRef sErr As String) As Long
    Dim st As Object,rs As Object
    WMSDBX_ScalarLong=0:sErr=""
    On Error GoTo EH
    st=oCon.createStatement():rs=st.executeQuery(sql)
    If rs.next() Then WMSDBX_ScalarLong=rs.getLong(1)
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="ScalarLong: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSDBX_ScalarDouble(oCon As Object,sql As String,ByRef sErr As String) As Double
    Dim st As Object,rs As Object
    WMSDBX_ScalarDouble=0:sErr=""
    On Error GoTo EH
    st=oCon.createStatement():rs=st.executeQuery(sql)
    If rs.next() Then WMSDBX_ScalarDouble=rs.getDouble(1)
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="ScalarDouble: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSDBX_ScalarText(oCon As Object,sql As String,ByRef sErr As String) As String
    Dim st As Object,rs As Object
    WMSDBX_ScalarText="":sErr=""
    On Error GoTo EH
    st=oCon.createStatement():rs=st.executeQuery(sql)
    If rs.next() Then WMSDBX_ScalarText=rs.getString(1)
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="ScalarText: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSDBX_CountByOperation(oCon As Object,opID As String,ByRef sErr As String) As Long
    Dim ps As Object,rs As Object
    WMSDBX_CountByOperation=0:sErr=""
    On Error GoTo EH
    ps=WMSDBX_Prepare(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE OPERATION_ID=?",sErr)
    If sErr<>"" Then Exit Function
    ps.setString(1,opID)
    rs=ps.executeQuery()
    If rs.next() Then WMSDBX_CountByOperation=rs.getLong(1)
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="Operation lookup: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSDBX_VerifyMovement(oCon As Object,movID As String,lotID As String,expectedQty As Double,ByRef sErr As String) As Boolean
    Dim ps As Object,rs As Object,n As Long,q As Double
    WMSDBX_VerifyMovement=False:sErr=""
    On Error GoTo EH
    ps=WMSDBX_Prepare(oCon,"SELECT COUNT(*),CAST(COALESCE(SUM(QTY),0) AS DOUBLE PRECISION) FROM WMS_STOCK_MOVEMENTS WHERE MOVEMENT_ID=? AND LOT_ID=? AND COALESCE(STATUS_NAME,'POSTED')='POSTED'",sErr)
    If sErr<>"" Then Exit Function
    ps.setString(1,movID):ps.setString(2,lotID)
    rs=ps.executeQuery()
    If rs.next() Then
        n=rs.getLong(1):q=rs.getDouble(2)
    End If
    If n<>1 Then sErr="Контрольное чтение движения не прошло.":GoTo Done
    If Abs(q-expectedQty)>0.0001 Then sErr="Контроль количества движения не прошёл.":GoTo Done
    WMSDBX_VerifyMovement=True
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="VerifyMovement: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSDBX_TestPrepared(oCon As Object,ByRef sErr As String) As Boolean
    Dim ps As Object,rs As Object
    WMSDBX_TestPrepared=False:sErr=""
    On Error GoTo EH
    ps=WMSDBX_Prepare(oCon,"SELECT CAST(? AS VARCHAR(40)) FROM RDB$DATABASE",sErr)
    If sErr<>"" Then Exit Function
    ps.setString(1,"WMS_OK")
    rs=ps.executeQuery()
    If rs.next() Then WMSDBX_TestPrepared=(rs.getString(1)="WMS_OK")
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt ps
    Exit Function
EH:
    sErr="PreparedStatement test: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSDBX_SyntaxProbe()
End Sub


Sub WMSDBX_BlockConnection(detail As String)
    gWMSDBX_UnsafeConnection=True
    If gWMSDBX_UnsafeReason="" Then gWMSDBX_UnsafeReason=detail
End Sub
