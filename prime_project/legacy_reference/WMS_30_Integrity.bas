Option Explicit

Global Const WMSINT_VERSION = "3.1.0-INTEGRITY-PERFORMANCE"

Function WMSINT_EnsureSchema(oCon As Object,ByRef sErr As String) As Boolean
    Dim st As Object
    WMSINT_EnsureSchema=False:sErr=""
    On Error GoTo EH
    st=oCon.createStatement()
    WMSINT_CreateIndex oCon,st,"IX_WMS_MOV_STATUS_PRODUCT_LOC","WMS_STOCK_MOVEMENTS(STATUS_NAME,PRODUCT_CODE,LOCATION_NAME)",sErr:If sErr<>"" Then GoTo Done
    WMSINT_CreateIndex oCon,st,"IX_WMS_MOV_FLOW","WMS_STOCK_MOVEMENTS(FLOW_CHANNEL)",sErr:If sErr<>"" Then GoTo Done
    WMSINT_CreateIndex oCon,st,"IX_WMS_MOV_LOT2","WMS_STOCK_MOVEMENTS(LOT_ID)",sErr:If sErr<>"" Then GoTo Done
    If WMSDB_TableExistsSimple(oCon,"WMS_PRODUCT_ALIASES",sErr) Then
        WMSINT_CreateIndex oCon,st,"IX_WMS_ALIAS_ARTICLE","WMS_PRODUCT_ALIASES(ARTICLE_CODE)",sErr:If sErr<>"" Then GoTo Done
    ElseIf sErr<>"" Then
        GoTo Done
    End If
    oCon.commit()
    WMSINT_EnsureSchema=True
Done:
    WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="Integrity schema: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSINT_CreateIndex(oCon As Object,st As Object,indexName As String,target As String,ByRef sErr As String)
    If WMSDBST_IndexExists(oCon,indexName,sErr) Then Exit Sub
    If sErr<>"" Then Exit Sub
    On Error GoTo EH
    st.executeUpdate("CREATE INDEX " & indexName & " ON " & target)
    Exit Sub
EH:
    sErr="Индекс " & indexName & ": " & CStr(Err) & " " & Error$
End Sub

Function WMSINT_CheckText(oCon As Object,ByRef nProblems As Long,ByRef sErr As String) As String
    Dim s As String,n As Long,e As String,d As Double
    nProblems=0:sErr="":s=""

    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE PRODUCT_CODE IS NULL OR TRIM(PRODUCT_CODE)=''",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Движения без кода: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE QTY=0",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Нулевые движения: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=m.PRODUCT_CODE WHERE p.PRODUCT_CODE IS NULL",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Движения без номенклатуры: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID IS NULL OR TRIM(LOT_ID)=''",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Движения без партии: " & CStr(n) & Chr(10)

    n=WMSDBX_ScalarLong(oCon,"SELECT COUNT(*) FROM (SELECT MOVEMENT_ID,COUNT(*) C FROM WMS_STOCK_MOVEMENTS GROUP BY MOVEMENT_ID HAVING COUNT(*)>1) X",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Дубли MOVEMENT_ID: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    d=WMSDBX_ScalarDouble(oCon,"SELECT CAST(COALESCE(SUM(QTY),0) AS DOUBLE PRECISION) FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED'",e)
    If e<>"" Then sErr=e:Exit Function
    s=s & "Контрольная сумма QTY: " & CStr(d) & Chr(10)

    WMSINT_CheckText=s
End Function
