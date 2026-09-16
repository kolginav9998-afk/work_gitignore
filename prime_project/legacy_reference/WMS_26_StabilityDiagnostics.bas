Option Explicit

Global Const WMS26_VERSION = "3.1.0-DEEP-DIAGNOSTICS"
Global Const WMS26_SHEET = "Диагностика WMS"

Sub WMS26_RunDiagnostics()
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String,r As Long
    Dim nProblems As Long,sIntegrity As String,t0 As Double,ms As Long,prepOK As Boolean,report As String,path As String
    oDoc=ThisComponent
    If Not WMSDBX_TryEnter("DIAGNOSTICS") Then MsgBox "Другая операция WMS ещё выполняется.",48,"WMS":Exit Sub
    On Error GoTo EH
    If oDoc.Sheets.hasByName(WMS26_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMS26_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMS26_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMS26_SHEET)
    End If
    oSh.IsVisible=True
    oSh.getCellRangeByPosition(0,0,3,120).clearContents(1023)
    oSh.getCellByPosition(0,0).String="ДИАГНОСТИКА WMS"
    oSh.getCellByPosition(0,1).String="Только чтение. Проверяет именно ODB, который открыл текущий Calc."
    oSh.getCellByPosition(0,3).String="Параметр":oSh.getCellByPosition(1,3).String="Значение"
    r=4
    WMS26_Put oSh,r,"Версия диагностики",WMS26_VERSION
    WMS26_Put oSh,r,"ODS",WMSDB_URLForDisplay(oDoc.URL)
    WMS26_Put oSh,r,"ODB, который использует Calc",WMSDB_URLForDisplay(WMSDB_ResolveBaseURL(oDoc))

    t0=Timer
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    ms=CLng((Timer-t0)*1000)
    If sErr<>"" Then WMS26_Put oSh,r,"Подключение","ОШИБКА: " & sErr:GoTo Done
    WMS26_Put oSh,r,"Подключение","OK | " & CStr(ms) & " мс"

    prepOK=WMSDBX_TestPrepared(oCon,sErr)
    If prepOK Then
        WMS26_Put oSh,r,"PreparedStatement","OK"
    Else
        WMS26_Put oSh,r,"PreparedStatement","ОШИБКА: " & sErr
    End If
    sErr=""

    WMS26_Put oSh,r,"Товаров",WMS26_Scalar(oCon,"SELECT COUNT(*) FROM WMS_PRODUCTS",sErr)
    WMS26_Put oSh,r,"Движений",WMS26_Scalar(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS",sErr)
    WMS26_Put oSh,r,"Сумма QTY POSTED",WMS26_Scalar(oCon,"SELECT CAST(COALESCE(SUM(QTY),0) AS DOUBLE PRECISION) FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED'",sErr)
    WMS26_Put oSh,r,"Партий",WMS26_Scalar(oCon,"SELECT COUNT(*) FROM WMS_STOCK_LOTS",sErr)
    WMS26_Put oSh,r,"Положительных групп остатка",WMS26_Scalar(oCon,"SELECT COUNT(*) FROM (SELECT PRODUCT_CODE,UNIT_NAME,LOCATION_NAME,SUM(QTY) S FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' GROUP BY PRODUCT_CODE,UNIT_NAME,LOCATION_NAME HAVING SUM(QTY)>0.000001) X",sErr)
    WMS26_Put oSh,r,"Отрицательных групп остатка",WMS26_Scalar(oCon,"SELECT COUNT(*) FROM (SELECT PRODUCT_CODE,UNIT_NAME,LOCATION_NAME,SUM(QTY) S FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' GROUP BY PRODUCT_CODE,UNIT_NAME,LOCATION_NAME HAVING SUM(QTY)<-0.000001) X",sErr)
    WMS26_Put oSh,r,"Последнее движение",WMS26_Scalar(oCon,"SELECT FIRST 1 MOVEMENT_ID || ' | ' || PRODUCT_NAME || ' | ' || CAST(QTY AS VARCHAR(40)) || ' ' || UNIT_NAME FROM WMS_STOCK_MOVEMENTS ORDER BY CREATED_AT DESC",sErr)

    t0=Timer
    WMS26_Scalar oCon,"SELECT COUNT(*) FROM (SELECT PRODUCT_CODE,UNIT_NAME,LOCATION_NAME,SUM(QTY) FROM WMS_STOCK_MOVEMENTS WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' GROUP BY PRODUCT_CODE,UNIT_NAME,LOCATION_NAME) X",sErr
    ms=CLng((Timer-t0)*1000)
    WMS26_Put oSh,r,"Тест агрегирования остатка",CStr(ms) & " мс"

    sIntegrity=WMSINT_CheckText(oCon,nProblems,sErr)
    WMS26_Put oSh,r,"Integrity проблемных групп",CStr(nProblems)
    WMS26_Put oSh,r,"Integrity","см. ниже"
    WMSDB_Close

    oSh.getCellByPosition(0,r).String="ПРОВЕРКА ЦЕЛОСТНОСТИ":r=r+1
    WMS26_WriteMultiline oSh,r,sIntegrity
    oSh.getCellByPosition(0,r).String="ВЕС / СЛОЖНОСТЬ CALC":r=r+1
    WMS26_AppendWorkbookStats oDoc,oSh,r
Done:
    oSh.Columns.getByIndex(0).Width=7200:oSh.Columns.getByIndex(1).Width=17500
    oSh.getCellRangeByPosition(0,0,1,0).CellBackColor=2500134:oSh.getCellRangeByPosition(0,0,1,0).CharColor=16777215:oSh.getCellRangeByPosition(0,0,1,0).CharWeight=150
    oSh.getCellRangeByPosition(0,3,1,3).CellBackColor=4473924:oSh.getCellRangeByPosition(0,3,1,3).CharColor=16777215:oSh.getCellRangeByPosition(0,3,1,3).CharWeight=150
    report=WMS26_BuildReport(oSh,r)
    path=WMS26_SaveReport(report)
    If path<>"" Then oSh.getCellByPosition(0,2).String="TXT: " & path
    oDoc.CurrentController.setActiveSheet(oSh)
    WMSDBX_Leave
    Exit Sub
EH:
    On Error Resume Next:WMSDB_Close
    WMSDBX_Leave
    MsgBox "Диагностика: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMS26_Put(oSh As Object,ByRef r As Long,k As String,v As String)
    oSh.getCellByPosition(0,r).String=k:oSh.getCellByPosition(1,r).String=v:r=r+1
End Sub

Sub WMS26_WriteMultiline(oSh As Object,ByRef r As Long,s As String)
    Dim a As Variant,i As Long
    a=Split(s,Chr(10))
    For i=LBound(a) To UBound(a)
        If Trim(CStr(a(i)))<>"" Then oSh.getCellByPosition(0,r).String=CStr(a(i)):r=r+1
    Next i
End Sub

Function WMS26_Scalar(oCon As Object,sql As String,ByRef sErr As String) As String
    WMS26_Scalar=WMSDBX_ScalarText(oCon,sql,sErr)
    If sErr<>"" Then WMS26_Scalar="ОШИБКА: " & sErr
End Function

Sub WMS26_AppendWorkbookStats(oDoc As Object,oOut As Object,ByRef r As Long)
    Dim i As Long,oSh As Object,cur As Object,lastR As Long,lastC As Long,nShapes As Long
    On Error Resume Next
    For i=0 To oDoc.Sheets.getCount()-1
        oSh=oDoc.Sheets.getByIndex(i)
        cur=oSh.createCursor():cur.gotoEndOfUsedArea(True)
        lastR=cur.RangeAddress.EndRow:lastC=cur.RangeAddress.EndColumn
        nShapes=oSh.DrawPage.getCount()
        oOut.getCellByPosition(0,r).String=oSh.Name
        oOut.getCellByPosition(1,r).String="used=" & CStr(lastR+1) & "x" & CStr(lastC+1) & " | shapes=" & CStr(nShapes)
        r=r+1
    Next i
    On Error GoTo 0
End Sub

Function WMS26_BuildReport(oSh As Object,lastRow As Long) As String
    Dim i As Long,s As String
    s="WMS DIAGNOSTICS " & Format(Now,"YYYY-MM-DD HH:MM:SS") & Chr(10)
    For i=0 To lastRow
        If Trim(oSh.getCellByPosition(0,i).String)<>"" Then
            s=s & oSh.getCellByPosition(0,i).String
            If Trim(oSh.getCellByPosition(1,i).String)<>"" Then s=s & " = " & oSh.getCellByPosition(1,i).String
            s=s & Chr(10)
        End If
    Next i
    WMS26_BuildReport=s
End Function

Function WMS26_SaveReport(report As String) As String
    Dim baseDir As String,dirPath As String,filePath As String,sfa As Object,outSt As Object,txt As Object
    WMS26_SaveReport=""
    On Error GoTo EH
    baseDir=WMSACT_BaseFolder():If baseDir="" Then Exit Function
    dirPath=baseDir & "Diagnostics" & WMSACT_PathSep()
    WMSACT_EnsureFolder dirPath
    filePath=dirPath & "WMS_DIAGNOSTICS_" & Format(Now,"YYYYMMDD_HHMMSS") & ".txt"
    sfa=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    outSt=sfa.openFileWrite(ConvertToURL(filePath))
    txt=CreateUnoService("com.sun.star.io.TextOutputStream")
    txt.setOutputStream(outSt):txt.setEncoding("UTF-8"):txt.writeString(report):txt.closeOutput()
    WMS26_SaveReport=filePath
    Exit Function
EH:
    WMS26_SaveReport=""
End Function

Sub WMS26_OptimizeWorkbook()
    Dim oDoc As Object,i As Long,oSh As Object
    oDoc=ThisComponent
    On Error Resume Next
    If oDoc.Sheets.hasByName(WMS26_SHEET) Then oDoc.Sheets.getByName(WMS26_SHEET).IsVisible=False
    For i=0 To oDoc.Sheets.getCount()-1
        oSh=oDoc.Sheets.getByIndex(i)
        WMS26_RemoveOrphanButtons oSh
    Next i
    On Error GoTo 0
End Sub

Sub WMS26_RemoveOrphanButtons(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i):nm="":ctl=shape.Control:nm=ctl.Name
        If Left(nm,8)="WMS_OLD_" Or Left(nm,8)="WMS_DBS_" Or Left(nm,8)="WMS_MAN_" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub
