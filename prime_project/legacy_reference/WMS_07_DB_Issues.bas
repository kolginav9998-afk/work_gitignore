Option Explicit

' ============================================================================
' WMS_07_DB_Issues.bas
' v0.1.0 — Safe Firebird persistence for sheet "Выдачи".
'
' Normal lifecycle:
'   Fill issue in Calc -> "Провести выдачу" -> validate -> Firebird -> COMMIT
'   -> close/reopen DB -> verify SOURCE_ID -> remove row from Calc.
'
' Returnable issues are NOT lost: RETURNABLE is stored in Firebird.
' Future return UI will read open returnable issues from Firebird and write
' return events into WMS_ISSUE_RETURNS.
' ============================================================================

Global Const WMSDBIU_VERSION = "0.4.3-GUARD-SYNTAX-FIX"
Global Const WMSDBIU_SHEET = "Выдачи"
Global Const WMSDBIU_TABLE = "WMS_ISSUES"
Global Const WMSDBIU_RETURNS = "WMS_ISSUE_RETURNS"

Sub WMSDBIu_Install()
    Dim s As String
    If WMSDBIu_InstallCore(ThisComponent,s) Then
        MsgBox s,64,"WMS — БД Выдачи"
    Else
        MsgBox s,16,"WMS — БД Выдачи"
    End If
End Sub

Sub WMSDBIu_SelfCheck()
    Dim oCon As Object,sErr As String,s As String,b1 As Boolean,b2 As Boolean
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — БД Выдачи":Exit Sub
    b1=WMSDB_TableExistsSimple(oCon,WMSDBIU_TABLE,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — БД Выдачи":Exit Sub
    b2=WMSDB_TableExistsSimple(oCon,WMSDBIU_RETURNS,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — БД Выдачи":Exit Sub
    s="Соединение: OK" & Chr(10) & _
      WMSDBIU_TABLE & ": " & IIf(b1,"OK","НЕТ") & Chr(10) & _
      WMSDBIU_RETURNS & ": " & IIf(b2,"OK","НЕТ") & Chr(10) & _
      "Модуль: " & WMSDBIU_VERSION
    WMSDB_Close
    MsgBox s,IIf(b1 And b2,64,48),"WMS — БД Выдачи"
End Sub

Function WMSDBIu_InstallCore(oDoc As Object,ByRef sReport As String) As Boolean
    Dim oCon As Object,sErr As String
    WMSDBIu_InstallCore=False:sReport=""
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBIU_SHEET) Then
        sReport="В книге нет листа 'Выдачи'."
        Exit Function
    End If

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then sReport=sErr:Exit Function

    If Not WMSDBIu_EnsureTables(oCon,sErr) Then sReport=sErr:Exit Function
    If Not WMSDB_SaveDatabaseDocument(sErr) Then sReport=sErr:Exit Function

    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then sReport=sErr:Exit Function
    If Not WMSDB_TableExistsSimple(oCon,WMSDBIU_TABLE,sErr) Then
        sReport="После повторного открытия таблица " & WMSDBIU_TABLE & " не найдена."
        WMSDB_Close
        Exit Function
    End If
    ' 0.2.0: preserve exactly which stock bucket was issued.
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"ALLOC_CATEGORY",sErr) Then
        If sErr<>"" Then Exit Function
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD ALLOC_CATEGORY VARCHAR(120)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"ALLOC_SUBCATEGORY",sErr) Then
        If sErr<>"" Then Exit Function
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD ALLOC_SUBCATEGORY VARCHAR(120)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"STOCK_ORIGIN",sErr) Then
        If sErr<>"" Then Exit Function
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD STOCK_ORIGIN VARCHAR(120)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"LOT_ID",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD LOT_ID VARCHAR(128)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"ENTRY_QTY",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD ENTRY_QTY DECIMAL(18,4)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"ENTRY_UNIT",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD ENTRY_UNIT VARCHAR(64)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"BASE_QTY",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD BASE_QTY DECIMAL(18,4)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBIU_TABLE,"BASE_UNIT",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBIU_TABLE & " ADD BASE_UNIT VARCHAR(64)",sErr
        If sErr<>"" Then Exit Function
    End If

    If Not WMSDB_TableExistsSimple(oCon,WMSDBIU_RETURNS,sErr) Then
        sReport="После повторного открытия таблица " & WMSDBIU_RETURNS & " не найдена."
        WMSDB_Close
        Exit Function
    End If
    WMSDB_Close

    sReport="БД для листа 'Выдачи' готова." & Chr(10) & _
            WMSDBIU_TABLE & ": OK" & Chr(10) & _
            WMSDBIU_RETURNS & ": OK" & Chr(10) & _
            "Модуль: " & WMSDBIU_VERSION
    WMSDBIu_InstallCore=True
    Exit Function
EH:
    sReport="Установка БД Выдачи: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

Function WMSDBIu_EnsureTables(oCon As Object,ByRef sErr As String) As Boolean
    Dim sql As String,n As Long
    WMSDBIu_EnsureTables=False:sErr=""

    If Not WMSDB_TableExistsSimple(oCon,WMSDBIU_TABLE,sErr) Then
        If sErr<>"" Then Exit Function
        sql="CREATE TABLE " & WMSDBIU_TABLE & " (" & _
            "SOURCE_ID VARCHAR(96) NOT NULL PRIMARY KEY," & _
            "ISSUE_NO INTEGER," & _
            "PRODUCT_CODE VARCHAR(128)," & _
            "PRODUCT_NAME VARCHAR(512) NOT NULL," & _
            "ISSUE_QTY DECIMAL(18,4) NOT NULL," & _
            "UNIT_NAME VARCHAR(64)," & _
            "EMPLOYEE_NAME VARCHAR(256) NOT NULL," & _
            "ISSUE_DATE DATE," & _
            "SOURCE_LOCATION VARCHAR(256)," & _
            "RETURNABLE VARCHAR(8)," & _
            "RETURNED_QTY DECIMAL(18,4) DEFAULT 0 NOT NULL," & _
            "RETURN_DATE DATE," & _
            "NOTE_TEXT VARCHAR(1024)," & _
            "ROW_MODE VARCHAR(64)," & _
            "ROW_STATE VARCHAR(64)," & _
            "RETURN_STATE VARCHAR(64)," & _
            "BUSINESS_HASH VARCHAR(128)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)"
        n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
        If sErr<>"" Then Exit Function
    End If

    If Not WMSDB_TableExistsSimple(oCon,WMSDBIU_RETURNS,sErr) Then
        If sErr<>"" Then Exit Function
        sql="CREATE TABLE " & WMSDBIU_RETURNS & " (" & _
            "RETURN_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "SOURCE_ID VARCHAR(96) NOT NULL," & _
            "RETURN_SEQ INTEGER NOT NULL," & _
            "RETURN_QTY DECIMAL(18,4) NOT NULL," & _
            "RETURN_DATE DATE," & _
            "NOTE_TEXT VARCHAR(1024)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)"
        n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
        If sErr<>"" Then Exit Function
    End If

    WMSDBIu_EnsureTables=True
End Function

Sub WMSDBIu_ConductSelectedIssue()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long
    Dim sErr As String,sid As String,msg As String,ans As Integer
    Dim a As Variant,productName As String,employeeName As String,qty As Double
    Dim code As String,unitText As String,loc As String,dIssue As Double,noteText As String
    Dim lotID As String,sCat As String,sSub As String,sOrigin As String
    Dim baseQty As Double,baseUnit As String
    Dim actRowsText As String,actPartyTo As String
    Dim guardStarted As Boolean

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSDBIU_SHEET) Then MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Провести выдачу":Exit Sub
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSDBIU_SHEET Then MsgBox "Откройте лист 'Выдачи' и выберите строку.",48,"WMS — Провести выдачу":Exit Sub

    On Error GoTo EH
    oSel=oDoc.CurrentSelection:r=oSel.CellAddress.Row
    If r<1 Or Not Issues_RowHasBusinessData(oSh,r) Then MsgBox "Выберите рабочую строку выдачи.",48,"WMS — Провести выдачу":Exit Sub

    If Not Issues_API_ValidateForPosting(r,msg) Then
        If msg="" Then msg="Строка не прошла проверку."
        MsgBox msg,48,"WMS — Провести выдачу":Exit Sub
    End If

    a=Issues_API_GetPayload(r)
    sid=Trim(CStr(a(0))):code=Trim(CStr(a(1))):productName=Trim(CStr(a(2)))
    qty=CDbl(a(3)):unitText=Trim(CStr(a(4))):employeeName=Trim(CStr(a(5)))
    dIssue=CDbl(a(6)):loc=Trim(CStr(a(7))):noteText=Trim(CStr(a(11)))

    If Not WMSDBIu_SelectLotForIssue(oDoc,code,loc,qty,unitText,True,lotID,baseQty,baseUnit,sCat,sSub,sOrigin,sErr) Then
        MsgBox sErr,48,"WMS — Партия выдачи":Exit Sub
    End If

    ans=MsgBox("Провести выдачу?" & Chr(10) & Chr(10) & _
               productName & Chr(10) & _
               "Выдача: " & CStr(qty) & " " & unitText & Chr(10) & _
               "Из остатка: " & CStr(baseQty) & " " & baseUnit & Chr(10) & _
               "Партия: " & lotID & Chr(10) & _
               "Получил: " & employeeName,36,"WMS — Провести выдачу")
    If ans<>6 Then Exit Sub

    If Not WMSSAFE_GuardBegin("ISSUE",sid,"CONDUCT",sErr) Then
        MsgBox sErr,48,"WMS — Защита операции"
        Exit Sub
    End If
    guardStarted=True

    If Not WMSDBIu_SaveIssue(oDoc,oSh,r,a,sErr) Then
        WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,sErr
        MsgBox sErr,16,"WMS — Провести выдачу"
        Exit Sub
    End If
    If Not WMSDBIu_SaveLotAllocation(oDoc,sid,lotID,qty,unitText,baseQty,baseUnit,sCat,sSub,sOrigin,sErr) Then
        WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,sErr
        MsgBox "Выдача записана, но партия не закреплена. Строка НЕ удалена." & Chr(10) & sErr,16,"WMS — Провести выдачу"
        Exit Sub
    End If
    If Not WMSDBIu_PostLotIssueMovement(oDoc,"OUT-" & sid,sid,lotID,code,productName,qty,unitText,baseQty,baseUnit, _
        loc,sCat,sSub,sOrigin,dIssue,"Выдано: " & employeeName & IIf(noteText="",""," | " & noteText),sErr) Then
        WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,sErr
        MsgBox "Списание партии НЕ подтверждено. Строка НЕ удалена." & Chr(10) & sErr,16,"WMS — Провести выдачу"
        Exit Sub
    End If
    If Not WMSDBIu_VerifyLotIssueAfterReconnect(oDoc,sid,"OUT-" & sid,lotID,baseQty,sErr) Then
        WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,sErr
        MsgBox "Финальная проверка Firebird не пройдена. Строка НЕ удалена." & Chr(10) & sErr,16,"WMS — Провести выдачу"
        Exit Sub
    End If

    WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",True,productName & " | " & employeeName
    guardStarted=False

    actRowsText=WMSACT_IssueRowText(oSh,r)
    actPartyTo=employeeName

    If Not WMSDBIu_RemoveIssueCellsSafely(oDoc,oSh,r,sErr) Then
        MsgBox "БД сохранена, но строку очистить не удалось. Вручную не удаляйте." & Chr(10) & sErr,48,"WMS — Провести выдачу":Exit Sub
    End If

    WMSSAFE_Audit "ISSUE_CONDUCTED","ISSUE",sid,"DONE",productName & " | " & employeeName

    MsgBox "Выдача проведена." & Chr(10) & _
           "Выдано: " & CStr(qty) & " " & unitText & Chr(10) & _
           "Списано: " & CStr(baseQty) & " " & baseUnit & Chr(10) & _
           "Партия: " & lotID,64,"WMS — Провести выдачу"

    If Trim(actRowsText)<>"" Then
        WMSACT_AskAndCreate "OUT","Склад",actPartyTo,"Акт приёма-передачи ТМЦ",actRowsText,sid
    End If
    Exit Sub
EH:
    If guardStarted And sid<>"" Then WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,CStr(Err) & " " & Error$
    gWMSISS_Busy=False
    MsgBox "Проведение выдачи остановлено: " & CStr(Err) & " " & Error$,16,"WMS — Провести выдачу"
End Sub


' ============================================================================
' BULK CONDUCT 0.3.0
'
' Safe algorithm:
'   1) No row is deleted while Firebird writes are in progress.
'   2) Each valid row is persisted using the already-tested issue + stock APIs.
'   3) Rows with ambiguous stock buckets are NOT guessed; they remain in Calc.
'   4) Firebird is closed/reopened ONCE after the batch.
'   5) Only rows whose issue + OUT movement are found after reconnect are removed.
'   6) Calc rows are removed from bottom to top and sheet repaint is done once.
'
' This deliberately prioritizes data safety over maximum raw speed.
' ============================================================================

Sub WMSDBIu_ConductAllIssues()
    Dim oDoc As Object,oSh As Object,last As Long,r As Long
    Dim nBusiness As Long,nPrepared As Long,nDeleted As Long,nErrors As Long
    Dim sErr As String,sRowErr As String,sReport As String
    Dim aRows() As Long,aSids() As String,aMovs() As String
    Dim cap As Long,oCon As Object,i As Long
    Dim ans As Integer

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSDBIU_SHEET) Then
        MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Провести все"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSDBIU_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    If last<1 Then
        MsgBox "Нет выдач для проведения.",48,"WMS — Провести все"
        Exit Sub
    End If

    For r=1 To last
        If Issues_RowHasBusinessData(oSh,r) Then nBusiness=nBusiness+1
    Next r
    If nBusiness=0 Then
        MsgBox "Нет выдач для проведения.",48,"WMS — Провести все"
        Exit Sub
    End If

    ans=MsgBox("Провести все готовые выдачи?" & Chr(10) & Chr(10) & _
               "Строк с данными: " & CStr(nBusiness) & Chr(10) & _
               "Строки с ошибками или несколькими доступными остатками останутся в Calc.", _
               36,"WMS — Провести все")
    If ans<>6 Then Exit Sub

    cap=nBusiness
    ReDim aRows(1 To cap)
    ReDim aSids(1 To cap)
    ReDim aMovs(1 To cap)

    gWMSISS_Busy=True
    On Error GoTo EH

    ' Phase 1: persist, but DO NOT remove rows.
    For r=1 To last
        If Issues_RowHasBusinessData(oSh,r) Then
            sRowErr=""
            Dim sid As String,movementID As String
            sid="":movementID=""
            If WMSDBIu_PersistIssueRowForBulk(oDoc,oSh,r,sid,movementID,sRowErr) Then
                nPrepared=nPrepared+1
                aRows(nPrepared)=r
                aSids(nPrepared)=sid
                aMovs(nPrepared)=movementID
            Else
                nErrors=nErrors+1
                If Len(sReport)<7000 Then
                    sReport=sReport & "Строка " & CStr(r+1) & ": " & sRowErr & Chr(10)
                End If
            End If
        End If
    Next r

    If nPrepared=0 Then
        gWMSISS_Busy=False
        MsgBox "Ни одна выдача не проведена." & _
               IIf(sReport="","",Chr(10) & Chr(10) & sReport),48,"WMS — Провести все"
        Exit Sub
    End If

    ' Phase 2: one durability reconnect for the whole batch.
    WMSDB_Close
    sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then
        gWMSISS_Busy=False
        MsgBox "Записи отправлены в Firebird, но финальная проверка после reconnect не выполнена." & _
               Chr(10) & "Строки НЕ удалены." & Chr(10) & Chr(10) & sErr,16,"WMS — Провести все"
        Exit Sub
    End If

    ' Mark rows that really exist after reconnect.
    Dim aVerified() As Boolean
    ReDim aVerified(1 To nPrepared)
    For i=1 To nPrepared
        sErr=""
        If WMSDBIu_RecordExists(oCon,aSids(i),sErr) Then
            If WMSDBST_VerifyMovementCon(oCon,aMovs(i),sErr) Then
                aVerified(i)=True
            End If
        End If
        If aVerified(i) Then
            WMSSAFE_GuardFinish "ISSUE",aSids(i),"CONDUCT",True,"Bulk verified"
        Else
            WMSSAFE_GuardFinish "ISSUE",aSids(i),"CONDUCT",False,"Bulk durability verification failed"
            nErrors=nErrors+1
            If Len(sReport)<7000 Then
                sReport=sReport & "Строка " & CStr(aRows(i)+1) & _
                    ": Firebird не подтвердил выдачу/движение после reconnect." & Chr(10)
            End If
        End If
    Next i
    WMSDB_Close

    ' Phase 3: delete only verified rows, bottom to top.
    For i=nPrepared To 1 Step -1
        If aVerified(i) Then
            sErr=""
            If WMSDBIu_RemoveIssueCellsBulk(oSh,aRows(i),sErr) Then
                nDeleted=nDeleted+1
            Else
                nErrors=nErrors+1
                If Len(sReport)<7000 Then
                    sReport=sReport & "Строка " & CStr(aRows(i)+1) & _
                        ": БД сохранена, но строка не очищена: " & sErr & Chr(10)
                End If
            End If
        End If
    Next i

    ' One final UI refresh instead of repainting after every deleted row.
    WMSDBIu_RenumberVisibleIssues oSh
    WMSDBIu_RepaintVisibleIssues oSh
    Issues_ResetQueueCache
    Issues_RefreshFilter oDoc,oSh
    gWMSISS_Busy=False
    WMSSAFE_Audit "ISSUE_BULK_CONDUCTED","BATCH","ISSUES-BULK-" & Format(Now,"YYYYMMDD-HHMMSS"), _
        IIf(nErrors=0,"DONE","PARTIAL"),"Проведено: " & CStr(nDeleted) & "; ошибок: " & CStr(nErrors)

    MsgBox "Массовое проведение завершено." & Chr(10) & _
           "Проведено и удалено из Calc: " & CStr(nDeleted) & Chr(10) & _
           "Осталось для проверки: " & CStr(nBusiness-nDeleted) & _
           IIf(sReport="","",Chr(10) & Chr(10) & "Причины:" & Chr(10) & sReport), _
           IIf(nErrors=0,64,48),"WMS — Провести все"
    Exit Sub
EH:
    gWMSISS_Busy=False
    On Error Resume Next
    WMSDB_Close
    MsgBox "Массовое проведение остановлено: " & CStr(Err) & " " & Error$ & Chr(10) & _
           "Неподтверждённые строки из Calc автоматически не удалялись.",16,"WMS — Провести все"
End Sub

Function WMSDBIu_PersistIssueRowForBulk(oDoc As Object,oSh As Object,r As Long, _
    ByRef sid As String,ByRef movementID As String,ByRef sErr As String) As Boolean

    Dim a As Variant,msg As String,code As String,productName As String
    Dim employeeName As String,unitText As String,loc As String,noteText As String
    Dim qty As Double,dIssue As Double
    Dim lotID As String,sCat As String,sSub As String,sOrigin As String
    Dim baseQty As Double,baseUnit As String

    WMSDBIu_PersistIssueRowForBulk=False:sErr="":sid="":movementID=""
    On Error GoTo EH

    If Not Issues_API_ValidateForPosting(r,msg) Then
        If msg="" Then msg="Строка не прошла проверку."
        sErr=msg:Exit Function
    End If

    a=Issues_API_GetPayload(r)
    sid=Trim(CStr(a(0))):code=Trim(CStr(a(1))):productName=Trim(CStr(a(2)))
    qty=CDbl(a(3)):unitText=Trim(CStr(a(4))):employeeName=Trim(CStr(a(5)))
    dIssue=CDbl(a(6)):loc=Trim(CStr(a(7))):noteText=Trim(CStr(a(11)))

    If Not WMSDBIu_SelectLotForIssue(oDoc,code,loc,qty,unitText,False,lotID,baseQty,baseUnit,sCat,sSub,sOrigin,sErr) Then GoTo GuardFail
    If Not WMSDBIu_SaveIssue(oDoc,oSh,r,a,sErr) Then GoTo GuardFail
    If Not WMSDBIu_SaveLotAllocation(oDoc,sid,lotID,qty,unitText,baseQty,baseUnit,sCat,sSub,sOrigin,sErr) Then GoTo GuardFail

    movementID="OUT-" & sid
    If Not WMSDBIu_PostLotIssueMovement(oDoc,movementID,sid,lotID,code,productName,qty,unitText,baseQty,baseUnit, _
        loc,sCat,sSub,sOrigin,dIssue,"Выдано: " & employeeName & IIf(noteText="",""," | " & noteText),sErr) Then GoTo GuardFail

    WMSDBIu_PersistIssueRowForBulk=True
    Exit Function
GuardFail:
    WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,sErr
    Exit Function
EH:
    If sid<>"" Then WMSSAFE_GuardFinish "ISSUE",sid,"CONDUCT",False,CStr(Err) & " " & Error$
    sErr="Строка " & CStr(r+1) & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_RemoveIssueCellsBulk(oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim lastCol As Long,lastRow As Long,addr As Variant
    WMSDBIu_RemoveIssueCellsBulk=False
    sErr=""
    On Error GoTo EH

    lastCol=Issues_LastHeaderCol(oSh)
    If lastCol<WMSISS_USER_LAST_COL Then lastCol=WMSISS_USER_LAST_COL
    lastRow=Issues_LastContentRow(oSh,lastCol)

    If r<1 Or r>lastRow Then
        sErr="Некорректная строка: " & CStr(r+1)
        Exit Function
    End If

    addr=oSh.getCellRangeByPosition(0,r,lastCol,r).RangeAddress
    oSh.removeRange(addr,com.sun.star.sheet.CellDeleteMode.UP)

    WMSDBIu_RemoveIssueCellsBulk=True
    Exit Function
EH:
    sErr=CStr(Err) & " " & Error$
End Function


' ============================================================================
' LOT-AWARE ISSUE HELPERS
' ============================================================================
Function WMSDBIu_SelectLotForIssue(oDoc As Object,code As String,loc As String,entryQty As Double,entryUnit As String, _
    interactive As Boolean,ByRef lotID As String,ByRef baseQty As Double,ByRef baseUnit As String, _
    ByRef sCat As String,ByRef sSub As String,ByRef sOrigin As String,ByRef sErr As String) As Boolean

    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String
    Dim ids() As String,bqs() As Double,bus() As String,cats() As String,subs() As String,orgs() As String
    Dim avs() As Double,n As Long,need As Double,sList As String,choice As String,idx As Long

    WMSDBIu_SelectLotForIssue=False:sErr=""
    lotID="":baseQty=0:baseUnit="":sCat="":sSub="":sOrigin=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    sql="SELECT l.LOT_ID,l.BASE_UNIT,u.FACTOR_TO_BASE," & _
        "COALESCE(l.DEST_CATEGORY,''),COALESCE(l.DEST_SUBCATEGORY,''),COALESCE(l.ORIGIN_NAME,'')," & _
        "COALESCE(SUM(m.QTY),0),l.RECEIPT_DATE " & _
        "FROM WMS_STOCK_LOTS l " & _
        "JOIN WMS_LOT_UNITS u ON u.LOT_ID=l.LOT_ID " & _
        "LEFT JOIN WMS_STOCK_MOVEMENTS m ON m.LOT_ID=l.LOT_ID " & _
        "WHERE l.PRODUCT_CODE=" & WMSDB_SQLText(code) & _
        " AND l.LOCATION_NAME=" & WMSDB_SQLText(loc) & _
        " AND u.UNIT_NAME=" & WMSDB_SQLText(entryUnit) & _
        " GROUP BY l.LOT_ID,l.BASE_UNIT,u.FACTOR_TO_BASE,l.DEST_CATEGORY,l.DEST_SUBCATEGORY,l.ORIGIN_NAME,l.RECEIPT_DATE " & _
        "HAVING COALESCE(SUM(m.QTY),0)>0 " & _
        "ORDER BY l.RECEIPT_DATE,l.LOT_ID"

    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        need=entryQty*oRS.getDouble(3)
        If oRS.getDouble(7)+0.0000001>=need Then
            n=n+1
            ReDim Preserve ids(1 To n):ReDim Preserve bqs(1 To n):ReDim Preserve bus(1 To n)
            ReDim Preserve cats(1 To n):ReDim Preserve subs(1 To n):ReDim Preserve orgs(1 To n):ReDim Preserve avs(1 To n)
            ids(n)=oRS.getString(1):bus(n)=oRS.getString(2)
            cats(n)=oRS.getString(4):subs(n)=oRS.getString(5):orgs(n)=oRS.getString(6):avs(n)=oRS.getDouble(7)
            bqs(n)=need
            If n<=30 Then sList=sList & CStr(n) & ". " & ids(n) & _
                " | " & WMSDBST_Num(entryQty) & " " & entryUnit & _
                " = " & WMSDBST_Num(need) & " " & bus(n) & _
                " | доступно " & WMSDBST_Num(avs(n)) & " " & bus(n) & Chr(10)
        End If
    Loop

    If n=0 Then
        sErr="Нет одной партии, которая может выдать " & WMSDBST_Num(entryQty) & " " & entryUnit & _
             " из места " & loc & "." & Chr(10) & _
             "Если количество лежит в нескольких партиях — разделите выдачу на две строки."
        Exit Function
    End If

    If n=1 Or Not interactive Then
        idx=1
    Else
        choice=InputBox("Доступно несколько партий. Выберите:" & Chr(10) & Chr(10) & sList, _
                        "WMS — Партия выдачи","1")
        If Trim(choice)="" Then sErr="Выбор партии отменён.":Exit Function
        If Not IsNumeric(choice) Then sErr="Введите номер партии.":Exit Function
        idx=CLng(choice)
        If idx<1 Or idx>n Or idx>30 Then sErr="Нет такой партии.":Exit Function
    End If

    lotID=ids(idx):baseQty=bqs(idx):baseUnit=bus(idx)
    sCat=cats(idx):sSub=subs(idx):sOrigin=orgs(idx)
    WMSDBIu_SelectLotForIssue=True
    Exit Function
EH:
    sErr="Выбор партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_SaveLotAllocation(oDoc As Object,sid As String,lotID As String,entryQty As Double,entryUnit As String, _
    baseQty As Double,baseUnit As String,sCat As String,sSub As String,sOrigin As String,ByRef sErr As String) As Boolean
    Dim oCon As Object,sql As String,n As Long
    WMSDBIu_SaveLotAllocation=False:sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    sql="UPDATE " & WMSDBIU_TABLE & " SET LOT_ID=" & WMSDB_SQLText(lotID) & _
        ",ENTRY_QTY=" & WMSDBIu_NumberSQL(entryQty) & _
        ",ENTRY_UNIT=" & WMSDB_SQLText(entryUnit) & _
        ",BASE_QTY=" & WMSDBIu_NumberSQL(baseQty) & _
        ",BASE_UNIT=" & WMSDB_SQLText(baseUnit) & _
        ",ALLOC_CATEGORY=" & WMSDB_SQLText(sCat) & _
        ",ALLOC_SUBCATEGORY=" & WMSDB_SQLText(sSub) & _
        ",STOCK_ORIGIN=" & WMSDB_SQLText(sOrigin) & _
        ",UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
    If sErr<>"" Or n<=0 Then If sErr="" Then sErr="Не удалось закрепить партию.":Exit Function
    oCon.commit()
    WMSDBIu_SaveLotAllocation=True
    Exit Function
EH:
    sErr="Сохранение партии выдачи: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_PostLotIssueMovement(oDoc As Object,movementID As String,sid As String,lotID As String, _
    code As String,nm As String,entryQty As Double,entryUnit As String,baseQty As Double,baseUnit As String, _
    loc As String,sCat As String,sSub As String,sOrigin As String,dIssue As Double,noteText As String,ByRef sErr As String) As Boolean

    Dim oCon As Object,oStmt As Object,sql As String,avail As Double
    WMSDBIu_PostLotIssueMovement=False:sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBIu_GetLotAvailableCon(oCon,lotID,avail,sErr) Then Exit Function
    If avail+0.0000001<baseQty Then
        sErr="В партии " & lotID & " осталось " & WMSDBST_Num(avail) & " " & baseUnit & _
             ", требуется " & WMSDBST_Num(baseQty) & " " & baseUnit & "."
        Exit Function
    End If

    If Not WMSDBST_PostMovementClassified(oDoc,movementID,sid,"ISSUE","OUT",code,nm,-baseQty,baseUnit, _
        loc,sOrigin,sCat,sSub,dIssue,noteText,sErr) Then Exit Function

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    sql="UPDATE WMS_STOCK_MOVEMENTS SET LOT_ID=" & WMSDB_SQLText(lotID) & _
        ",ENTRY_QTY=" & WMSDBIu_NumberSQL(entryQty) & _
        ",ENTRY_UNIT=" & WMSDB_SQLText(entryUnit) & _
        ",BASE_QTY=" & WMSDBIu_NumberSQL(baseQty) & _
        ",BASE_UNIT=" & WMSDB_SQLText(baseUnit) & _
        " WHERE MOVEMENT_ID=" & WMSDB_SQLText(movementID)
    oStmt=oCon.createStatement():oStmt.executeUpdate(sql):oCon.commit()
    WMSDBIu_PostLotIssueMovement=True
    Exit Function
EH:
    sErr="Списание партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_GetLotAvailableCon(oCon As Object,lotID As String,ByRef avail As Double,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object
    WMSDBIu_GetLotAvailableCon=False:sErr="":avail=0
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COALESCE(SUM(QTY),0) FROM WMS_STOCK_MOVEMENTS WHERE LOT_ID=" & WMSDB_SQLText(lotID))
    If oRS.next() Then avail=oRS.getDouble(1)
    WMSDBIu_GetLotAvailableCon=True
    Exit Function
EH:
    sErr="Остаток партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_VerifyLotIssueAfterReconnect(oDoc As Object,sid As String,movementID As String,lotID As String,expectedBaseQty As Double,ByRef sErr As String) As Boolean
    Dim oCon As Object,oStmt As Object,oRS As Object
    WMSDBIu_VerifyLotIssueAfterReconnect=False:sErr=""
    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COALESCE(LOT_ID,''),BASE_QTY FROM " & WMSDBIU_TABLE & " WHERE SOURCE_ID=" & WMSDB_SQLText(sid))
    If Not oRS.next() Then sErr="Выдача не найдена после reconnect.":Exit Function
    If Trim(oRS.getString(1))<>lotID Or Abs(oRS.getDouble(2)-expectedBaseQty)>0.0001 Then
        sErr="Партия или базовое количество выдачи не совпало.":Exit Function
    End If
    oRS=oStmt.executeQuery("SELECT COALESCE(LOT_ID,''),QTY FROM WMS_STOCK_MOVEMENTS WHERE MOVEMENT_ID=" & WMSDB_SQLText(movementID))
    If Not oRS.next() Then sErr="OUT-движение не найдено после reconnect.":Exit Function
    If Trim(oRS.getString(1))<>lotID Or Abs(oRS.getDouble(2)+expectedBaseQty)>0.0001 Then
        sErr="OUT-движение не подтверждает нужную партию.":Exit Function
    End If
    WMSDB_Close
    WMSDBIu_VerifyLotIssueAfterReconnect=True
    Exit Function
EH:
    sErr="Проверка партии: " & CStr(Err) & " " & Error$
    On Error Resume Next:WMSDB_Close
End Function

Function WMSDBIu_SaveIssue(oDoc As Object,oSh As Object,r As Long,a As Variant,ByRef sErr As String) As Boolean
    Dim oCon As Object,sid As String,sql As String,n As Long,bExists As Boolean,sHash As String
    WMSDBIu_SaveIssue=False:sErr=""
    On Error GoTo EH

    sid=Trim(CStr(a(0)))
    sHash=Issues_CellText(oSh,r,"_WMS_BaseFingerprint")

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBIu_EnsureTables(oCon,sErr) Then Exit Function

    bExists=WMSDBIu_RecordExists(oCon,sid,sErr)
    If sErr<>"" Then Exit Function

    On Error Resume Next
    oCon.setAutoCommit(False)
    On Error GoTo EH

    If bExists Then
        sql="UPDATE " & WMSDBIU_TABLE & " SET " & WMSDBIu_IssueSetSQL(oSh,r,a,sHash) & _
            ",UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    Else
        sql="INSERT INTO " & WMSDBIU_TABLE & " (" & _
            "SOURCE_ID,ISSUE_NO,PRODUCT_CODE,PRODUCT_NAME,ISSUE_QTY,UNIT_NAME,EMPLOYEE_NAME," & _
            "ISSUE_DATE,SOURCE_LOCATION,RETURNABLE,RETURNED_QTY,RETURN_DATE,NOTE_TEXT,ROW_MODE,ROW_STATE," & _
            "RETURN_STATE,BUSINESS_HASH,CREATED_AT,UPDATED_AT) VALUES (" & _
            WMSDB_SQLText(sid) & "," & WMSDBIu_IntSQL(oSh.getCellByPosition(0,r).Value) & "," & _
            WMSDB_SQLText(CStr(a(1))) & "," & WMSDB_SQLText(CStr(a(2))) & "," & _
            WMSDBIu_NumberSQL(CDbl(a(3))) & "," & WMSDB_SQLText(CStr(a(4))) & "," & _
            WMSDB_SQLText(CStr(a(5))) & "," & WMSDBIu_DateSQL(CDbl(a(6))) & "," & _
            WMSDB_SQLText(CStr(a(7))) & "," & WMSDB_SQLText(CStr(a(8))) & "," & _
            WMSDBIu_NumberSQL(CDbl(a(9))) & "," & WMSDBIu_DateSQL(CDbl(a(10))) & "," & _
            WMSDB_SQLText(CStr(a(11))) & "," & WMSDB_SQLText(CStr(a(12))) & "," & _
            WMSDB_SQLText(CStr(a(13))) & "," & WMSDB_SQLText(CStr(a(15))) & "," & _
            WMSDB_SQLText(sHash) & ",CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)"
    End If

    n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
    If sErr<>"" Or n<0 Then
        On Error Resume Next:oCon.rollback():oCon.setAutoCommit(True):On Error GoTo EH
        If sErr="" Then sErr="Firebird не подтвердил запись."
        Exit Function
    End If

    oCon.commit()
    On Error Resume Next:oCon.setAutoCommit(True):On Error GoTo EH

    If Not WMSDBIu_RecordExists(oCon,sid,sErr) Then
        If sErr="" Then sErr="После COMMIT запись не найдена по SOURCE_ID=" & sid
        Exit Function
    End If

    Issues_API_MarkPosted r,sHash
    If Not WMSDB_SaveDatabaseDocument(sErr) Then Exit Function

    WMSDBIu_SaveIssue=True
    Exit Function
EH:
    On Error Resume Next
    If Not (IsNull(oCon) Or IsEmpty(oCon)) Then oCon.rollback():oCon.setAutoCommit(True)
    sErr="Запись выдачи: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_IssueSetSQL(oSh As Object,r As Long,a As Variant,sHash As String) As String
    WMSDBIu_IssueSetSQL = _
        "ISSUE_NO=" & WMSDBIu_IntSQL(oSh.getCellByPosition(0,r).Value) & _
        ",PRODUCT_CODE=" & WMSDB_SQLText(CStr(a(1))) & _
        ",PRODUCT_NAME=" & WMSDB_SQLText(CStr(a(2))) & _
        ",ISSUE_QTY=" & WMSDBIu_NumberSQL(CDbl(a(3))) & _
        ",UNIT_NAME=" & WMSDB_SQLText(CStr(a(4))) & _
        ",EMPLOYEE_NAME=" & WMSDB_SQLText(CStr(a(5))) & _
        ",ISSUE_DATE=" & WMSDBIu_DateSQL(CDbl(a(6))) & _
        ",SOURCE_LOCATION=" & WMSDB_SQLText(CStr(a(7))) & _
        ",RETURNABLE=" & WMSDB_SQLText(CStr(a(8))) & _
        ",RETURNED_QTY=" & WMSDBIu_NumberSQL(CDbl(a(9))) & _
        ",RETURN_DATE=" & WMSDBIu_DateSQL(CDbl(a(10))) & _
        ",NOTE_TEXT=" & WMSDB_SQLText(CStr(a(11))) & _
        ",ROW_MODE=" & WMSDB_SQLText(CStr(a(12))) & _
        ",ROW_STATE=" & WMSDB_SQLText(CStr(a(13))) & _
        ",RETURN_STATE=" & WMSDB_SQLText(CStr(a(15))) & _
        ",BUSINESS_HASH=" & WMSDB_SQLText(sHash)
End Function

Function WMSDBIu_RecordExists(oCon As Object,sid As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBIu_RecordExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT SOURCE_ID FROM " & WMSDBIU_TABLE & " WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBIu_RecordExists=(Trim(oRS.getString(1))=Trim(sid))
    Exit Function
EH:
    sErr=Error$
End Function

Function WMSDBIu_VerifyAfterReconnect(oDoc As Object,sid As String,ByRef sErr As String) As Boolean
    Dim oCon As Object
    WMSDBIu_VerifyAfterReconnect=False:sErr=""
    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBIu_RecordExists(oCon,sid,sErr) Then
        If sErr="" Then sErr="SOURCE_ID не найден после повторного подключения: " & sid
        WMSDB_Close
        Exit Function
    End If
    WMSDB_Close
    WMSDBIu_VerifyAfterReconnect=True
End Function

Sub WMSDBIu_RenumberVisibleIssues(oSh As Object)
    Dim last As Long,r As Long,n As Long
    last=Issues_LastContentRow(oSh,11)
    n=0
    For r=1 To last
        If Issues_RowHasBusinessData(oSh,r) Then
            n=n+1
            oSh.getCellByPosition(0,r).Value=n
            Issues_ApplyRowVisual oSh,r
        End If
    Next r
End Sub


' ============================================================================
' SAFE WORK-AREA DELETE 0.1.1
'
' IMPORTANT:
' Never delete a whole Calc row here. Whole-row deletion also moves/warps
' controls and other objects placed to the right of the work table.
'
' We remove ONLY the cells belonging to the WMS Issues table (A through the
' last WMS technical header) and shift those cells upward. Shapes/buttons and
' unrelated cells to the right stay exactly where they are.
' ============================================================================
Function WMSDBIu_RemoveIssueCellsSafely(oDoc As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim lastCol As Long,lastRow As Long,addr As Variant
    WMSDBIu_RemoveIssueCellsSafely=False:sErr=""
    On Error GoTo EH

    lastCol=Issues_LastHeaderCol(oSh)
    If lastCol<WMSISS_USER_LAST_COL Then lastCol=WMSISS_USER_LAST_COL

    lastRow=Issues_LastContentRow(oSh,lastCol)
    If r<1 Or r>lastRow Then
        sErr="Некорректная рабочая строка: " & CStr(r+1)
        Exit Function
    End If

    gWMSISS_Busy=True

    ' XCellRangeMovement.removeRange with UP affects only this rectangular
    ' table area; it does NOT remove the spreadsheet row itself.
    addr=oSh.getCellRangeByPosition(0,r,lastCol,r).RangeAddress
    oSh.removeRange(addr,com.sun.star.sheet.CellDeleteMode.UP)

    WMSDBIu_RenumberVisibleIssues oSh
    WMSDBIu_RepaintVisibleIssues oSh

    gWMSISS_Busy=False
    Issues_ResetQueueCache
    Issues_RefreshFilter oDoc,oSh

    WMSDBIu_RemoveIssueCellsSafely=True
    Exit Function
EH:
    gWMSISS_Busy=False
    sErr=CStr(Err) & " " & Error$
End Function

Sub WMSDBIu_RepaintVisibleIssues(oSh As Object)
    Dim last As Long,r As Long
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    For r=1 To last
        If Issues_RowHasBusinessData(oSh,r) Then Issues_ApplyRowVisual oSh,r
    Next r
End Sub

' One-time repair for a sheet that was already visually shifted by v0.1.0.
' It does not touch Firebird records and does not delete any business data.
Sub WMSDBIu_RepairIssueLayout()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBIU_SHEET) Then
        MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Выдачи: ремонт"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSDBIU_SHEET)

    gWMSISS_Busy=True
    WMSDBIu_RenumberVisibleIssues oSh
    WMSDBIu_RepaintVisibleIssues oSh
    If Not Issues_InstallButtonsCore(oDoc,oSh,sErr) Then
        gWMSISS_Busy=False
        MsgBox "Данные не затронуты, но панель кнопок не восстановлена:" & Chr(10) & sErr,48,"WMS — Выдачи: ремонт"
        Exit Sub
    End If
    gWMSISS_Busy=False

    Issues_ResetQueueCache
    Issues_RefreshFilter oDoc,oSh

    MsgBox "Лист 'Выдачи' восстановлен после старого способа удаления строк." & Chr(10) & _
           "Данные Firebird не изменялись.",64,"WMS — Выдачи: ремонт"
    Exit Sub
EH:
    gWMSISS_Busy=False
    MsgBox "Ремонт оформления: " & CStr(Err) & " " & Error$,16,"WMS — Выдачи: ремонт"
End Sub


Function WMSDBIu_ColumnExists(oCon As Object,sTable As String,sColumn As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBIu_ColumnExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE " & _
        "TRIM(RDB$RELATION_NAME)='" & UCase(Replace(sTable,"'","''")) & "' AND " & _
        "TRIM(RDB$FIELD_NAME)='" & UCase(Replace(sColumn,"'","''")) & "'"
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBIu_ColumnExists=(oRS.getInt(1)>0)
    Exit Function
EH:
    sErr="Проверка поля " & sColumn & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_SaveAllocation(oDoc As Object,sid As String,sCat As String,sSub As String,sOrigin As String,ByRef sErr As String) As Boolean
    Dim oCon As Object,sql As String,n As Long
    WMSDBIu_SaveAllocation=False:sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBIu_EnsureTables(oCon,sErr) Then Exit Function
    sql="UPDATE " & WMSDBIU_TABLE & " SET " & _
        "ALLOC_CATEGORY=" & WMSDB_SQLText(sCat) & "," & _
        "ALLOC_SUBCATEGORY=" & WMSDB_SQLText(sSub) & "," & _
        "STOCK_ORIGIN=" & WMSDB_SQLText(sOrigin) & "," & _
        "UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
    If sErr<>"" Then Exit Function
    oCon.commit()
    If Not WMSDB_SaveDatabaseDocument(sErr) Then Exit Function
    WMSDBIu_SaveAllocation=True
    Exit Function
EH:
    sErr="Сохранение назначения остатка: " & CStr(Err) & " " & Error$
End Function

Function WMSDBIu_VerifyIssueAndMovementAfterReconnect(oDoc As Object,sid As String,movementID As String,ByRef sErr As String) As Boolean
    Dim oCon As Object
    WMSDBIu_VerifyIssueAndMovementAfterReconnect=False:sErr=""
    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBIu_RecordExists(oCon,sid,sErr) Then
        If sErr="" Then sErr="Выдача не найдена после reconnect."
        WMSDB_Close
        Exit Function
    End If
    If Not WMSDBST_VerifyMovementCon(oCon,movementID,sErr) Then
        If sErr="" Then sErr="Движение списания не найдено после reconnect."
        WMSDB_Close
        Exit Function
    End If
    WMSDB_Close
    WMSDBIu_VerifyIssueAndMovementAfterReconnect=True
End Function

Function WMSDBIu_NumberSQL(v As Double) As String
    WMSDBIu_NumberSQL=Replace(CStr(v),",",".")
End Function

Function WMSDBIu_IntSQL(v As Double) As String
    WMSDBIu_IntSQL=CStr(CLng(v))
End Function

Function WMSDBIu_DateSQL(v As Double) As String
    Dim d As Date
    If v<=0 Then WMSDBIu_DateSQL="NULL":Exit Function
    d=CDate(v)
    WMSDBIu_DateSQL="DATE '" & Right("0000" & CStr(Year(d)),4) & "-" & _
                     Right("00" & CStr(Month(d)),2) & "-" & _
                     Right("00" & CStr(Day(d)),2) & "'"
End Function

