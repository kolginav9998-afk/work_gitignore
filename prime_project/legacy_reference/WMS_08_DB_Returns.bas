Option Explicit

' ============================================================================
' WMS_08_DB_Returns.bas
' v0.1.0 — Returns for already conducted returnable issues.
'
' Source of truth: Firebird.
' Calc does NOT keep historical issue rows.
'
' Flow:
'   button "Возврат"
'   -> query open returnable issues from WMS_ISSUES
'   -> choose one by number
'   -> enter returned quantity
'   -> insert immutable event into WMS_ISSUE_RETURNS
'   -> update aggregate RETURNED_QTY / RETURN_STATE in WMS_ISSUES
'   -> COMMIT + save Base document
'   -> close/reopen Firebird
'   -> verify event and aggregate totals
' ============================================================================

Global Const WMSDBR_VERSION = "0.3.1-PRODUCTION-RETURN-RECOVERY"
Global Const WMSDBR_ISSUES = "WMS_ISSUES"
Global Const WMSDBR_RETURNS = "WMS_ISSUE_RETURNS"

Sub WMSDBR_Install()
    Dim oCon As Object,sErr As String
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Возвраты":Exit Sub

    If Not WMSDB_TableExistsSimple(oCon,WMSDBR_ISSUES,sErr) Then
        WMSDB_Close
        MsgBox "Таблица WMS_ISSUES не найдена. Сначала установите WMS_07_DB_Issues.",16,"WMS — Возвраты"
        Exit Sub
    End If
    If Not WMSDB_TableExistsSimple(oCon,WMSDBR_RETURNS,sErr) Then
        WMSDB_Close
        MsgBox "Таблица WMS_ISSUE_RETURNS не найдена. Запустите WMSDBIu_Install.",16,"WMS — Возвраты"
        Exit Sub
    End If
    If Not WMSDBR_EnsureLotColumns(oCon,sErr) Then
        WMSDB_Close
        MsgBox "Не удалось обновить схему возвратов:" & Chr(10) & sErr,16,"WMS — Возвраты"
        Exit Sub
    End If
    WMSDB_Close

    MsgBox "Модуль возвратов готов." & Chr(10) & _
           "Версия: " & WMSDBR_VERSION & Chr(10) & _
           "История: Firebird / WMS_ISSUE_RETURNS",64,"WMS — Возвраты"
End Sub

Sub WMSDBR_SelfCheck()
    Dim oCon As Object,sErr As String,n As Long
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Возвраты":Exit Sub
    n=WMSDBR_CountOpen(oCon,sErr)
    WMSDB_Close
    If sErr<>"" Then
        MsgBox sErr,16,"WMS — Возвраты"
    Else
        MsgBox "Соединение: OK" & Chr(10) & _
               "Открытых возвратных выдач: " & CStr(n) & Chr(10) & _
               "Модуль: " & WMSDBR_VERSION,64,"WMS — Возвраты"
    End If
End Sub

Function WMSDBR_EnsureLotColumns(oCon As Object,ByRef sErr As String) As Boolean
    WMSDBR_EnsureLotColumns=False:sErr=""
    On Error GoTo EH
    If Not WMSDBIu_ColumnExists(oCon,WMSDBR_RETURNS,"LOT_ID",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBR_RETURNS & " ADD LOT_ID VARCHAR(128)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBR_RETURNS,"ENTRY_UNIT",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBR_RETURNS & " ADD ENTRY_UNIT VARCHAR(64)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBR_RETURNS,"BASE_QTY",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBR_RETURNS & " ADD BASE_QTY DECIMAL(18,4)",sErr
        If sErr<>"" Then Exit Function
    End If
    If Not WMSDBIu_ColumnExists(oCon,WMSDBR_RETURNS,"BASE_UNIT",sErr) Then
        WMSDB_ExecuteUpdate oCon,"ALTER TABLE " & WMSDBR_RETURNS & " ADD BASE_UNIT VARCHAR(64)",sErr
        If sErr<>"" Then Exit Function
    End If
    oCon.commit()
    WMSDBR_EnsureLotColumns=True
    Exit Function
EH:
    sErr="Схема возвратов: " & CStr(Err) & " " & Error$
End Function

Sub WMSDBR_ReturnDialog()
    Dim oCon As Object,oStmt As Object,oRS As Object,sErr As String,sql As String
    Dim sList As String,choice As String,idx As Long,n As Long
    Dim ids() As String,names() As String,emps() As String,units() As String
    Dim issued() As Double,returned() As Double,openQty() As Double
    Dim sid As String,qtyText As String,qty As Double,note As String
    Dim dText As String,dSQL As String,ans As Integer

    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Возврат":Exit Sub

    sql="SELECT SOURCE_ID,PRODUCT_NAME,EMPLOYEE_NAME,UNIT_NAME,ISSUE_QTY,RETURNED_QTY " & _
        "FROM " & WMSDBR_ISSUES & " " & _
        "WHERE UPPER(TRIM(COALESCE(RETURNABLE,''))) NOT IN ('','НЕТ','NO','FALSE','0') " & _
        "AND COALESCE(RETURNED_QTY,0) < ISSUE_QTY " & _
        "ORDER BY ISSUE_DATE, CREATED_AT"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)

    n=0:sList=""
    Do While oRS.next()
        n=n+1
        ReDim Preserve ids(1 To n)
        ReDim Preserve names(1 To n)
        ReDim Preserve emps(1 To n)
        ReDim Preserve units(1 To n)
        ReDim Preserve issued(1 To n)
        ReDim Preserve returned(1 To n)
        ReDim Preserve openQty(1 To n)

        ids(n)=oRS.getString(1)
        names(n)=oRS.getString(2)
        emps(n)=oRS.getString(3)
        units(n)=oRS.getString(4)
        issued(n)=oRS.getDouble(5)
        returned(n)=oRS.getDouble(6)
        openQty(n)=issued(n)-returned(n)

        If n<=30 Then
            sList=sList & CStr(n) & ". " & names(n) & " — " & emps(n) & _
                  " — осталось " & WMSDBR_Num(openQty(n)) & " " & units(n) & Chr(10)
        End If
    Loop
    WMSDB_Close

    If n=0 Then
        MsgBox "Открытых возвратных выдач нет.",64,"WMS — Возврат"
        Exit Sub
    End If

    If n>30 Then sList=sList & Chr(10) & "Показаны первые 30 из " & CStr(n) & "."

    choice=InputBox("Введите номер выдачи из списка:" & Chr(10) & Chr(10) & sList, _
                    "WMS — Возврат")
    If Trim(choice)="" Then Exit Sub
    If Not IsNumeric(choice) Then
        MsgBox "Нужно ввести номер из списка.",48,"WMS — Возврат"
        Exit Sub
    End If
    idx=CLng(choice)
    If idx<1 Or idx>n Or idx>30 Then
        MsgBox "Нет такой позиции в показанном списке.",48,"WMS — Возврат"
        Exit Sub
    End If

    sid=ids(idx)
    qtyText=InputBox("Товар: " & names(idx) & Chr(10) & _
                     "Получил: " & emps(idx) & Chr(10) & _
                     "Выдано: " & WMSDBR_Num(issued(idx)) & " " & units(idx) & Chr(10) & _
                     "Уже возвращено: " & WMSDBR_Num(returned(idx)) & " " & units(idx) & Chr(10) & _
                     "Осталось вернуть: " & WMSDBR_Num(openQty(idx)) & " " & units(idx) & Chr(10) & Chr(10) & _
                     "Сколько возвращают сейчас?", _
                     "WMS — Количество возврата",WMSDBR_Num(openQty(idx)))
    If Trim(qtyText)="" Then Exit Sub

    qty=WMSDBR_ParseNumber(qtyText,sErr)
    If sErr<>"" Then MsgBox sErr,48,"WMS — Возврат":Exit Sub
    If qty<=0 Then MsgBox "Количество должно быть больше нуля.",48,"WMS — Возврат":Exit Sub
    If qty>openQty(idx)+0.0000001 Then
        MsgBox "Нельзя вернуть больше остатка." & Chr(10) & _
               "Осталось: " & WMSDBR_Num(openQty(idx)) & " " & units(idx),48,"WMS — Возврат"
        Exit Sub
    End If

    dText=InputBox("Дата возврата (ДД.ММ.ГГГГ)." & Chr(10) & _
                   "Оставьте пустым — сегодняшняя дата.", _
                   "WMS — Дата возврата")
    dSQL=WMSDBR_DateInputSQL(dText,sErr)
    If sErr<>"" Then MsgBox sErr,48,"WMS — Возврат":Exit Sub

    note=InputBox("Примечание к возврату (необязательно):","WMS — Возврат")

    ans=MsgBox("Подтвердить возврат?" & Chr(10) & Chr(10) & _
               names(idx) & Chr(10) & _
               emps(idx) & Chr(10) & _
               "Возврат сейчас: " & WMSDBR_Num(qty) & " " & units(idx) & Chr(10) & _
               "После возврата останется: " & WMSDBR_Num(openQty(idx)-qty) & " " & units(idx), _
               36,"WMS — Возврат")
    If ans<>6 Then Exit Sub

    If Not WMSDBR_PostReturn(ThisComponent,sid,qty,dSQL,note,sErr) Then
        MsgBox "Возврат НЕ проведён." & Chr(10) & Chr(10) & sErr,16,"WMS — Возврат"
        Exit Sub
    End If

    MsgBox "Возврат проведён." & Chr(10) & _
           "Товар: " & names(idx) & Chr(10) & _
           "Возвращено сейчас: " & WMSDBR_Num(qty) & " " & units(idx) & Chr(10) & _
           "Осталось вернуть: " & WMSDBR_Num(openQty(idx)-qty) & " " & units(idx), _
           64,"WMS — Возврат"
    Exit Sub

EH:
    On Error Resume Next
    WMSDB_Close
    MsgBox "Ошибка возврата: " & CStr(Err) & " " & Error$,16,"WMS — Возврат"
End Sub

Function WMSDBR_PostReturn(oDoc As Object,sid As String,qty As Double,dSQL As String,note As String,ByRef sErr As String) As Boolean
    Dim oCon As Object,oStmt As Object,oRS As Object
    Dim issueQty As Double,oldReturned As Double,newReturned As Double
    Dim seq As Long,retID As String,state As String,sql As String,n As Long
    Dim code As String,nm As String,unitText As String,loc As String,cat As String,subcat As String,origin As String
    Dim lotID As String,baseUnit As String,baseQty As Double,dRet As Double

    WMSDBR_PostReturn=False:sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBR_EnsureLotColumns(oCon,sErr) Then
        WMSDB_Close
        Exit Function
    End If

    ' Never create a second return while a previously committed return event
    ' for the same issue still has no stock movement. Repair it first.
    If WMSDBR_HasPendingReturnCon(oCon,sid,sErr) Then
        WMSDB_Close
        If sErr<>"" Then Exit Function
        If WMSDBR_RepairPendingForSource(oDoc,sid,sErr) Then
            sErr="Обнаружен и восстановлен незавершённый предыдущий возврат. " & _
                 "Откройте «Возврат» ещё раз — остаток уже пересчитан."
        End If
        Exit Function
    End If
    If sErr<>"" Then
        WMSDB_Close
        Exit Function
    End If

    sql="SELECT ISSUE_QTY,RETURNED_QTY,PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,SOURCE_LOCATION," & _
        "COALESCE(ALLOC_CATEGORY,''),COALESCE(ALLOC_SUBCATEGORY,''),COALESCE(STOCK_ORIGIN,'')," & _
        "COALESCE(LOT_ID,''),COALESCE(BASE_UNIT,'') " & _
        "FROM " & WMSDBR_ISSUES & " WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If Not oRS.next() Then sErr="Выдача не найдена: " & sid:WMSDB_Close:Exit Function

    issueQty=oRS.getDouble(1):oldReturned=oRS.getDouble(2)
    code=oRS.getString(3):nm=oRS.getString(4):unitText=oRS.getString(5):loc=oRS.getString(6)
    cat=oRS.getString(7):subcat=oRS.getString(8):origin=oRS.getString(9)
    lotID=Trim(oRS.getString(10)):baseUnit=Trim(oRS.getString(11))

    newReturned=oldReturned+qty
    If qty<=0 Or newReturned>issueQty+0.0000001 Then sErr="Некорректное количество возврата.":WMSDB_Close:Exit Function

    If lotID<>"" Then
        If Not WMSDBUL_ConvertToBaseCon(oCon,lotID,qty,unitText,baseQty,baseUnit,sErr) Then WMSDB_Close:Exit Function
    Else
        baseQty=qty:baseUnit=unitText
    End If

    seq=WMSDBR_NextSeq(oCon,sid,sErr)
    If sErr<>"" Then WMSDB_Close:Exit Function
    retID=WMSDBR_NewReturnID(sid,seq)
    If newReturned>=issueQty-0.0000001 Then state="CLOSED" Else state="PARTIAL"

    On Error Resume Next:oCon.setAutoCommit(False):On Error GoTo EH

    sql="INSERT INTO " & WMSDBR_RETURNS & _
        " (RETURN_ID,SOURCE_ID,RETURN_SEQ,RETURN_QTY,RETURN_DATE,NOTE_TEXT,LOT_ID,ENTRY_UNIT,BASE_QTY,BASE_UNIT,CREATED_AT) VALUES (" & _
        WMSDB_SQLText(retID) & "," & WMSDB_SQLText(sid) & "," & CStr(seq) & "," & _
        WMSDBR_NumSQL(qty) & "," & dSQL & "," & WMSDB_SQLText(note) & "," & _
        WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(unitText) & "," & WMSDBR_NumSQL(baseQty) & "," & _
        WMSDB_SQLText(baseUnit) & ",CURRENT_TIMESTAMP)"
    n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
    If sErr<>"" Then GoTo RollbackFail

    sql="UPDATE " & WMSDBR_ISSUES & " SET RETURNED_QTY=" & WMSDBR_NumSQL(newReturned) & _
        ",RETURN_DATE=" & dSQL & ",RETURN_STATE=" & WMSDB_SQLText(state) & _
        ",UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sid)
    n=WMSDB_ExecuteUpdate(oCon,sql,sErr)
    If sErr<>"" Then GoTo RollbackFail

    oCon.commit()
    On Error Resume Next:oCon.setAutoCommit(True):On Error GoTo EH
    If Not WMSDB_SaveDatabaseDocument(sErr) Then WMSDB_Close:Exit Function
    WMSDB_Close

    If Not WMSDBR_VerifyAfterReconnect(oDoc,sid,retID,newReturned,sErr) Then Exit Function

    dRet=WMSDBR_SQLDateToCalc(dSQL)
    If lotID<>"" Then
        If Not WMSDBST_PostMovementClassified(oDoc,"RET-" & retID,retID,"RETURN","IN", _
            code,nm,baseQty,baseUnit,loc,origin,cat,subcat,dRet, _
            "Возврат по выдаче " & sid & IIf(note="",""," | " & note),sErr) Then Exit Function

        oCon=WMSDB_GetConnectionEx(oDoc,sErr)
        If sErr<>"" Then Exit Function
        sql="UPDATE WMS_STOCK_MOVEMENTS SET LOT_ID=" & WMSDB_SQLText(lotID) & _
            ",ENTRY_QTY=" & WMSDBR_NumSQL(qty) & _
            ",ENTRY_UNIT=" & WMSDB_SQLText(unitText) & _
            ",BASE_QTY=" & WMSDBR_NumSQL(baseQty) & _
            ",BASE_UNIT=" & WMSDB_SQLText(baseUnit) & _
            " WHERE MOVEMENT_ID=" & WMSDB_SQLText("RET-" & retID)
        oStmt=oCon.createStatement():oStmt.executeUpdate(sql):oCon.commit():WMSDB_Close
    Else
        If Not WMSDBST_PostReturnMovement(oDoc,"RET-" & retID,retID,code,nm,qty,unitText,loc, _
            cat,subcat,origin,dRet,"Возврат по выдаче " & sid & IIf(note="",""," | " & note),sErr) Then Exit Function
    End If

    If Not WMSDBR_VerifyLotReturnMovement(oDoc,"RET-" & retID,lotID,baseQty,sErr) Then Exit Function

    WMSDBR_PostReturn=True
    Exit Function

RollbackFail:
    On Error Resume Next:oCon.rollback():oCon.setAutoCommit(True):WMSDB_Close
    Exit Function
EH:
    On Error Resume Next:oCon.rollback():oCon.setAutoCommit(True):WMSDB_Close
    sErr=CStr(Err) & " " & Error$
End Function


Function WMSDBR_HasPendingReturnCon(oCon As Object,sid As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBR_HasPendingReturnCon=False
    sErr=""
    On Error GoTo EH

    oStmt=oCon.createStatement()
    sql="SELECT FIRST 1 R.RETURN_ID " & _
        "FROM " & WMSDBR_RETURNS & " R " & _
        "LEFT JOIN WMS_STOCK_MOVEMENTS M ON M.MOVEMENT_ID='RET-' || R.RETURN_ID " & _
        "WHERE R.SOURCE_ID=" & WMSDB_SQLText(sid) & " AND M.MOVEMENT_ID IS NULL " & _
        "ORDER BY R.CREATED_AT"
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBR_HasPendingReturnCon=True
    Exit Function
EH:
    sErr="Проверка незавершённых возвратов: " & CStr(Err) & " " & Error$
End Function

Function WMSDBR_CountPendingCon(oCon As Object,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBR_CountPendingCon=0
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    sql="SELECT COUNT(*) FROM " & WMSDBR_RETURNS & " R " & _
        "LEFT JOIN WMS_STOCK_MOVEMENTS M ON M.MOVEMENT_ID='RET-' || R.RETURN_ID " & _
        "WHERE M.MOVEMENT_ID IS NULL"
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBR_CountPendingCon=oRS.getInt(1)
    Exit Function
EH:
    sErr="Подсчёт незавершённых возвратов: " & CStr(Err) & " " & Error$
End Function

Function WMSDBR_RepairPendingForSource(oDoc As Object,sid As String,ByRef sErr As String) As Boolean
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String
    Dim retID As String,code As String,nm As String,loc As String
    Dim cat As String,subcat As String,origin As String,lotID As String
    Dim entryUnit As String,baseUnit As String,noteText As String
    Dim entryQty As Double,baseQty As Double,dRet As Double,movementID As String

    WMSDBR_RepairPendingForSource=False
    sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBR_EnsureLotColumns(oCon,sErr) Then
        WMSDB_Close
        Exit Function
    End If

    sql="SELECT FIRST 1 R.RETURN_ID,R.RETURN_QTY,R.RETURN_DATE,COALESCE(R.NOTE_TEXT,'')," & _
        "COALESCE(R.LOT_ID,''),COALESCE(R.ENTRY_UNIT,''),COALESCE(R.BASE_QTY,R.RETURN_QTY)," & _
        "COALESCE(R.BASE_UNIT,''),I.PRODUCT_CODE,I.PRODUCT_NAME,I.SOURCE_LOCATION," & _
        "COALESCE(I.ALLOC_CATEGORY,''),COALESCE(I.ALLOC_SUBCATEGORY,''),COALESCE(I.STOCK_ORIGIN,'') " & _
        "FROM " & WMSDBR_RETURNS & " R " & _
        "JOIN " & WMSDBR_ISSUES & " I ON I.SOURCE_ID=R.SOURCE_ID " & _
        "LEFT JOIN WMS_STOCK_MOVEMENTS M ON M.MOVEMENT_ID='RET-' || R.RETURN_ID " & _
        "WHERE R.SOURCE_ID=" & WMSDB_SQLText(sid) & " AND M.MOVEMENT_ID IS NULL " & _
        "ORDER BY R.CREATED_AT"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If Not oRS.next() Then
        WMSDB_Close
        WMSDBR_RepairPendingForSource=True
        Exit Function
    End If

    retID=oRS.getString(1)
    entryQty=oRS.getDouble(2)
    dRet=oRS.getDouble(3)
    noteText=oRS.getString(4)
    lotID=Trim(oRS.getString(5))
    entryUnit=Trim(oRS.getString(6))
    baseQty=oRS.getDouble(7)
    baseUnit=Trim(oRS.getString(8))
    code=oRS.getString(9)
    nm=oRS.getString(10)
    loc=oRS.getString(11)
    cat=oRS.getString(12)
    subcat=oRS.getString(13)
    origin=oRS.getString(14)
    WMSDB_Close

    If entryUnit="" Then entryUnit=baseUnit
    If baseUnit="" Then baseUnit=entryUnit
    If baseQty=0 Then baseQty=entryQty
    movementID="RET-" & retID

    If Not WMSDBST_PostMovementClassified(oDoc,movementID,retID,"RETURN","IN", _
        code,nm,baseQty,baseUnit,loc,origin,cat,subcat,dRet, _
        "Восстановленный возврат по выдаче " & sid & IIf(noteText="",""," | " & noteText),sErr) Then
        Exit Function
    End If

    If lotID<>"" Then
        oCon=WMSDB_GetConnectionEx(oDoc,sErr)
        If sErr<>"" Then Exit Function
        sql="UPDATE WMS_STOCK_MOVEMENTS SET LOT_ID=" & WMSDB_SQLText(lotID) & _
            ",ENTRY_QTY=" & WMSDBR_NumSQL(entryQty) & _
            ",ENTRY_UNIT=" & WMSDB_SQLText(entryUnit) & _
            ",BASE_QTY=" & WMSDBR_NumSQL(baseQty) & _
            ",BASE_UNIT=" & WMSDB_SQLText(baseUnit) & _
            " WHERE MOVEMENT_ID=" & WMSDB_SQLText(movementID)
        oStmt=oCon.createStatement()
        oStmt.executeUpdate(sql)
        oCon.commit()
        WMSDB_Close
    End If

    If Not WMSDBR_VerifyLotReturnMovement(oDoc,movementID,lotID,baseQty,sErr) Then Exit Function
    WMSDBR_RepairPendingForSource=True
    Exit Function
EH:
    On Error Resume Next
    WMSDB_Close
    sErr="Восстановление возврата: " & CStr(Err) & " " & Error$
End Function

Sub WMSDBR_RepairPendingReturns()
    Dim oCon As Object,oStmt As Object,oRS As Object,sErr As String
    Dim ids() As String,n As Long,i As Long,repaired As Long

    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then
        MsgBox sErr,16,"WMS — Возвраты"
        Exit Sub
    End If

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT DISTINCT R.SOURCE_ID FROM " & WMSDBR_RETURNS & " R " & _
        "LEFT JOIN WMS_STOCK_MOVEMENTS M ON M.MOVEMENT_ID='RET-' || R.RETURN_ID " & _
        "WHERE M.MOVEMENT_ID IS NULL")
    Do While oRS.next()
        n=n+1
        ReDim Preserve ids(1 To n)
        ids(n)=oRS.getString(1)
    Loop
    WMSDB_Close

    If n=0 Then
        MsgBox "Незавершённых возвратов нет.",64,"WMS — Возвраты"
        Exit Sub
    End If

    For i=1 To n
        sErr=""
        If WMSDBR_RepairPendingForSource(ThisComponent,ids(i),sErr) Then
            repaired=repaired+1
        Else
            MsgBox "Не удалось восстановить возврат по выдаче " & ids(i) & ":" & Chr(10) & sErr,16,"WMS — Возвраты"
            Exit Sub
        End If
    Next i

    MsgBox "Восстановление завершено." & Chr(10) & _
           "Исправлено выдач: " & CStr(repaired),64,"WMS — Возвраты"
    Exit Sub
EH:
    On Error Resume Next
    WMSDB_Close
    MsgBox "Ошибка восстановления возвратов: " & CStr(Err) & " " & Error$,16,"WMS — Возвраты"
End Sub

Function WMSDBR_VerifyLotReturnMovement(oDoc As Object,movementID As String,lotID As String,expectedBaseQty As Double,ByRef sErr As String) As Boolean
    Dim oCon As Object,oStmt As Object,oRS As Object
    WMSDBR_VerifyLotReturnMovement=False:sErr=""
    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COALESCE(LOT_ID,''),QTY FROM WMS_STOCK_MOVEMENTS WHERE MOVEMENT_ID=" & WMSDB_SQLText(movementID))
    If Not oRS.next() Then sErr="Движение возврата не найдено.":WMSDB_Close:Exit Function
    If lotID<>"" And Trim(oRS.getString(1))<>lotID Then sErr="Возврат попал не в исходную партию.":WMSDB_Close:Exit Function
    If Abs(oRS.getDouble(2)-expectedBaseQty)>0.0001 Then sErr="Базовое количество возврата не совпало.":WMSDB_Close:Exit Function
    WMSDB_Close
    WMSDBR_VerifyLotReturnMovement=True
End Function

Function WMSDBR_VerifyAfterReconnect(oDoc As Object,sid As String,retID As String,expectedReturned As Double,ByRef sErr As String) As Boolean
    Dim oCon As Object,oStmt As Object,oRS As Object,actual As Double
    WMSDBR_VerifyAfterReconnect=False:sErr=""

    WMSDB_Close
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT RETURN_ID FROM " & WMSDBR_RETURNS & _
                           " WHERE RETURN_ID=" & WMSDB_SQLText(retID))
    If Not oRS.next() Then
        sErr="Событие возврата не найдено после повторного подключения."
        WMSDB_Close
        Exit Function
    End If

    oRS=oStmt.executeQuery("SELECT RETURNED_QTY FROM " & WMSDBR_ISSUES & _
                           " WHERE SOURCE_ID=" & WMSDB_SQLText(sid))
    If Not oRS.next() Then
        sErr="Исходная выдача не найдена после повторного подключения."
        WMSDB_Close
        Exit Function
    End If
    actual=oRS.getDouble(1)
    WMSDB_Close

    If Abs(actual-expectedReturned)>0.0001 Then
        sErr="Контроль суммы возврата не совпал. Ожидалось " & _
             WMSDBR_Num(expectedReturned) & ", в БД " & WMSDBR_Num(actual)
        Exit Function
    End If

    WMSDBR_VerifyAfterReconnect=True
End Function

Function WMSDBR_NextSeq(oCon As Object,sid As String,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object
    WMSDBR_NextSeq=1:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT MAX(RETURN_SEQ) FROM " & WMSDBR_RETURNS & _
                           " WHERE SOURCE_ID=" & WMSDB_SQLText(sid))
    If oRS.next() Then
        If Not oRS.wasNull() Then WMSDBR_NextSeq=oRS.getInt(1)+1
    End If
    Exit Function
EH:
    sErr=Error$
End Function

Function WMSDBR_CountOpen(oCon As Object,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object
    WMSDBR_CountOpen=0:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COUNT(*) FROM " & WMSDBR_ISSUES & _
        " WHERE UPPER(TRIM(COALESCE(RETURNABLE,''))) NOT IN ('','НЕТ','NO','FALSE','0')" & _
        " AND COALESCE(RETURNED_QTY,0)<ISSUE_QTY")
    If oRS.next() Then WMSDBR_CountOpen=oRS.getInt(1)
    Exit Function
EH:
    sErr=Error$
End Function

Function WMSDBR_NewReturnID(sid As String,seq As Long) As String
    WMSDBR_NewReturnID="RET-" & Format(Now,"YYYYMMDDHHMMSS") & "-" & _
                       Right("000" & CStr(seq),3) & "-" & Right(sid,12)
End Function

Function WMSDBR_ParseNumber(s As String,ByRef sErr As String) As Double
    Dim t As String
    WMSDBR_ParseNumber=0:sErr=""
    t=Trim(s)
    t=Replace(t," ","")
    t=Replace(t,",",".")
    If t="" Or Not IsNumeric(t) Then
        sErr="Некорректное количество: " & s
        Exit Function
    End If
    ' Val is locale-independent for dot decimal.
    WMSDBR_ParseNumber=Val(t)
End Function

Function WMSDBR_DateInputSQL(s As String,ByRef sErr As String) As String
    Dim t As String,p As Variant,dd As Long,mm As Long,yy As Long,d As Date
    sErr="":t=Trim(s)

    If t="" Then
        d=Date
    Else
        t=Replace(t,"/","." )
        t=Replace(t,"-","." )
        p=Split(t,".")
        If UBound(p)<>2 Then
            sErr="Дата должна быть в формате ДД.ММ.ГГГГ."
            WMSDBR_DateInputSQL="NULL"
            Exit Function
        End If
        If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Or Not IsNumeric(p(2)) Then
            sErr="Некорректная дата."
            WMSDBR_DateInputSQL="NULL"
            Exit Function
        End If
        dd=CLng(p(0)):mm=CLng(p(1)):yy=CLng(p(2))
        On Error GoTo BadDate
        d=DateSerial(yy,mm,dd)
        If Day(d)<>dd Or Month(d)<>mm Or Year(d)<>yy Then GoTo BadDate
        On Error GoTo 0
    End If

    WMSDBR_DateInputSQL="DATE '" & Right("0000" & CStr(Year(d)),4) & "-" & _
                         Right("00" & CStr(Month(d)),2) & "-" & _
                         Right("00" & CStr(Day(d)),2) & "'"
    Exit Function

BadDate:
    sErr="Некорректная календарная дата."
    WMSDBR_DateInputSQL="NULL"
End Function

Function WMSDBR_Num(v As Double) As String
    Dim s As String
    s=CStr(v)
    If InStr(s,",")>0 Then
        Do While Right(s,1)="0":s=Left(s,Len(s)-1):Loop
        If Right(s,1)="," Then s=Left(s,Len(s)-1)
    ElseIf InStr(s,".")>0 Then
        Do While Right(s,1)="0":s=Left(s,Len(s)-1):Loop
        If Right(s,1)="." Then s=Left(s,Len(s)-1)
    End If
    WMSDBR_Num=s
End Function


Function WMSDBR_SQLDateToCalc(dSQL As String) As Double
    Dim s As String,y As Integer,m As Integer,d As Integer
    s=Replace(Replace(Trim(dSQL),"DATE ",""),"'","")
    If Len(s)>=10 Then
        y=CInt(Left(s,4)):m=CInt(Mid(s,6,2)):d=CInt(Mid(s,9,2))
        WMSDBR_SQLDateToCalc=CDbl(DateSerial(y,m,d))
    Else
        WMSDBR_SQLDateToCalc=CDbl(Date)
    End If
End Function

Function WMSDBR_NumSQL(v As Double) As String
    WMSDBR_NumSQL=Replace(CStr(v),",",".")
End Function

