Option Explicit

Global Const WMSOPS_VERSION = "2.0.0-RETURNS-INVENTORY"
Global Const WMSOPS_RET_SHEET = "Возвраты"
Global Const WMSOPS_INV_SHEET = "Инвентаризация"

Sub WMSOPS_Install()
    Dim oDoc As Object
    oDoc=ThisComponent
    On Error GoTo EH
    WMSOPS_PrepareReturns oDoc
    WMSOPS_PrepareInventory oDoc
    WMSOPS_RefreshReturns
    Exit Sub
EH:
    MsgBox "Возвраты/инвентаризация: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSOPS_PrepareReturns(oDoc As Object)
    Dim oSh As Object,h As Variant,i As Long
    If oDoc.Sheets.hasByName(WMSOPS_RET_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSOPS_RET_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSOPS_RET_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSOPS_RET_SHEET)
    End If
    oSh.IsVisible=True
    oSh.getCellRangeByPosition(0,0,12,4).clearContents(1023)
    oSh.getCellByPosition(0,0).String="ВОЗВРАТЫ"
    oSh.getCellByPosition(0,1).String="Открытые возвратные выдачи. Введите количество в колонке «Вернуть сейчас» и нажмите «Провести возвраты»."
    oSh.getCellRangeByPosition(0,0,10,0).CellBackColor=10040064
    oSh.getCellRangeByPosition(0,0,10,0).CharColor=16777215:oSh.getCellRangeByPosition(0,0,10,0).CharWeight=150
    h=Array("Дата выдачи","Наименование","Артикул","Выдано","Возвращено","Осталось","Ед.","Кому","Партия","Вернуть сейчас","Комментарий","SOURCE_ID")
    For i=0 To UBound(h):oSh.getCellByPosition(i,4).String=CStr(h(i)):Next i
    oSh.getCellRangeByPosition(0,4,11,4).CellBackColor=4473924:oSh.getCellRangeByPosition(0,4,11,4).CharColor=16777215:oSh.getCellRangeByPosition(0,4,11,4).CharWeight=150
    oSh.Columns.getByIndex(1).Width=6500:oSh.Columns.getByIndex(7).Width=4500:oSh.Columns.getByIndex(10).Width=6000
    oSh.Columns.getByIndex(11).IsVisible=False
    WMSOPS_InstallReturnButtons oDoc,oSh
End Sub

Sub WMSOPS_RefreshReturns()
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim sErr As String,r As Long,c As Object,lastR As Long,sql As String
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSOPS_RET_SHEET)
    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then MsgBox sErr,16,"WMS — Возвраты":Exit Sub
    On Error GoTo EH
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow:If lastR<5 Then lastR=5
    oSh.getCellRangeByPosition(0,5,11,lastR).clearContents(1023)
    sql="SELECT i.ISSUE_DATE,i.PRODUCT_NAME,COALESCE(p.SUPPLIER_ARTICLE,''),i.ISSUE_QTY,i.RETURNED_QTY," & _
        "(i.ISSUE_QTY-i.RETURNED_QTY),i.UNIT_NAME,i.EMPLOYEE_NAME,COALESCE(i.LOT_ID,''),i.SOURCE_ID " & _
        "FROM WMS_ISSUES i LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=i.PRODUCT_CODE " & _
        "WHERE UPPER(COALESCE(i.RETURNABLE,'')) IN ('ДА','YES','1','TRUE') AND i.ISSUE_QTY>i.RETURNED_QTY ORDER BY i.ISSUE_DATE,i.EMPLOYEE_NAME"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    r=5
    Do While oRS.next()
        oSh.getCellByPosition(0,r).String=oRS.getString(1)
        oSh.getCellByPosition(1,r).String=oRS.getString(2)
        oSh.getCellByPosition(2,r).String=oRS.getString(3)
        oSh.getCellByPosition(3,r).Value=oRS.getDouble(4)
        oSh.getCellByPosition(4,r).Value=oRS.getDouble(5)
        oSh.getCellByPosition(5,r).Value=oRS.getDouble(6)
        oSh.getCellByPosition(6,r).String=oRS.getString(7)
        oSh.getCellByPosition(7,r).String=oRS.getString(8)
        oSh.getCellByPosition(8,r).String=oRS.getString(9)
        oSh.getCellByPosition(11,r).String=oRS.getString(10)
        r=r+1
    Loop
    WMSDB_Close
    oSh.getCellByPosition(0,2).String="Открытых позиций: " & CStr(r-5)
    Exit Sub
EH:
    WMSDB_Close
    MsgBox "Обновление возвратов: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSOPS_ConductReturns()
    Dim oDoc As Object,oSh As Object,r As Long,lastR As Long,c As Object,nOK As Long,nErr As Long,sErr As String
    If Not WMSDBX_TryEnter("RETURNS") Then MsgBox "Дождитесь завершения другой операции WMS.",48,"WMS":Exit Sub
    On Error GoTo BatchFailed
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSOPS_RET_SHEET)
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow
    For r=5 To lastR
        If oSh.getCellByPosition(9,r).Value>0 Then
            sErr=""
            If WMSOPS_ReturnRow(oDoc,oSh,r,sErr) Then
                nOK=nOK+1
                oSh.getCellByPosition(9,r).Value=0
            Else
                nErr=nErr+1
                MsgBox "Строка " & CStr(r+1) & ": " & sErr,48,"WMS — Возвраты"
                Exit For
            End If
        End If
    Next r
    WMSDBX_Leave
    If gWMSDBX_UnsafeConnection Then Exit Sub
    If nErr=0 Then WMSOPS_RefreshReturns
    On Error Resume Next:WMSDBST_RefreshStock:On Error GoTo 0
    MsgBox "Возвратов проведено: " & CStr(nOK) & Chr(10) & "Ошибок: " & CStr(nErr),IIf(nErr=0,64,48),"WMS — Возвраты"
    Exit Sub
BatchFailed:
    sErr="Обработка остановлена: " & CStr(Err) & " " & Error$ & ". Успешные строки могли быть проведены; проверьте результат перед повтором."
    WMSDBX_Leave
    MsgBox sErr,16,"WMS"
End Sub

Function WMSOPS_ReturnRow(oDoc As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim commitConfirmed As Boolean,changedRows As Long
    Dim sourceID As String,retQty As Double,remaining As Double,unitName As String,lotID As String,note As String
    Dim oCon As Object,oStmt As Object,oRS As Object,code As String,nm As String,loc As String,origin As String,cat As String,subcat As String,baseUnit As String,baseQty As Double
    Dim retID As String,movID As String,retSeq As Long
    WMSOPS_ReturnRow=False:sErr=""
    sourceID=Trim(oSh.getCellByPosition(11,r).String):retQty=oSh.getCellByPosition(9,r).Value:remaining=oSh.getCellByPosition(5,r).Value
    unitName=Trim(oSh.getCellByPosition(6,r).String):lotID=Trim(oSh.getCellByPosition(8,r).String):note=Trim(oSh.getCellByPosition(10,r).String)
    If sourceID="" Then sErr="Нет SOURCE_ID.":Exit Function
    If retQty<=0 Then sErr="Количество возврата должно быть >0.":Exit Function
    If retQty>remaining+0.000001 Then sErr="Возврат больше невозвращённого количества.":Exit Function

    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then Exit Function
    On Error GoTo EH
    If Not WMSDBX_Begin(oCon,sErr) Then Exit Function
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT PRODUCT_CODE,PRODUCT_NAME,SOURCE_LOCATION,COALESCE(STOCK_ORIGIN,''),COALESCE(ALLOC_CATEGORY,''),COALESCE(ALLOC_SUBCATEGORY,''),COALESCE(BASE_UNIT,UNIT_NAME),COALESCE(BASE_QTY,ISSUE_QTY),ISSUE_QTY,COALESCE(RETURNED_QTY,0),UNIT_NAME,COALESCE(LOT_ID,''),COALESCE(RETURNABLE,'') FROM WMS_ISSUES WHERE SOURCE_ID=" & WMSDB_SQLText(sourceID))
    If Not oRS.next() Then sErr="Выдача не найдена.":GoTo Fail
    code=oRS.getString(1):nm=oRS.getString(2):loc=oRS.getString(3):origin=oRS.getString(4):cat=oRS.getString(5):subcat=oRS.getString(6):baseUnit=oRS.getString(7)
    ' Re-read current state; the editable sheet is not authoritative.
    remaining=oRS.getDouble(9)-oRS.getDouble(10)
    If retQty>remaining+0.000001 Then sErr="Количество возврата устарело. Обновите список и повторите ввод.":GoTo Fail
    If UCase(Trim(unitName))<>UCase(Trim(oRS.getString(11))) Then sErr="Единица выдачи изменилась. Обновите список.":GoTo Fail
    If lotID<>Trim(oRS.getString(12)) Then sErr="Партия выдачи изменилась. Обновите список.":GoTo Fail
    Select Case UCase(Trim(oRS.getString(13)))
        Case "ДА","YES","1","TRUE"
        Case Else:sErr="Эта выдача не помечена как возвратная.":GoTo Fail
    End Select
    WMSDBX_CloseRS oRS
    If UCase(Trim(unitName))=UCase(Trim(baseUnit)) Then
        baseQty=retQty
    Else
        baseQty=retQty*WMSWF_LotFactorCon(oCon,lotID,unitName,sErr)
        If sErr<>"" Or baseQty<=0 Then GoTo Fail
    End If

    retID=WMSWF_NewID("RET"):movID=WMSWF_NewID("MOV")
    oRS=oStmt.executeQuery("SELECT COALESCE(MAX(RETURN_SEQ),0)+1 FROM WMS_ISSUE_RETURNS WHERE SOURCE_ID=" & WMSDB_SQLText(sourceID))
    If oRS.next() Then
        retSeq=oRS.getLong(1)
    Else
        retSeq=1
    End If
    WMSDBX_CloseRS oRS
    oStmt.executeUpdate("INSERT INTO WMS_ISSUE_RETURNS (RETURN_ID,SOURCE_ID,RETURN_SEQ,RETURN_QTY,RETURN_DATE,NOTE_TEXT,LOT_ID,ENTRY_UNIT,BASE_QTY,BASE_UNIT) VALUES (" & _
        WMSDB_SQLText(retID) & "," & WMSDB_SQLText(sourceID) & "," & CStr(retSeq) & "," & WMSWF_Num(retQty) & ",CURRENT_DATE," & WMSDB_SQLText(note) & "," & WMSDB_SQLText(lotID) & "," & _
        WMSDB_SQLText(unitName) & "," & WMSWF_Num(baseQty) & "," & WMSDB_SQLText(baseUnit) & ")")
    changedRows=oStmt.executeUpdate("UPDATE WMS_ISSUES SET RETURNED_QTY=COALESCE(RETURNED_QTY,0)+" & WMSWF_Num(retQty) & ",RETURN_DATE=CURRENT_DATE,RETURN_STATE=CASE WHEN COALESCE(RETURNED_QTY,0)+" & WMSWF_Num(retQty) & ">=ISSUE_QTY THEN 'CLOSED' ELSE 'PARTIAL' END,UPDATED_AT=CURRENT_TIMESTAMP WHERE SOURCE_ID=" & WMSDB_SQLText(sourceID) & " AND ISSUE_QTY-COALESCE(RETURNED_QTY,0)>=" & WMSWF_Num(retQty))
    If changedRows<>1 Then sErr="Возврат не записан: доступное количество изменилось.":GoTo Fail
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,NOTE_TEXT,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT,STATUS_NAME,FLOW_CHANNEL,OPERATION_ID) VALUES (" & _
        WMSDB_SQLText(movID) & "," & WMSDB_SQLText(retID) & ",'RETURN','IN'," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(baseQty) & "," & WMSDB_SQLText(baseUnit) & "," & WMSDB_SQLText(loc) & ",CURRENT_DATE," & _
        WMSDB_SQLText(note) & "," & WMSDB_SQLText(origin) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSDB_SQLText(lotID) & "," & WMSWF_Num(retQty) & "," & WMSDB_SQLText(unitName) & "," & _
        WMSWF_Num(baseQty) & "," & WMSDB_SQLText(baseUnit) & ",'POSTED','RETURN'," & WMSDB_SQLText(retID) & ")")
    If Not WMSWF_VerifyMovementAndLotCon(oCon,movID,lotID,baseQty,sErr) Then GoTo Fail
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    If Not WMSDBX_Commit(oCon,sErr) Then
        sErr=sErr & " | Операция: " & retID
        Exit Function
    End If
    commitConfirmed=True
    If Not WMSDBX_SaveODB(sErr) Then
        sErr="Операция " & retID & " зафиксирована в открытой базе, но сохранение ODB не подтверждено. Не повторяйте её. " & sErr
        WMSDBX_BlockConnection sErr
        Exit Function
    End If
    WMSDB_Close
    WMSOPS_ReturnRow=True
    Exit Function
Fail:
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    WMSDBX_RollbackQuiet oCon
    If gWMSDBX_UnsafeConnection Then sErr=sErr & " | " & gWMSDBX_UnsafeReason
    WMSDB_Close
    Exit Function
EH:
    If commitConfirmed Then
        sErr="Операция " & retID & " уже зафиксирована; ошибка последующего обновления: " & CStr(Err) & " " & Error$ & ". Не повторяйте проведение."
        WMSDBX_BlockConnection sErr
        Exit Function
    End If
    sErr="Возврат: " & CStr(Err) & " " & Error$:Resume Fail
End Function

Sub WMSOPS_PrepareInventory(oDoc As Object)
    Dim oSh As Object,h As Variant,i As Long
    If oDoc.Sheets.hasByName(WMSOPS_INV_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSOPS_INV_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSOPS_INV_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSOPS_INV_SHEET)
    End If
    oSh.IsVisible=True
    oSh.getCellRangeByPosition(0,0,12,4).clearContents(1023)
    oSh.getCellByPosition(0,0).String="ИНВЕНТАРИЗАЦИЯ"
    oSh.getCellByPosition(0,1).String="Загрузите остатки по партиям, внесите факт и проведите только расхождения."
    oSh.getCellRangeByPosition(0,0,10,0).CellBackColor=10027008:oSh.getCellRangeByPosition(0,0,10,0).CharColor=16777215:oSh.getCellRangeByPosition(0,0,10,0).CharWeight=150
    h=Array("Партия","Код","Наименование","Место","Источник","Категория","Подкатегория","Учёт","Факт","Разница","Ед.","Комментарий")
    For i=0 To UBound(h):oSh.getCellByPosition(i,4).String=CStr(h(i)):Next i
    oSh.getCellRangeByPosition(0,4,11,4).CellBackColor=4473924:oSh.getCellRangeByPosition(0,4,11,4).CharColor=16777215:oSh.getCellRangeByPosition(0,4,11,4).CharWeight=150
    oSh.Columns.getByIndex(2).Width=6500:oSh.Columns.getByIndex(3).Width=4000:oSh.Columns.getByIndex(11).Width=6000
    WMSOPS_InstallInventoryButtons oDoc,oSh
End Sub

Sub WMSOPS_LoadInventory()
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object,oRS As Object,sErr As String,r As Long,c As Object,lastR As Long
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSOPS_INV_SHEET)
    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then MsgBox sErr,16,"WMS — Инвентаризация":Exit Sub
    On Error GoTo EH
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow:If lastR<5 Then lastR=5
    oSh.getCellRangeByPosition(0,5,11,lastR).clearContents(1023)
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT LOT_ID,PRODUCT_CODE,PRODUCT_NAME,LOCATION_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME,QTY,UNIT_NAME FROM WMS_V_LOT_BALANCE WHERE QTY<>0 ORDER BY LOCATION_NAME,PRODUCT_NAME,LOT_ID")
    r=5
    Do While oRS.next()
        oSh.getCellByPosition(0,r).String=oRS.getString(1):oSh.getCellByPosition(1,r).String=oRS.getString(2):oSh.getCellByPosition(2,r).String=oRS.getString(3)
        oSh.getCellByPosition(3,r).String=oRS.getString(4):oSh.getCellByPosition(4,r).String=oRS.getString(5):oSh.getCellByPosition(5,r).String=oRS.getString(6):oSh.getCellByPosition(6,r).String=oRS.getString(7)
        oSh.getCellByPosition(7,r).Value=oRS.getDouble(8):oSh.getCellByPosition(8,r).Value=oRS.getDouble(8):oSh.getCellByPosition(9,r).Value=0:oSh.getCellByPosition(10,r).String=oRS.getString(9)
        r=r+1
    Loop
    WMSDB_Close
    oSh.getCellByPosition(0,2).String="Загружено партий: " & CStr(r-5) & ". Измените только колонку Факт."
    Exit Sub
EH:
    WMSDB_Close
    MsgBox "Загрузка инвентаризации: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSOPS_RecalcInventory()
    Dim oSh As Object,r As Long,lastR As Long,c As Object
    oSh=ThisComponent.Sheets.getByName(WMSOPS_INV_SHEET):c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow
    For r=5 To lastR:oSh.getCellByPosition(9,r).Value=oSh.getCellByPosition(8,r).Value-oSh.getCellByPosition(7,r).Value:Next r
End Sub

Sub WMSOPS_ConductInventory()
    Dim oDoc As Object,oSh As Object,r As Long,lastR As Long,c As Object,diff As Double,nOK As Long,nErr As Long,sErr As String
    If Not WMSDBX_TryEnter("INVENTORY") Then MsgBox "Дождитесь завершения другой операции WMS.",48,"WMS":Exit Sub
    On Error GoTo BatchFailed
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSOPS_INV_SHEET):WMSOPS_RecalcInventory
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow
    For r=5 To lastR
        diff=oSh.getCellByPosition(9,r).Value
        If Abs(diff)>0.000001 Then
            sErr=""
            If WMSOPS_InventoryRow(oDoc,oSh,r,diff,sErr) Then
                nOK=nOK+1
            Else
                nErr=nErr+1
                MsgBox "Строка " & CStr(r+1) & ": " & sErr,48,"WMS — Инвентаризация"
                Exit For
            End If
        End If
    Next r
    WMSDBX_Leave
    If gWMSDBX_UnsafeConnection Then Exit Sub
    On Error Resume Next:WMSDBST_RefreshStock:On Error GoTo 0
    MsgBox "Корректировок: " & CStr(nOK) & Chr(10) & "Ошибок: " & CStr(nErr),IIf(nErr=0,64,48),"WMS — Инвентаризация"
    If nErr=0 Then WMSOPS_LoadInventory
    Exit Sub
BatchFailed:
    sErr="Обработка остановлена: " & CStr(Err) & " " & Error$ & ". Успешные строки могли быть проведены; проверьте результат перед повтором."
    WMSDBX_Leave
    MsgBox sErr,16,"WMS"
End Sub

Function WMSOPS_InventoryRow(oDoc As Object,oSh As Object,r As Long,diff As Double,ByRef sErr As String) As Boolean
    Dim commitConfirmed As Boolean,changedRows As Long
    Dim lotID As String,code As String,nm As String,loc As String,origin As String,cat As String,subcat As String,unitName As String,note As String,movID As String,invID As String
    Dim oCon As Object,oStmt As Object,oRS As Object,currentQty As Double
    WMSOPS_InventoryRow=False:sErr=""
    lotID=Trim(oSh.getCellByPosition(0,r).String):code=Trim(oSh.getCellByPosition(1,r).String):nm=Trim(oSh.getCellByPosition(2,r).String)
    loc=Trim(oSh.getCellByPosition(3,r).String):origin=Trim(oSh.getCellByPosition(4,r).String):cat=Trim(oSh.getCellByPosition(5,r).String):subcat=Trim(oSh.getCellByPosition(6,r).String)
    unitName=Trim(oSh.getCellByPosition(10,r).String):note=Trim(oSh.getCellByPosition(11,r).String)
    If lotID="" Or code="" Or loc="" Or unitName="" Then sErr="Неполная строка инвентаризации.":Exit Function

    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then Exit Function
    On Error GoTo EH
    If Not WMSDBX_Begin(oCon,sErr) Then Exit Function
    oStmt=oCon.createStatement():movID=WMSWF_NewID("MOV"):invID=WMSWF_NewID("INV")
    ' Reject an old count sheet, including a retry after a prior committed row.
    ' Match all dimensions used by the lot balance view.
    oRS=oStmt.executeQuery("SELECT COALESCE(SUM(QTY),0) FROM WMS_V_LOT_BALANCE WHERE LOT_ID=" & WMSDB_SQLText(lotID) & _
        " AND PRODUCT_CODE=" & WMSDB_SQLText(code) & " AND UNIT_NAME=" & WMSDB_SQLText(unitName) & _
        " AND LOCATION_NAME=" & WMSDB_SQLText(loc) & " AND ORIGIN_NAME=" & WMSDB_SQLText(origin) & _
        " AND CATEGORY_NAME=" & WMSDB_SQLText(cat) & " AND SUBCATEGORY_NAME=" & WMSDB_SQLText(subcat))
    If Not oRS.next() Then sErr="Не удалось проверить текущий остаток партии.":GoTo Fail
    currentQty=oRS.getDouble(1)
    WMSDBX_CloseRS oRS
    If Abs(currentQty-oSh.getCellByPosition(7,r).Value)>0.000001 Then
        sErr="Остаток партии изменился после загрузки. Сохраните результаты подсчёта и загрузите свежий остаток."
        GoTo Fail
    End If
    If oSh.getCellByPosition(8,r).Type<>com.sun.star.table.CellContentType.VALUE Then
        sErr="Факт должен быть введён числом, а не пустой ячейкой или текстом.":GoTo Fail
    End If
    If oSh.getCellByPosition(8,r).Value<0 Then sErr="Фактическое количество не может быть отрицательным.":GoTo Fail
    oStmt.executeUpdate("INSERT INTO WMS_CORRECTIONS (CORRECTION_ID,SOURCE_ENTITY_TYPE,SOURCE_ENTITY_ID,CORRECTION_TYPE,REASON_TEXT) VALUES (" & _
        WMSDB_SQLText(invID) & ",'LOT'," & WMSDB_SQLText(lotID) & ",'INVENTORY'," & WMSDB_SQLText(IIf(note="","Инвентаризация",note)) & ")")
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,NOTE_TEXT,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT,STATUS_NAME,FLOW_CHANNEL,OPERATION_ID) VALUES (" & _
        WMSDB_SQLText(movID) & "," & WMSDB_SQLText(invID) & ",'INVENTORY','ADJUST'," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(diff) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & ",CURRENT_DATE," & _
        WMSDB_SQLText(note) & "," & WMSDB_SQLText(origin) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSDB_SQLText(lotID) & "," & WMSWF_Num(diff) & "," & WMSDB_SQLText(unitName) & "," & WMSWF_Num(diff) & "," & WMSDB_SQLText(unitName) & ",'POSTED','INVENTORY'," & WMSDB_SQLText(invID) & ")")
    If Not WMSWF_VerifyMovementAndLotCon(oCon,movID,lotID,diff,sErr) Then GoTo Fail
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    If Not WMSDBX_Commit(oCon,sErr) Then
        sErr=sErr & " | Операция: " & invID
        Exit Function
    End If
    commitConfirmed=True
    If Not WMSDBX_SaveODB(sErr) Then
        sErr="Операция " & invID & " зафиксирована в открытой базе, но сохранение ODB не подтверждено. Не повторяйте её. " & sErr
        WMSDBX_BlockConnection sErr
        Exit Function
    End If
    WMSDB_Close
    ' Successful rows no longer retain a difference on a partially failed batch.
    oSh.getCellByPosition(7,r).Value=oSh.getCellByPosition(8,r).Value
    oSh.getCellByPosition(9,r).Value=0
    WMSOPS_InventoryRow=True
    Exit Function
Fail:
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    WMSDBX_RollbackQuiet oCon
    If gWMSDBX_UnsafeConnection Then sErr=sErr & " | " & gWMSDBX_UnsafeReason
    WMSDB_Close
    Exit Function
EH:
    If commitConfirmed Then
        sErr="Операция " & invID & " уже зафиксирована; ошибка последующего обновления: " & CStr(Err) & " " & Error$ & ". Не повторяйте проведение."
        WMSDBX_BlockConnection sErr
        Exit Function
    End If
    sErr="Инвентаризация: " & CStr(Err) & " " & Error$:Resume Fail
End Function

Sub WMSOPS_InstallReturnButtons(oDoc As Object,oSh As Object)
    WMSOPS_Buttons oDoc,oSh,"WMS_RET_FORM",Array("Обновить","Провести возвраты"),Array("WMSOPS_RefreshReturns","WMSOPS_ConductReturns"),"RET"
End Sub

Sub WMSOPS_InstallInventoryButtons(oDoc As Object,oSh As Object)
    WMSOPS_Buttons oDoc,oSh,"WMS_INV_FORM",Array("Загрузить остаток","Пересчитать","Провести расхождения"),Array("WMSOPS_LoadInventory","WMSOPS_RecalcInventory","WMSOPS_ConductInventory"),"INV"
End Sub

Sub WMSOPS_Buttons(oDoc As Object,oSh As Object,formName As String,labels As Variant,macros As Variant,prefix As String)
    Dim forms As Object,form As Object,i As Long
    WMSOPS_RemoveButtons oSh,prefix
    forms=oSh.DrawPage.Forms
    On Error Resume Next:If forms.hasByName(formName) Then forms.removeByName(formName):On Error GoTo EH
    form=oDoc.createInstance("com.sun.star.form.component.Form"):form.Name=formName:forms.insertByName(formName,form)
    For i=0 To UBound(labels):WMSOPS_AddButton oDoc,oSh,form,"WMS_" & prefix & "_" & CStr(i),CStr(labels(i)),i,3,4200,800,CStr(macros(i)):Next i
    Exit Sub
EH:
    MsgBox "Кнопки: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSOPS_AddButton(oDoc As Object,oSh As Object,oForm As Object,sName As String,sLabel As String,nCol As Long,nRow As Long,nW As Long,nH As Long,sMacro As String)
    Dim model As Object,shape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    model=oDoc.createInstance("com.sun.star.form.component.CommandButton"):model.Name=sName:model.Label=sLabel
    oForm.insertByName(sName,model):idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=model:a=oSh.getCellByPosition(nCol,nRow).Position:shape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=nW:sz.Height=nH:shape.Size=sz:oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed":ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_23_ReturnsInventory." & sMacro & "?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSOPS_RemoveButtons(oSh As Object,prefix As String)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i):nm="":ctl=shape.Control:nm=ctl.Name
        If Left(nm,5+Len(prefix))="WMS_" & prefix & "_" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub

Sub WMSOPS_SyntaxProbe()
End Sub
