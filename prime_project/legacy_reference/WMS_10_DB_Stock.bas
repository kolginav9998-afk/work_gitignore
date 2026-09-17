Option Explicit

' ============================================================
' WMS_10_DB_Stock
' Движения + Номенклатура + Остатки
' Версия 0.1.0 FOUNDATION
'
' Зависимость:
'   Standard.WMS_04_DB_Connection (PORTABLE)
'
' Принцип:
'   PRODUCT_CODE = обязательный внутренний код пользователя.
'   Номенклатура создаётся/обновляется по PRODUCT_CODE.
'   Остаток НЕ хранится отдельным числом — считается из
'   неизменяемого журнала WMS_STOCK_MOVEMENTS.
' ============================================================

Global Const WMSDBST_VERSION = "3.1.0-NARROW-BATCH-STOCK"
Global Const WMSDBST_SHEET = "Остаток"
Global Const WMSDBST_FORM = "WMS_STOCK_PANEL"

Sub WMSDBST_Install()
    Dim oDoc As Object,oCon As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail

    If Not WMSDBST_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail
    WMSDB_Close

    WMSDBST_InstallStockSheet oDoc
    WMSDBST_RefreshStock

    MsgBox "WMS_10 установлен." & Chr(10) & _
           "Лист Остаток обновлён." & Chr(10) & _
           "Версия: " & WMSDBST_VERSION,64,"WMS — Движения + Остатки"
    Exit Sub
Fail:
    WMSDB_Close
    MsgBox sErr,16,"WMS — Движения + Остатки"
    Exit Sub
EH:
    sErr="Установка WMS_10: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSDBST_EnsureSchema(oCon As Object, ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSDBST_EnsureSchema = False
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()

    If Not WMSDB_TableExistsSimple(oCon, "WMS_PRODUCTS", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate( _
            "CREATE TABLE WMS_PRODUCTS (" & _
            "PRODUCT_CODE VARCHAR(80) NOT NULL PRIMARY KEY," & _
            "PRODUCT_NAME VARCHAR(255) NOT NULL," & _
            "UNIT_NAME VARCHAR(40) NOT NULL," & _
            "DEFAULT_LOCATION VARCHAR(120)," & _
            "SUPPLIER_ARTICLE VARCHAR(120)," & _
            "ACTIVE_FLAG SMALLINT DEFAULT 1 NOT NULL," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)")
    End If

    If Not WMSDB_TableExistsSimple(oCon, "WMS_STOCK_MOVEMENTS", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate( _
            "CREATE TABLE WMS_STOCK_MOVEMENTS (" & _
            "MOVEMENT_ID VARCHAR(100) NOT NULL PRIMARY KEY," & _
            "SOURCE_ID VARCHAR(100) NOT NULL," & _
            "SOURCE_TYPE VARCHAR(40) NOT NULL," & _
            "MOVEMENT_TYPE VARCHAR(40) NOT NULL," & _
            "PRODUCT_CODE VARCHAR(80) NOT NULL," & _
            "PRODUCT_NAME VARCHAR(255) NOT NULL," & _
            "QTY DECIMAL(18,4) NOT NULL," & _
            "UNIT_NAME VARCHAR(40) NOT NULL," & _
            "LOCATION_NAME VARCHAR(120) NOT NULL," & _
            "ORIGIN_NAME VARCHAR(120)," & _
            "DEST_CATEGORY VARCHAR(120)," & _
            "DEST_SUBCATEGORY VARCHAR(120)," & _
            "MOVEMENT_DATE DATE NOT NULL," & _
            "NOTE_TEXT VARCHAR(500)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "CONSTRAINT FK_WMS_MOV_PRODUCT FOREIGN KEY (PRODUCT_CODE) " & _
            "REFERENCES WMS_PRODUCTS(PRODUCT_CODE))")
    End If

    ' Миграция существующей схемы 0.1.0 без потери данных.
    If Not WMSDBST_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "ORIGIN_NAME", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD ORIGIN_NAME VARCHAR(120)")
    End If
    If Not WMSDBST_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "DEST_CATEGORY", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD DEST_CATEGORY VARCHAR(120)")
    End If
    If Not WMSDBST_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "DEST_SUBCATEGORY", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD DEST_SUBCATEGORY VARCHAR(120)")
    End If

    If Not WMSDBST_IndexExists(oCon, "UX_WMS_MOV_SOURCE", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE UNIQUE INDEX UX_WMS_MOV_SOURCE ON WMS_STOCK_MOVEMENTS(SOURCE_TYPE,SOURCE_ID,MOVEMENT_TYPE,PRODUCT_CODE,LOCATION_NAME)")
    End If

    If Not WMSDBST_IndexExists(oCon, "IX_WMS_MOV_PRODUCT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE INDEX IX_WMS_MOV_PRODUCT ON WMS_STOCK_MOVEMENTS(PRODUCT_CODE)")
    End If

    If Not WMSDBST_IndexExists(oCon, "IX_WMS_MOV_LOCATION", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE INDEX IX_WMS_MOV_LOCATION ON WMS_STOCK_MOVEMENTS(LOCATION_NAME)")
    End If

    If Not WMSARCH_EnsureSchema(oCon,sErr) Then Exit Function
    If Not WMSINT_EnsureSchema(oCon,sErr) Then Exit Function
    oCon.commit()
    WMSDBST_EnsureSchema = True
    Exit Function
EH:
    sErr = "Не удалось создать схему остатков: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_ColumnExists(oCon As Object, sTable As String, sColumn As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sql As String
    WMSDBST_ColumnExists = False
    sErr = ""
    On Error GoTo EH
    sql = "SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE " & _
          "TRIM(RDB$RELATION_NAME)='" & WMSDBST_SQL(UCase(sTable)) & "' AND " & _
          "TRIM(RDB$FIELD_NAME)='" & WMSDBST_SQL(UCase(sColumn)) & "'"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBST_ColumnExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Ошибка проверки поля " & sTable & "." & sColumn & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_IndexExists(oCon As Object, sIndex As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sql As String
    WMSDBST_IndexExists = False
    sErr = ""
    On Error GoTo EH
    sql = "SELECT COUNT(*) FROM RDB$INDICES WHERE TRIM(RDB$INDEX_NAME)='" & _
          WMSDBST_SQL(sIndex) & "'"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBST_IndexExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Ошибка проверки индекса " & sIndex & ": " & CStr(Err) & " " & Error$
End Function

' ============================================================
' API ДЛЯ БУДУЩЕГО ПОДКЛЮЧЕНИЯ ORDERS / ISSUES / RETURNS
' ============================================================

Function WMSDBST_PostMovement(oDoc As Object, sMovementID As String, _
    sSourceID As String, sSourceType As String, sMovementType As String, _
    sProductCode As String, sProductName As String, dQty As Double, _
    sUnit As String, sLocation As String, dMovementDate As Double, _
    sNote As String, ByRef sErr As String) As Boolean

    Dim oCon As Object, oStmt As Object, sql As String
    WMSDBST_PostMovement = False
    sErr = ""

    If Trim(sProductCode) = "" Then
        sErr = "Проведение запрещено: внутренний Код товара обязателен."
        Exit Function
    End If
    If Trim(sProductName) = "" Then
        sErr = "Проведение запрещено: наименование товара пустое."
        Exit Function
    End If
    If Trim(sUnit) = "" Then
        sErr = "Проведение запрещено: единица измерения пустая."
        Exit Function
    End If
    If Trim(sLocation) = "" Then
        sErr = "Проведение запрещено: место хранения пустое."
        Exit Function
    End If
    If dQty = 0 Then
        sErr = "Проведение запрещено: движение с количеством 0."
        Exit Function
    End If

    On Error GoTo EH
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then Exit Function
    If Not WMSDBST_EnsureSchema(oCon, sErr) Then GoTo Fail

    If Not WMSDBST_EnsureProductCon(oCon, sProductCode, sProductName, sUnit, _
                                    sLocation, "", sErr) Then GoTo Fail

    ' Идемпотентность: один и тот же MOVEMENT_ID второй раз не проводится.
    If WMSDBST_MovementExists(oCon, sMovementID, sErr) Then
        If sErr <> "" Then GoTo Fail
        WMSDBST_PostMovement = True
        WMSDB_Close
        Exit Function
    End If

    sql = "INSERT INTO WMS_STOCK_MOVEMENTS " & _
          "(MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,NOTE_TEXT) VALUES (" & _
          "'" & WMSDBST_SQL(sMovementID) & "'," & _
          "'" & WMSDBST_SQL(sSourceID) & "'," & _
          "'" & WMSDBST_SQL(sSourceType) & "'," & _
          "'" & WMSDBST_SQL(sMovementType) & "'," & _
          "'" & WMSDBST_SQL(sProductCode) & "'," & _
          "'" & WMSDBST_SQL(sProductName) & "'," & _
          WMSDBST_SQLNumber(dQty) & "," & _
          "'" & WMSDBST_SQL(sUnit) & "'," & _
          "'" & WMSDBST_SQL(sLocation) & "'," & _
          "'" & WMSDBST_SQLDate(dMovementDate) & "'," & _
          "'" & WMSDBST_SQL(sNote) & "')"

    oStmt = oCon.createStatement()
    oStmt.executeUpdate(sql)
    oCon.commit()

    If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail
    WMSDB_Close

    ' Повторное открытие и проверка — только после этого считаем движение надёжно записанным.
    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then Exit Function
    If Not WMSDBST_MovementExists(oCon, sMovementID, sErr) Then
        If sErr = "" Then sErr = "Движение записалось, но не прошло контрольное чтение."
        GoTo Fail
    End If

    WMSDBST_PostMovement = True
    WMSDB_Close
    Exit Function
Fail:
    WMSDB_Close
    Exit Function
EH:
    sErr = "Ошибка проведения движения: " & CStr(Err) & " " & Error$
    WMSDB_Close
End Function

Function WMSDBST_EnsureProductCon(oCon As Object, sCode As String, sName As String, _
    sUnit As String, sLocation As String, sSupplierArticle As String, _
    ByRef sErr As String) As Boolean

    Dim oStmt As Object, sql As String
    WMSDBST_EnsureProductCon = False
    sErr = ""
    If Trim(sCode) = "" Then
        sErr = "Номенклатура не может быть создана без внутреннего кода."
        Exit Function
    End If
    On Error GoTo EH

    sql = "UPDATE OR INSERT INTO WMS_PRODUCTS " & _
          "(PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION,SUPPLIER_ARTICLE,ACTIVE_FLAG,UPDATED_AT) VALUES (" & _
          "'" & WMSDBST_SQL(sCode) & "'," & _
          "'" & WMSDBST_SQL(sName) & "'," & _
          "'" & WMSDBST_SQL(sUnit) & "'," & _
          "'" & WMSDBST_SQL(sLocation) & "'," & _
          "'" & WMSDBST_SQL(sSupplierArticle) & "',1,CURRENT_TIMESTAMP) " & _
          "MATCHING (PRODUCT_CODE)"
    oStmt = oCon.createStatement()
    oStmt.executeUpdate(sql)
    WMSDBST_EnsureProductCon = True
    Exit Function
EH:
    sErr = "Ошибка номенклатуры " & sCode & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_MovementExists(oCon As Object, sMovementID As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object
    WMSDBST_MovementExists = False
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS WHERE MOVEMENT_ID='" & WMSDBST_SQL(sMovementID) & "'")
    If oRS.next() Then WMSDBST_MovementExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Ошибка проверки движения: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_PostMovementClassified(oDoc As Object, sMovementID As String, _
    sSourceID As String, sSourceType As String, sMovementType As String, _
    sProductCode As String, sProductName As String, dQty As Double, _
    sUnit As String, sLocation As String, sOrigin As String, _
    sDestCategory As String, sDestSubcategory As String, _
    dMovementDate As Double, sNote As String, ByRef sErr As String) As Boolean

    Dim oCon As Object
    WMSDBST_PostMovementClassified=False
    sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    If Not WMSDBST_PostMovementClassifiedCon(oCon,sMovementID,sSourceID,sSourceType,sMovementType, _
        sProductCode,sProductName,dQty,sUnit,sLocation,sOrigin,sDestCategory,sDestSubcategory, _
        dMovementDate,sNote,sErr) Then
        Exit Function
    End If

    If Not WMSDB_SaveDatabaseDocument(sErr) Then Exit Function

    ' Important: do not close/reopen Base here. The caller may already be in the
    ' middle of a conduct operation. Repeated embedded-Base open/close cycles
    ' caused LibreOffice freezes on Windows.
    If Not WMSDBST_MovementExists(oCon,sMovementID,sErr) Then
        If sErr="" Then sErr="Движение не найдено после COMMIT."
        Exit Function
    End If

    WMSDBST_PostMovementClassified=True
    Exit Function
EH:
    sErr="Ошибка проведения движения: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_PostMovementClassifiedCon(oCon As Object, sMovementID As String, _
    sSourceID As String, sSourceType As String, sMovementType As String, _
    sProductCode As String, sProductName As String, dQty As Double, _
    sUnit As String, sLocation As String, sOrigin As String, _
    sDestCategory As String, sDestSubcategory As String, _
    dMovementDate As Double, sNote As String, ByRef sErr As String) As Boolean

    Dim oStmt As Object,sql As String
    WMSDBST_PostMovementClassifiedCon=False
    sErr=""

    If Trim(sProductCode)="" Then sErr="Проведение запрещено: внутренний Код товара обязателен.":Exit Function
    If Trim(sProductName)="" Then sErr="Проведение запрещено: наименование товара пустое.":Exit Function
    If Trim(sUnit)="" Then sErr="Проведение запрещено: единица измерения пустая.":Exit Function
    If Trim(sLocation)="" Then sErr="Проведение запрещено: место хранения пустое.":Exit Function
    If dQty=0 Then sErr="Проведение запрещено: движение с количеством 0.":Exit Function

    On Error GoTo EH

    If Not WMSDBST_EnsureSchema(oCon,sErr) Then Exit Function
    If Not WMSDBST_EnsureProductCon(oCon,sProductCode,sProductName,sUnit,sLocation,"",sErr) Then Exit Function

    ' Idempotency: a repeated click must never double the stock.
    If WMSDBST_MovementExists(oCon,sMovementID,sErr) Then
        If sErr<>"" Then Exit Function
        WMSDBST_PostMovementClassifiedCon=True
        Exit Function
    End If

    sql="INSERT INTO WMS_STOCK_MOVEMENTS " & _
        "(MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME," & _
        "ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,MOVEMENT_DATE,NOTE_TEXT) VALUES (" & _
        "'"&WMSDBST_SQL(sMovementID)&"','"&WMSDBST_SQL(sSourceID)&"','"&WMSDBST_SQL(sSourceType)&"'," & _
        "'"&WMSDBST_SQL(sMovementType)&"','"&WMSDBST_SQL(sProductCode)&"','"&WMSDBST_SQL(sProductName)&"'," & _
        WMSDBST_SQLNumber(dQty)&",'"&WMSDBST_SQL(sUnit)&"','"&WMSDBST_SQL(sLocation)&"'," & _
        "'"&WMSDBST_SQL(sOrigin)&"','"&WMSDBST_SQL(sDestCategory)&"','"&WMSDBST_SQL(sDestSubcategory)&"'," & _
        "'"&WMSDBST_SQLDate(dMovementDate)&"','"&WMSDBST_SQL(sNote)&"')"

    oStmt=oCon.createStatement()
    oStmt.executeUpdate(sql)
    oCon.commit()

    If Not WMSDBST_MovementExists(oCon,sMovementID,sErr) Then
        If sErr="" Then sErr="Движение не найдено после COMMIT."
        Exit Function
    End If

    WMSDBST_PostMovementClassifiedCon=True
    Exit Function
EH:
    sErr="Ошибка записи движения: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_VerifyMovementCon(oCon As Object,sMovementID As String,ByRef sErr As String) As Boolean
    WMSDBST_VerifyMovementCon=WMSDBST_MovementExists(oCon,sMovementID,sErr)
End Function

 ' ============================================================
' ЛИСТ ОСТАТОК — ПОЛНОЦЕННЫЙ FIREBIRD UI
' ============================================================

Sub WMSDBST_RefreshStock()
    WMSDBST_RefreshStockFilter "ALL"
End Sub

Sub WMSDBST_ShowAll()
    WMSDBST_RefreshStockFilter "ALL"
End Sub

Sub WMSDBST_ShowOffice()
    WMSDBST_RefreshStockFilter "OFFICE"
End Sub

Sub WMSDBST_ShowProduction()
    WMSDBST_RefreshStockFilter "PRODUCTION"
End Sub

Sub WMSDBST_ShowParts()
    WMSDBST_RefreshStockFilter "PARTS"
End Sub

Sub WMSDBST_ShowNegative()
    WMSDBST_RefreshStockFilter "NEGATIVE"
End Sub

Sub WMSDBST_RefreshStockFilter(mode As String)
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim sErr As String,sql As String,whereSQL As String,havingSQL As String,title As String
    Dim r As Long,nRows As Long,batch(0 To 199) As Variant,bn As Long,t0 As Double,hasMore As Boolean
    oDoc=ThisComponent
    If Not WMSDBX_TryEnter("STOCK") Then MsgBox "Другая операция WMS ещё выполняется.",48,"WMS — Остаток":Exit Sub
    On Error GoTo EH
    t0=Timer
    oDoc.lockControllers()
    If Not WMSDBST_GetSheet(oDoc,oSh) Then sErr="Лист Остаток не найден.":GoTo Fail
    oCon=WMSDB_GetConnectionEx(oDoc,sErr):If sErr<>"" Then GoTo Fail

    whereSQL=" WHERE COALESCE(m.STATUS_NAME,'POSTED')='POSTED'"
    havingSQL=" HAVING ABS(SUM(m.QTY))>0.000001"
    Select Case UCase(Trim(mode))
        Case "OFFICE":whereSQL=whereSQL & " AND UPPER(TRIM(COALESCE(m.ORIGIN_NAME,'')))='OFFICE'":title="Из офиса"
        Case "PRODUCTION":whereSQL=whereSQL & " AND UPPER(TRIM(COALESCE(m.ORIGIN_NAME,''))) IN ('PRODUCTION','PARTS','WORKSHOP')":title="Из производства"
        Case "PARTS":whereSQL=whereSQL & " AND UPPER(TRIM(COALESCE(m.ORIGIN_NAME,'')))='PARTS'":title="Детали"
        Case "NEGATIVE":havingSQL=" HAVING SUM(m.QTY)<-0.000001":title="Отрицательные"
        Case Else:title="Весь склад"
    End Select

    ' Narrow result set on purpose: only seven typed fields.
    ' No nullable category/origin fields are fetched by Calc; they are only SQL filters.
    sql="SELECT m.PRODUCT_CODE,CAST(MAX(m.PRODUCT_NAME) AS VARCHAR(255))," & _
        "CAST(COALESCE(MAX(p.SUPPLIER_ARTICLE),'') AS VARCHAR(120))," & _
        "CAST(m.UNIT_NAME AS VARCHAR(40)),CAST(m.LOCATION_NAME AS VARCHAR(120))," & _
        "CAST(SUM(m.QTY) AS DOUBLE PRECISION),CAST(MAX(m.MOVEMENT_DATE) AS VARCHAR(10)) " & _
        "FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=m.PRODUCT_CODE" & whereSQL & _
        " GROUP BY m.PRODUCT_CODE,m.UNIT_NAME,m.LOCATION_NAME" & havingSQL & _
        " ORDER BY MAX(m.PRODUCT_NAME),m.LOCATION_NAME"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)

    oSh.getCellRangeByPosition(0,6,6,2005).clearContents(1023)
    r=6:bn=0
    Do While oRS.next()
        batch(bn)=Array(oRS.getString(1),oRS.getString(2),oRS.getString(3),oRS.getString(4),oRS.getString(5),oRS.getDouble(6),oRS.getString(7))
        bn=bn+1:nRows=nRows+1
        If bn=200 Then
            WMSDBST_WriteBatch oSh,r,batch,bn
            r=r+bn:bn=0
        End If
        If nRows>=2000 Then
            hasMore=oRS.next()
            Exit Do
        End If
    Loop
    If bn>0 Then WMSDBST_WriteBatch oSh,r,batch,bn

    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt:WMSDB_Close
    oSh.getCellByPosition(0,3).String="Режим: " & title
    oSh.getCellByPosition(0,4).String="Позиций: " & CStr(nRows) & " | " & CStr(CLng((Timer-t0)*1000)) & " мс | " & Format(Now,"DD.MM.YYYY HH:MM:SS")
    If hasMore Then oSh.getCellByPosition(0,4).String="ВНИМАНИЕ: показаны первые 2000 позиций; список неполный. Режим: " & title & " | " & Format(Now,"DD.MM.YYYY HH:MM:SS")
Done:
    On Error Resume Next:oDoc.unlockControllers():On Error GoTo 0
    WMSDBX_Leave
    Exit Sub
Fail:
    On Error Resume Next
    oSh.getCellByPosition(0,4).String="ОБНОВЛЕНИЕ НЕ ЗАВЕРШЕНО. Данные на экране могут быть неполными или устаревшими."
    On Error GoTo 0
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    On Error Resume Next:WMSDB_Close:oDoc.unlockControllers():On Error GoTo 0
    WMSDBX_Leave
    MsgBox sErr,16,"WMS — Остаток"
    Exit Sub
EH:
    sErr="Остаток: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSDBST_WriteBatch(oSh As Object,startRow As Long,batch As Variant,n As Long)
    Dim out() As Variant,i As Long
    If n<=0 Then Exit Sub
    ReDim out(0 To n-1) As Variant
    For i=0 To n-1:out(i)=batch(i):Next i
    oSh.getCellRangeByPosition(0,startRow,6,startRow+n-1).setDataArray(out())
End Sub

Sub WMSDBST_Diagnostics()
    WMS26_RunDiagnostics
End Sub

Sub WMSDBST_InstallStockSheet(oDoc As Object)
    Dim oSh As Object,h As Variant,i As Long
    If oDoc.Sheets.hasByName(WMSDBST_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSDBST_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSDBST_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSDBST_SHEET)
    End If
    oSh.IsVisible=True
    On Error GoTo EH

    oSh.getCellRangeByPosition(0,0,10,5).clearContents(1023)
    oSh.getCellByPosition(0,0).String="ОСТАТОК СКЛАДА"
    oSh.getCellByPosition(0,1).String="Источник данных: Firebird / журнал движений. Остаток считается напрямую, без формул Calc и без промежуточного VIEW."
    oSh.getCellByPosition(0,3).String="Режим: Весь склад"
    oSh.getCellByPosition(0,4).String=""

    h=Array("Код","Наименование","Артикул","Ед.","Место хранения","Остаток","Последнее движение")
    For i=0 To UBound(h):oSh.getCellByPosition(i,5).String=CStr(h(i)):Next i

    oSh.getCellRangeByPosition(0,0,6,0).CellBackColor=2500134:oSh.getCellRangeByPosition(0,0,6,0).CharColor=16777215:oSh.getCellRangeByPosition(0,0,6,0).CharWeight=150
    oSh.getCellRangeByPosition(0,1,6,1).CellBackColor=15132390
    oSh.getCellRangeByPosition(0,5,6,5).CellBackColor=4473924:oSh.getCellRangeByPosition(0,5,6,5).CharColor=16777215:oSh.getCellRangeByPosition(0,5,6,5).CharWeight=150
    oSh.Columns.getByIndex(0).Width=3600:oSh.Columns.getByIndex(1).Width=7600:oSh.Columns.getByIndex(2).Width=4200:oSh.Columns.getByIndex(3).Width=2400
    oSh.Columns.getByIndex(4).Width=4800:oSh.Columns.getByIndex(5).Width=3000:oSh.Columns.getByIndex(6).Width=3600

    WMSDBST_InstallButtons oDoc,oSh
    WMSDBST_RefreshStock
    Exit Sub
EH:
    MsgBox "Интерфейс остатка: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSDBST_InstallButtons(oDoc As Object,oSh As Object)
    Dim oForms As Object,oForm As Object
    On Error GoTo EH

    WMSDBST_RemoveButtons oSh
    oForms=oSh.DrawPage.Forms

    If oForms.hasByName(WMSDBST_FORM) Then oForms.removeByName(WMSDBST_FORM)

    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name=WMSDBST_FORM
    oForms.insertByName(WMSDBST_FORM,oForm)

    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_ALL","Весь склад",0,2,2800,800,"WMSDBST_ShowAll"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_OFFICE","Из офиса",1,2,2800,800,"WMSDBST_ShowOffice"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_PROD","Из производства",2,2,4000,800,"WMSDBST_ShowProduction"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_PARTS","Детали",3,2,2500,800,"WMSDBST_ShowParts"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_NEG","Отрицательные",4,2,3400,800,"WMSDBST_ShowNegative"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_HISTORY","История",5,2,2600,800,"WMSDBST_StockHistorySelected"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_LOTS","Партии",6,2,2400,800,"WMSDBST_LotsSelected"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_SEARCH","Поиск",7,2,2200,800,"WMSDBST_OpenSearch"
    WMSDBST_AddButton oDoc,oSh,oForm,"WMS_STOCK_DIAG","Диагностика",8,2,3000,800,"WMSDBST_Diagnostics"
    Exit Sub
EH:
    MsgBox "Не удалось создать панель Остаток: " & CStr(Err) & " " & Error$,16,"WMS — Остаток"
End Sub

Sub WMSDBST_AddButton(oDoc As Object,oSh As Object,oForm As Object,sName As String, _
    sLabel As String,nCol As Long,nRow As Long,nW As Long,nH As Long,sMacroName As String)

    Dim oModel As Object,oShape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor

    oModel=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    oModel.Name=sName
    oModel.Label=sLabel
    oModel.Tabstop=False
    oForm.insertByName(sName,oModel)
    idx=oForm.Count-1

    oShape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    oShape.Control=oModel
    a=oSh.getCellByPosition(nCol,nRow).Position
    oShape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size")
    sz.Width=nW
    sz.Height=nH
    oShape.Size=sz
    oSh.DrawPage.add(oShape)

    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.AddListenerParam=""
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_10_DB_Stock." & sMacroName & "?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
End Sub

Sub WMSDBST_RemoveButtons(oSh As Object)
    Dim i As Long,oShape As Object,sName As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        oShape=oSh.DrawPage.getByIndex(i)
        sName=""
        sName=oShape.Control.Name
        If Left(sName,10)="WMS_STOCK_" Then oSh.DrawPage.remove(oShape)
    Next i
    On Error GoTo 0
End Sub

Function WMSDBST_GetSheet(oDoc As Object,ByRef oSh As Object) As Boolean
    WMSDBST_GetSheet=False
    If oDoc.Sheets.hasByName(WMSDBST_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSDBST_SHEET)
        WMSDBST_GetSheet=True
    End If
End Function

Sub WMSDBST_RefreshDiagnostic()
    Dim oDoc As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim sErr As String,sql As String,n As Long,t As Double
    oDoc=ThisComponent
    On Error GoTo EH

    t=Timer
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail

    sql="SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS"
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then n=oRS.getInt(1)

    MsgBox "Firebird отвечает." & Chr(10) & _
           "Движений: " & CStr(n) & Chr(10) & _
           "Время запроса: " & Format(Timer-t,"0.00") & " сек." & Chr(10) & _
           "Версия: " & WMSDBST_VERSION,64,"WMS — Диагностика остатка"
    Exit Sub

Fail:
    MsgBox sErr,16,"WMS — Диагностика остатка"
    Exit Sub
EH:
    MsgBox "Диагностика: " & CStr(Err) & " " & Error$,16,"WMS — Диагностика остатка"
End Sub

Sub WMSDBST_RepairButtons()
    Dim oDoc As Object,oSh As Object
    oDoc=ThisComponent
    On Error GoTo EH

    If Not WMSDBST_GetSheet(oDoc,oSh) Then
        MsgBox "Лист 'Остаток' не найден.",16,"WMS — Остаток"
        Exit Sub
    End If

    WMSDBST_InstallButtons oDoc,oSh
    MsgBox "Кнопки листа 'Остаток' перепривязаны." & Chr(10) & _
           "Версия: " & WMSDBST_VERSION & Chr(10) & _
           "Кнопка 'Обновить остаток' теперь вызывает макрос через ActionListener.", _
           64,"WMS — Остаток"
    Exit Sub
EH:
    MsgBox "Ремонт кнопок: " & CStr(Err) & " " & Error$,16,"WMS — Остаток"
End Sub

Sub WMSDBST_LotsSelected()
    Dim oDoc As Object,oSh As Object,oSel As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim r As Long,code As String,loc As String,sErr As String,sql As String,s As String,n As Long
    oDoc=ThisComponent
    If Not WMSDBST_GetSheet(oDoc,oSh) Then Exit Sub
    On Error GoTo EH
    oSel=oDoc.CurrentSelection:r=oSel.CellAddress.Row
    If r<5 Then MsgBox "Выберите строку товара.",48,"WMS — Партии":Exit Sub
    code=Trim(oSh.getCellByPosition(0,r).String)
    loc=Trim(oSh.getCellByPosition(3,r).String)
    If code="" Then Exit Sub

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Партии":Exit Sub

    sql="SELECT l.LOT_ID,l.RECEIPT_DATE,l.DOC_QTY,l.DOC_UNIT,l.BASE_QTY,l.BASE_UNIT," & _
        "COALESCE(SUM(m.QTY),0),COALESCE(l.ORIGIN_NAME,'') " & _
        "FROM WMS_STOCK_LOTS l LEFT JOIN WMS_STOCK_MOVEMENTS m ON m.LOT_ID=l.LOT_ID " & _
        "WHERE l.PRODUCT_CODE=" & WMSDB_SQLText(code) & _
        " AND l.LOCATION_NAME=" & WMSDB_SQLText(loc) & _
        " GROUP BY l.LOT_ID,l.RECEIPT_DATE,l.DOC_QTY,l.DOC_UNIT,l.BASE_QTY,l.BASE_UNIT,l.ORIGIN_NAME " & _
        "ORDER BY l.RECEIPT_DATE,l.LOT_ID"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)

    s="Товар: " & code & Chr(10) & "Место: " & loc & Chr(10) & Chr(10)
    Do While oRS.next()
        n=n+1
        If n<=25 Then
            s=s & oRS.getString(1) & " | " & oRS.getString(2) & Chr(10) & _
                "Документ: " & WMSDBST_Num(oRS.getDouble(3)) & " " & oRS.getString(4) & _
                " | Приход в остаток: " & WMSDBST_Num(oRS.getDouble(5)) & " " & oRS.getString(6) & _
                " | Осталось: " & WMSDBST_Num(oRS.getDouble(7)) & " " & oRS.getString(6) & Chr(10) & _
                "Источник: " & oRS.getString(8) & Chr(10) & Chr(10)
        End If
    Loop
    WMSDB_Close
    If n=0 Then s=s & "Партий не найдено."
    If n>25 Then s=s & "... ещё партий: " & CStr(n-25)
    MsgBox s,64,"WMS — Партии товара"
    Exit Sub
EH:
    On Error Resume Next:WMSDB_Close
    MsgBox "Партии товара: " & CStr(Err) & " " & Error$,16,"WMS — Партии"
End Sub

Sub WMSDBST_StockHistorySelected()
    Dim oDoc As Object,oSh As Object,oSel As Object,code As String,r As Long
    oDoc=ThisComponent
    If Not WMSDBST_GetSheet(oDoc,oSh) Then Exit Sub

    On Error GoTo EH
    oSel=oDoc.CurrentSelection
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        MsgBox "Выберите любую ячейку в строке товара.",48,"WMS — История товара"
        Exit Sub
    End If

    r=oSel.CellAddress.Row
    If r<5 Then
        MsgBox "Выберите строку товара в таблице остатка.",48,"WMS — История товара"
        Exit Sub
    End If

    code=Trim(oSh.getCellByPosition(0,r).String)
    If code="" Then
        MsgBox "В выбранной строке нет кода товара.",48,"WMS — История товара"
        Exit Sub
    End If

    If oDoc.Sheets.hasByName("База - Поиск") Then
        oDoc.CurrentController.setActiveSheet(oDoc.Sheets.getByName("База - Поиск"))
        MsgBox "Открыт поиск." & Chr(10) & "Код товара: " & code,64,"WMS — История товара"
    Else
        MsgBox "Лист 'База - Поиск' не найден." & Chr(10) & "Код товара: " & code,48,"WMS — История товара"
    End If
    Exit Sub
EH:
    MsgBox "История товара: " & CStr(Err) & " " & Error$,16,"WMS — История товара"
End Sub

Sub WMSDBST_OpenSearch()
    Dim oDoc As Object
    oDoc=ThisComponent
    If oDoc.Sheets.hasByName("База - Поиск") Then
        oDoc.CurrentController.setActiveSheet(oDoc.Sheets.getByName("База - Поиск"))
    Else
        MsgBox "Лист 'База - Поиск' не найден.",48,"WMS — Поиск"
    End If
End Sub


Function WMSDBST_GetBucketList(oDoc As Object,sProductCode As String,sLocation As String, _
    ByRef sList As String,ByRef aCats As Variant,ByRef aSubs As Variant,ByRef aOrigins As Variant, _
    ByRef aQtys As Variant,ByRef nCount As Long,ByRef sErr As String) As Boolean

    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String
    Dim cats() As String,subs() As String,origins() As String,qtys() As Double
    Dim c As String,sc As String,org As String,q As Double

    WMSDBST_GetBucketList=False:sErr="":sList="":nCount=0
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBST_EnsureSchema(oCon,sErr) Then Exit Function

    sql="SELECT COALESCE(DEST_CATEGORY,''),COALESCE(DEST_SUBCATEGORY,''),COALESCE(ORIGIN_NAME,''),SUM(QTY) " & _
        "FROM WMS_STOCK_MOVEMENTS WHERE PRODUCT_CODE='" & WMSDBST_SQL(sProductCode) & "' " & _
        "AND LOCATION_NAME='" & WMSDBST_SQL(sLocation) & "' " & _
        "GROUP BY DEST_CATEGORY,DEST_SUBCATEGORY,ORIGIN_NAME HAVING SUM(QTY)>0 " & _
        "ORDER BY DEST_CATEGORY,DEST_SUBCATEGORY,ORIGIN_NAME"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)

    Do While oRS.next()
        nCount=nCount+1
        ReDim Preserve cats(1 To nCount)
        ReDim Preserve subs(1 To nCount)
        ReDim Preserve origins(1 To nCount)
        ReDim Preserve qtys(1 To nCount)

        c=oRS.getString(1):sc=oRS.getString(2):org=oRS.getString(3):q=oRS.getDouble(4)
        cats(nCount)=c:subs(nCount)=sc:origins(nCount)=org:qtys(nCount)=q

        If nCount<=30 Then
            sList=sList & CStr(nCount) & ". " & _
                  IIf(c="","Без категории",c) & _
                  IIf(sc="",""," → " & sc) & _
                  IIf(org="",""," | " & org) & _
                  " | доступно " & WMSDBST_Num(q) & Chr(10)
        End If
    Loop

    aCats=cats:aSubs=subs:aOrigins=origins:aQtys=qtys
    WMSDBST_GetBucketList=True
    Exit Function
EH:
    sErr="Получение остатков по партиям: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_GetBucketAvailable(oDoc As Object,sProductCode As String,sLocation As String, _
    sCat As String,sSub As String,sOrigin As String,ByRef dAvailable As Double,ByRef sErr As String) As Boolean

    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String
    WMSDBST_GetBucketAvailable=False:sErr="":dAvailable=0
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSDBST_EnsureSchema(oCon,sErr) Then Exit Function

    sql="SELECT COALESCE(SUM(QTY),0) FROM WMS_STOCK_MOVEMENTS WHERE " & _
        "PRODUCT_CODE='" & WMSDBST_SQL(sProductCode) & "' AND " & _
        "LOCATION_NAME='" & WMSDBST_SQL(sLocation) & "' AND " & _
        "COALESCE(DEST_CATEGORY,'')='" & WMSDBST_SQL(sCat) & "' AND " & _
        "COALESCE(DEST_SUBCATEGORY,'')='" & WMSDBST_SQL(sSub) & "' AND " & _
        "COALESCE(ORIGIN_NAME,'')='" & WMSDBST_SQL(sOrigin) & "'"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then dAvailable=oRS.getDouble(1)

    WMSDBST_GetBucketAvailable=True
    Exit Function
EH:
    sErr="Проверка доступного остатка: " & CStr(Err) & " " & Error$
End Function

Function WMSDBST_PostIssueMovement(oDoc As Object,sMovementID As String,sSourceID As String, _
    sProductCode As String,sProductName As String,dQty As Double,sUnit As String,sLocation As String, _
    sCat As String,sSub As String,sOrigin As String,dMovementDate As Double,sNote As String, _
    ByRef sErr As String) As Boolean

    Dim dAvailable As Double
    WMSDBST_PostIssueMovement=False:sErr=""

    If dQty<=0 Then sErr="Количество выдачи должно быть больше 0.":Exit Function

    If Not WMSDBST_GetBucketAvailable(oDoc,sProductCode,sLocation,sCat,sSub,sOrigin,dAvailable,sErr) Then Exit Function

    If dAvailable+0.0000001<dQty Then
        sErr="Недостаточный остаток. Доступно " & WMSDBST_Num(dAvailable) & " " & sUnit & _
             ", требуется " & WMSDBST_Num(dQty) & " " & sUnit & "."
        Exit Function
    End If

    If Not WMSDBST_PostMovementClassified(oDoc,sMovementID,sSourceID,"ISSUE","OUT", _
        sProductCode,sProductName,-dQty,sUnit,sLocation,sOrigin,sCat,sSub,dMovementDate,sNote,sErr) Then Exit Function

    WMSDBST_PostIssueMovement=True
End Function

Function WMSDBST_PostReturnMovement(oDoc As Object,sMovementID As String,sSourceID As String, _
    sProductCode As String,sProductName As String,dQty As Double,sUnit As String,sLocation As String, _
    sCat As String,sSub As String,sOrigin As String,dMovementDate As Double,sNote As String, _
    ByRef sErr As String) As Boolean

    WMSDBST_PostReturnMovement=False:sErr=""
    If dQty<=0 Then sErr="Количество возврата должно быть больше 0.":Exit Function

    If Not WMSDBST_PostMovementClassified(oDoc,sMovementID,sSourceID,"RETURN","IN", _
        sProductCode,sProductName,dQty,sUnit,sLocation,sOrigin,sCat,sSub,dMovementDate,sNote,sErr) Then Exit Function

    WMSDBST_PostReturnMovement=True
End Function

Function WMSDBST_Num(v As Double) As String
    WMSDBST_Num=Replace(CStr(v),",",".")
End Function

Sub WMSDBST_SelfCheck()
    Dim oDoc As Object,oCon As Object,sErr As String,sReport As String
    Dim p As Boolean,m As Boolean
    oDoc=ThisComponent
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail

    Dim c1 As Boolean,c2 As Boolean,c3 As Boolean
    p=WMSDB_TableExistsSimple(oCon,"WMS_PRODUCTS",sErr)
    If sErr<>"" Then GoTo Fail
    m=WMSDB_TableExistsSimple(oCon,"WMS_STOCK_MOVEMENTS",sErr)
    If sErr<>"" Then GoTo Fail
    c1=WMSDBST_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","ORIGIN_NAME",sErr)
    If sErr<>"" Then GoTo Fail
    c2=WMSDBST_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","DEST_CATEGORY",sErr)
    If sErr<>"" Then GoTo Fail
    c3=WMSDBST_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","DEST_SUBCATEGORY",sErr)
    If sErr<>"" Then GoTo Fail

    sReport="WMS_10 SelfCheck" & Chr(10) & _
            "WMS_PRODUCTS: " & WMSDBST_YesNo(p) & Chr(10) & _
            "WMS_STOCK_MOVEMENTS: " & WMSDBST_YesNo(m) & Chr(10) & _
            "Источник: " & WMSDBST_YesNo(c1) & Chr(10) & _
            "Категория назначения: " & WMSDBST_YesNo(c2) & Chr(10) & _
            "Подкатегория назначения: " & WMSDBST_YesNo(c3) & Chr(10) & _
            "Версия: " & WMSDBST_VERSION

    If p And m And c1 And c2 And c3 Then
        sReport=sReport & Chr(10) & "Результат: OK"
        MsgBox sReport,64,"WMS — Движения + Остатки"
    Else
        sReport=sReport & Chr(10) & "Результат: НЕ ГОТОВО"
        MsgBox sReport,48,"WMS — Движения + Остатки"
    End If
    WMSDB_Close
    Exit Sub
Fail:
    WMSDB_Close
    MsgBox sErr,16,"WMS — Движения + Остатки"
    Exit Sub
EH:
    sErr="SelfCheck WMS_10: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSDBST_YesNo(v As Boolean) As String
    If v Then
        WMSDBST_YesNo="есть"
    Else
        WMSDBST_YesNo="НЕТ"
    End If
End Function

Function WMSDBST_SQL(s As String) As String
    WMSDBST_SQL = Replace(CStr(s), "'", "''")
End Function

Function WMSDBST_SQLNumber(d As Double) As String
    Dim s As String
    s=CStr(d)
    s=Replace(s,",",".")
    WMSDBST_SQLNumber=s
End Function

Function WMSDBST_SQLDate(d As Double) As String
    Dim y As Integer,m As Integer,dd As Integer
    If d<=0 Then
        WMSDBST_SQLDate=Format(Date,"YYYY-MM-DD")
        Exit Function
    End If
    y=Year(CDate(d)):m=Month(CDate(d)):dd=Day(CDate(d))
    WMSDBST_SQLDate=Right("0000"&CStr(y),4)&"-"&Right("00"&CStr(m),2)&"-"&Right("00"&CStr(dd),2)
End Function

