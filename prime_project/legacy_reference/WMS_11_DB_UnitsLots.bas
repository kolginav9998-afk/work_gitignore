Option Explicit

' ============================================================
' WMS_11_DB_UnitsLots
' Универсальные единицы измерения + партии
' Версия 0.1.0 FOUNDATION
'
' Зависимости:
'   Standard.WMS_04_DB_Connection
'   существующая WMS_PRODUCTS / WMS_STOCK_MOVEMENTS (WMS_10)
'
' ВАЖНО:
' - миграция только ДОБАВЛЯЕТ таблицы/поля;
' - существующие движения и остатки не переписываются;
' - WMS_STOCK_MOVEMENTS.QTY + UNIT_NAME остаются базовым
'   количеством/единицей для обратной совместимости;
' - LOT_ID у старых движений может быть NULL.
' ============================================================

Global Const WMSDBUL_VERSION = "0.1.0-FOUNDATION"
Global Const WMSDBUL_SCHEMA = "1"

Sub WMSDBUL_Install()
    Dim oDoc As Object, oCon As Object, sErr As String, sReport As String
    oDoc = ThisComponent
    On Error GoTo EH

    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then GoTo Fail

    If Not WMSDBUL_EnsureSchema(oCon, sErr) Then GoTo Fail
    If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail

    If Not WMSDBUL_SelfCheckCore(oCon, sReport) Then
        sErr = "Схема создана, но SelfCheck не пройден:" & Chr(10) & sReport
        GoTo Fail
    End If

    MsgBox "Units & Lots Foundation установлен." & Chr(10) & _
           "Данные существующего склада не изменялись." & Chr(10) & _
           "Версия: " & WMSDBUL_VERSION & Chr(10) & _
           "Схема Units/Lots: " & WMSDBUL_SCHEMA, 64, "WMS — Единицы и партии"
    Exit Sub

Fail:
    MsgBox sErr, 16, "WMS — Единицы и партии"
    Exit Sub
EH:
    sErr = "Установка Units/Lots: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSDBUL_EnsureSchema(oCon As Object, ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSDBUL_EnsureSchema = False
    sErr = ""
    On Error GoTo EH

    oStmt = oCon.createStatement()

    ' --------------------------------------------------------
    ' Партия прихода.
    ' DOC_* = как в первичном документе.
    ' BASE_* = физическое складское количество.
    ' --------------------------------------------------------
    If Not WMSDBUL_TableExists(oCon, "WMS_STOCK_LOTS", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate( _
            "CREATE TABLE WMS_STOCK_LOTS (" & _
            "LOT_ID VARCHAR(100) NOT NULL PRIMARY KEY," & _
            "PRODUCT_CODE VARCHAR(80) NOT NULL," & _
            "SOURCE_ID VARCHAR(100)," & _
            "SOURCE_TYPE VARCHAR(40)," & _
            "DOC_QTY DECIMAL(18,4)," & _
            "DOC_UNIT VARCHAR(40)," & _
            "BASE_QTY DECIMAL(18,4) NOT NULL," & _
            "BASE_UNIT VARCHAR(40) NOT NULL," & _
            "LOCATION_NAME VARCHAR(120) NOT NULL," & _
            "ORIGIN_NAME VARCHAR(120)," & _
            "DEST_CATEGORY VARCHAR(120)," & _
            "DEST_SUBCATEGORY VARCHAR(120)," & _
            "RECEIPT_DATE DATE NOT NULL," & _
            "NOTE_TEXT VARCHAR(500)," & _
            "ACTIVE_FLAG SMALLINT DEFAULT 1 NOT NULL," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "CONSTRAINT FK_WMS_LOT_PRODUCT FOREIGN KEY (PRODUCT_CODE) " & _
            "REFERENCES WMS_PRODUCTS(PRODUCT_CODE))")
    End If

    ' --------------------------------------------------------
    ' Любое число альтернативных единиц для конкретной партии.
    ' FACTOR_TO_BASE означает:
    '   1 UNIT_NAME = FACTOR_TO_BASE * BASE_UNIT
    '
    ' Примеры:
    '   партия винтов, BASE=шт:
    '       шт   -> 1
    '       упак -> 1000
    '
    '   партия металла, BASE=кг:
    '       кг -> 1
    '       шт -> 10
    ' --------------------------------------------------------
    If Not WMSDBUL_TableExists(oCon, "WMS_LOT_UNITS", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate( _
            "CREATE TABLE WMS_LOT_UNITS (" & _
            "LOT_ID VARCHAR(100) NOT NULL," & _
            "UNIT_NAME VARCHAR(40) NOT NULL," & _
            "FACTOR_TO_BASE DECIMAL(18,6) NOT NULL," & _
            "UNIT_ROLE VARCHAR(20)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "CONSTRAINT FK_WMS_LOTUNIT_LOT FOREIGN KEY (LOT_ID) " & _
            "REFERENCES WMS_STOCK_LOTS(LOT_ID))")
    End If

    If Not WMSDBUL_IndexExists(oCon, "UX_WMS_LOT_UNIT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE UNIQUE INDEX UX_WMS_LOT_UNIT ON WMS_LOT_UNITS(LOT_ID,UNIT_NAME)")
    End If

    If Not WMSDBUL_IndexExists(oCon, "IX_WMS_LOT_PRODUCT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE INDEX IX_WMS_LOT_PRODUCT ON WMS_STOCK_LOTS(PRODUCT_CODE)")
    End If

    If Not WMSDBUL_IndexExists(oCon, "IX_WMS_LOT_SOURCE", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE INDEX IX_WMS_LOT_SOURCE ON WMS_STOCK_LOTS(SOURCE_TYPE,SOURCE_ID)")
    End If

    ' --------------------------------------------------------
    ' Расширение существующего журнала движений.
    ' QTY/UNIT_NAME НЕ удаляются и НЕ меняются.
    ' Для новых движений:
    '   QTY/UNIT_NAME = подписанное базовое движение,
    '   ENTRY_QTY/ENTRY_UNIT = что реально ввёл кладовщик,
    '   LOT_ID = партия.
    ' --------------------------------------------------------
    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "LOT_ID", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD LOT_ID VARCHAR(100)")
    End If

    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "ENTRY_QTY", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD ENTRY_QTY DECIMAL(18,4)")
    End If

    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "ENTRY_UNIT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD ENTRY_UNIT VARCHAR(40)")
    End If

    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "BASE_QTY", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD BASE_QTY DECIMAL(18,4)")
    End If

    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "BASE_UNIT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD BASE_UNIT VARCHAR(40)")
    End If

    If Not WMSDBUL_IndexExists(oCon, "IX_WMS_MOV_LOT", sErr) Then
        If sErr <> "" Then Exit Function
        oStmt.executeUpdate("CREATE INDEX IX_WMS_MOV_LOT ON WMS_STOCK_MOVEMENTS(LOT_ID)")
    End If

    WMSDBUL_PutMeta oCon, "UNITS_LOTS_SCHEMA", WMSDBUL_SCHEMA
    WMSDBUL_PutMeta oCon, "UNITS_LOTS_VERSION", WMSDBUL_VERSION

    oCon.commit()
    WMSDBUL_EnsureSchema = True
    Exit Function
EH:
    sErr = "Миграция Units/Lots: " & CStr(Err) & " " & Error$
    On Error Resume Next
    oCon.rollback()
End Function

Sub WMSDBUL_SelfCheck()
    Dim oDoc As Object, oCon As Object, sErr As String, sReport As String
    oDoc = ThisComponent
    On Error GoTo EH

    oCon = WMSDB_GetConnectionEx(oDoc, sErr)
    If sErr <> "" Then GoTo Fail

    If WMSDBUL_SelfCheckCore(oCon, sReport) Then
        MsgBox "SelfCheck Units/Lots: OK" & Chr(10) & _
               sReport, 64, "WMS — Единицы и партии"
    Else
        MsgBox "SelfCheck Units/Lots: ОШИБКА" & Chr(10) & _
               sReport, 16, "WMS — Единицы и партии"
    End If
    Exit Sub

Fail:
    MsgBox sErr, 16, "WMS — Единицы и партии"
    Exit Sub
EH:
    MsgBox "SelfCheck: " & CStr(Err) & " " & Error$, 16, "WMS — Единицы и партии"
End Sub

Function WMSDBUL_SelfCheckCore(oCon As Object, ByRef sReport As String) As Boolean
    Dim sErr As String, sMeta As String
    WMSDBUL_SelfCheckCore = False
    sReport = ""

    If Not WMSDBUL_TableExists(oCon, "WMS_STOCK_LOTS", sErr) Then
        If sErr = "" Then sErr = "WMS_STOCK_LOTS отсутствует."
        sReport = sErr
        Exit Function
    End If

    If Not WMSDBUL_TableExists(oCon, "WMS_LOT_UNITS", sErr) Then
        If sErr = "" Then sErr = "WMS_LOT_UNITS отсутствует."
        sReport = sErr
        Exit Function
    End If

    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "LOT_ID", sErr) Then GoTo Missing
    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "ENTRY_QTY", sErr) Then GoTo Missing
    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "ENTRY_UNIT", sErr) Then GoTo Missing
    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "BASE_QTY", sErr) Then GoTo Missing
    If Not WMSDBUL_ColumnExists(oCon, "WMS_STOCK_MOVEMENTS", "BASE_UNIT", sErr) Then GoTo Missing

    sMeta = WMSDB_ScalarString(oCon, _
        "SELECT META_VALUE FROM WMS_META WHERE META_KEY='UNITS_LOTS_SCHEMA'", sErr)
    If sErr <> "" Then
        sReport = sErr
        Exit Function
    End If
    If sMeta <> WMSDBUL_SCHEMA Then
        sReport = "UNITS_LOTS_SCHEMA=" & sMeta & ", ожидалось " & WMSDBUL_SCHEMA
        Exit Function
    End If

    sReport = "WMS_STOCK_LOTS: OK" & Chr(10) & _
              "WMS_LOT_UNITS: OK" & Chr(10) & _
              "Поля движений Units/Lots: OK" & Chr(10) & _
              "Существующие движения не изменялись." & Chr(10) & _
              "Версия: " & WMSDBUL_VERSION
    WMSDBUL_SelfCheckCore = True
    Exit Function

Missing:
    If sErr = "" Then sErr = "Не найдено обязательное поле WMS_STOCK_MOVEMENTS."
    sReport = sErr
End Function

' ============================================================
' API: создание партии. Пока не подключается автоматически
' к Заказам — это FOUNDATION, чтобы сначала проверить схему.
' ============================================================

Function WMSDBUL_CreateLotCon(oCon As Object, sLotID As String, sProductCode As String, _
    sSourceID As String, sSourceType As String, dDocQty As Double, sDocUnit As String, _
    dBaseQty As Double, sBaseUnit As String, sLocation As String, sOrigin As String, _
    sDestCategory As String, sDestSubcategory As String, dReceiptDate As Double, _
    sNote As String, ByRef sErr As String) As Boolean

    Dim oStmt As Object, sql As String
    WMSDBUL_CreateLotCon = False
    sErr = ""

    If Trim(sLotID) = "" Then sErr = "LOT_ID пустой.": Exit Function
    If Trim(sProductCode) = "" Then sErr = "Код товара пустой.": Exit Function
    If dBaseQty <= 0 Then sErr = "Базовое количество партии должно быть > 0.": Exit Function
    If Trim(sBaseUnit) = "" Then sErr = "Базовая единица пустая.": Exit Function
    If Trim(sLocation) = "" Then sErr = "Место хранения пустое.": Exit Function

    On Error GoTo EH

    If WMSDBUL_LotExists(oCon, sLotID, sErr) Then
        If sErr <> "" Then Exit Function
        WMSDBUL_CreateLotCon = True
        Exit Function
    End If

    sql = "INSERT INTO WMS_STOCK_LOTS " & _
          "(LOT_ID,PRODUCT_CODE,SOURCE_ID,SOURCE_TYPE,DOC_QTY,DOC_UNIT,BASE_QTY,BASE_UNIT," & _
          "LOCATION_NAME,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,RECEIPT_DATE,NOTE_TEXT) VALUES (" & _
          WMSDB_SQLText(sLotID) & "," & _
          WMSDB_SQLText(sProductCode) & "," & _
          WMSDB_SQLText(sSourceID) & "," & _
          WMSDB_SQLText(sSourceType) & "," & _
          WMSDBUL_SQLNumber(dDocQty) & "," & _
          WMSDB_SQLText(sDocUnit) & "," & _
          WMSDBUL_SQLNumber(dBaseQty) & "," & _
          WMSDB_SQLText(sBaseUnit) & "," & _
          WMSDB_SQLText(sLocation) & "," & _
          WMSDB_SQLText(sOrigin) & "," & _
          WMSDB_SQLText(sDestCategory) & "," & _
          WMSDB_SQLText(sDestSubcategory) & "," & _
          WMSDB_SQLText(WMSDBUL_SQLDate(dReceiptDate)) & "," & _
          WMSDB_SQLText(sNote) & ")"

    oStmt = oCon.createStatement()
    oStmt.executeUpdate(sql)

    ' Базовая единица всегда имеет коэффициент 1.
    If Not WMSDBUL_SetLotUnitCon(oCon, sLotID, sBaseUnit, 1, "BASE", sErr) Then Exit Function

    ' Если документальная единица отличается, сохраняем её коэффициент позже
    ' отдельным вызовом SetLotUnitCon, когда пользователь подтвердит соотношение.

    WMSDBUL_CreateLotCon = True
    Exit Function
EH:
    sErr = "Создание партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBUL_SetLotUnitCon(oCon As Object, sLotID As String, sUnit As String, _
    dFactorToBase As Double, sRole As String, ByRef sErr As String) As Boolean

    Dim oStmt As Object, sql As String
    WMSDBUL_SetLotUnitCon = False
    sErr = ""

    If Trim(sLotID) = "" Then sErr = "LOT_ID пустой.": Exit Function
    If Trim(sUnit) = "" Then sErr = "Единица измерения пустая.": Exit Function
    If dFactorToBase <= 0 Then sErr = "Коэффициент должен быть > 0.": Exit Function

    On Error GoTo EH
    sql = "UPDATE OR INSERT INTO WMS_LOT_UNITS " & _
          "(LOT_ID,UNIT_NAME,FACTOR_TO_BASE,UNIT_ROLE) VALUES (" & _
          WMSDB_SQLText(sLotID) & "," & _
          WMSDB_SQLText(sUnit) & "," & _
          WMSDBUL_SQLNumber(dFactorToBase) & "," & _
          WMSDB_SQLText(sRole) & ") MATCHING (LOT_ID,UNIT_NAME)"
    oStmt = oCon.createStatement()
    oStmt.executeUpdate(sql)
    WMSDBUL_SetLotUnitCon = True
    Exit Function
EH:
    sErr = "Единица партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBUL_ConvertToBaseCon(oCon As Object, sLotID As String, _
    dEntryQty As Double, sEntryUnit As String, ByRef dBaseQty As Double, _
    ByRef sBaseUnit As String, ByRef sErr As String) As Boolean

    Dim oStmt As Object, oRS As Object, sql As String, dFactor As Double
    WMSDBUL_ConvertToBaseCon = False
    sErr = ""
    dBaseQty = 0
    sBaseUnit = ""

    If dEntryQty <= 0 Then sErr = "Количество должно быть > 0.": Exit Function
    If Trim(sEntryUnit) = "" Then sErr = "Единица выдачи пустая.": Exit Function

    On Error GoTo EH

    sql = "SELECT l.BASE_UNIT,u.FACTOR_TO_BASE " & _
          "FROM WMS_STOCK_LOTS l JOIN WMS_LOT_UNITS u ON u.LOT_ID=l.LOT_ID " & _
          "WHERE l.LOT_ID=" & WMSDB_SQLText(sLotID) & _
          " AND u.UNIT_NAME=" & WMSDB_SQLText(sEntryUnit)
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)

    If Not oRS.next() Then
        sErr = "Для партии " & sLotID & " не задан пересчёт единицы '" & sEntryUnit & "'."
        Exit Function
    End If

    sBaseUnit = oRS.getString(1)
    dFactor = oRS.getDouble(2)
    dBaseQty = dEntryQty * dFactor

    If dBaseQty <= 0 Then
        sErr = "Некорректный результат пересчёта."
        Exit Function
    End If

    WMSDBUL_ConvertToBaseCon = True
    Exit Function
EH:
    sErr = "Пересчёт единицы: " & CStr(Err) & " " & Error$
End Function

Function WMSDBUL_LotExists(oCon As Object, sLotID As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object
    WMSDBUL_LotExists = False
    sErr = ""
    On Error GoTo EH
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT COUNT(*) FROM WMS_STOCK_LOTS WHERE LOT_ID=" & WMSDB_SQLText(sLotID))
    If oRS.next() Then WMSDBUL_LotExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Проверка партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBUL_TableExists(oCon As Object, sTable As String, ByRef sErr As String) As Boolean
    WMSDBUL_TableExists = WMSDB_TableExistsSimple(oCon, sTable, sErr)
End Function

Function WMSDBUL_ColumnExists(oCon As Object, sTable As String, sColumn As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sql As String
    WMSDBUL_ColumnExists = False
    sErr = ""
    On Error GoTo EH
    sql = "SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE " & _
          "TRIM(RDB$RELATION_NAME)=" & WMSDB_SQLText(UCase(Trim(sTable))) & _
          " AND TRIM(RDB$FIELD_NAME)=" & WMSDB_SQLText(UCase(Trim(sColumn)))
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBUL_ColumnExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Проверка поля " & sTable & "." & sColumn & ": " & CStr(Err) & " " & Error$
End Function

Function WMSDBUL_IndexExists(oCon As Object, sIndex As String, ByRef sErr As String) As Boolean
    Dim oStmt As Object, oRS As Object, sql As String
    WMSDBUL_IndexExists = False
    sErr = ""
    On Error GoTo EH
    sql = "SELECT COUNT(*) FROM RDB$INDICES WHERE TRIM(RDB$INDEX_NAME)=" & _
          WMSDB_SQLText(UCase(Trim(sIndex)))
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBUL_IndexExists = (oRS.getInt(1) > 0)
    Exit Function
EH:
    sErr = "Проверка индекса " & sIndex & ": " & CStr(Err) & " " & Error$
End Function

Sub WMSDBUL_PutMeta(oCon As Object, sKey As String, sValue As String)
    Dim oStmt As Object, n As Long
    On Error GoTo EH
    oStmt = oCon.createStatement()
    n = oStmt.executeUpdate("UPDATE WMS_META SET META_VALUE=" & WMSDB_SQLText(sValue) & _
        ",UPDATED_AT=CURRENT_TIMESTAMP WHERE META_KEY=" & WMSDB_SQLText(sKey))
    If n = 0 Then
        oStmt.executeUpdate("INSERT INTO WMS_META (META_KEY,META_VALUE,UPDATED_AT) VALUES (" & _
            WMSDB_SQLText(sKey) & "," & WMSDB_SQLText(sValue) & ",CURRENT_TIMESTAMP)")
    End If
    Exit Sub
EH:
    Err = 0
End Sub

Function WMSDBUL_SQLNumber(d As Double) As String
    Dim s As String
    s = CStr(d)
    s = Replace(s, ",", ".")
    WMSDBUL_SQLNumber = s
End Function

Function WMSDBUL_SQLDate(d As Double) As String
    Dim dt As Date
    dt = CDate(d)
    WMSDBUL_SQLDate = Format(Year(dt), "0000") & "-" & _
                      Format(Month(dt), "00") & "-" & _
                      Format(Day(dt), "00")
End Function

