Option Explicit

' PRIME_03_Catalog
' Товары, алиасы поставщиков, единицы измерения, генерация ЕИ-кодов, живой lookup по коду.
' database_connection_allowed=false: весь lookup идёт по memory index (Collection), построенному
' из DB_PRIME_PRODUCTS/DB_PRIME_ALIASES/DB_PRIME_PRODUCT_UNITS, а не поячейковым сканированием.

Private gProductIndex As Object   ' Collection: PRODUCT_CODE -> Variant(строка таблицы)
Private gProductIndexBuilt As Boolean

Public Sub PRIME_BuildProductIndex()
    Set gProductIndex = New Collection
    If Not PRIME_SheetExists(SH_DB_PRODUCTS) Then
        gProductIndexBuilt = True
        Exit Sub
    End If
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_PRODUCTS)
    If UBound(table) >= 1 Then
        Dim headers As Variant
        headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
        Dim colCode As Long
        colCode = PRIME_ColIndex(headers, "PRODUCT_CODE")
        Dim colOpId As Long
        colOpId = PRIME_ColIndex(headers, "OP_ID")
        Dim i As Long
        For i = 1 To UBound(table)
            Dim code As String
            code = CStr(table(i)(colCode))
            Dim visible As Boolean
            visible = True
            If colOpId >= 0 Then
                Dim opId As String
                opId = CStr(table(i)(colOpId))
                ' Пусто = легаси/миграция/прямое создание - видим всегда. Непусто = товар создан
                ' внутри проведения документа - видим, только если его OP_ID уже COMMITTED (иначе
                ' это "призрачный" товар из ещё не завершённой или упавшей транзакции).
                If opId <> "" Then visible = PRIME_IsOpIdCommitted(opId)
            End If
            If code <> "" And visible And Not PRIME_CollectionHasKey(gProductIndex, code) Then
                ' headless_repeat_array_cache_read_crash (2.1.2, см. PRIME_02_Store.PRIME_HeaderMap) -
                ' тот же класс дефекта: массив как элемент Collection падает при повторном .Item()
                ' в одной цепочке вызовов. PRIME_Orders_AutofillByCode читает 5 разных полей ОДНОГО
                ' и того же товара подряд (PRIME_GetProductField x5 -> PRIME_GetProduct x5) - без
                ' этого обхода 2-5-е чтения детерминированно падали бы. Единственный потребитель
                ' (PRIME_GetProductField) и так всегда делает CStr(rec(col)), потери типов нет.
                gProductIndex.Add(PRIME_JoinRowAsTabString(table(i)), code)
            End If
        Next i
    End If
    gProductIndexBuilt = True
End Sub

' См. комментарий headless_repeat_array_cache_read_crash в PRIME_BuildProductIndex выше -
' Chr(9) выбран как разделитель, потому что ни одно реальное значение колонки товара (коды,
' наименования, категории, числа/даты как текст) не содержит символ табуляции.
Private Function PRIME_JoinRowAsTabString(ByVal row As Variant) As String
    Dim s As String
    s = ""
    Dim i As Long
    For i = LBound(row) To UBound(row)
        If i > LBound(row) Then s = s & Chr(9)
        s = s & CStr(row(i))
    Next i
    PRIME_JoinRowAsTabString = s
End Function

Public Sub PRIME_InvalidateProductIndex()
    gProductIndexBuilt = False
    Set gProductIndex = Nothing
End Sub

Private Sub PRIME_EnsureProductIndex()
    If Not gProductIndexBuilt Then PRIME_BuildProductIndex()
End Sub

Public Function PRIME_ProductExists(ByVal productCode As String) As Boolean
    PRIME_EnsureProductIndex()
    PRIME_ProductExists = PRIME_CollectionHasKey(gProductIndex, productCode)
End Function

' Возвращает строку товара (Variant-массив по колонкам DB_PRIME_PRODUCTS) или Empty, если не найден.
Public Function PRIME_GetProduct(ByVal productCode As String) As Variant
    PRIME_EnsureProductIndex()
    Dim rec As Variant
    If PRIME_CollectionTryGet(gProductIndex, productCode, rec) Then
        PRIME_GetProduct = Split(CStr(rec), Chr(9))
    Else
        PRIME_GetProduct = Empty
    End If
End Function

Public Function PRIME_GetProductField(ByVal productCode As String, ByVal fieldName As String) As String
    Dim rec As Variant
    rec = PRIME_GetProduct(productCode)
    If IsEmpty(rec) Then
        PRIME_GetProductField = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim col As Long
    col = PRIME_ColIndex(headers, fieldName)
    If col = -1 Then
        PRIME_GetProductField = ""
    Else
        PRIME_GetProductField = CStr(rec(col))
    End If
End Function

' Новый код ЕИ-00000001... Вызывать ТОЛЬКО когда у товара нет внутреннего кода -
' forbidden: "creating a new product code when existing internal code is supplied".
Public Function PRIME_NextProductCode() As String
    Dim n As Long
    n = PRIME_SequenceNext("PRODUCT_CODE")
    PRIME_NextProductCode = PRIME_FormatProductCode(n)
End Function

Public Function PRIME_FormatProductCode(ByVal n As Long) As String
    Dim s As String
    s = CStr(n)
    Do While Len(s) < PRODUCT_CODE_DIGITS
        s = "0" & s
    Loop
    PRIME_FormatProductCode = PRODUCT_CODE_PREFIX & s
End Function

' Создаёт новую карточку товара. name_is_unique_key=false: совпадение названий НЕ объединяет
' товары автоматически - это осознанное решение вызывающего кода, не эвристика каталога.
Public Function PRIME_CreateProduct(ByVal productName As String, ByVal baseUnit As String, _
        ByVal location As String, ByVal category As String, ByVal subcategory As String, _
        ByVal returnable As Boolean, ByVal accountType As String) As String
    Dim code As String
    code = PRIME_NextProductCode()

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim row(UBound(headers)) As Variant
    Dim now As String
    now = Format(Now, "YYYY-MM-DD HH:MM:SS")

    row(PRIME_ColIndex(headers, "PRODUCT_CODE")) = code
    row(PRIME_ColIndex(headers, "PRODUCT_NAME")) = productName
    row(PRIME_ColIndex(headers, "BASE_UNIT")) = baseUnit
    row(PRIME_ColIndex(headers, "DEFAULT_LOCATION")) = location
    row(PRIME_ColIndex(headers, "CATEGORY")) = category
    row(PRIME_ColIndex(headers, "SUBCATEGORY")) = subcategory
    row(PRIME_ColIndex(headers, "RETURNABLE")) = IIf(returnable, 1, 0)
    row(PRIME_ColIndex(headers, "ACCOUNT_TYPE")) = accountType
    row(PRIME_ColIndex(headers, "ACTIVE")) = 1
    row(PRIME_ColIndex(headers, "CREATED_AT")) = now
    row(PRIME_ColIndex(headers, "UPDATED_AT")) = now

    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_DB_PRODUCTS, rows)
    PRIME_InvalidateProductIndex()

    PRIME_CreateProduct = code
End Function

' --- Единицы измерения / конвертация (product_units) ---
' Возвращает Empty, если для товара/единицы нет фактора - вызывающий обязан заблокировать
' проведение целиком (missing_conversion_behavior = "Block entire posting").
Public Function PRIME_GetUnitFactor(ByVal productCode As String, ByVal unitName As String) As Variant
    If Not PRIME_SheetExists(SH_DB_PRODUCT_UNITS) Then
        PRIME_GetUnitFactor = Empty
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCT_UNITS)
    Dim colCode As Long, colUnit As Long, colFactor As Long, colActive As Long
    colCode = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colUnit = PRIME_ColIndex(headers, "UNIT_NAME")
    colFactor = PRIME_ColIndex(headers, "FACTOR_TO_BASE")
    colActive = PRIME_ColIndex(headers, "ACTIVE")

    ' Базовая единица товара всегда имеет фактор 1, даже без явной строки в таблице.
    If LCase(PRIME_GetProductField(productCode, "BASE_UNIT")) = LCase(unitName) Then
        PRIME_GetUnitFactor = 1#
        Exit Function
    End If

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_PRODUCT_UNITS)
    Dim i As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(colCode)) = productCode And LCase(CStr(table(i)(colUnit))) = LCase(unitName) _
                And CLng(table(i)(colActive)) <> 0 Then
            PRIME_GetUnitFactor = CDbl(table(i)(colFactor))
            Exit Function
        End If
    Next i
    PRIME_GetUnitFactor = Empty
End Function

Public Function PRIME_ConvertQtyToBase(ByVal productCode As String, ByVal unitName As String, ByVal qty As Double) As Variant
    Dim factor As Variant
    factor = PRIME_GetUnitFactor(productCode, unitName)
    If IsEmpty(factor) Then
        PRIME_ConvertQtyToBase = Empty
        Exit Function
    End If
    Dim result As Double
    result = qty * CDbl(factor)
    PRIME_ConvertQtyToBase = PRIME_RoundQty(result)
End Function

Public Function PRIME_RoundQty(ByVal qty As Double) As Double
    Dim mul As Double
    mul = 10 ^ PRIME_QTY_PRECISION_DIGITS
    PRIME_RoundQty = Int(qty * mul + 0.5) / mul
End Function

' --- Алиасы поставщиков (product_aliases) ---
' ambiguous_match_behavior = "Не угадывать. Требовать выбор пользователя." - при >1 совпадении
' возвращаем специальный маркер, вызывающий обязан показать выбор, а не брать первый попавшийся.
Public Const PRIME_ALIAS_AMBIGUOUS As String = "#AMBIGUOUS#"
Public Const PRIME_ALIAS_NOT_FOUND As String = ""

Public Function PRIME_FindProductByAlias(ByVal platform As String, ByVal seller As String, _
        ByVal supplierCode As String, ByVal supplierArticle As String) As String
    If Not PRIME_SheetExists(SH_DB_ALIASES) Then
        PRIME_FindProductByAlias = PRIME_ALIAS_NOT_FOUND
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_ALIASES)
    Dim colPlatform As Long, colSeller As Long, colSupCode As Long, colSupArt As Long, colProduct As Long, colActive As Long
    colPlatform = PRIME_ColIndex(headers, "PLATFORM")
    colSeller = PRIME_ColIndex(headers, "SELLER")
    colSupCode = PRIME_ColIndex(headers, "SUPPLIER_CODE")
    colSupArt = PRIME_ColIndex(headers, "SUPPLIER_ARTICLE")
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colActive = PRIME_ColIndex(headers, "ACTIVE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_ALIASES)
    Dim matches As New Collection
    Dim i As Long
    For i = 1 To UBound(table)
        If CLng(table(i)(colActive)) <> 0 _
                And PRIME_SameOrBlank(CStr(table(i)(colPlatform)), platform) _
                And PRIME_SameOrBlank(CStr(table(i)(colSeller)), seller) _
                And PRIME_SameOrBlank(CStr(table(i)(colSupCode)), supplierCode) _
                And PRIME_SameOrBlank(CStr(table(i)(colSupArt)), supplierArticle) Then
            matches.Add(CStr(table(i)(colProduct)))
        End If
    Next i

    If matches.Count = 0 Then
        PRIME_FindProductByAlias = PRIME_ALIAS_NOT_FOUND
    ElseIf matches.Count = 1 Then
        PRIME_FindProductByAlias = matches.Item(1)
    Else
        PRIME_FindProductByAlias = PRIME_ALIAS_AMBIGUOUS
    End If
End Function

Private Function PRIME_SameOrBlank(ByVal storedValue As String, ByVal queryValue As String) As Boolean
    If queryValue = "" Then
        PRIME_SameOrBlank = True
    Else
        PRIME_SameOrBlank = (LCase(storedValue) = LCase(queryValue))
    End If
End Function

' --- Местоположение с положительным остатком (для автоподстановки "Откуда" в Выдачах) ---
' location_rule: подставить, только если положительный остаток ровно в одном месте.
' Возвращает "" если мест 0 или >1 (multiple_locations => оставить пустым и показать подсказку).
' single_physical_warehouse (FINAL): раньше здесь был жёстко зашит SC_GENERAL - "Выдачи"
' работали только с позициями, пришедшими через "Заказы", и не находили остаток по EI,
' пришедшему через "Приход — Офис/Производство/Детали" (сам EI существовал, но автоподстановка
' молчала "остатка нет"). Один физический склад - без фильтра по источнику вообще.
Public Function PRIME_SingleLocationWithStock(ByVal productCode As String) As String
    Dim locations() As String
    Dim quantities() As Double
    Dim contours() As String
    PRIME_StockByLocation(productCode, "", locations, quantities, contours)

    Dim result As String
    result = ""
    Dim countPositive As Long
    countPositive = 0

    Dim i As Long
    If UBound(locations) >= LBound(locations) Then
        For i = LBound(locations) To UBound(locations)
            If quantities(i) > 0 Then
                countPositive = countPositive + 1
                result = locations(i)
            End If
        Next i
    End If

    If countPositive = 1 Then
        PRIME_SingleLocationWithStock = result
    Else
        PRIME_SingleLocationWithStock = ""
    End If
End Function

' Остаток товара по месту хранения, посчитанный из COMMITTED-движений (DB_PRIME_MOVEMENTS).
' Заполняет параллельные массивы locations()/quantities()/contours() - Collection в StarBasic не
' отдаёт свои ключи обратно, поэтому агрегация ведётся через явные массивы, а не через Collection.
' FINAL mega-task (single_physical_warehouse): группировка была по (место, контур) - тот же класс
' бага, что и в PRIME_09_StockSearch.PRIME_Stock_Rebuild (см. подробное объяснение там) - две
' строки с одним и тем же местом, но разным STOCK_CONTOUR (например, обычное поступление и
' перемещение с пустым контуром), задваивались в остатке по месту вместо суммирования в одну
' позицию. Теперь группировка строго по месту; contourFilter оставлен в сигнатуре ради
' совместимости вызывающих (оба текущих вызова передают "") - contours() на выходе несёт только
' последний увиденный тег как информационный, не влияет на суммирование.
' committed_only_stock (2.0.1): фильтр по PRIME_IsOpIdCommitted - см. комментарий у
' PRIME_04_Posting.PRIME_LotBalance, здесь та же независимая реализация суммирования по
' движениям, поэтому фильтр нужно было продублировать отдельно.
Public Sub PRIME_StockByLocation(ByVal productCode As String, ByVal contourFilter As String, ByRef locations() As String, ByRef quantities() As Double, ByRef contours() As String)
    Dim locCount As Long
    locCount = 0

    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        ReDim locations(-1)
        ReDim quantities(-1)
        ReDim contours(-1)
        Exit Sub
    End If

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colProduct As Long, colLoc As Long, colQty As Long, colOpId As Long, colContour As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colLoc = PRIME_ColIndex(headers, "LOCATION")
    colQty = PRIME_ColIndex(headers, "QTY_BASE")
    colOpId = PRIME_ColIndex(headers, "OP_ID")
    colContour = PRIME_ColIndex(headers, "STOCK_CONTOUR")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    If UBound(table) < 1 Then
        ReDim locations(-1)
        ReDim quantities(-1)
        ReDim contours(-1)
        Exit Sub
    End If

    ReDim locations(UBound(table))
    ReDim quantities(UBound(table))
    ReDim contours(UBound(table))

    Dim i As Long, j As Long, foundIdx As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(colProduct)) = productCode And PRIME_IsOpIdCommitted(CStr(table(i)(colOpId))) Then
            Dim rowContour As String
            rowContour = ""
            If colContour >= 0 Then rowContour = CStr(table(i)(colContour))
            If contourFilter <> "" And rowContour <> contourFilter Then GoTo ContinueLoop
            Dim loc As String
            loc = CStr(table(i)(colLoc))
            foundIdx = -1
            For j = 0 To locCount - 1
                If locations(j) = loc Then
                    foundIdx = j
                    Exit For
                End If
            Next j
            If foundIdx = -1 Then
                locations(locCount) = loc
                contours(locCount) = rowContour
                quantities(locCount) = CDbl(table(i)(colQty))
                locCount = locCount + 1
            Else
                quantities(foundIdx) = quantities(foundIdx) + CDbl(table(i)(colQty))
                If rowContour <> "" Then contours(foundIdx) = rowContour
            End If
        End If
ContinueLoop:
    Next i

    If locCount = 0 Then
        ReDim locations(-1)
        ReDim quantities(-1)
        ReDim contours(-1)
    Else
        ReDim Preserve locations(locCount - 1)
        ReDim Preserve quantities(locCount - 1)
        ReDim Preserve contours(locCount - 1)
    End If
End Sub
