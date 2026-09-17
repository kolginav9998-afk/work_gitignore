Option Explicit

' PRIME_09_StockSearch
' Единственная реализация остатков и поиска (в 1.4.1 их было по два параллельных набора -
' модуль 09 неиспользуемый и модуль 24 живой для поиска; здесь только одна реализация каждого).

' === Остаток ====================================================================================
' Группировка по PRODUCT_CODE + BASE_UNIT + LOCATION (stock_view.group_by), источник - COMMITTED
' движения. Без скрытого лимита строк (silent_2000_row_limit=false) - пишем пакетно, сколько есть.

Public Sub PRIME_Stock_ShowAllButton()
    PRIME_Stock_Rebuild("ALL", "")
End Sub

Public Sub PRIME_Stock_ShowNonZeroButton()
    PRIME_Stock_Rebuild("NONZERO", "")
End Sub

Public Sub PRIME_Stock_ShowNegativeButton()
    PRIME_Stock_Rebuild("NEGATIVE", "")
End Sub

Public Sub PRIME_Stock_SearchByCodeButton()
    Dim code As String
    code = InputBox("Внутренний код товара:", "Остаток по коду")
    If code = "" Then Exit Sub
    PRIME_Stock_Rebuild("CODE", Trim(code))
End Sub

' Дополнительные фильтры остатка (recommendation §32): категория/подкатегория/место хранения.
Public Sub PRIME_Stock_ShowByCategoryButton()
    Dim value As String
    value = InputBox("Категория:", "Остаток по категории")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("CATEGORY", Trim(value))
End Sub

Public Sub PRIME_Stock_ShowBySubcategoryButton()
    Dim value As String
    value = InputBox("Подкатегория:", "Остаток по подкатегории")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("SUBCATEGORY", Trim(value))
End Sub

Public Sub PRIME_Stock_ShowByLocationButton()
    Dim value As String
    value = InputBox("Место хранения:", "Остаток по месту хранения")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("LOCATION", Trim(value))
End Sub

Private Sub PRIME_Stock_Rebuild(ByVal filterMode As String, ByVal filterValue As String)
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_STOCK)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_STOCK)

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)

    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then Exit Sub
    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colProduct As Long, colLoc As Long, colQty As Long, colDate As Long, colOpId As Long
    colProduct = PRIME_ColIndex(moveHeaders, "PRODUCT_CODE")
    colLoc = PRIME_ColIndex(moveHeaders, "LOCATION")
    colQty = PRIME_ColIndex(moveHeaders, "QTY_BASE")
    colDate = PRIME_ColIndex(moveHeaders, "MOVE_DATE")
    colOpId = PRIME_ColIndex(moveHeaders, "OP_ID")

    Dim moveTable As Variant
    moveTable = PRIME_ReadTable(SH_DB_MOVEMENTS)
    If UBound(moveTable) < 1 Then Exit Sub

    ' Агрегация product|location -> (qty, lastDate), через параллельные массивы (см. PRIME_03_Catalog
    ' для объяснения, почему не Collection с перечислением ключей).
    Dim keys() As String
    Dim qtys() As Double
    Dim lastDates() As String
    Dim n As Long
    n = 0
    ReDim keys(UBound(moveTable))
    ReDim qtys(UBound(moveTable))
    ReDim lastDates(UBound(moveTable))

    Dim i As Long, j As Long
    For i = 1 To UBound(moveTable)
        Dim pc As String, loc As String
        pc = CStr(moveTable(i)(colProduct))
        If filterMode = "CODE" And pc <> filterValue Then GoTo ContinueLoop
        If filterMode = "CATEGORY" And LCase(PRIME_GetProductField(pc, "CATEGORY")) <> LCase(filterValue) Then GoTo ContinueLoop
        If filterMode = "SUBCATEGORY" And LCase(PRIME_GetProductField(pc, "SUBCATEGORY")) <> LCase(filterValue) Then GoTo ContinueLoop
        If filterMode = "LOCATION" And LCase(CStr(moveTable(i)(colLoc))) <> LCase(filterValue) Then GoTo ContinueLoop
        ' committed_only_stock (2.0.1): лист "Остаток" не должен показывать PREPARED/FAILED
        ' движения как реальный остаток - см. PRIME_04_Posting.PRIME_LotBalance.
        If Not PRIME_IsOpIdCommitted(CStr(moveTable(i)(colOpId))) Then GoTo ContinueLoop
        loc = CStr(moveTable(i)(colLoc))
        Dim k As String
        k = pc & "|" & loc
        Dim foundIdx As Long
        foundIdx = -1
        For j = 0 To n - 1
            If keys(j) = k Then
                foundIdx = j
                Exit For
            End If
        Next j
        Dim md As String
        md = CStr(moveTable(i)(colDate))
        If foundIdx = -1 Then
            keys(n) = k
            qtys(n) = CDbl(moveTable(i)(colQty))
            lastDates(n) = md
            n = n + 1
        Else
            qtys(foundIdx) = qtys(foundIdx) + CDbl(moveTable(i)(colQty))
            If md > lastDates(foundIdx) Then lastDates(foundIdx) = md
        End If
ContinueLoop:
    Next i

    Dim outRows() As Variant
    ReDim outRows(n)
    Dim outN As Long
    outN = 0
    For i = 0 To n - 1
        Dim include As Boolean
        Select Case filterMode
            Case "NONZERO" : include = (qtys(i) <> 0)
            Case "NEGATIVE" : include = (qtys(i) < 0)
            Case Else : include = True
        End Select
        If include Then
            Dim parts() As String
            parts = Split(keys(i), "|")
            Dim row(UBound(headers)) As Variant
            row(PRIME_ColIndex(headers, "Код")) = parts(0)
            row(PRIME_ColIndex(headers, "Наименование")) = PRIME_GetProductField(parts(0), "PRODUCT_NAME")
            row(PRIME_ColIndex(headers, "Ед. изм.")) = PRIME_GetProductField(parts(0), "BASE_UNIT")
            row(PRIME_ColIndex(headers, "Место хранения")) = parts(1)
            row(PRIME_ColIndex(headers, "Остаток")) = qtys(i)
            row(PRIME_ColIndex(headers, "Последняя операция")) = lastDates(i)
            outRows(outN) = row
            outN = outN + 1
        End If
    Next i

    If outN > 0 Then
        ReDim Preserve outRows(outN - 1)
        PRIME_AppendRowsBatch(SH_STOCK, outRows)
    End If
End Sub

' === Остаток — Заказы ==========================================================================
' Все сохранённые снимки заказа + текущий остаток соответствующей партии, включая нулевой
' (include_zero_lot_balance) - общий остаток товара НЕ дублируется как итог каждой поставки.
Public Sub PRIME_StockOrders_RefreshButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_STOCK_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_STOCK_ORDERS)

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)

    If Not PRIME_SheetExists(SH_DB_ORDER_SNAPSHOT) Then Exit Sub
    Dim snapHeaders As Variant
    snapHeaders = PRIME_HeaderMap(SH_DB_ORDER_SNAPSHOT)
    Dim snapTable As Variant
    snapTable = PRIME_ReadTable(SH_DB_ORDER_SNAPSHOT)
    If UBound(snapTable) < 1 Then Exit Sub

    Dim colOrderId As Long, colDate As Long, colProduct As Long, colQty As Long, colLot As Long
    colOrderId = PRIME_ColIndex(snapHeaders, "ORDER_ID")
    colDate = PRIME_ColIndex(snapHeaders, "DOC_DATE")
    colProduct = PRIME_ColIndex(snapHeaders, "PRODUCT_CODE")
    colQty = PRIME_ColIndex(snapHeaders, "DELIVERY_QTY")
    colLot = PRIME_ColIndex(snapHeaders, "LOT_ID")

    Dim outRows(UBound(snapTable) - 1) As Variant
    Dim i As Long
    For i = 1 To UBound(snapTable)
        Dim row(UBound(headers)) As Variant
        Dim productCode As String
        productCode = CStr(snapTable(i)(colProduct))
        Dim lotId As String
        lotId = CStr(snapTable(i)(colLot))
        row(PRIME_ColIndex(headers, "ORDER_ID")) = CStr(snapTable(i)(colOrderId))
        row(PRIME_ColIndex(headers, "Дата")) = CStr(snapTable(i)(colDate))
        row(PRIME_ColIndex(headers, "Код товара")) = productCode
        row(PRIME_ColIndex(headers, "Наименование")) = PRIME_GetProductField(productCode, "PRODUCT_NAME")
        row(PRIME_ColIndex(headers, "Количество поставки")) = CDbl(snapTable(i)(colQty))
        row(PRIME_ColIndex(headers, "LOT_ID")) = lotId
        row(PRIME_ColIndex(headers, "Текущий остаток партии")) = PRIME_LotBalance(lotId)
        outRows(i - 1) = row
    Next i

    PRIME_AppendRowsBatch(SH_STOCK_ORDERS, outRows)
End Sub

' Список партий конкретного товара с их текущим балансом (WMS_STOCK_LOTS-аналог) - пишет в лист
' "База - Поиск" в том же формате, что и обычный поиск, чтобы не заводить ещё один лист.
Public Sub PRIME_Stock_ShowLotsByCodeButton()
    Dim code As String
    code = InputBox("Внутренний код товара:", "Партии товара")
    If Trim(code) = "" Then Exit Sub
    code = Trim(code)

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_SEARCH)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SEARCH)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)

    Dim lots() As String
    Dim balances() As Double
    PRIME_FifoLotsForProduct(code, lots, balances)
    If UBound(lots) < LBound(lots) Then
        MsgBox "У товара " & code & " нет партий."
        Exit Sub
    End If

    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim lotTable As Variant
    lotTable = PRIME_ReadTable(SH_DB_LOTS)
    Dim colLotId As Long, colDate As Long, colLoc As Long
    colLotId = PRIME_ColIndex(lotHeaders, "LOT_ID")
    colDate = PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")
    colLoc = PRIME_ColIndex(lotHeaders, "LOCATION")

    Dim outRows(UBound(lots)) As Variant
    Dim i As Long
    For i = LBound(lots) To UBound(lots)
        Dim idx As Long
        idx = PRIME_FindRowByKey(lotTable, colLotId, lots(i))
        Dim row(UBound(headers)) As Variant
        row(PRIME_ColIndex(headers, "Тип")) = "LOT"
        row(PRIME_ColIndex(headers, "Код")) = code
        row(PRIME_ColIndex(headers, "Наименование")) = PRIME_GetProductField(code, "PRODUCT_NAME")
        row(PRIME_ColIndex(headers, "Количество")) = balances(i)
        If idx >= 0 Then
            row(PRIME_ColIndex(headers, "Дата")) = CStr(lotTable(idx)(colDate))
            row(PRIME_ColIndex(headers, "Подробности")) = "LOT_ID: " & lots(i) & "; Место: " & CStr(lotTable(idx)(colLoc))
        Else
            row(PRIME_ColIndex(headers, "Подробности")) = "LOT_ID: " & lots(i)
        End If
        outRows(i) = row
    Next i
    PRIME_AppendRowsBatch(SH_SEARCH, outRows)
End Sub

Public Sub PRIME_Search_ClearButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_SEARCH)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SEARCH)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)
End Sub

' === База - Поиск ===============================================================================
' batch_search по документам/строкам/партиям сразу по нескольким полям (search.search_fields).
Public Sub PRIME_Search_RunButton()
    Dim query As String
    query = InputBox("Поиск (код, наименование, ORDER_ID, DOC_ID, LOT_ID, получатель...):", "База - Поиск")
    If Trim(query) = "" Then Exit Sub
    PRIME_Search_Execute(LCase(Trim(query)))
End Sub

' "Показать всё" - тот же вывод, без фильтра по подстроке.
Public Sub PRIME_Search_ShowAllButton()
    PRIME_Search_Execute("")
End Sub

Private Sub PRIME_Search_Execute(ByVal query As String)
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_SEARCH)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SEARCH)

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)

    If Not PRIME_SheetExists(SH_DB_DOC_LINES) Or Not PRIME_SheetExists(SH_DB_DOCUMENTS) Then Exit Sub

    Dim lineHeaders As Variant, docHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim lineTable As Variant, docTable As Variant
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)

    Dim colLineDocId As Long, colProduct As Long, colQty As Long, colRecipient As Long, colDest As Long
    colLineDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
    colProduct = PRIME_ColIndex(lineHeaders, "PRODUCT_CODE")
    colQty = PRIME_ColIndex(lineHeaders, "QTY_BASE")
    colRecipient = PRIME_ColIndex(lineHeaders, "RECIPIENT")
    colDest = PRIME_ColIndex(lineHeaders, "DESTINATION_PROJECT")

    Dim colDocDocId As Long, colDocType As Long, colDocDate As Long, colOrderId As Long
    colDocDocId = PRIME_ColIndex(docHeaders, "DOC_ID")
    colDocType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
    colDocDate = PRIME_ColIndex(docHeaders, "DOC_DATE")
    colOrderId = PRIME_ColIndex(docHeaders, "ORDER_ID")

    Dim outRows() As Variant
    ReDim outRows(200)
    Dim n As Long
    n = 0

    If UBound(lineTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(lineTable)
            Dim docId As String
            docId = CStr(lineTable(i)(colLineDocId))
            Dim productCode As String
            productCode = CStr(lineTable(i)(colProduct))
            Dim recipient As String
            recipient = CStr(lineTable(i)(colRecipient))
            Dim dest As String
            dest = CStr(lineTable(i)(colDest))
            Dim productName As String
            productName = PRIME_GetProductField(productCode, "PRODUCT_NAME")

            Dim docIdx As Long
            docIdx = PRIME_FindRowByKey(docTable, colDocDocId, docId)
            Dim orderId As String
            orderId = ""
            If docIdx >= 0 Then orderId = CStr(docTable(docIdx)(colOrderId))

            Dim haystack As String
            haystack = LCase(docId & " " & productCode & " " & productName & " " & recipient & " " & dest & " " & orderId)
            If InStr(haystack, query) > 0 Then
                Dim row(UBound(headers)) As Variant
                row(PRIME_ColIndex(headers, "Тип")) = IIf(docIdx >= 0, CStr(docTable(docIdx)(colDocType)), "")
                row(PRIME_ColIndex(headers, "Дата")) = IIf(docIdx >= 0, CStr(docTable(docIdx)(colDocDate)), "")
                row(PRIME_ColIndex(headers, "DOC_ID")) = docId
                row(PRIME_ColIndex(headers, "Код")) = productCode
                row(PRIME_ColIndex(headers, "Наименование")) = productName
                row(PRIME_ColIndex(headers, "Количество")) = CDbl(lineTable(i)(colQty))
                row(PRIME_ColIndex(headers, "Подробности")) = "Получатель: " & recipient & "; Назначение: " & dest & "; ORDER_ID: " & orderId
                If n > UBound(outRows) Then ReDim Preserve outRows(UBound(outRows) + 200)
                outRows(n) = row
                n = n + 1
            End If
        Next i
    End If

    If n > 0 Then
        ReDim Preserve outRows(n - 1)
        PRIME_AppendRowsBatch(SH_SEARCH, outRows)
        MsgBox "Найдено строк: " & n
    Else
        MsgBox "Ничего не найдено."
    End If
End Sub
