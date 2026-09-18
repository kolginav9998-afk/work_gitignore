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
    code = InputBox("Внутренний код товара:", "Наличие по коду")
    If code = "" Then Exit Sub
    PRIME_Stock_Rebuild("CODE", Trim(code))
End Sub

' Дополнительные фильтры (recommendation §32): категория/подкатегория/место хранения.
Public Sub PRIME_Stock_ShowByCategoryButton()
    Dim value As String
    value = InputBox("Категория:", "Наличие по категории")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("CATEGORY", Trim(value))
End Sub

Public Sub PRIME_Stock_ShowBySubcategoryButton()
    Dim value As String
    value = InputBox("Подкатегория:", "Наличие по подкатегории")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("SUBCATEGORY", Trim(value))
End Sub

Public Sub PRIME_Stock_ShowByLocationButton()
    Dim value As String
    value = InputBox("Место хранения:", "Наличие по месту хранения")
    If Trim(value) = "" Then Exit Sub
    PRIME_Stock_Rebuild("LOCATION", Trim(value))
End Sub

' FINAL mega-task (single_physical_warehouse): контур - неотъемлемая метка происхождения, а не
' физический раздел склада, поэтому быстрые фильтры "Наличие" переведены с контура на
' происхождение EI (LOT.ORIGIN = лист, создавший позицию: "Заказы"/"Приход — Офис"/
' "Приход — Производство"/"Приход — Детали") - см. PRIME_Stock_Rebuild ниже, filterMode="ORIGIN".
' critical_fix (confirmed bug #3): раньше "Из производства" вызывал ShowContourGeneralButton
' (SC_GENERAL) - показывал общий контур вместо реального производства.
Public Sub PRIME_Stock_ShowOriginOrdersButton()
    PRIME_Stock_Rebuild("ORIGIN", SH_ORDERS)
End Sub

Public Sub PRIME_Stock_ShowOriginOfficeButton()
    PRIME_Stock_Rebuild("ORIGIN", SH_RECEIPT_OFFICE)
End Sub

Public Sub PRIME_Stock_ShowOriginProductionButton()
    PRIME_Stock_Rebuild("ORIGIN", SH_RECEIPT_PRODUCTION)
End Sub

Public Sub PRIME_Stock_ShowOriginDetailsButton()
    PRIME_Stock_Rebuild("ORIGIN", SH_RECEIPT_DETAILS)
End Sub

' Для партии, полученной перемещением/инвентаризацией (ORIGIN партии в DB_PRIME_LOTS хранит
' исходный лист прихода - см. PRIME_04_Posting.PRIME_PostTransferLines/PostAdjustmentLines),
' поднимаемся по цепочке PARENT_LOT_ID до первой партии без родителя - это и есть настоящее
' происхождение EI (заказ/офис/производство/детали), которое transfers.rule запрещает менять.
' lotTable/lotHeaders читаются один раз вызывающей стороной (PRIME_Stock_Rebuild) - здесь только
' поиск по уже загруженной в память таблице, без повторных обращений к листу на партию.
Private Sub PRIME_Stock_ResolveLotRoot(ByVal lotTable As Variant, ByVal lotHeaders As Variant, ByVal startLotId As String, _
        ByRef outOrigin As String, ByRef outDocId As String, ByRef outDate As String)
    outOrigin = "" : outDocId = "" : outDate = ""
    If startLotId = "" Then Exit Sub

    Dim colLotId As Long, colOrigin As Long, colRecDoc As Long, colRecDate As Long, colParent As Long
    colLotId = PRIME_ColIndex(lotHeaders, "LOT_ID")
    colOrigin = PRIME_ColIndex(lotHeaders, "ORIGIN")
    colRecDoc = PRIME_ColIndex(lotHeaders, "RECEIPT_DOC_ID")
    colRecDate = PRIME_ColIndex(lotHeaders, "RECEIPT_DATE")
    colParent = PRIME_ColIndex(lotHeaders, "PARENT_LOT_ID")

    Dim cur As String
    cur = startLotId
    Dim guard As Long
    guard = 0
    Do While cur <> "" And guard < 1000
        Dim idx As Long
        idx = PRIME_FindRowByKey(lotTable, colLotId, cur)
        If idx < 0 Then Exit Do
        outOrigin = CStr(lotTable(idx)(colOrigin))
        outDocId = CStr(lotTable(idx)(colRecDoc))
        outDate = CStr(lotTable(idx)(colRecDate))
        Dim parent As String
        parent = ""
        If colParent >= 0 Then parent = CStr(lotTable(idx)(colParent))
        If parent = "" Then Exit Do
        cur = parent
        guard = guard + 1
    Loop
End Sub

' R24/FINAL: "Наличие" - сводный обзор по (EI_CODE=PRODUCT_CODE, место), одна строка на реальную
' позицию, БЕЗ агрегации одноимённых/разных EI (one_physical_stock, primary_grain=EI_CODE) - не
' обязательный ежедневный экран (актуальный остаток дублируется inline на рабочих листах - см.
' PRIME_05_Orders/06_Issues/07_Workflows). Источник - те же COMMITTED-движения, что и весь
' остальной остаток/FIFO. "Тип/источник"/"Дата прихода"/"Исходный DOC_ID" берутся из корневой
' партии (PRIME_Stock_ResolveLotRoot), а не из отдельных движений - иначе после разрешённого
' cross-tag списания (Выдачи снимают любой валидный EI) строки того же EI расходились бы на
' несколько неверных "источников" (см. PRIME_Stock_Rebuild диагностику в KNOWN_ISSUES).
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
    Dim colProduct As Long, colLoc As Long, colQty As Long, colDate As Long, colOpId As Long, colMoveLot As Long, colMoveDoc As Long
    colProduct = PRIME_ColIndex(moveHeaders, "PRODUCT_CODE")
    colLoc = PRIME_ColIndex(moveHeaders, "LOCATION")
    colQty = PRIME_ColIndex(moveHeaders, "QTY_BASE")
    colDate = PRIME_ColIndex(moveHeaders, "MOVE_DATE")
    colOpId = PRIME_ColIndex(moveHeaders, "OP_ID")
    colMoveLot = PRIME_ColIndex(moveHeaders, "LOT_ID")
    colMoveDoc = PRIME_ColIndex(moveHeaders, "DOC_ID")

    Dim moveTable As Variant
    moveTable = PRIME_ReadTable(SH_DB_MOVEMENTS)
    If UBound(moveTable) < 1 Then Exit Sub

    Dim lotTable As Variant, lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    lotTable = PRIME_ReadTable(SH_DB_LOTS)

    Dim docTable As Variant, docHeaders As Variant
    Dim colDocDocId As Long, colDocType As Long
    Dim hasDocTable As Boolean
    hasDocTable = PRIME_SheetExists(SH_DB_DOCUMENTS)
    If hasDocTable Then
        docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
        docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
        colDocDocId = PRIME_ColIndex(docHeaders, "DOC_ID")
        colDocType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
    End If

    ' Агрегация product|location -> (qty, arrived, issued, returned, origin, rootDocId, rootDate),
    ' через параллельные массивы (см. PRIME_03_Catalog для объяснения, почему не Collection с
    ' перечислением ключей).
    Dim keys() As String
    Dim qtys() As Double
    Dim arrived() As Double
    Dim issued() As Double
    Dim returned() As Double
    Dim origins() As String
    Dim rootDocIds() As String
    Dim rootDates() As String
    Dim n As Long
    n = 0
    ReDim keys(UBound(moveTable))
    ReDim qtys(UBound(moveTable))
    ReDim arrived(UBound(moveTable))
    ReDim issued(UBound(moveTable))
    ReDim returned(UBound(moveTable))
    ReDim origins(UBound(moveTable))
    ReDim rootDocIds(UBound(moveTable))
    ReDim rootDates(UBound(moveTable))

    Dim i As Long, j As Long
    For i = 1 To UBound(moveTable)
        Dim pc As String, loc As String
        pc = CStr(moveTable(i)(colProduct))

        Dim lotOrigin As String, lotRootDocId As String, lotRootDate As String
        PRIME_Stock_ResolveLotRoot lotTable, lotHeaders, CStr(moveTable(i)(colMoveLot)), lotOrigin, lotRootDocId, lotRootDate

        If filterMode = "CODE" And pc <> filterValue Then GoTo ContinueLoop
        If filterMode = "CATEGORY" And LCase(PRIME_GetProductField(pc, "CATEGORY")) <> LCase(filterValue) Then GoTo ContinueLoop
        If filterMode = "SUBCATEGORY" And LCase(PRIME_GetProductField(pc, "SUBCATEGORY")) <> LCase(filterValue) Then GoTo ContinueLoop
        If filterMode = "LOCATION" And LCase(CStr(moveTable(i)(colLoc))) <> LCase(filterValue) Then GoTo ContinueLoop
        If filterMode = "ORIGIN" And lotOrigin <> filterValue Then GoTo ContinueLoop
        ' committed_only_stock (2.0.1): лист "Наличие" не должен показывать PREPARED/FAILED
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
        If foundIdx = -1 Then
            foundIdx = n
            keys(n) = k
            qtys(n) = 0
            arrived(n) = 0
            issued(n) = 0
            returned(n) = 0
            origins(n) = lotOrigin
            rootDocIds(n) = lotRootDocId
            rootDates(n) = lotRootDate
            n = n + 1
        End If

        Dim qty As Double
        qty = CDbl(moveTable(i)(colQty))
        qtys(foundIdx) = qtys(foundIdx) + qty

        Dim docType As String
        docType = ""
        If hasDocTable And colMoveDoc >= 0 Then
            Dim docIdx As Long
            docIdx = PRIME_FindRowByKey(docTable, colDocDocId, CStr(moveTable(i)(colMoveDoc)))
            If docIdx >= 0 Then docType = CStr(docTable(docIdx)(colDocType))
        End If

        If docType = DOC_RETURN And qty > 0 Then
            returned(foundIdx) = returned(foundIdx) + qty
        ElseIf qty > 0 Then
            arrived(foundIdx) = arrived(foundIdx) + qty
        ElseIf qty < 0 Then
            issued(foundIdx) = issued(foundIdx) - qty
        End If
ContinueLoop:
    Next i

    ' headless_buffered_array_readback_corruption (2.1.2, см. PRIME_04_Posting.PRIME_PostIssueLines) -
    ' раньше строки накапливались в outRows()-буфер (Dim row() внутри цикла, ReDim Preserve в конце)
    ' и передавались одним PRIME_AppendRowsBatch; фикс - писать каждую строку сразу поячейково
    ' внутри цикла, без буферизации.
    Dim colOutCode As Long, colOutName As Long, colOutOrigin As Long, colOutLoc As Long
    Dim colOutArrived As Long, colOutIssued As Long, colOutReturned As Long, colOutBalance As Long
    Dim colOutRecDate As Long, colOutRootDoc As Long
    colOutCode = PRIME_ColIndex(headers, "Внутренний код")
    colOutName = PRIME_ColIndex(headers, "Наименование")
    colOutOrigin = PRIME_ColIndex(headers, "Тип/источник")
    colOutLoc = PRIME_ColIndex(headers, "Место хранения")
    colOutArrived = PRIME_ColIndex(headers, "Пришло")
    colOutIssued = PRIME_ColIndex(headers, "Выдано/списано")
    colOutReturned = PRIME_ColIndex(headers, "Возвращено")
    colOutBalance = PRIME_ColIndex(headers, "Остаток")
    colOutRecDate = PRIME_ColIndex(headers, "Дата прихода")
    colOutRootDoc = PRIME_ColIndex(headers, "Исходный DOC_ID")

    Dim outRow As Long
    outRow = PRIME_FormSchemaFirstDataRow(SH_STOCK)
    Dim outN As Long
    outN = 0
    Dim parts() As String
    For i = 0 To n - 1
        Dim include As Boolean
        Select Case filterMode
            Case "NONZERO" : include = (qtys(i) <> 0)
            Case "NEGATIVE" : include = (qtys(i) < 0)
            Case Else : include = True
        End Select
        If include Then
            parts = Split(keys(i), "|")
            oSheet.getCellByPosition(colOutCode, outRow).setString(parts(0))
            oSheet.getCellByPosition(colOutName, outRow).setString(PRIME_GetProductField(parts(0), "PRODUCT_NAME"))
            oSheet.getCellByPosition(colOutOrigin, outRow).setString(origins(i))
            oSheet.getCellByPosition(colOutLoc, outRow).setString(parts(1))
            oSheet.getCellByPosition(colOutArrived, outRow).setValue(arrived(i))
            oSheet.getCellByPosition(colOutIssued, outRow).setValue(issued(i))
            oSheet.getCellByPosition(colOutReturned, outRow).setValue(returned(i))
            oSheet.getCellByPosition(colOutBalance, outRow).setValue(qtys(i))
            oSheet.getCellByPosition(colOutRecDate, outRow).setString(rootDates(i))
            oSheet.getCellByPosition(colOutRootDoc, outRow).setString(rootDocIds(i))
            outRow = outRow + 1
            outN = outN + 1
        End If
    Next i
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
    PRIME_FifoLotsForProductAny(code, lots, balances)
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
    query = InputBox("Поиск (код, наименование, ORDER_ID, DOC_ID, LOT_ID, получатель...):", "Поиск")
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

    ' headless_buffered_array_readback_corruption (2.1.2, см. PRIME_04_Posting.PRIME_PostIssueLines) -
    ' раньше строки накапливались в outRows()-буфер (Dim row() внутри цикла, растущий ReDim Preserve)
    ' и передавались одним PRIME_AppendRowsBatch; фикс - писать каждую строку сразу поячейково
    ' внутри цикла, без буферизации.
    Dim colOutType As Long, colOutDate As Long, colOutDocId As Long, colOutCode As Long
    Dim colOutName As Long, colOutQty As Long, colOutDetails As Long
    colOutType = PRIME_ColIndex(headers, "Тип")
    colOutDate = PRIME_ColIndex(headers, "Дата")
    colOutDocId = PRIME_ColIndex(headers, "DOC_ID")
    colOutCode = PRIME_ColIndex(headers, "Код")
    colOutName = PRIME_ColIndex(headers, "Наименование")
    colOutQty = PRIME_ColIndex(headers, "Количество")
    colOutDetails = PRIME_ColIndex(headers, "Подробности")

    Dim outRow As Long
    outRow = PRIME_FormSchemaFirstDataRow(SH_SEARCH)
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

            ' committed_only_everywhere (R07): DB_PRIME_DOCUMENTS/DOC_LINES строки пишутся ДО
            ' TX=COMMITTED (см. PRIME_04_Posting.PRIME_WriteDocumentHeader) - поиск не должен
            ' показывать PREPARED/FAILED документы как реально существующие.
            If Not PRIME_IsDocIdCommitted(docId) Then GoTo ContinueSearchLoop

            Dim docIdx As Long
            docIdx = PRIME_FindRowByKey(docTable, colDocDocId, docId)
            Dim orderId As String
            orderId = ""
            If docIdx >= 0 Then orderId = CStr(docTable(docIdx)(colOrderId))

            Dim haystack As String
            haystack = LCase(docId & " " & productCode & " " & productName & " " & recipient & " " & dest & " " & orderId)
            If InStr(haystack, query) > 0 Then
                oSheet.getCellByPosition(colOutType, outRow).setString(IIf(docIdx >= 0, CStr(docTable(docIdx)(colDocType)), ""))
                oSheet.getCellByPosition(colOutDate, outRow).setString(IIf(docIdx >= 0, CStr(docTable(docIdx)(colDocDate)), ""))
                oSheet.getCellByPosition(colOutDocId, outRow).setString(docId)
                oSheet.getCellByPosition(colOutCode, outRow).setString(productCode)
                oSheet.getCellByPosition(colOutName, outRow).setString(productName)
                oSheet.getCellByPosition(colOutQty, outRow).setValue(CDbl(lineTable(i)(colQty)))
                oSheet.getCellByPosition(colOutDetails, outRow).setString("Получатель: " & recipient & "; Назначение: " & dest & "; ORDER_ID: " & orderId)
                outRow = outRow + 1
                n = n + 1
            End If
ContinueSearchLoop:
        Next i
    End If

    If n > 0 Then
        MsgBox "Найдено строк: " & n
    Else
        MsgBox "Ничего не найдено."
    End If
End Sub
