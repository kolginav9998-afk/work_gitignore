Option Explicit

' PRIME_02_Store
' Единственный слой, которому разрешено напрямую читать/писать скрытые DB_PRIME_*/SYS_PRIME_*
' листы (data_access_layer.only_store_layer_directly_accesses_hidden_db_sheets). Все остальные
' модули работают через эти функции, а не через getCellByPosition на системных листах напрямую.
'
' Никаких commit/rollback/DDL/store() здесь нет - это чистый доступ к диапазонам Calc
' (performance.preferred_range_api: getDataArray/setDataArray, батчами).

' As Object, не As Collection: StarBasic не позволяет отложенное объявление без New для
' типа Collection (компилируется только "As New Collection" или "As Object" + Set при инициализации).
Private gHeaderCache As Object      ' Collection: имя листа -> Variant(массив заголовков)
Private gCommittedKeyCache As Object ' Collection: SOURCE_KEY -> DOC_ID, только для COMMITTED
Private gCommittedOpIdCache As Object ' Collection: OP_ID -> True, только для COMMITTED (committed_only_stock)
Private gCommittedDocIdCache As Object ' Collection: DOC_ID -> True, только для COMMITTED (2.1.0, committed_only_everywhere)

Private Function PRIME_Doc() As Object
    PRIME_Doc = ThisComponent
End Function

Public Function PRIME_GetSheet(ByVal sheetName As String) As Object
    Dim oSheets As Object
    oSheets = PRIME_Doc().getSheets()
    If Not oSheets.hasByName(sheetName) Then
        Err.Raise 1001, "PRIME_Store.PRIME_GetSheet", "PRIME_Store: лист не найден: " & sheetName
    End If
    PRIME_GetSheet = oSheets.getByName(sheetName)
End Function

Public Function PRIME_SheetExists(ByVal sheetName As String) As Boolean
    PRIME_SheetExists = PRIME_Doc().getSheets().hasByName(sheetName)
End Function

' Последняя фактически использованная строка данных (0-based), headerRow-1 если данных нет
' (см. PRIME_FormSchemaHeaderRow - header_schema_registry: заголовок не всегда в строке 0).
Public Function PRIME_FindLastRow(ByVal oSheet As Object) As Long
    Dim headerRow As Long
    headerRow = PRIME_FormSchemaHeaderRow(oSheet.Name)
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    Dim lastRow As Long
    lastRow = oCursor.RangeAddress.EndRow
    If lastRow <= headerRow Then
        PRIME_FindLastRow = headerRow - 1 ' нет строк данных
        Exit Function
    End If
    PRIME_FindLastRow = lastRow
End Function

Public Function PRIME_FindLastCol(ByVal oSheet As Object) As Long
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    PRIME_FindLastCol = oCursor.RangeAddress.EndColumn
End Function

' Заголовки (строка 0) листа, с кэшем. Инвалидация - PRIME_InvalidateHeaderCache(sheetName)
' при структурных изменениях (миграция/установка), не на каждый ввод.
Public Function PRIME_HeaderMap(ByVal sheetName As String) As Variant
    If gHeaderCache Is Nothing Then
        Set gHeaderCache = New Collection
    End If

    Dim cached As Variant
    If PRIME_CollectionTryGet(gHeaderCache, sheetName, cached) Then
        PRIME_HeaderMap = cached
        Exit Function
    End If

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim lastCol As Long
    lastCol = PRIME_FindLastCol(oSheet)
    If lastCol < 0 Then lastCol = 0

    Dim headerRow As Long
    headerRow = PRIME_FormSchemaHeaderRow(sheetName)
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, headerRow, lastCol, headerRow)
    Dim data As Variant
    data = oRange.getDataArray()

    Dim headers(lastCol) As String
    Dim i As Long
    For i = 0 To lastCol
        headers(i) = CStr(data(0)(i))
    Next i

    gHeaderCache.Add(headers, sheetName)
    PRIME_HeaderMap = headers
End Function

Public Sub PRIME_InvalidateHeaderCache(ByVal sheetName As String)
    If gHeaderCache Is Nothing Then Exit Sub
    On Error Resume Next
    gHeaderCache.Remove(sheetName)
    On Error Goto 0
End Sub

Public Function PRIME_ColIndex(ByVal headers As Variant, ByVal columnName As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If headers(i) = columnName Then
            PRIME_ColIndex = i
            Exit Function
        End If
    Next i
    PRIME_ColIndex = -1
End Function

' Читает весь лист начиная с реальной строки заголовка (header_schema_registry) одним
' getDataArray. Возвращаемый массив ВСЕГДА 0-индексирован относительно заголовка: table(0) -
' заголовок, table(1..) - данные, независимо от физической строки на листе - весь остальной
' код (PRIME_FindRowByKey, циклы "For i = 1 To UBound(table)") работает с этим массивом, а не
' с физическими номерами строк, и поэтому не меняется при разных HeaderRow у разных листов.
Public Function PRIME_ReadTable(ByVal sheetName As String) As Variant
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim headerRow As Long
    headerRow = PRIME_FormSchemaHeaderRow(sheetName)
    Dim lastRow As Long, lastCol As Long
    lastRow = PRIME_FindLastRow(oSheet)
    lastCol = PRIME_FindLastCol(oSheet)
    If lastCol < 0 Then lastCol = 0
    If lastRow < headerRow Then
        ' 2.1.0 critical fix: table(0) must be a PROPER (lastCol+1)-wide phantom row of Empty
        ' values, not a bare Empty scalar - several callers (PRIME_SequenceNext, PRIME_MetaSet,
        ' and any future code following the same pattern) size a brand-new row via
        ' "Dim newRow(UBound(table(0))) As Variant" when no data row exists yet to copy from.
        ' A bare Empty scalar there made UBound(table(0)) resolve to -1 (StarBasic's "empty
        ' array" convention, used throughout this codebase for genuinely empty result arrays -
        ' see the many "ReDim x(-1)" calls elsewhere), so newRow was declared with ZERO elements
        ' and the very next "newRow(colName) = ..." wrote out of bounds. Confirmed by direct
        ' function invocation: PRIME_SequenceNext("DOC_ID") - or any sequence name - crashed the
        ' FIRST TIME EVER it ran on a genuinely empty SYS_PRIME_SEQ, which in practice meant the
        ' very first order/receipt/issue/etc. a brand-new PRIME workbook ever tried to create.
        Dim phantomRow(lastCol) As Variant
        Dim empty1(0) As Variant
        empty1(0) = phantomRow
        PRIME_ReadTable = empty1
        Exit Function
    End If
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, headerRow, lastCol, lastRow)
    PRIME_ReadTable = oRange.getDataArray()
End Function

' Дописывает строки в конец таблицы одним батчем (performance: setDataArray, не поячейково).
' rows - 2D Variant-массив (rows(rowIdx)(colIdx)), количество столбцов должно совпадать с шапкой.
Public Sub PRIME_AppendRowsBatch(ByVal sheetName As String, ByVal rows As Variant)
    Dim rowCount As Long
    rowCount = UBound(rows) - LBound(rows) + 1
    If rowCount <= 0 Then Exit Sub

    Dim colCount As Long
    colCount = UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(sheetName)
    Dim startRow As Long
    startRow = PRIME_FindLastRow(oSheet) + 1
    If startRow < firstDataRow Then startRow = firstDataRow ' никогда не писать поверх заголовка/панели

    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, startRow, colCount - 1, startRow + rowCount - 1)
    oRange.setDataArray(rows)
End Sub

' Перезаписывает существующие строки ТАБЛИЦЫ (как её возвращает PRIME_ReadTable), начиная с
' tableRowIndex - 0-based индекс ВНУТРИ массива table (не физический номер строки листа!),
' обычно результат PRIME_FindRowByKey. Сама функция переводит его в физическую строку через
' header_schema_registry - вызывающему коду не нужно знать реальный HeaderRow листа.
Public Sub PRIME_UpdateRowsBatch(ByVal sheetName As String, ByVal tableRowIndex As Long, ByVal rows As Variant)
    Dim rowCount As Long
    rowCount = UBound(rows) - LBound(rows) + 1
    If rowCount <= 0 Then Exit Sub
    Dim colCount As Long
    colCount = UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim physicalStartRow As Long
    physicalStartRow = PRIME_FormSchemaHeaderRow(sheetName) + tableRowIndex

    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, physicalStartRow, colCount - 1, physicalStartRow + rowCount - 1)
    oRange.setDataArray(rows)
End Sub

' Очищает только строки данных (firstDataRow..lastRow), не трогая заголовок/декоративную
' панель над ним - общая замена разбросанных по коду "clearContents от строки 1".
Public Sub PRIME_ClearDataRows(ByVal oSheet As Object, ByVal headers As Variant)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(oSheet.Name)
    If lastRow >= firstDataRow Then
        oSheet.getCellRangeByPosition(0, firstDataRow, UBound(headers), lastRow).clearContents(1023)
    End If
End Sub

' Линейный поиск по значению колонки (для системных справочников умеренного размера:
' products, aliases, units, lots, kits). Для DB_PRIME_MOVEMENTS/DOCUMENTS используем
' PRIME_FindCommittedBySourceKey (кэшированный), а не эту функцию.
' Возвращает индекс строки в table (0 = заголовок), -1 если не найдено.
Public Function PRIME_FindRowByKey(ByVal table As Variant, ByVal keyCol As Long, ByVal keyValue As String) As Long
    Dim i As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(keyCol)) = keyValue Then
            PRIME_FindRowByKey = i
            Exit Function
        End If
    Next i
    PRIME_FindRowByKey = -1
End Function

' --- Кэш COMMITTED SOURCE_KEY -> DOC_ID (идемпотентность, см. ARCHITECTURE §3) ---
' Инвалидируется при каждом успешном commit (PRIME_RegisterCommittedKey) и может быть
' полностью перестроен PRIME_RebuildCommittedKeyCache при открытии документа.

' Строит ОБА кэша (SOURCE_KEY->DOC_ID и OP_ID->True) одним проходом по SYS_PRIME_TX -
' committed_only_stock (2.0.1): PRIME_LotBalance/PRIME_StockByLocation обязаны считать
' только движения, чей OP_ID реально COMMITTED, иначе PREPARED/FAILED-движения (которые
' физически уже могли быть дописаны в DB_PRIME_MOVEMENTS до сбоя после записи PREPARED,
' см. PRIME_04_Posting.PRIME_PostDocument) ошибочно влияли бы на остаток.
Public Sub PRIME_RebuildCommittedKeyCache()
    Set gCommittedKeyCache = New Collection
    Set gCommittedOpIdCache = New Collection
    Set gCommittedDocIdCache = New Collection
    If Not PRIME_SheetExists(SH_SYS_TX) Then Exit Sub

    Dim tx As Variant
    tx = PRIME_ReadTable(SH_SYS_TX)
    If UBound(tx) < 1 Then Exit Sub

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim colSourceKey As Long, colDocId As Long, colState As Long, colOpId As Long
    colSourceKey = PRIME_ColIndex(headers, "SOURCE_KEY")
    colDocId = PRIME_ColIndex(headers, "DOC_ID")
    colState = PRIME_ColIndex(headers, "STATE")
    colOpId = PRIME_ColIndex(headers, "OP_ID")

    Dim i As Long
    For i = 1 To UBound(tx)
        If CStr(tx(i)(colState)) = TX_COMMITTED Then
            Dim k As String
            k = CStr(tx(i)(colSourceKey))
            If Not PRIME_CollectionHasKey(gCommittedKeyCache, k) Then
                gCommittedKeyCache.Add(CStr(tx(i)(colDocId)), k)
            End If
            Dim opId As String
            opId = CStr(tx(i)(colOpId))
            If Not PRIME_CollectionHasKey(gCommittedOpIdCache, opId) Then
                gCommittedOpIdCache.Add(True, opId)
            End If
            If Not PRIME_CollectionHasKey(gCommittedDocIdCache, CStr(tx(i)(colDocId))) Then
                gCommittedDocIdCache.Add(True, CStr(tx(i)(colDocId)))
            End If
        End If
    Next i
End Sub

' Возвращает существующий DOC_ID, если SOURCE_KEY уже COMMITTED, иначе "".
Public Function PRIME_FindCommittedBySourceKey(ByVal sourceKey As String) As String
    If gCommittedKeyCache Is Nothing Then
        PRIME_RebuildCommittedKeyCache()
    End If
    Dim existing As Variant
    If PRIME_CollectionTryGet(gCommittedKeyCache, sourceKey, existing) Then
        PRIME_FindCommittedBySourceKey = CStr(existing)
    Else
        PRIME_FindCommittedBySourceKey = ""
    End If
End Function

' 2.1.0 critical fix (R01): раньше регистрировал ТОЛЬКО SOURCE_KEY->DOC_ID и не трогал
' gCommittedOpIdCache/gCommittedDocIdCache - PRIME_IsOpIdCommitted(только что закоммиченного
' opId) возвращал False до следующего PRIME_RebuildCommittedKeyCache (обычно - до переоткрытия
' документа), поэтому только что проведённый приход мог не попасть в остаток/FIFO в ТЕКУЩЕМ
' сеансе. Теперь все три кэша обновляются в одном месте синхронно.
Public Sub PRIME_RegisterCommittedKey(ByVal sourceKey As String, ByVal docId As String, ByVal opId As String)
    If gCommittedKeyCache Is Nothing Then
        PRIME_RebuildCommittedKeyCache()
    End If
    If Not PRIME_CollectionHasKey(gCommittedKeyCache, sourceKey) Then
        gCommittedKeyCache.Add(docId, sourceKey)
    End If
    If Not PRIME_CollectionHasKey(gCommittedOpIdCache, opId) Then
        gCommittedOpIdCache.Add(True, opId)
    End If
    If Not PRIME_CollectionHasKey(gCommittedDocIdCache, docId) Then
        gCommittedDocIdCache.Add(True, docId)
    End If
End Sub

' transaction_cache_store_consistency (2.0.1, расширено в 2.1.0): вызывается, когда COMMITTED
' уже был выставлен и SOURCE_KEY/OP_ID/DOC_ID уже зарегистрированы здесь, но последующий
' ThisComponent.store() всё же провалился и TX откатывается на FAILED (см.
' PRIME_PostDocument.PostFailed) - без этого отката повторная попытка того же SOURCE_KEY
' ошибочно получила бы ответ "уже проведено", а Returns/Documents/Lines той же операции
' продолжали бы ошибочно считаться COMMITTED (committed_only_everywhere).
Public Sub PRIME_UnregisterCommittedKey(ByVal sourceKey As String, ByVal docId As String, ByVal opId As String)
    On Error Resume Next
    If Not (gCommittedKeyCache Is Nothing) Then gCommittedKeyCache.Remove(sourceKey)
    If Not (gCommittedOpIdCache Is Nothing) Then gCommittedOpIdCache.Remove(opId)
    If Not (gCommittedDocIdCache Is Nothing) Then gCommittedDocIdCache.Remove(docId)
    On Error Goto 0
End Sub

' committed_only_stock: True, только если этот OP_ID зафиксирован как COMMITTED в SYS_PRIME_TX.
Public Function PRIME_IsOpIdCommitted(ByVal opId As String) As Boolean
    If gCommittedOpIdCache Is Nothing Then
        PRIME_RebuildCommittedKeyCache()
    End If
    PRIME_IsOpIdCommitted = PRIME_CollectionHasKey(gCommittedOpIdCache, opId)
End Function

' committed_only_everywhere (2.1.0, R07): True, только если DOC_ID реально COMMITTED - нужно
' всем чтениям DB_PRIME_DOCUMENTS/DOC_LINES/RETURNS/ALLOCATIONS/ORDER_SNAPSHOT, которые сами по
' себе не хранят STATE и раньше могли увидеть строки, физически записанные ДО отката на FAILED
' (см. PRIME_04_Posting.PRIME_WriteDocumentHeader - пишет строки на шаге 4, до TX_COMMITTED).
Public Function PRIME_IsDocIdCommitted(ByVal docId As String) As Boolean
    If docId = "" Then
        PRIME_IsDocIdCommitted = False
        Exit Function
    End If
    If gCommittedDocIdCache Is Nothing Then
        PRIME_RebuildCommittedKeyCache()
    End If
    PRIME_IsDocIdCommitted = PRIME_CollectionHasKey(gCommittedDocIdCache, docId)
End Function

' DOC_LINE_ID всегда имеет вид "<DOC_ID>-L<n>" (см. PRIME_04_Posting.PRIME_WriteDocLines) -
' DOC_ID сам по себе может содержать дефисы ("DOC-00000123"), поэтому ищем ПОСЛЕДНЕЕ "-L", а не
' режем по первому дефису.
Public Function PRIME_DocIdFromLineId(ByVal docLineId As String) As String
    Dim pos As Long
    pos = InStrRev(docLineId, "-L")
    If pos <= 0 Then
        PRIME_DocIdFromLineId = ""
    Else
        PRIME_DocIdFromLineId = Left(docLineId, pos - 1)
    End If
End Function

' --- Générique: следующий номер последовательности (SYS_PRIME_SEQ), пакетно безопасно ---
Public Function PRIME_SequenceNext(ByVal seqName As String) As Long
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_SEQ)
    Dim colName As Long, colValue As Long
    colName = PRIME_ColIndex(headers, "SEQ_NAME")
    colValue = PRIME_ColIndex(headers, "NEXT_VALUE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_SEQ)
    Dim rowIdx As Long
    rowIdx = PRIME_FindRowByKey(table, colName, seqName)

    Dim nextVal As Long
    If rowIdx = -1 Then
        nextVal = 1
        Dim newRow(UBound(table(0))) As Variant
        newRow(colName) = seqName
        newRow(colValue) = nextVal
        Dim rows(0) As Variant
        rows(0) = newRow
        PRIME_AppendRowsBatch(SH_SYS_SEQ, rows)
    Else
        nextVal = CLng(table(rowIdx)(colValue)) + 1
        Dim updRow(UBound(table(rowIdx))) As Variant
        Dim c As Long
        For c = 0 To UBound(table(rowIdx))
            updRow(c) = table(rowIdx)(c)
        Next c
        updRow(colValue) = nextVal
        Dim updRows(0) As Variant
        updRows(0) = updRow
        PRIME_UpdateRowsBatch(SH_SYS_SEQ, rowIdx, updRows)
    End If
    PRIME_SequenceNext = nextVal
End Function

' Только чтение: последнее выданное значение последовательности (0, если ни разу не
' использовалась) - НИЧЕГО не пишет. Нужна для планирования (validation) новых ЕИ-кодов ДО того,
' как транзакция дошла до PREPARED: transaction protocol запрещает физическую запись до
' SYS_PRIME_TX.STATE=PREPARED, но код товара должен быть решён заранее, чтобы попасть в план.
' Следующий код = PRIME_PeekSequenceValue(...) + 1 (то же самое "+1 от текущего", что и
' PRIME_SequenceNext, чтобы peek и реальная выдача никогда не расходились).
Public Function PRIME_PeekSequenceValue(ByVal seqName As String) As Long
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_SEQ)
    Dim colName As Long, colValue As Long
    colName = PRIME_ColIndex(headers, "SEQ_NAME")
    colValue = PRIME_ColIndex(headers, "NEXT_VALUE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_SEQ)
    Dim rowIdx As Long
    rowIdx = PRIME_FindRowByKey(table, colName, seqName)

    If rowIdx = -1 Then
        PRIME_PeekSequenceValue = 0
    Else
        PRIME_PeekSequenceValue = CLng(table(rowIdx)(colValue))
    End If
End Function

' Персистентно фиксирует значение последовательности РОВНО в newValue (создаёт строку, если её
' ещё нет). Используется на стадии физической записи (после PREPARED), чтобы зафиксировать сдвиг
' счётчика, уже решённый на стадии планирования через PRIME_PeekSequenceValue - в отличие от
' PRIME_SequenceNext (семантика "+1 от текущего"), здесь пишется точное заранее вычисленное число.
Public Sub PRIME_AdvanceSequenceTo(ByVal seqName As String, ByVal newValue As Long)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_SEQ)
    Dim colName As Long, colValue As Long
    colName = PRIME_ColIndex(headers, "SEQ_NAME")
    colValue = PRIME_ColIndex(headers, "NEXT_VALUE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_SEQ)
    Dim rowIdx As Long
    rowIdx = PRIME_FindRowByKey(table, colName, seqName)

    If rowIdx = -1 Then
        Dim newRow(UBound(table(0))) As Variant
        newRow(colName) = seqName
        newRow(colValue) = newValue
        Dim rows(0) As Variant
        rows(0) = newRow
        PRIME_AppendRowsBatch(SH_SYS_SEQ, rows)
    Else
        Dim updRow(UBound(table(rowIdx))) As Variant
        Dim c As Long
        For c = 0 To UBound(table(rowIdx))
            updRow(c) = table(rowIdx)(c)
        Next c
        updRow(colValue) = newValue
        Dim updRows(0) As Variant
        updRows(0) = updRow
        PRIME_UpdateRowsBatch(SH_SYS_SEQ, rowIdx, updRows)
    End If
End Sub

' --- SYS_PRIME_META: простой key/value (версии, текущая сессия инвентаризации и т.п.) ---
Public Function PRIME_MetaGet(ByVal key As String) As String
    If Not PRIME_SheetExists(SH_SYS_META) Then
        PRIME_MetaGet = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_META)
    Dim colKey As Long, colValue As Long
    colKey = PRIME_ColIndex(headers, "KEY")
    colValue = PRIME_ColIndex(headers, "VALUE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_META)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colKey, key)
    If idx = -1 Then
        PRIME_MetaGet = ""
    Else
        PRIME_MetaGet = CStr(table(idx)(colValue))
    End If
End Function

Public Sub PRIME_MetaSet(ByVal key As String, ByVal value As String)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_META)
    Dim colKey As Long, colValue As Long
    colKey = PRIME_ColIndex(headers, "KEY")
    colValue = PRIME_ColIndex(headers, "VALUE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_META)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, colKey, key)

    If idx = -1 Then
        Dim newRow(UBound(table(0))) As Variant
        newRow(colKey) = key
        newRow(colValue) = value
        Dim rows(0) As Variant
        rows(0) = newRow
        PRIME_AppendRowsBatch(SH_SYS_META, rows)
    Else
        Dim updRow(UBound(table(idx))) As Variant
        Dim c As Long
        For c = 0 To UBound(table(idx))
            updRow(c) = table(idx)(c)
        Next c
        updRow(colValue) = value
        Dim updRows(0) As Variant
        updRows(0) = updRow
        PRIME_UpdateRowsBatch(SH_SYS_META, idx, updRows)
    End If
End Sub

' --- Вспомогательные обёртки над Collection (в StarBasic нет .Contains/.TryGetValue) ---
Public Function PRIME_CollectionHasKey(ByVal coll As Object, ByVal key As String) As Boolean
    Dim dummy As Variant
    PRIME_CollectionHasKey = PRIME_CollectionTryGet(coll, key, dummy)
End Function

Public Function PRIME_CollectionTryGet(ByVal coll As Object, ByVal key As String, ByRef outValue As Variant) As Boolean
    If coll Is Nothing Then
        PRIME_CollectionTryGet = False
        Exit Function
    End If
    On Error Goto NotFound
    outValue = coll.Item(key)
    PRIME_CollectionTryGet = True
    Exit Function
NotFound:
    PRIME_CollectionTryGet = False
End Function
