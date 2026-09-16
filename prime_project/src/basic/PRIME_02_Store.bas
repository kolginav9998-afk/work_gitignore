Option Explicit

' PRIME_02_Store
' Единственный слой, которому разрешено напрямую читать/писать скрытые DB_PRIME_*/SYS_PRIME_*
' листы (data_access_layer.only_store_layer_directly_accesses_hidden_db_sheets). Все остальные
' модули работают через эти функции, а не через getCellByPosition на системных листах напрямую.
'
' Никаких commit/rollback/DDL/store() здесь нет - это чистый доступ к диапазонам Calc
' (performance.preferred_range_api: getDataArray/setDataArray, батчами).

Private gHeaderCache As Collection      ' key: имя листа -> Variant(массив заголовков)
Private gCommittedKeyCache As Collection ' key: SOURCE_KEY -> DOC_ID, только для COMMITTED

Private Function PRIME_Doc() As Object
    PRIME_Doc = ThisComponent
End Function

Public Function PRIME_GetSheet(ByVal sheetName As String) As Object
    Dim oSheets As Object
    oSheets = PRIME_Doc().getSheets()
    If Not oSheets.hasByName(sheetName) Then
        Err.Raise(1001, "PRIME_Store.PRIME_GetSheet", "PRIME_Store: лист не найден: " & sheetName)
    End If
    PRIME_GetSheet = oSheets.getByName(sheetName)
End Function

Public Function PRIME_SheetExists(ByVal sheetName As String) As Boolean
    PRIME_SheetExists = PRIME_Doc().getSheets().hasByName(sheetName)
End Function

' Последняя фактически использованная строка (0-based), -1 если лист пуст (только заголовок или ничего).
Public Function PRIME_FindLastRow(ByVal oSheet As Object) As Long
    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    Dim lastRow As Long
    lastRow = oCursor.RangeAddress.EndRow
    ' Пустой лист без данных: используемая область может быть единственной ячейкой A1 пустой
    If lastRow = 0 Then
        Dim v As Variant
        v = oSheet.getCellByPosition(0, 0).getString()
        If v = "" Then
            PRIME_FindLastRow = -1
            Exit Function
        End If
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
        gHeaderCache = New Collection
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

    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, 0, lastCol, 0)
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

' Читает весь лист (включая заголовок в строке 0) одним getDataArray.
Public Function PRIME_ReadTable(ByVal sheetName As String) As Variant
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim lastRow As Long, lastCol As Long
    lastRow = PRIME_FindLastRow(oSheet)
    lastCol = PRIME_FindLastCol(oSheet)
    If lastRow < 0 Then
        Dim empty1(0) As Variant
        PRIME_ReadTable = empty1
        Exit Function
    End If
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, 0, lastCol, lastRow)
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
    Dim startRow As Long
    startRow = PRIME_FindLastRow(oSheet) + 1
    If startRow < 1 Then startRow = 1 ' строка 0 - всегда заголовок

    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, startRow, colCount - 1, startRow + rowCount - 1)
    oRange.setDataArray(rows)
End Sub

' Перезаписывает существующие строки, начиная с startRow (0-based, включая возможность
' перезаписи строки заголовка при явном намерении - вызывающий отвечает за startRow >= 1
' для обычных данных).
Public Sub PRIME_UpdateRowsBatch(ByVal sheetName As String, ByVal startRow As Long, ByVal rows As Variant)
    Dim rowCount As Long
    rowCount = UBound(rows) - LBound(rows) + 1
    If rowCount <= 0 Then Exit Sub
    Dim colCount As Long
    colCount = UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, startRow, colCount - 1, startRow + rowCount - 1)
    oRange.setDataArray(rows)
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

Public Sub PRIME_RebuildCommittedKeyCache()
    Set gCommittedKeyCache = New Collection
    If Not PRIME_SheetExists(SH_SYS_TX) Then Exit Sub

    Dim tx As Variant
    tx = PRIME_ReadTable(SH_SYS_TX)
    If UBound(tx) < 1 Then Exit Sub

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim colSourceKey As Long, colDocId As Long, colState As Long
    colSourceKey = PRIME_ColIndex(headers, "SOURCE_KEY")
    colDocId = PRIME_ColIndex(headers, "DOC_ID")
    colState = PRIME_ColIndex(headers, "STATE")

    Dim i As Long
    For i = 1 To UBound(tx)
        If CStr(tx(i)(colState)) = TX_COMMITTED Then
            Dim k As String
            k = CStr(tx(i)(colSourceKey))
            If Not PRIME_CollectionHasKey(gCommittedKeyCache, k) Then
                gCommittedKeyCache.Add(CStr(tx(i)(colDocId)), k)
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

Public Sub PRIME_RegisterCommittedKey(ByVal sourceKey As String, ByVal docId As String)
    If gCommittedKeyCache Is Nothing Then
        PRIME_RebuildCommittedKeyCache()
    End If
    If Not PRIME_CollectionHasKey(gCommittedKeyCache, sourceKey) Then
        gCommittedKeyCache.Add(docId, sourceKey)
    End If
End Sub

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
