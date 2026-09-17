Option Explicit

' PRIME_16_Journal (новый в 2.1.0, R25)
' Лист "Журнал" - read-only список ТОЛЬКО COMMITTED документов (committed_only_everywhere),
' обновляется по кнопке (batch_output), без построчного пересчёта на каждое открытие книги.

Public Sub PRIME_Journal_RefreshButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_JOURNAL)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_JOURNAL)

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    PRIME_ClearDataRows(oSheet, headers)

    If Not PRIME_SheetExists(SH_DB_DOCUMENTS) Or Not PRIME_SheetExists(SH_SYS_TX) Then Exit Sub

    Dim docHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim docTable As Variant
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
    Dim colDocId As Long, colDocType As Long, colDocDate As Long, colSourceSheet As Long, colOpId As Long
    colDocId = PRIME_ColIndex(docHeaders, "DOC_ID")
    colDocType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
    colDocDate = PRIME_ColIndex(docHeaders, "DOC_DATE")
    colSourceSheet = PRIME_ColIndex(docHeaders, "SOURCE_SHEET")
    colOpId = PRIME_ColIndex(docHeaders, "OP_ID")

    Dim txHeaders As Variant
    Dim hasTxOpId As Boolean
    txHeaders = PRIME_HeaderMap(SH_SYS_TX)
    Dim txTable As Variant
    txTable = PRIME_ReadTable(SH_SYS_TX)
    Dim colTxOpId As Long, colTxCommittedAt As Long
    colTxOpId = PRIME_ColIndex(txHeaders, "OP_ID")
    colTxCommittedAt = PRIME_ColIndex(txHeaders, "COMMITTED_AT")

    Dim lineHeaders As Variant
    Dim hasLines As Boolean
    hasLines = PRIME_SheetExists(SH_DB_DOC_LINES)
    Dim lineTable As Variant
    Dim colLineDocId As Long, colRecipient As Long
    If hasLines Then
        lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
        lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
        colLineDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
        colRecipient = PRIME_ColIndex(lineHeaders, "RECIPIENT")
    End If

    Dim outRows() As Variant
    ReDim outRows(200)
    Dim n As Long
    n = 0

    If UBound(docTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(docTable)
            Dim docId As String
            docId = CStr(docTable(i)(colDocId))
            Dim opId As String
            opId = ""
            If colOpId >= 0 Then opId = CStr(docTable(i)(colOpId))
            ' committed_only_everywhere (R07): DOC_ID должен быть реально COMMITTED - иначе
            ' документ, физически записанный до отката на FAILED, попал бы в журнал как рабочий.
            If Not PRIME_IsDocIdCommitted(docId) Then GoTo ContinueDoc

            Dim sourceSheet As String
            sourceSheet = CStr(docTable(i)(colSourceSheet))
            Dim committedAt As String
            committedAt = ""
            If opId <> "" And colTxOpId >= 0 And UBound(txTable) >= 1 Then
                Dim txIdx As Long
                txIdx = PRIME_FindRowByKey(txTable, colTxOpId, opId)
                If txIdx >= 0 And colTxCommittedAt >= 0 Then committedAt = CStr(txTable(txIdx)(colTxCommittedAt))
            End If

            Dim lineCount As Long
            Dim recipient As String
            lineCount = 0
            recipient = ""
            If hasLines And UBound(lineTable) >= 1 Then
                Dim j As Long
                For j = 1 To UBound(lineTable)
                    If CStr(lineTable(j)(colLineDocId)) = docId Then
                        lineCount = lineCount + 1
                        If recipient = "" Then recipient = CStr(lineTable(j)(colRecipient))
                    End If
                Next j
            End If

            Dim row(UBound(headers)) As Variant
            row(PRIME_ColIndex(headers, "DOC_ID")) = docId
            row(PRIME_ColIndex(headers, "OP_ID")) = opId
            row(PRIME_ColIndex(headers, "Тип")) = CStr(docTable(i)(colDocType))
            row(PRIME_ColIndex(headers, "Дата/время")) = IIf(committedAt <> "", committedAt, CStr(docTable(i)(colDocDate)))
            row(PRIME_ColIndex(headers, "Статус")) = "COMMITTED"
            row(PRIME_ColIndex(headers, "Источник")) = sourceSheet
            row(PRIME_ColIndex(headers, "Контур")) = PRIME_ContourDisplayName(PRIME_ContourForSheet(sourceSheet))
            row(PRIME_ColIndex(headers, "Поставщик/Получатель")) = recipient
            row(PRIME_ColIndex(headers, "Количество строк")) = lineCount
            If n > UBound(outRows) Then ReDim Preserve outRows(UBound(outRows) + 200)
            outRows(n) = row
            n = n + 1
ContinueDoc:
        Next i
    End If

    If n > 0 Then
        ReDim Preserve outRows(n - 1)
        PRIME_AppendRowsBatch(SH_JOURNAL, outRows)
    End If
End Sub
