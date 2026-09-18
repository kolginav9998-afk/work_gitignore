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
    Dim colLineDocId As Long, colRecipient As Long, colLineLocTo As Long
    If hasLines Then
        lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
        lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
        colLineDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
        colRecipient = PRIME_ColIndex(lineHeaders, "RECIPIENT")
        colLineLocTo = PRIME_ColIndex(lineHeaders, "LOCATION_TO")
    End If

    ' headless_buffered_array_readback_corruption (2.1.2, см. PRIME_04_Posting.PRIME_PostIssueLines) -
    ' раньше строки накапливались в outRows()-буфер (Dim row() внутри цикла, растущий ReDim Preserve)
    ' и передавались одним PRIME_AppendRowsBatch; фикс - писать каждую строку сразу поячейково
    ' внутри цикла, без буферизации.
    Dim colOutDocId As Long, colOutOpId As Long, colOutType As Long, colOutDate As Long, colOutStatus As Long
    Dim colOutSource As Long, colOutDest As Long, colOutRecipient As Long, colOutLineCount As Long
    colOutDocId = PRIME_ColIndex(headers, "DOC_ID")
    colOutOpId = PRIME_ColIndex(headers, "OP_ID")
    colOutType = PRIME_ColIndex(headers, "Тип")
    colOutDate = PRIME_ColIndex(headers, "Дата/время")
    colOutStatus = PRIME_ColIndex(headers, "Статус")
    colOutSource = PRIME_ColIndex(headers, "Источник")
    colOutDest = PRIME_ColIndex(headers, "Куда")
    colOutRecipient = PRIME_ColIndex(headers, "Поставщик/Получатель")
    colOutLineCount = PRIME_ColIndex(headers, "Количество строк")

    Dim outRow As Long
    outRow = PRIME_FormSchemaFirstDataRow(SH_JOURNAL)
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
            Dim recipient As String, destination As String
            lineCount = 0
            recipient = ""
            destination = ""
            If hasLines And UBound(lineTable) >= 1 Then
                Dim j As Long
                For j = 1 To UBound(lineTable)
                    If CStr(lineTable(j)(colLineDocId)) = docId Then
                        lineCount = lineCount + 1
                        If recipient = "" Then recipient = CStr(lineTable(j)(colRecipient))
                        If destination = "" And colLineLocTo >= 0 Then destination = CStr(lineTable(j)(colLineLocTo))
                    End If
                Next j
            End If

            oSheet.getCellByPosition(colOutDocId, outRow).setString(docId)
            oSheet.getCellByPosition(colOutOpId, outRow).setString(opId)
            oSheet.getCellByPosition(colOutType, outRow).setString(CStr(docTable(i)(colDocType)))
            oSheet.getCellByPosition(colOutDate, outRow).setString(IIf(committedAt <> "", committedAt, CStr(docTable(i)(colDocDate))))
            oSheet.getCellByPosition(colOutStatus, outRow).setString("COMMITTED")
            oSheet.getCellByPosition(colOutSource, outRow).setString(sourceSheet)
            oSheet.getCellByPosition(colOutDest, outRow).setString(destination)
            oSheet.getCellByPosition(colOutRecipient, outRow).setString(recipient)
            oSheet.getCellByPosition(colOutLineCount, outRow).setValue(lineCount)
            outRow = outRow + 1
            n = n + 1
ContinueDoc:
        Next i
    End If
End Sub
