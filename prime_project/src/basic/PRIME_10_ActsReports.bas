Option Explicit

' PRIME_10_ActsReports
' Акты (ODT по DOC_ID из журнала, не из текущего выделения), отчёт руководителю, дашборд.
' writer_failure_must_not_rollback_stock_document: акт строится ПОСЛЕ проведения, ошибка Writer
' не трогает уже COMMITTED документ. false_success_message_allowed=false - при ошибке никогда
' не сообщаем об успехе.

' === Акты ========================================================================================
Public Sub PRIME_Acts_CreateFromDocButton()
    Dim docId As String
    docId = InputBox("DOC_ID документа, по которому создать акт:", "Создать акт")
    If Trim(docId) = "" Then Exit Sub
    PRIME_Acts_CreateAct(Trim(docId))
End Sub

Public Function PRIME_Acts_CreateAct(ByVal docId As String) As String
    On Error Goto Fail

    Dim docHeaders As Variant, lineHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim docTable As Variant, lineTable As Variant
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)

    Dim docIdx As Long
    docIdx = PRIME_FindRowByKey(docTable, PRIME_ColIndex(docHeaders, "DOC_ID"), docId)
    If docIdx = -1 Then
        MsgBox "Документ " & docId & " не найден в журнале - акт не создан."
        PRIME_Acts_CreateAct = ""
        Exit Function
    End If

    Dim oDesktop As Object
    oDesktop = createUnoService("com.sun.star.frame.Desktop")
    Dim oArgs(0) As New com.sun.star.beans.PropertyValue
    oArgs(0).Name = "Hidden"
    oArgs(0).Value = True
    Dim oDoc As Object
    oDoc = oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, oArgs())

    Dim oText As Object
    oText = oDoc.getText()
    Dim oCur As Object
    oCur = oText.createTextCursor()

    oCur.CharHeight = 14
    oCur.CharWeight = com.sun.star.awt.FontWeight.BOLD
    oText.insertString(oCur, "АКТ по документу " & docId, False)
    oText.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    oCur.CharWeight = com.sun.star.awt.FontWeight.NORMAL
    oCur.CharHeight = 11
    oText.insertString(oCur, "Тип: " & CStr(docTable(docIdx)(PRIME_ColIndex(docHeaders, "DOC_TYPE"))) & _
        "   Дата: " & CStr(docTable(docIdx)(PRIME_ColIndex(docHeaders, "DOC_DATE"))), False)
    oText.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    oText.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)

    Dim colDocId As Long
    colDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
    Dim lineCount As Long
    lineCount = 0
    If UBound(lineTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(lineTable)
            If CStr(lineTable(i)(colDocId)) = docId Then
                Dim productCode As String
                productCode = CStr(lineTable(i)(PRIME_ColIndex(lineHeaders, "PRODUCT_CODE")))
                lineCount = lineCount + 1
                Dim txt As String
                txt = lineCount & ". Код: " & productCode & "   " & PRIME_GetProductField(productCode, "PRODUCT_NAME") & _
                    "   Кол-во: " & CStr(lineTable(i)(PRIME_ColIndex(lineHeaders, "QTY_BASE"))) & " " & _
                    CStr(lineTable(i)(PRIME_ColIndex(lineHeaders, "UNIT")))
                oText.insertString(oCur, txt, False)
                oText.insertControlCharacter(oCur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
            End If
        Next i
    End If

    If lineCount = 0 Then
        oDoc.close(False)
        MsgBox "У документа " & docId & " нет строк - акт не создан."
        PRIME_Acts_CreateAct = ""
        Exit Function
    End If

    Dim actsDir As String
    actsDir = PRIME_EnsureDir(PRIME_DIR_ACTS)
    Dim fullPath As String
    fullPath = actsDir & "Act_" & docId & ".odt"

    Dim saveArgs(0) As New com.sun.star.beans.PropertyValue
    saveArgs(0).Name = "FilterName"
    saveArgs(0).Value = "writer8"
    oDoc.storeToURL(ConvertToURL(fullPath), saveArgs())
    oDoc.close(False)

    Dim actHeaders As Variant
    actHeaders = PRIME_HeaderMap(SH_DB_ACTS)
    Dim row(UBound(actHeaders)) As Variant
    row(PRIME_ColIndex(actHeaders, "ACT_ID")) = "ACT-" & Format(PRIME_SequenceNext("ACT_ID"), "00000000")
    row(PRIME_ColIndex(actHeaders, "DOC_ID")) = docId
    row(PRIME_ColIndex(actHeaders, "FILE_PATH")) = fullPath
    row(PRIME_ColIndex(actHeaders, "GENERATED_AT")) = Format(Now, "YYYY-MM-DD HH:MM:SS")
    row(PRIME_ColIndex(actHeaders, "STATUS")) = "OK"
    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_DB_ACTS, rows)

    MsgBox "Акт создан: " & fullPath
    PRIME_Acts_CreateAct = fullPath
    Exit Function

Fail:
    MsgBox "Акт НЕ создан из-за ошибки: " & Error$
    PRIME_Acts_CreateAct = ""
End Function

Public Sub PRIME_Acts_OpenByDocIdButton()
    Dim docId As String
    docId = InputBox("DOC_ID акта для открытия:", "Открыть акт")
    If Trim(docId) = "" Then Exit Sub

    Dim actHeaders As Variant
    actHeaders = PRIME_HeaderMap(SH_DB_ACTS)
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_ACTS)
    Dim idx As Long
    idx = PRIME_FindRowByKey(table, PRIME_ColIndex(actHeaders, "DOC_ID"), Trim(docId))
    If idx = -1 Then
        MsgBox "Акт для документа " & docId & " не найден. Создайте его повторно."
        Exit Sub
    End If

    Dim path As String
    path = CStr(table(idx)(PRIME_ColIndex(actHeaders, "FILE_PATH")))
    Dim oDesktop As Object
    oDesktop = createUnoService("com.sun.star.frame.Desktop")
    Dim noArgs()
    oDesktop.loadComponentFromURL(ConvertToURL(path), "_blank", 0, noArgs())
End Sub

' === Каталоги (пути относительно каталога документа) ===========================================
Public Function PRIME_DocDir() As String
    Dim p As String
    p = ConvertFromURL(ThisComponent.getURL())
    Dim i As Long
    i = PRIME_LastInStr(p, GetPathSeparator())
    PRIME_DocDir = Left(p, i)
End Function

Public Function PRIME_EnsureDir(ByVal relDirName As String) As String
    Dim full As String
    full = PRIME_DocDir() & relDirName & GetPathSeparator()
    If Not FileExists(full) Then MkDir(full)
    PRIME_EnsureDir = full
End Function

' === Дашборд =====================================================================================
' Явное обновление (explicit_refresh), видимая дата снимка, ошибка не публикует частичный снимок
' (failed_new_snapshot_must_not_replace_previous_successful_snapshot).
Public Sub PRIME_Dashboard_RefreshButton()
    On Error Goto Fail
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_DASHBOARD)

    Dim totalDocs As Long, receipts As Long, issues As Long, returnsCount As Long
    If PRIME_SheetExists(SH_DB_DOCUMENTS) Then
        Dim docHeaders As Variant
        docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
        Dim docTable As Variant
        docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
        Dim colType As Long
        colType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
        If UBound(docTable) >= 1 Then
            Dim i As Long
            For i = 1 To UBound(docTable)
                totalDocs = totalDocs + 1
                Select Case CStr(docTable(i)(colType))
                    Case DOC_RECEIPT : receipts = receipts + 1
                    Case DOC_ISSUE : issues = issues + 1
                    Case DOC_RETURN : returnsCount = returnsCount + 1
                End Select
            Next i
        End If
    End If

    Dim uniqueProducts As Long
    If PRIME_SheetExists(SH_DB_PRODUCTS) Then
        Dim prodTable As Variant
        prodTable = PRIME_ReadTable(SH_DB_PRODUCTS)
        If UBound(prodTable) >= 1 Then uniqueProducts = UBound(prodTable)
    End If

    ' Все ячейки вычислены успешно - публикуем снимок одним проходом.
    PRIME_Dashboard_SetLabelValue(oSheet, "Документов всего", CStr(totalDocs))
    PRIME_Dashboard_SetLabelValue(oSheet, "Приходов", CStr(receipts))
    PRIME_Dashboard_SetLabelValue(oSheet, "Выдач", CStr(issues))
    PRIME_Dashboard_SetLabelValue(oSheet, "Возвратов", CStr(returnsCount))
    PRIME_Dashboard_SetLabelValue(oSheet, "Товаров в каталоге", CStr(uniqueProducts))
    PRIME_Dashboard_SetLabelValue(oSheet, "Снимок обновлён", Format(Now, "YYYY-MM-DD HH:MM:SS"))
    Exit Sub

Fail:
    MsgBox "Обновление дашборда прервано ошибкой (" & Error$ & "). Предыдущий снимок не изменён."
End Sub

Private Sub PRIME_Dashboard_SetLabelValue(ByVal oSheet As Object, ByVal label As String, ByVal value As String)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 0 To lastRow
        If oSheet.getCellByPosition(0, r).getString() = label Then
            oSheet.getCellByPosition(1, r).setString(value)
            Exit Sub
        End If
    Next r
    Dim newRow As Long
    newRow = lastRow + 1
    oSheet.getCellByPosition(0, newRow).setString(label)
    oSheet.getCellByPosition(1, newRow).setString(value)
End Sub

' === Отчёт руководителю ==========================================================================
' rows_are_not_documents: считаем документы по DB_PRIME_DOCUMENTS, а не строки листа ввода.
' saved_report_is_snapshot: "Сохранить" фиксирует момент, последующие изменения склада его не переписывают.
Public Sub PRIME_Report_RefreshFactsButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_REPORT_INPUT)

    Dim receipts As Long, issues As Long, returnsCount As Long
    If PRIME_SheetExists(SH_DB_DOCUMENTS) Then
        Dim docHeaders As Variant
        docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
        Dim docTable As Variant
        docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
        Dim colType As Long
        colType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
        If UBound(docTable) >= 1 Then
            Dim i As Long
            For i = 1 To UBound(docTable)
                Select Case CStr(docTable(i)(colType))
                    Case DOC_RECEIPT : receipts = receipts + 1
                    Case DOC_ISSUE : issues = issues + 1
                    Case DOC_RETURN : returnsCount = returnsCount + 1
                End Select
            Next i
        End If
    End If

    PRIME_Dashboard_SetLabelValue(oSheet, "Приёмок (документов)", CStr(receipts))
    PRIME_Dashboard_SetLabelValue(oSheet, "Выдач (документов)", CStr(issues))
    PRIME_Dashboard_SetLabelValue(oSheet, "Возвратов (документов)", CStr(returnsCount))
End Sub

Public Sub PRIME_Report_BuildButton()
    Dim oInput As Object
    oInput = PRIME_GetSheet(SH_REPORT_INPUT)
    Dim oOutput As Object
    oOutput = PRIME_GetSheet(SH_REPORT_FINAL)

    Dim reportDate As String
    reportDate = PRIME_Dashboard_GetLabelValue(oInput, "Дата отчёта")
    If reportDate = "" Then reportDate = Format(Now, "YYYY-MM-DD")
    Dim author As String
    author = PRIME_Dashboard_GetLabelValue(oInput, "Автор")

    Dim lines As String
    lines = "ОТЧЁТ РУКОВОДИТЕЛЮ" & Chr(10)
    lines = lines & "Дата: " & reportDate & "   Автор: " & author & Chr(10) & Chr(10)
    lines = lines & "Приёмок: " & PRIME_Dashboard_GetLabelValue(oInput, "Приёмок (документов)") & Chr(10)
    lines = lines & "Выдач: " & PRIME_Dashboard_GetLabelValue(oInput, "Выдач (документов)") & Chr(10)
    lines = lines & "Возвратов: " & PRIME_Dashboard_GetLabelValue(oInput, "Возвратов (документов)") & Chr(10)

    oOutput.getCellByPosition(0, 0).setString(lines)
    MsgBox "Черновик отчёта сформирован на листе """ & SH_REPORT_FINAL & """. Проверьте и нажмите ""Сохранить отчёт""."
End Sub

Public Sub PRIME_Report_SaveButton()
    Dim oOutput As Object
    oOutput = PRIME_GetSheet(SH_REPORT_FINAL)
    Dim current As String
    current = oOutput.getCellByPosition(0, 0).getString()
    oOutput.getCellByPosition(0, 1).setString("Сохранено (снимок): " & Format(Now, "YYYY-MM-DD HH:MM:SS"))
    MsgBox "Отчёт сохранён как снимок. Дальнейшие изменения склада не переписывают этот текст автоматически."
End Sub

Private Function PRIME_Dashboard_GetLabelValue(ByVal oSheet As Object, ByVal label As String) As String
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 0 To lastRow
        If oSheet.getCellByPosition(0, r).getString() = label Then
            PRIME_Dashboard_GetLabelValue = oSheet.getCellByPosition(1, r).getString()
            Exit Function
        End If
    Next r
    PRIME_Dashboard_GetLabelValue = ""
End Function
