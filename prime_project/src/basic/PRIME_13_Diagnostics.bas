Option Explicit

' PRIME_13_Diagnostics
' Действительно read-only (в 1.4.1 диагностика была помечена "только чтение", но вызывала общий
' WMS_Close с commit/store - store_allowed=false, modify_data_allowed=false здесь буквально:
' эта функция никогда не пишет ни в один системный лист и не вызывает ThisComponent.store().

Public Sub PRIME_Diagnostics_RunButton()
    Dim report As String
    report = "ДИАГНОСТИКА PRIME " & Format(Now, "YYYY-MM-DD HH:MM:SS") & Chr(10) & Chr(10)

    Dim problems As Long
    problems = 0

    report = report & PRIME_Check_DuplicateProductCodes(problems)
    report = report & PRIME_Check_DuplicateDocIds(problems)
    report = report & PRIME_Check_DuplicateMoveIds(problems)
    report = report & PRIME_Check_MovementsWithoutProduct(problems)
    report = report & PRIME_Check_LotsWithoutProduct(problems)
    report = report & PRIME_Check_AllocationsWithoutIssueOrLot(problems)
    report = report & PRIME_Check_NegativeStock(problems)
    report = report & PRIME_Check_ReturnGreaterThanIssue(problems)
    report = report & PRIME_Check_PreparedTransactions(problems)
    report = report & PRIME_Check_UnitMismatches(problems)
    report = report & PRIME_Check_KitIntegrity(problems)
    report = report & PRIME_Check_DanglingDocLines(problems)
    report = report & PRIME_Check_DocumentsWithoutLines(problems)
    report = report & PRIME_Check_DuplicateCommittedSourceKey(problems)

    report = report & Chr(10) & "ИТОГО проблем: " & problems

    PRIME_Diagnostics_WriteToSheet(report)
    MsgBox report
End Sub

Private Sub PRIME_Diagnostics_WriteToSheet(ByVal report As String)
    On Error Resume Next ' диагностика не должна упасть, если лист недоступен - но и не пишет в DB_PRIME_*
    If Not PRIME_SheetExists(SH_DIAGNOSTICS) Then Exit Sub
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_DIAGNOSTICS)
    oSheet.getCellByPosition(0, 0).setString(report)
End Sub

' --- duplicate PRODUCT_CODE ---
Private Function PRIME_Check_DuplicateProductCodes(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_PRODUCTS) Then
        PRIME_Check_DuplicateProductCodes = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "PRODUCT_CODE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_PRODUCTS)
    Dim dupCount As Long
    dupCount = PRIME_CountDuplicates(table, colCode)
    If dupCount > 0 Then
        result = "[ПРОБЛЕМА] Дублирующихся PRODUCT_CODE: " & dupCount & Chr(10)
        problems = problems + dupCount
    End If
    PRIME_Check_DuplicateProductCodes = result
End Function

Private Function PRIME_Check_DuplicateDocIds(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_DOCUMENTS) Then
        PRIME_Check_DuplicateDocIds = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_DOCUMENTS)
    Dim dupCount As Long
    dupCount = PRIME_CountDuplicates(table, PRIME_ColIndex(headers, "DOC_ID"))
    If dupCount > 0 Then
        result = "[ПРОБЛЕМА] Дублирующихся DOC_ID: " & dupCount & Chr(10)
        problems = problems + dupCount
    End If
    PRIME_Check_DuplicateDocIds = result
End Function

Private Function PRIME_Check_DuplicateMoveIds(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        PRIME_Check_DuplicateMoveIds = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    Dim dupCount As Long
    dupCount = PRIME_CountDuplicates(table, PRIME_ColIndex(headers, "MOVE_ID"))
    If dupCount > 0 Then
        result = "[ПРОБЛЕМА] Дублирующихся MOVE_ID: " & dupCount & Chr(10)
        problems = problems + dupCount
    End If
    PRIME_Check_DuplicateMoveIds = result
End Function

Private Function PRIME_CountDuplicates(ByVal table As Variant, ByVal keyCol As Long) As Long
    If UBound(table) < 2 Then
        PRIME_CountDuplicates = 0
        Exit Function
    End If
    Dim seen() As String
    ReDim seen(UBound(table))
    Dim seenCount As Long
    seenCount = 0
    Dim dupCount As Long
    dupCount = 0
    Dim i As Long, j As Long, isDup As Boolean
    For i = 1 To UBound(table)
        Dim k As String
        k = CStr(table(i)(keyCol))
        isDup = False
        For j = 0 To seenCount - 1
            If seen(j) = k Then
                isDup = True
                Exit For
            End If
        Next j
        If isDup Then
            dupCount = dupCount + 1
        Else
            seen(seenCount) = k
            seenCount = seenCount + 1
        End If
    Next i
    PRIME_CountDuplicates = dupCount
End Function

Private Function PRIME_Check_MovementsWithoutProduct(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        PRIME_Check_MovementsWithoutProduct = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colProduct As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    Dim cnt As Long
    cnt = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If Not PRIME_ProductExists(CStr(table(i)(colProduct))) Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Движений без существующего товара: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_MovementsWithoutProduct = result
End Function

Private Function PRIME_Check_LotsWithoutProduct(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_LOTS) Then
        PRIME_Check_LotsWithoutProduct = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_LOTS)
    Dim colProduct As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_LOTS)
    Dim cnt As Long
    cnt = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If Not PRIME_ProductExists(CStr(table(i)(colProduct))) Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Партий без существующего товара: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_LotsWithoutProduct = result
End Function

Private Function PRIME_Check_AllocationsWithoutIssueOrLot(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_ALLOCATIONS) Then
        PRIME_Check_AllocationsWithoutIssueOrLot = ""
        Exit Function
    End If
    Dim allocHeaders As Variant
    allocHeaders = PRIME_HeaderMap(SH_DB_ALLOCATIONS)
    Dim allocTable As Variant
    allocTable = PRIME_ReadTable(SH_DB_ALLOCATIONS)
    Dim lineHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim lineTable As Variant
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim lotTable As Variant
    lotTable = PRIME_ReadTable(SH_DB_LOTS)

    Dim cnt As Long
    cnt = 0
    If UBound(allocTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(allocTable)
            Dim lineId As String, lotId As String
            lineId = CStr(allocTable(i)(PRIME_ColIndex(allocHeaders, "DOC_LINE_ID")))
            lotId = CStr(allocTable(i)(PRIME_ColIndex(allocHeaders, "LOT_ID")))
            If PRIME_FindRowByKey(lineTable, PRIME_ColIndex(lineHeaders, "DOC_LINE_ID"), lineId) = -1 Then cnt = cnt + 1
            If PRIME_FindRowByKey(lotTable, PRIME_ColIndex(lotHeaders, "LOT_ID"), lotId) = -1 Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Allocations без строки выдачи/партии: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_AllocationsWithoutIssueOrLot = result
End Function

' aggregate_quantities_separately_by_unit: считаем отрицательный остаток отдельно по каждой единице.
Private Function PRIME_Check_NegativeStock(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_LOTS) Then
        PRIME_Check_NegativeStock = ""
        Exit Function
    End If
    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim lotTable As Variant
    lotTable = PRIME_ReadTable(SH_DB_LOTS)
    Dim colProduct As Long
    colProduct = PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")

    Dim seen() As String
    ReDim seen(UBound(lotTable))
    Dim seenCount As Long
    seenCount = 0
    Dim cnt As Long
    cnt = 0

    If UBound(lotTable) >= 1 Then
        Dim i As Long, j As Long, already As Boolean
        For i = 1 To UBound(lotTable)
            Dim code As String
            code = CStr(lotTable(i)(colProduct))
            already = False
            For j = 0 To seenCount - 1
                If seen(j) = code Then already = True : Exit For
            Next j
            If Not already Then
                seen(seenCount) = code
                seenCount = seenCount + 1
                If PRIME_TotalLotBalance(code) < -0.0000005 Then cnt = cnt + 1
            End If
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Товаров с отрицательным суммарным остатком: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_NegativeStock = result
End Function

Private Function PRIME_Check_ReturnGreaterThanIssue(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_RETURNS) Then
        PRIME_Check_ReturnGreaterThanIssue = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_RETURNS)
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_RETURNS)
    Dim colOrig As Long
    colOrig = PRIME_ColIndex(headers, "ORIGINAL_ISSUE_DOC_LINE_ID")

    Dim cnt As Long
    cnt = 0
    Dim checked() As String
    ReDim checked(UBound(table))
    Dim checkedCount As Long
    checkedCount = 0

    If UBound(table) >= 1 Then
        Dim i As Long, j As Long, already As Boolean
        For i = 1 To UBound(table)
            Dim origId As String
            origId = CStr(table(i)(colOrig))
            already = False
            For j = 0 To checkedCount - 1
                If checked(j) = origId Then already = True : Exit For
            Next j
            If Not already Then
                checked(checkedCount) = origId
                checkedCount = checkedCount + 1
                If PRIME_AlreadyReturnedQtyBase(origId) > PRIME_DocLineQtyBase(origId) + 0.0000005 Then
                    cnt = cnt + 1
                End If
            End If
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Строк выдачи с возвратом больше выданного: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_ReturnGreaterThanIssue = result
End Function

' crash_recovery: PREPARED без COMMITTED виден в диагностике, но не доводится автоматически.
Private Function PRIME_Check_PreparedTransactions(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_SYS_TX) Then
        PRIME_Check_PreparedTransactions = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_TX)
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "STATE")

    Dim cnt As Long
    cnt = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If CStr(table(i)(colState)) = TX_PREPARED Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ВНИМАНИЕ] Незавершённых (PREPARED) операций: " & cnt & " - не учтены в остатке, требуют ручной проверки." & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_PreparedTransactions = result
End Function

Private Function PRIME_Check_UnitMismatches(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_DOC_LINES) Then
        PRIME_Check_UnitMismatches = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim colProduct As Long, colUnit As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colUnit = PRIME_ColIndex(headers, "UNIT")

    Dim cnt As Long
    cnt = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            Dim factor As Variant
            factor = PRIME_GetUnitFactor(CStr(table(i)(colProduct)), CStr(table(i)(colUnit)))
            If IsEmpty(factor) Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Строк документов с единицей без коэффициента пересчёта: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_UnitMismatches = result
End Function

' dangling_line (recommendation §43): строка документа, чей DOC_ID не существует в
' DB_PRIME_DOCUMENTS - не должно происходить при нормальной работе единого posting engine
' (PRIME_WriteDocumentHeader пишет шапку до строк), но проверяется отдельно на случай ручного
' редактирования скрытых листов или незавершённой миграции.
Private Function PRIME_Check_DanglingDocLines(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_DOC_LINES) Or Not PRIME_SheetExists(SH_DB_DOCUMENTS) Then
        PRIME_Check_DanglingDocLines = ""
        Exit Function
    End If
    Dim lineHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim lineTable As Variant
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim docHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim docTable As Variant
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
    Dim colLineDocId As Long, colDocDocId As Long
    colLineDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
    colDocDocId = PRIME_ColIndex(docHeaders, "DOC_ID")

    Dim cnt As Long
    cnt = 0
    If UBound(lineTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(lineTable)
            If PRIME_FindRowByKey(docTable, colDocDocId, CStr(lineTable(i)(colLineDocId))) = -1 Then
                cnt = cnt + 1
            End If
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Строк документов со ссылкой на несуществующий DOC_ID: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_DanglingDocLines = result
End Function

' document_without_lines (recommendation §43): документ без единой строки - зафиксированная
' операция обязана иметь хотя бы одну строку (PRIME_ValidateAndExpandPlan отклоняет пустой
' план ДО записи, поэтому это тоже "не должно происходить", а не ожидаемое состояние).
Private Function PRIME_Check_DocumentsWithoutLines(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_DB_DOCUMENTS) Or Not PRIME_SheetExists(SH_DB_DOC_LINES) Then
        PRIME_Check_DocumentsWithoutLines = ""
        Exit Function
    End If
    Dim docHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim docTable As Variant
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)
    Dim lineHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim lineTable As Variant
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim colDocDocId As Long, colLineDocId As Long
    colDocDocId = PRIME_ColIndex(docHeaders, "DOC_ID")
    colLineDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")

    Dim cnt As Long
    cnt = 0
    If UBound(docTable) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(docTable)
            If PRIME_FindRowByKey(lineTable, colLineDocId, CStr(docTable(i)(colDocDocId))) = -1 Then
                cnt = cnt + 1
            End If
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Документов без единой строки: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_DocumentsWithoutLines = result
End Function

' Дубли SOURCE_KEY среди COMMITTED-транзакций - при нормальной работе идемпотентности
' (PRIME_FindCommittedBySourceKey проверяется ДО записи) не должно возникать вовсе; проверка
' нужна, чтобы обнаружить ручную порчу SYS_PRIME_TX или скрытый дефект идемпотентности.
Private Function PRIME_Check_DuplicateCommittedSourceKey(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_SYS_TX) Then
        PRIME_Check_DuplicateCommittedSourceKey = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim table As Variant
    table = PRIME_ReadTable(SH_SYS_TX)
    Dim colSourceKey As Long, colState As Long
    colSourceKey = PRIME_ColIndex(headers, "SOURCE_KEY")
    colState = PRIME_ColIndex(headers, "STATE")

    Dim seen() As String
    ReDim seen(UBound(table))
    Dim seenCount As Long
    seenCount = 0
    Dim cnt As Long
    cnt = 0

    If UBound(table) >= 1 Then
        Dim i As Long, j As Long, isDup As Boolean
        For i = 1 To UBound(table)
            If CStr(table(i)(colState)) = TX_COMMITTED Then
                Dim k As String
                k = CStr(table(i)(colSourceKey))
                isDup = False
                For j = 0 To seenCount - 1
                    If seen(j) = k Then isDup = True : Exit For
                Next j
                If isDup Then
                    cnt = cnt + 1
                Else
                    seen(seenCount) = k
                    seenCount = seenCount + 1
                End If
            End If
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Повторных COMMITTED SOURCE_KEY в SYS_PRIME_TX: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_DuplicateCommittedSourceKey = result
End Function

Private Function PRIME_Check_KitIntegrity(ByRef problems As Long) As String
    Dim result As String
    result = ""
    If Not PRIME_SheetExists(SH_KITS) Then
        PRIME_Check_KitIntegrity = ""
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_KITS)
    Dim table As Variant
    table = PRIME_ReadTable(SH_KITS)
    Dim colProduct As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")

    Dim cnt As Long
    cnt = 0
    If UBound(table) >= 1 Then
        Dim i As Long
        For i = 1 To UBound(table)
            If Not PRIME_ProductExists(CStr(table(i)(colProduct))) Then cnt = cnt + 1
        Next i
    End If
    If cnt > 0 Then
        result = "[ПРОБЛЕМА] Компонентов комплектов со ссылкой на несуществующий товар: " & cnt & Chr(10)
        problems = problems + cnt
    End If
    PRIME_Check_KitIntegrity = result
End Function
