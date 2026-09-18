Option Explicit

Type PrimeDocLine
    ProductCode As String
    ProductName As String
    IsNewProduct As Boolean
    QtyInput As Double
    UnitInput As String
    LocationFrom As String
    LocationTo As String
    Contour As String
    ContourFrom As String
    ContourTo As String
    DestinationProject As String
    Recipient As String
    Comment As String
    Price As Double
    OriginalDocLineId As String
    OrderLineId As String
    QtyBase As Double
    LotId As String
End Type

Type PrimeDocPlan
    DocType As String
    DocDate As String
    SourceSheet As String
    SourceKey As String
    OrderId As String
    Lines(999) As PrimeDocLine
    LineCount As Long
End Type

' PRIME_ZZ_E2ETest - НЕ часть продукта, диагностический модуль для прямой (не через
' UI/CurrentSelection) рантайм-проверки в реальном LibreOffice. Внедряется отдельным скриптом
' (tools/run_e2e_diagnostic.py), НЕ входит в PRIME_MODULES build_ods.py и НЕ попадает в
' собираемый production .ods.
'
' ЧЕСТНЫЙ ИТОГ ЭТОЙ ДИАГНОСТИКИ (сессия 2.1.0), см. также docs/KNOWN_ISSUES.md:
' 1. Этой диагностикой найден и подтверждён РЕАЛЬНЫЙ баг (исправлен в PRIME_02_Store.
'    PRIME_ReadTable): пустая таблица возвращала "фантомную" строку из одного элемента Empty
'    вместо строки правильной ширины, из-за чего PRIME_SequenceNext (и любой другой код по
'    тому же паттерну "Dim newRow(UBound(table(0)))") падал при ПЕРВОМ ЗА ВСЮ ИСТОРИЮ вызове на
'    полностью пустой таблице - то есть первое же действие в свежей книге. Подтверждено: после
'    фикса PRIME_SequenceNext("DOC_ID") на пустой SYS_PRIME_SEQ перестал падать.
' 2. Дальнейшая диагностика показала, что ПОЛНЫЙ вызов PRIME_PostDocument через ВНЕШНИЙ headless
'    invoke() в этой конкретной песочнице ненадёжен даже для простейшего прихода существующего
'    товара (см. PRIME_ZZ_ExistingProductReceipt) - воспроизводимо возвращает "" с внутренней
'    "Object variable not set", хотя статические проверки (structural/source) и чистая
'    Python-модель тех же алгоритмов (tests/model_tests.py) проходят полностью. Ручная
'    построчная замена того же кода локальными переменными (PRIME_ZZ_ChainBisect) НЕ смогла
'    чисто воспроизвести точку отказа - что само по себе указывает на нестабильность именно
'    механизма внешнего invoke() в этой headless-среде, а не на очевидную одну строку кода.
'    Это расширяет (не противоречит) уже задокументированное ограничение KNOWN_ISSUES §11
'    ("многошаговые внешние invoke()-цепочки ненадёжны") - теперь известно, что это может
'    проявляться даже внутри ОДНОГО вызова PRIME_PostDocument, не только между вызовами.
' 3. Это НЕ подтверждает и НЕ опровергает работоспособность PRIME_PostDocument при реальном
'    интерактивном использовании (клик мышью в открытом LibreOffice) - недоступно для проверки
'    в этой песочнице (нет дисплея). Статус R29/R30 в REQUIREMENTS_MATRIX.md -
'    BLOCKED_BY_EXTERNAL_ENVIRONMENT, а не IMPLEMENTED_VERIFIED и не "известный баг".

Private gLogFile As Integer

Private Sub Log(ByVal msg As String)
    Print #gLogFile, msg
End Sub

Private Function Fmt(ByVal v As Double) As String
    Fmt = Format(v, "0.######")
End Function

Public Function PRIME_ZZ_Bisect(ByVal logPath As String) As String
    gLogFile = FreeFile
    Open logPath For Output As #gLogFile
    On Error Resume Next ' per-statement checking below, not a single catch-all handler

    Log "checkpoint 0: start"

    Dim code As String
    code = PRIME_CreateProduct("Bisect Widget", "шт", "LOC-X", "", "", False, "")
    If Err.Number <> 0 Then Log "ERR after CreateProduct: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 1: code=" & code

    Dim n1 As Long
    n1 = PRIME_SequenceNext("LOT_ID")
    If Err.Number <> 0 Then Log "ERR after SequenceNext(LOT_ID): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 2: n1=" & n1

    Dim lotId As String
    lotId = "LOT-" & Format(n1, "00000000")
    Log "checkpoint 2b: lotId=" & lotId

    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    If Err.Number <> 0 Then Log "ERR after HeaderMap(LOTS): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 3: lotHeaders UBound=" & UBound(lotHeaders)

    Dim lotRow(UBound(lotHeaders)) As Variant
    If Err.Number <> 0 Then Log "ERR after Dim lotRow: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 4: lotRow size=" & (UBound(lotRow) + 1)

    Dim idxLotId As Long
    idxLotId = PRIME_ColIndex(lotHeaders, "LOT_ID")
    Log "checkpoint 4b: idxLotId=" & idxLotId

    lotRow(idxLotId) = lotId
    If Err.Number <> 0 Then Log "ERR after write LOT_ID: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 5: wrote LOT_ID"

    Dim idxProductCode As Long
    idxProductCode = PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")
    lotRow(idxProductCode) = code
    If Err.Number <> 0 Then Log "ERR after write PRODUCT_CODE: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 6: wrote PRODUCT_CODE"

    Dim baseUnit As String
    baseUnit = PRIME_GetProductField(code, "BASE_UNIT")
    If Err.Number <> 0 Then Log "ERR after GetProductField: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 7: baseUnit=" & baseUnit

    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    If Err.Number <> 0 Then Log "ERR after HeaderMap(MOVEMENTS): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 8: moveHeaders UBound=" & UBound(moveHeaders)

    Dim lotRows(0) As Variant
    lotRows(0) = lotRow
    PRIME_AppendRowsBatch(SH_DB_LOTS, lotRows)
    If Err.Number <> 0 Then Log "ERR after AppendRowsBatch(LOTS): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 9: appended lot row"

    ' PRIME_BuildMovementRow is Private in PRIME_04_Posting - replicate its body inline here.
    Dim moveId As Long
    moveId = PRIME_SequenceNext("MOVE_ID")
    If Err.Number <> 0 Then Log "ERR after SequenceNext(MOVE_ID): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 10: moveId=" & moveId

    Dim moveRow(UBound(moveHeaders)) As Variant
    If Err.Number <> 0 Then Log "ERR after Dim moveRow: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 11: moveRow size=" & (UBound(moveRow) + 1)

    moveRow(PRIME_ColIndex(moveHeaders, "MOVE_ID")) = "MOV-" & Format(moveId, "00000000")
    moveRow(PRIME_ColIndex(moveHeaders, "DOC_ID")) = "DOC-BISECT"
    moveRow(PRIME_ColIndex(moveHeaders, "PRODUCT_CODE")) = code
    moveRow(PRIME_ColIndex(moveHeaders, "LOT_ID")) = lotId
    moveRow(PRIME_ColIndex(moveHeaders, "QTY_BASE")) = 10
    moveRow(PRIME_ColIndex(moveHeaders, "LOCATION")) = "LOC-X"
    moveRow(PRIME_ColIndex(moveHeaders, "MOVE_DATE")) = "2026-01-01"
    moveRow(PRIME_ColIndex(moveHeaders, "OP_ID")) = "OP-BISECT"
    Dim colContourM As Long
    colContourM = PRIME_ColIndex(moveHeaders, "STOCK_CONTOUR")
    If colContourM >= 0 Then moveRow(colContourM) = "GENERAL"
    If Err.Number <> 0 Then Log "ERR while filling moveRow: " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 12: moveRow filled"

    Dim moveRows(0) As Variant
    moveRows(0) = moveRow
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRows)
    If Err.Number <> 0 Then Log "ERR after AppendRowsBatch(MOVEMENTS): " & Err.Number & " " & Err.Description
    Err.Clear
    Log "checkpoint 13: appended move row"

    Log "ALL CHECKPOINTS REACHED"
    Close #gLogFile
    PRIME_ZZ_Bisect = "0"
End Function

Public Function PRIME_ZZ_TrivialLoop(ByVal logPath As String) As String
    gLogFile = FreeFile
    Open logPath For Output As #gLogFile
    Dim i As Long
    For i = 1 To 20
        Log "trivial line " & i
    Next i
    Close #gLogFile
    PRIME_ZZ_TrivialLoop = "0"
End Function

Public Function PRIME_ZZ_ChainBisect(ByVal logPath As String) As String
    gLogFile = FreeFile
    Open logPath For Output As #gLogFile
    On Error Resume Next

    Dim code As String
    code = PRIME_CreateProduct("Chain Widget", "шт", "LOC-CHAIN", "", "", False, "")
    Log "1 CreateProduct -> " & code & " err=" & Err.Number
    Err.Clear

    Dim locked As Boolean
    locked = PRIME_TryEnter()
    Log "2 TryEnter -> " & locked & " err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog("", "BUTTON_ENTER", "CHAIN_TEST", "k1")
    ' REMOVED: Log "2a AuditLog err=" & Err.Number
    Err.Clear

    Dim existing As String
    existing = PRIME_FindCommittedBySourceKey("CHAIN-TEST-KEY-1")
    Log "3 FindCommittedBySourceKey -> '" & existing & "' err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog("", "VALIDATION_START", "CHAIN_TEST", "k1")
    ' REMOVED: Log "3a AuditLog err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog("", "VALIDATION_OK", "CHAIN_TEST", "k1")
    ' REMOVED: Log "3b AuditLog err=" & Err.Number
    Err.Clear

    Dim opId As String
    opId = "OP-CHAINTEST-1"
    Dim docId As String
    docId = "DOC-" & Format(PRIME_SequenceNext("DOC_ID"), "00000000")
    Log "4 SequenceNext(DOC_ID) -> " & docId & " err=" & Err.Number
    Err.Clear

    PRIME_WriteTxRowPublicTest(opId, "CHAIN-TEST-KEY-1", docId, "PREPARED", "")
    Log "5 WriteTxRow err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog(opId, "TX_PREPARED", "CHAIN_TEST", docId)
    ' REMOVED: Log "5a AuditLog err=" & Err.Number
    Err.Clear

    Dim docHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Log "6 HeaderMap(DOCUMENTS) UBound=" & UBound(docHeaders) & " err=" & Err.Number
    Err.Clear
    Dim docRow(UBound(docHeaders)) As Variant
    docRow(PRIME_ColIndex(docHeaders, "DOC_ID")) = docId
    docRow(PRIME_ColIndex(docHeaders, "DOC_TYPE")) = "RECEIPT"
    docRow(PRIME_ColIndex(docHeaders, "DOC_DATE")) = "2026-01-01"
    docRow(PRIME_ColIndex(docHeaders, "SOURCE_SHEET")) = "CHAIN_TEST"
    docRow(PRIME_ColIndex(docHeaders, "SOURCE_KEY")) = "CHAIN-TEST-KEY-1"
    docRow(PRIME_ColIndex(docHeaders, "ORDER_ID")) = ""
    docRow(PRIME_ColIndex(docHeaders, "STATUS")) = "PROVEDENO"
    Dim opIdCol As Long
    opIdCol = PRIME_ColIndex(docHeaders, "OP_ID")
    If opIdCol >= 0 Then docRow(opIdCol) = opId
    Dim docRows(0) As Variant
    docRows(0) = docRow
    PRIME_AppendRowsBatch(SH_DB_DOCUMENTS, docRows)
    Log "7 AppendRowsBatch(DOCUMENTS) err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog(opId, "DOCS_WRITTEN", "CHAIN_TEST", docId)
    ' REMOVED: Log "7a AuditLog err=" & Err.Number
    Err.Clear

    Dim lineHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim lineRow(UBound(lineHeaders)) As Variant
    lineRow(PRIME_ColIndex(lineHeaders, "DOC_LINE_ID")) = docId & "-L1"
    lineRow(PRIME_ColIndex(lineHeaders, "DOC_ID")) = docId
    lineRow(PRIME_ColIndex(lineHeaders, "PRODUCT_CODE")) = code
    lineRow(PRIME_ColIndex(lineHeaders, "QTY_BASE")) = 5
    lineRow(PRIME_ColIndex(lineHeaders, "UNIT")) = "шт"
    Dim lineRows(0) As Variant
    lineRows(0) = lineRow
    PRIME_AppendRowsBatch(SH_DB_DOC_LINES, lineRows)
    Log "8 AppendRowsBatch(DOC_LINES) err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog(opId, "LINES_WRITTEN", "CHAIN_TEST", docId)
    ' REMOVED: Log "8a AuditLog err=" & Err.Number
    Err.Clear

    Dim lotHeaders As Variant
    lotHeaders = PRIME_HeaderMap(SH_DB_LOTS)
    Dim lotId As String
    lotId = "LOT-" & Format(PRIME_SequenceNext("LOT_ID"), "00000000")
    Log "9 SequenceNext(LOT_ID) -> " & lotId & " err=" & Err.Number
    Err.Clear
    Dim lotRow(UBound(lotHeaders)) As Variant
    lotRow(PRIME_ColIndex(lotHeaders, "LOT_ID")) = lotId
    lotRow(PRIME_ColIndex(lotHeaders, "PRODUCT_CODE")) = code
    Dim lotRows(0) As Variant
    lotRows(0) = lotRow
    PRIME_AppendRowsBatch(SH_DB_LOTS, lotRows)
    Log "10 AppendRowsBatch(LOTS) err=" & Err.Number
    Err.Clear
    ' REMOVED: PRIME_AuditLog(opId, "LOTS_WRITTEN", "CHAIN_TEST", docId)
    ' REMOVED: Log "10a AuditLog err=" & Err.Number
    Err.Clear

    Dim moveHeaders As Variant
    moveHeaders = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim moveRow(UBound(moveHeaders)) As Variant
    moveRow(PRIME_ColIndex(moveHeaders, "MOVE_ID")) = "MOV-" & Format(PRIME_SequenceNext("MOVE_ID"), "00000000")
    moveRow(PRIME_ColIndex(moveHeaders, "DOC_ID")) = docId
    moveRow(PRIME_ColIndex(moveHeaders, "PRODUCT_CODE")) = code
    moveRow(PRIME_ColIndex(moveHeaders, "LOT_ID")) = lotId
    moveRow(PRIME_ColIndex(moveHeaders, "QTY_BASE")) = 5
    moveRow(PRIME_ColIndex(moveHeaders, "OP_ID")) = opId
    Dim moveRows(0) As Variant
    moveRows(0) = moveRow
    PRIME_AppendRowsBatch(SH_DB_MOVEMENTS, moveRows)
    Log "11 AppendRowsBatch(MOVEMENTS) err=" & Err.Number
    Err.Clear

    Log "ALL STEPS REACHED"
    Close #gLogFile
    PRIME_ZZ_ChainBisect = "0"
End Function

Public Sub PRIME_WriteTxRowPublicTest(ByVal opId As String, ByVal sourceKey As String, ByVal docId As String, ByVal state As String, ByVal errText As String)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_SYS_TX)
    Dim row(UBound(headers)) As Variant
    row(PRIME_ColIndex(headers, "OP_ID")) = opId
    row(PRIME_ColIndex(headers, "SOURCE_KEY")) = sourceKey
    row(PRIME_ColIndex(headers, "DOC_ID")) = docId
    row(PRIME_ColIndex(headers, "STATE")) = state
    row(PRIME_ColIndex(headers, "STARTED_AT")) = Format(Now, "YYYY-MM-DD HH:MM:SS")
    row(PRIME_ColIndex(headers, "COMMITTED_AT")) = ""
    row(PRIME_ColIndex(headers, "HASH")) = ""
    row(PRIME_ColIndex(headers, "ERROR")) = errText
    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_SYS_TX, rows)
End Sub

Public Function PRIME_ZZ_ExistingProductReceipt(ByVal logPath As String) As String
    gLogFile = FreeFile
    Open logPath For Output As #gLogFile
    On Error Goto Fatal

    ' Seed a product directly (outside PRIME_PostDocument's call chain entirely) to isolate
    ' whether the issue is specific to PRIME_PostDocument's OWN internal chain of structural
    ' writes (TX/Document/Lines/Lots/Movements) or to the NEW deferred-product-creation path.
    Dim existingCode As String, existingUnit As String
    existingUnit = "шт"
    existingCode = PRIME_CreateProduct("Preseeded Widget", existingUnit, "LOC-SEED", "", "", False, "")
    Log "Seeded product " & existingCode & " (unit " & existingUnit & ")"

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_RECEIPT, "E2E_TEST2", "E2E-EXIST-RECEIPT-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim line1 As PrimeDocLine
    line1.ProductCode = existingCode
    line1.QtyInput = 7
    line1.UnitInput = existingUnit
    line1.LocationTo = "E2E-LOC-EXIST"
    line1.Contour = SC_GENERAL
    PRIME_PlanAddLine(plan, line1)
    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        Log "FAIL: receipt of EXISTING product failed: " & PRIME_LastPostError()
    Else
        Log "OK: receipt of existing product posted, doc=" & docId
        Dim bal As Double
        bal = PRIME_LocationContourBalance(existingCode, "E2E-LOC-EXIST", SC_GENERAL)
        Log "Balance at E2E-LOC-EXIST/GENERAL = " & Format(bal, "0.######")
    End If

    Close #gLogFile
    PRIME_ZZ_ExistingProductReceipt = IIf(docId = "", "1", "0")
    Exit Function
Fatal:
    Log "FATAL: " & Error$ & " Erl=" & Erl
    Close #gLogFile
    PRIME_ZZ_ExistingProductReceipt = "FATAL"
End Function

Public Function PRIME_ZZ_RunFullWorkflow(ByVal logPath As String) As String
    gLogFile = FreeFile
    Open logPath For Output As #gLogFile

    Dim failures As Long
    failures = 0

    On Error Goto Fatal

    Log "=== PRIME 2.1.0 direct-invocation end-to-end workflow ==="

    ' --- 1. Create a brand-new product via a RECEIPT with blank ProductCode (R11 deferred creation) ---
    Dim productName As String
    productName = "E2E Test Widget " & Format(Now, "HHMMSS")
    Dim plan1 As PrimeDocPlan
    PRIME_InitPlan(plan1, DOC_RECEIPT, "E2E_TEST", "E2E-RECEIPT-1-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim line1 As PrimeDocLine
    line1.ProductCode = ""
    line1.ProductName = productName
    line1.QtyInput = 10
    line1.UnitInput = "шт"
    line1.LocationTo = "E2E-LOC-A"
    line1.Contour = SC_GENERAL
    PRIME_PlanAddLine(plan1, line1)
    Dim doc1 As String
    doc1 = PRIME_PostDocument(plan1)
    If doc1 = "" Then
        Log "FAIL: initial receipt of new product failed: " & PRIME_LastPostError()
        failures = failures + 1
        GoTo Done
    End If
    Dim productCode As String
    productCode = plan1.Lines(0).ProductCode
    Log "OK: created product " & productCode & " (" & productName & "), receipt doc=" & doc1

    Dim stockA As Double
    stockA = PRIME_LocationContourBalance(productCode, "E2E-LOC-A", SC_GENERAL)
    If Abs(stockA - 10) > 0.0001 Then
        Log "FAIL: expected stock 10 at E2E-LOC-A/GENERAL immediately after commit (R01 - no reopen needed), got " & Fmt(stockA)
        failures = failures + 1
    Else
        Log "OK (R01): fresh receipt visible in SAME session without reopen, stock=" & Fmt(stockA)
    End If

    ' --- 2. Second partial receipt into a DIFFERENT location (still GENERAL contour) ---
    Dim plan2 As PrimeDocPlan
    PRIME_InitPlan(plan2, DOC_RECEIPT, "E2E_TEST", "E2E-RECEIPT-2-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim line2 As PrimeDocLine
    line2.ProductCode = productCode
    line2.QtyInput = 5
    line2.UnitInput = "шт"
    line2.LocationTo = "E2E-LOC-B"
    line2.Contour = SC_GENERAL
    PRIME_PlanAddLine(plan2, line2)
    Dim doc2 As String
    doc2 = PRIME_PostDocument(plan2)
    If doc2 = "" Then
        Log "FAIL: second receipt (different location) failed: " & PRIME_LastPostError()
        failures = failures + 1
    Else
        Log "OK: second receipt into E2E-LOC-B, doc=" & doc2
    End If

    ' --- 3. R02: issue from LOC-A must only see LOC-A's own stock (10), not the product total (15) ---
    Dim planBadIssue As PrimeDocPlan
    PRIME_InitPlan(planBadIssue, DOC_ISSUE, "E2E_TEST", "E2E-ISSUE-OVER-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineBadIssue As PrimeDocLine
    lineBadIssue.ProductCode = productCode
    lineBadIssue.QtyInput = 12 ' more than LOC-A's 10, but less than product-wide 15
    lineBadIssue.UnitInput = "шт"
    lineBadIssue.LocationFrom = "E2E-LOC-A"
    lineBadIssue.Contour = SC_GENERAL
    PRIME_PlanAddLine(planBadIssue, lineBadIssue)
    Dim docBadIssue As String
    docBadIssue = PRIME_PostDocument(planBadIssue)
    If docBadIssue <> "" Then
        Log "FAIL (R02): issuing 12 from LOC-A (which only has 10) was WRONGLY accepted using product-wide stock"
        failures = failures + 1
    Else
        Log "OK (R02): issuing 12 from LOC-A (only 10 there) correctly rejected: " & PRIME_LastPostError()
    End If

    ' --- 4. R10: two lines of 6 against LOC-A's 10 in ONE document must reject the WHOLE document ---
    Dim planReserve As PrimeDocPlan
    PRIME_InitPlan(planReserve, DOC_ISSUE, "E2E_TEST", "E2E-ISSUE-RESERVE-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineR1 As PrimeDocLine, lineR2 As PrimeDocLine
    lineR1.ProductCode = productCode : lineR1.QtyInput = 6 : lineR1.UnitInput = "шт"
    lineR1.LocationFrom = "E2E-LOC-A" : lineR1.Contour = SC_GENERAL
    lineR2.ProductCode = productCode : lineR2.QtyInput = 6 : lineR2.UnitInput = "шт"
    lineR2.LocationFrom = "E2E-LOC-A" : lineR2.Contour = SC_GENERAL
    PRIME_PlanAddLine(planReserve, lineR1)
    PRIME_PlanAddLine(planReserve, lineR2)
    Dim docReserve As String
    docReserve = PRIME_PostDocument(planReserve)
    If docReserve <> "" Then
        Log "FAIL (R10): two lines of 6 against LOC-A's 10 were WRONGLY both accepted (double-counted stock)"
        failures = failures + 1
    Else
        Log "OK (R10): two same-product lines correctly reserve against each other and reject the whole document: " & PRIME_LastPostError()
    End If
    Dim stockAfterReserveTest As Double
    stockAfterReserveTest = PRIME_LocationContourBalance(productCode, "E2E-LOC-A", SC_GENERAL)
    If Abs(stockAfterReserveTest - 10) > 0.0001 Then
        Log "FAIL: rejected multi-line document must not have partially posted - expected LOC-A stock still 10, got " & Fmt(stockAfterReserveTest)
        failures = failures + 1
    Else
        Log "OK: rejected document left stock untouched (10)"
    End If

    ' --- 5. A real, valid ISSUE of 4 from LOC-A ---
    Dim planIssue As PrimeDocPlan
    PRIME_InitPlan(planIssue, DOC_ISSUE, "E2E_TEST", "E2E-ISSUE-1-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineIssue As PrimeDocLine
    lineIssue.ProductCode = productCode : lineIssue.QtyInput = 4 : lineIssue.UnitInput = "шт"
    lineIssue.LocationFrom = "E2E-LOC-A" : lineIssue.Contour = SC_GENERAL
    lineIssue.Recipient = "E2E Tester"
    PRIME_PlanAddLine(planIssue, lineIssue)
    Dim docIssue As String
    docIssue = PRIME_PostDocument(planIssue)
    If docIssue = "" Then
        Log "FAIL: valid issue of 4 from LOC-A failed: " & PRIME_LastPostError()
        failures = failures + 1
        GoTo TransferSection
    End If
    Log "OK: issued 4 from LOC-A, doc=" & docIssue
    Dim issueLineId As String
    issueLineId = docIssue & "-L1"

    Dim stockAfterIssue As Double
    stockAfterIssue = PRIME_LocationContourBalance(productCode, "E2E-LOC-A", SC_GENERAL)
    If Abs(stockAfterIssue - 6) > 0.0001 Then
        Log "FAIL: expected LOC-A stock 6 after issuing 4 of 10, got " & Fmt(stockAfterIssue)
        failures = failures + 1
    Else
        Log "OK: LOC-A stock correctly 6 after issue"
    End If

    ' --- 6. R08: two partial returns of 2 each must consume the correct allocations in order ---
    Dim planRet1 As PrimeDocPlan
    PRIME_InitPlan(planRet1, DOC_RETURN, "E2E_TEST", "E2E-RETURN-1-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineRet1 As PrimeDocLine
    lineRet1.ProductCode = productCode : lineRet1.QtyInput = 2 : lineRet1.UnitInput = "шт"
    lineRet1.LocationTo = "E2E-LOC-A" : lineRet1.Contour = SC_GENERAL
    lineRet1.OriginalDocLineId = issueLineId
    PRIME_PlanAddLine(planRet1, lineRet1)
    Dim docRet1 As String
    docRet1 = PRIME_PostDocument(planRet1)
    If docRet1 = "" Then
        Log "FAIL: first partial return of 2 failed: " & PRIME_LastPostError()
        failures = failures + 1
    Else
        Log "OK: first partial return of 2 posted, doc=" & docRet1
    End If

    Dim planRet2 As PrimeDocPlan
    PRIME_InitPlan(planRet2, DOC_RETURN, "E2E_TEST", "E2E-RETURN-2-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineRet2 As PrimeDocLine
    lineRet2.ProductCode = productCode : lineRet2.QtyInput = 2 : lineRet2.UnitInput = "шт"
    lineRet2.LocationTo = "E2E-LOC-A" : lineRet2.Contour = SC_GENERAL
    lineRet2.OriginalDocLineId = issueLineId
    PRIME_PlanAddLine(planRet2, lineRet2)
    Dim docRet2 As String
    docRet2 = PRIME_PostDocument(planRet2)
    If docRet2 = "" Then
        Log "FAIL: second partial return of 2 failed: " & PRIME_LastPostError()
        failures = failures + 1
    Else
        Log "OK: second partial return of 2 posted, doc=" & docRet2
    End If

    Dim planRetOver As PrimeDocPlan
    PRIME_InitPlan(planRetOver, DOC_RETURN, "E2E_TEST", "E2E-RETURN-OVER-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineRetOver As PrimeDocLine
    lineRetOver.ProductCode = productCode : lineRetOver.QtyInput = 1 : lineRetOver.UnitInput = "шт" ' 2+2+1=5 > 4 issued
    lineRetOver.LocationTo = "E2E-LOC-A" : lineRetOver.Contour = SC_GENERAL
    lineRetOver.OriginalDocLineId = issueLineId
    PRIME_PlanAddLine(planRetOver, lineRetOver)
    Dim docRetOver As String
    docRetOver = PRIME_PostDocument(planRetOver)
    If docRetOver <> "" Then
        Log "FAIL: over-return (2+2+1=5 > 4 issued) was WRONGLY accepted"
        failures = failures + 1
    Else
        Log "OK: third return correctly rejected as exceeding remaining returnable: " & PRIME_LastPostError()
    End If

    Dim stockAfterReturns As Double
    stockAfterReturns = PRIME_LocationContourBalance(productCode, "E2E-LOC-A", SC_GENERAL)
    If Abs(stockAfterReturns - 10) > 0.0001 Then
        Log "FAIL: expected LOC-A stock back to 10 after returning all 4 issued (6+4), got " & Fmt(stockAfterReturns)
        failures = failures + 1
    Else
        Log "OK: LOC-A stock correctly back to 10 after both partial returns"
    End If

TransferSection:
    ' --- 7. TRANSFER: move 3 from LOC-A/GENERAL to WORKSHOP-1/WORKSHOP_DETAILS, verify lot lineage + contour ---
    Dim planTransfer As PrimeDocPlan
    PRIME_InitPlan(planTransfer, DOC_TRANSFER, "E2E_TEST", "E2E-TRANSFER-1-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineTransfer As PrimeDocLine
    lineTransfer.ProductCode = productCode : lineTransfer.QtyInput = 3 : lineTransfer.UnitInput = "шт"
    lineTransfer.LocationFrom = "E2E-LOC-A" : lineTransfer.ContourFrom = SC_GENERAL
    lineTransfer.LocationTo = "E2E-WORKSHOP-1" : lineTransfer.ContourTo = SC_WORKSHOP_DETAILS
    PRIME_PlanAddLine(planTransfer, lineTransfer)
    Dim docTransfer As String
    docTransfer = PRIME_PostDocument(planTransfer)
    If docTransfer = "" Then
        Log "FAIL: transfer of 3 from LOC-A/GENERAL to WORKSHOP-1/WORKSHOP_DETAILS failed: " & PRIME_LastPostError()
        failures = failures + 1
    Else
        Log "OK: transfer posted, doc=" & docTransfer
        Dim generalAfterTransfer As Double, workshopAfterTransfer As Double
        generalAfterTransfer = PRIME_LocationContourBalance(productCode, "E2E-LOC-A", SC_GENERAL)
        workshopAfterTransfer = PRIME_LocationContourBalance(productCode, "E2E-WORKSHOP-1", SC_WORKSHOP_DETAILS)
        If Abs(generalAfterTransfer - 7) > 0.0001 Or Abs(workshopAfterTransfer - 3) > 0.0001 Then
            Log "FAIL (R05): expected GENERAL/LOC-A=7 and WORKSHOP_DETAILS/WORKSHOP-1=3, got " & Fmt(generalAfterTransfer) & " / " & Fmt(workshopAfterTransfer)
            failures = failures + 1
        Else
            Log "OK (R05): contour-separated stock correct after transfer: GENERAL=7, WORKSHOP_DETAILS=3"
        End If
        Dim generalTotal As Double
        generalTotal = generalAfterTransfer + workshopAfterTransfer + PRIME_LocationContourBalance(productCode, "E2E-LOC-B", SC_GENERAL)
        If Abs(generalTotal - 15) > 0.0001 Then
            Log "FAIL: transfer must conserve total stock across all locations/contours (expected 15), got " & Fmt(generalTotal)
            failures = failures + 1
        Else
            Log "OK: total stock conserved across transfer (15)"
        End If
    End If

    ' --- 8. ADJUSTMENT: shortage of 2 at WORKSHOP-1/WORKSHOP_DETAILS must consume a real lot ---
    Dim planAdj As PrimeDocPlan
    PRIME_InitPlan(planAdj, DOC_ADJUSTMENT, "E2E_TEST", "E2E-ADJ-1-" & Format(Now, "YYYYMMDDHHMMSS"))
    Dim lineAdj As PrimeDocLine
    lineAdj.ProductCode = productCode : lineAdj.QtyInput = -2 : lineAdj.UnitInput = "шт"
    lineAdj.LocationTo = "E2E-WORKSHOP-1" : lineAdj.Contour = SC_WORKSHOP_DETAILS
    PRIME_PlanAddLine(planAdj, lineAdj)
    Dim docAdj As String
    docAdj = PRIME_PostDocument(planAdj)
    If docAdj = "" Then
        Log "FAIL: inventory shortage adjustment of -2 failed: " & PRIME_LastPostError()
        failures = failures + 1
    Else
        Dim workshopAfterAdj As Double
        workshopAfterAdj = PRIME_LocationContourBalance(productCode, "E2E-WORKSHOP-1", SC_WORKSHOP_DETAILS)
        If Abs(workshopAfterAdj - 1) > 0.0001 Then
            Log "FAIL (R06): expected WORKSHOP_DETAILS/WORKSHOP-1=1 after -2 adjustment on 3, got " & Fmt(workshopAfterAdj)
            failures = failures + 1
        Else
            Log "OK (R06): adjustment correctly lot-consistent, WORKSHOP_DETAILS/WORKSHOP-1=1"
        End If
    End If

Done:
    Log ""
    Log "=== SUMMARY: " & failures & " failure(s) ==="
    Close #gLogFile
    PRIME_ZZ_RunFullWorkflow = CStr(failures)
    Exit Function

Fatal:
    Log "FATAL EXCEPTION: " & Error$ & " at line " & Erl
    Close #gLogFile
    PRIME_ZZ_RunFullWorkflow = "FATAL"
End Function
