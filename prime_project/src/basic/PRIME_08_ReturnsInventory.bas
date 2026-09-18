Option Explicit

' Дублируем определение типов из PRIME_04_Posting - StarBasic не поддерживает совместное
' использование Type между модулями (Type виден только в модуле, где объявлен), поэтому
' идентичное объявление нужно в каждом модуле, который строит PrimeDocPlan/PrimeDocLine.
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


' PRIME_08_ReturnsInventory
' "Возвраты" - единственный механизм возврата (заменяет конфликтующие модули 08/23 из 1.4.1).
' "Инвентаризация" - снимок остатков, ввод факта, разница проводится как ADJUSTMENT только
' при отсутствии движений после снимка (inventory.workflow).

' === Возвраты =================================================================================

' Перестраивает список открытых возвратов из DB_PRIME_DOC_LINES (ISSUE-строки возвратного товара
' с остатком к возврату > 0). Чисто отображение - не пишет в движения.
Public Sub PRIME_Returns_RefreshButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_RETURNS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_RETURNS)

    ' Очищаем старые строки (только данные, не заголовок) - список полностью пересчитываемый.
    PRIME_ClearDataRows(oSheet, headers)

    If Not PRIME_SheetExists(SH_DB_DOC_LINES) Or Not PRIME_SheetExists(SH_DB_DOCUMENTS) Then Exit Sub

    Dim lineHeaders As Variant
    lineHeaders = PRIME_HeaderMap(SH_DB_DOC_LINES)
    Dim docHeaders As Variant
    docHeaders = PRIME_HeaderMap(SH_DB_DOCUMENTS)

    Dim lineTable As Variant
    lineTable = PRIME_ReadTable(SH_DB_DOC_LINES)
    Dim docTable As Variant
    docTable = PRIME_ReadTable(SH_DB_DOCUMENTS)

    Dim colLineId As Long, colDocId As Long, colProduct As Long, colQty As Long, colRecipient As Long
    colLineId = PRIME_ColIndex(lineHeaders, "DOC_LINE_ID")
    colDocId = PRIME_ColIndex(lineHeaders, "DOC_ID")
    colProduct = PRIME_ColIndex(lineHeaders, "PRODUCT_CODE")
    colQty = PRIME_ColIndex(lineHeaders, "QTY_BASE")
    colRecipient = PRIME_ColIndex(lineHeaders, "RECIPIENT")

    Dim colDocDocId As Long, colDocType As Long, colDocDate As Long
    colDocDocId = PRIME_ColIndex(docHeaders, "DOC_ID")
    colDocType = PRIME_ColIndex(docHeaders, "DOC_TYPE")
    colDocDate = PRIME_ColIndex(docHeaders, "DOC_DATE")

    ' headless_buffered_array_readback_corruption (2.1.2, см. PRIME_04_Posting.PRIME_PostIssueLines) -
    ' здесь раньше строки накапливались в outRows()-буфер (нарастающий Variant()-массив с Dim row()
    ' внутри цикла) и передавались одним PRIME_AppendRowsBatch в конце; после хотя бы одной более
    ' ранней записи на другой лист в этой же цепочке invoke это детерминированно писало на "Возвраты"
    ' пустые/нулевые строки без единой ошибки (подтверждено прямым чтением листа). Фикс - как и в
    ' PostIssueLines: писать каждую строку сразу поячейково внутри цикла, без буферизации.
    Dim colOutDate As Long, colOutCode As Long, colOutName As Long, colOutIssued As Long
    Dim colOutReturned As Long, colOutRemaining As Long, colOutUnit As Long, colOutWhom As Long
    Dim colOutOrigLine As Long
    colOutDate = PRIME_ColIndex(headers, "Дата выдачи")
    colOutCode = PRIME_ColIndex(headers, "Код")
    colOutName = PRIME_ColIndex(headers, "Наименование")
    colOutIssued = PRIME_ColIndex(headers, "Выдано")
    colOutReturned = PRIME_ColIndex(headers, "Уже возвращено")
    colOutRemaining = PRIME_ColIndex(headers, "Осталось к возврату")
    colOutUnit = PRIME_ColIndex(headers, "Ед. изм.")
    colOutWhom = PRIME_ColIndex(headers, "Кому")
    colOutOrigLine = PRIME_ColIndex(headers, "_PRIME_OriginalLineId")

    Dim outRow As Long
    outRow = PRIME_FormSchemaFirstDataRow(SH_RETURNS)

    Dim i As Long
    If UBound(lineTable) >= 1 Then
        For i = 1 To UBound(lineTable)
            Dim docId As String
            docId = CStr(lineTable(i)(colDocId))
            Dim docIdx As Long
            docIdx = PRIME_FindRowByKey(docTable, colDocDocId, docId)
            If docIdx >= 0 Then
                ' committed_only_everywhere (R07): DB_PRIME_DOCUMENTS/DOC_LINES строки пишутся ДО
                ' TX=COMMITTED (см. PRIME_04_Posting.PRIME_WriteDocumentHeader) - PREPARED/FAILED
                ' документ не должен предлагать "выдачу" к возврату.
                If CStr(docTable(docIdx)(colDocType)) = DOC_ISSUE And PRIME_IsDocIdCommitted(docId) Then
                    Dim productCode As String
                    productCode = CStr(lineTable(i)(colProduct))
                    If LCase(PRIME_GetProductField(productCode, "RETURNABLE")) = "1" Then
                        Dim lineId As String
                        lineId = CStr(lineTable(i)(colLineId))
                        Dim issuedQty As Double
                        issuedQty = CDbl(lineTable(i)(colQty))
                        Dim returnedQty As Double
                        returnedQty = PRIME_AlreadyReturnedQtyBase(lineId)
                        Dim remaining As Double
                        remaining = issuedQty - returnedQty
                        If remaining > 0.0000005 Then
                            oSheet.getCellByPosition(colOutDate, outRow).setString(CStr(docTable(docIdx)(colDocDate)))
                            oSheet.getCellByPosition(colOutCode, outRow).setString(productCode)
                            oSheet.getCellByPosition(colOutName, outRow).setString(PRIME_GetProductField(productCode, "PRODUCT_NAME"))
                            oSheet.getCellByPosition(colOutIssued, outRow).setValue(issuedQty)
                            oSheet.getCellByPosition(colOutReturned, outRow).setValue(returnedQty)
                            oSheet.getCellByPosition(colOutRemaining, outRow).setValue(remaining)
                            oSheet.getCellByPosition(colOutUnit, outRow).setString(PRIME_GetProductField(productCode, "BASE_UNIT"))
                            oSheet.getCellByPosition(colOutWhom, outRow).setString(CStr(lineTable(i)(colRecipient)))
                            oSheet.getCellByPosition(colOutOrigLine, outRow).setString(lineId)
                            outRow = outRow + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
End Sub

' Событие: ввод "Вернуть сейчас" клеймит строку стабильным ReturnState-маркером (по аналогии
' с Orders' pending-delivery, чтобы двойной клик "Провести возвраты" не задвоил возврат).
Public Sub PRIME_OnContentChanged_Returns(ByVal oRangeAddr As Variant)
    If Not PRIME_EventEnter() Then Exit Sub
    On Error Goto CleanExit
    If IsNull(oRangeAddr) Then GoTo CleanExit

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_RETURNS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_RETURNS)
    Dim colReturnNow As Long, colState As Long
    colReturnNow = PRIME_ColIndex(headers, "Вернуть сейчас")
    colState = PRIME_ColIndex(headers, "_PRIME_ReturnState")

    Dim firstDataRow As Long
    firstDataRow = PRIME_FormSchemaFirstDataRow(SH_RETURNS)
    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= firstDataRow Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colReturnNow And colState >= 0 Then
                    Dim v As String
                    v = Trim(oSheet.getCellByPosition(colReturnNow, r).getString())
                    Dim state As String
                    state = oSheet.getCellByPosition(colState, r).getString()
                    If v = "" Then
                        If Left(state, 4) = "RET:" Then oSheet.getCellByPosition(colState, r).setString("")
                    ElseIf Left(state, 4) <> "RET:" Then
                        oSheet.getCellByPosition(colState, r).setString("RET:" & Format(Now, "YYYYMMDDHHMMSS") & Int(Rnd * 9999))
                    End If
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
End Sub

' R09 (2.1.0): все отмеченные строки возврата проводятся ОДНИМ документом (один DOC_ID/OP_ID/
' один store()), а не циклом отдельных PostDocument-вызовов на каждую строку - раньше N строк
' возврата означали N документов и N store(), без атомарности "весь батч либо целиком, либо
' никак" и с лишними файловыми сохранениями.
Public Sub PRIME_Returns_ConductButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_RETURNS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_RETURNS)
    Dim colReturnNow As Long, colState As Long, colOriginalLine As Long
    colReturnNow = PRIME_ColIndex(headers, "Вернуть сейчас")
    colState = PRIME_ColIndex(headers, "_PRIME_ReturnState")
    colOriginalLine = PRIME_ColIndex(headers, "_PRIME_OriginalLineId")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_RETURN, SH_RETURNS, "")
    Dim rowForLine(999) As Long
    Dim batchKey As String
    batchKey = ""

    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(SH_RETURNS) To lastRow
        Dim qtyStr As String
        qtyStr = Trim(oSheet.getCellByPosition(colReturnNow, r).getString())
        ' batch_invalid_line fix (2.1.1): раньше заполненное, но невалидное (нечисловое/<=0)
        ' "Вернуть сейчас" молча (без единого сообщения) исключалось из батча - теперь
        ' останавливает построение всего батча целиком.
        If qtyStr <> "" And (Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0) Then
            MsgBox "Проведение не выполнено: строка " & (r + 1) & " ""Вернуть сейчас"" заполнена, но невалидна. Исправьте её или очистите перед проведением всего батча."
            Exit Sub
        End If
        If qtyStr <> "" And IsNumeric(qtyStr) And CDbl(qtyStr) > 0 Then
            Dim state As String
            state = oSheet.getCellByPosition(colState, r).getString()
            If Left(state, 4) <> "RET:" Then
                state = "RET:" & Format(Now, "YYYYMMDDHHMMSS") & Int(Rnd * 9999) & "-" & r
                oSheet.getCellByPosition(colState, r).setString(state)
            End If
            batchKey = batchKey & state & ","

            Dim docLine As PrimeDocLine
            docLine.ProductCode = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код"), r).getString()
            docLine.QtyInput = CDbl(qtyStr)
            docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), r).getString()
            docLine.LocationTo = PRIME_GetProductField(docLine.ProductCode, "DEFAULT_LOCATION")
            docLine.Contour = SC_GENERAL
            docLine.OriginalDocLineId = oSheet.getCellByPosition(colOriginalLine, r).getString()
            docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), r).getString()
            rowForLine(plan.LineCount) = r
            PRIME_PlanAddLine(plan, docLine)
        End If
    Next r

    If plan.LineCount = 0 Then
        MsgBox "Нет строк с заполненным ""Вернуть сейчас""."
        Exit Sub
    End If
    plan.SourceKey = "BATCH:" & batchKey

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Возврат не проведён: " & PRIME_LastPostError()
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        oSheet.getCellByPosition(colReturnNow, rowForLine(i)).setString("")
        oSheet.getCellByPosition(colState, rowForLine(i)).setString("")
    Next i

    MsgBox "Проведено возвратов: " & plan.LineCount & " (документ " & docId & ")"
    PRIME_Returns_RefreshButton()
End Sub

' === Инвентаризация ============================================================================

Public Sub PRIME_Inventory_LoadButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_INVENTORY)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_INVENTORY)

    PRIME_ClearDataRows(oSheet, headers)

    Dim sessionId As String
    sessionId = "INV-" & Format(PRIME_SequenceNext("INVENTORY_SESSION_ID"), "00000000")
    PRIME_MetaSet("CURRENT_INVENTORY_SESSION", sessionId)
    PRIME_MetaSet("INV_SNAPSHOT_" & sessionId, Format(Now, "YYYY-MM-DD HH:MM:SS"))

    Dim prodHeaders As Variant
    prodHeaders = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim prodTable As Variant
    prodTable = PRIME_ReadTable(SH_DB_PRODUCTS)
    Dim colCode As Long, colName As Long, colCat As Long, colSub As Long
    colCode = PRIME_ColIndex(prodHeaders, "PRODUCT_CODE")
    colName = PRIME_ColIndex(prodHeaders, "PRODUCT_NAME")
    colCat = PRIME_ColIndex(prodHeaders, "CATEGORY")
    colSub = PRIME_ColIndex(prodHeaders, "SUBCATEGORY")

    ' headless_buffered_array_readback_corruption (2.1.2, см. PRIME_04_Posting.PRIME_PostIssueLines) -
    ' раньше строки накапливались в outRows()-буфер (нарастающий Variant()-массив с Dim row() внутри
    ' цикла) и передавались одним PRIME_AppendRowsBatch в конце; это тот же ненадёжный паттерн, что
    ' был найден и исправлен в PostIssueLines и в PRIME_Returns_RefreshButton - пишем каждую строку
    ' сразу поячейково внутри цикла, без буферизации.
    Dim colOutSession As Long, colOutCode As Long, colOutName As Long, colOutLoc As Long
    Dim colOutContour As Long, colOutCat As Long, colOutSub As Long, colOutUchet As Long, colOutUnit As Long
    colOutSession = PRIME_ColIndex(headers, "Сессия")
    colOutCode = PRIME_ColIndex(headers, "Код")
    colOutName = PRIME_ColIndex(headers, "Наименование")
    colOutLoc = PRIME_ColIndex(headers, "Место")
    colOutContour = PRIME_ColIndex(headers, "Контур")
    colOutCat = PRIME_ColIndex(headers, "Категория")
    colOutSub = PRIME_ColIndex(headers, "Подкатегория")
    colOutUchet = PRIME_ColIndex(headers, "Учёт")
    colOutUnit = PRIME_ColIndex(headers, "Ед. изм.")

    Dim outRow As Long
    outRow = PRIME_FormSchemaFirstDataRow(SH_INVENTORY)
    Dim n As Long
    n = 0

    Dim locations() As String
    Dim quantities() As Double
    Dim contours() As String
    If UBound(prodTable) >= 1 Then
        Dim p As Long
        For p = 1 To UBound(prodTable)
            Dim code As String
            code = CStr(prodTable(p)(colCode))
            ' contourFilter="" - снимаем остаток по КАЖДОМУ (месту, контуру) отдельно (R06):
            ' один и тот же код на одном месте, но в разных контурах, не должен задваивать разницу.
            PRIME_StockByLocation(code, "", locations, quantities, contours)
            If UBound(locations) >= LBound(locations) Then
                Dim l As Long
                For l = LBound(locations) To UBound(locations)
                    If quantities(l) <> 0 Then
                        oSheet.getCellByPosition(colOutSession, outRow).setString(sessionId)
                        oSheet.getCellByPosition(colOutCode, outRow).setString(code)
                        oSheet.getCellByPosition(colOutName, outRow).setString(CStr(prodTable(p)(colName)))
                        oSheet.getCellByPosition(colOutLoc, outRow).setString(locations(l))
                        oSheet.getCellByPosition(colOutContour, outRow).setString(contours(l))
                        oSheet.getCellByPosition(colOutCat, outRow).setString(CStr(prodTable(p)(colCat)))
                        oSheet.getCellByPosition(colOutSub, outRow).setString(CStr(prodTable(p)(colSub)))
                        oSheet.getCellByPosition(colOutUchet, outRow).setValue(quantities(l))
                        oSheet.getCellByPosition(colOutUnit, outRow).setString(PRIME_GetProductField(code, "BASE_UNIT"))
                        outRow = outRow + 1
                        n = n + 1
                    End If
                Next l
            End If
        Next p
    End If

    MsgBox "Загружено строк остатка: " & n & ". Сессия: " & sessionId
End Sub

Public Sub PRIME_Inventory_RecalcButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_INVENTORY)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_INVENTORY)
    Dim colFact As Long, colUchet As Long, colDiff As Long
    colFact = PRIME_ColIndex(headers, "Факт")
    colUchet = PRIME_ColIndex(headers, "Учёт")
    colDiff = PRIME_ColIndex(headers, "Разница")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(SH_INVENTORY) To lastRow
        Dim factStr As String
        factStr = Trim(oSheet.getCellByPosition(colFact, r).getString())
        If factStr <> "" And IsNumeric(factStr) Then
            Dim uchet As Double
            uchet = CDbl(oSheet.getCellByPosition(colUchet, r).getString())
            oSheet.getCellByPosition(colDiff, r).setValue(CDbl(factStr) - uchet)
        End If
    Next r
End Sub

' inventory.workflow: перед проведением проверяем движения после снимка сессии - конфликтную
' строку пропускаем с сообщением, а не проводим вслепую.
' R06/R09 (2.1.0): все неконфликтные строки разницы проводятся ОДНИМ ADJUSTMENT-документом
' (один DOC_ID/OP_ID/один store()), контур берётся из колонки "Контур" (см.
' PRIME_Inventory_LoadButton) - PostAdjustmentLines теперь консервативно списывает/создаёт
' реальные партии, поэтому stock == sum(lots) гарантированно после проведения (R06).
Public Sub PRIME_Inventory_ConductButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_INVENTORY)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_INVENTORY)
    Dim colSession As Long, colCode As Long, colLoc As Long, colContour As Long, colDiff As Long, colUnit As Long
    colSession = PRIME_ColIndex(headers, "Сессия")
    colCode = PRIME_ColIndex(headers, "Код")
    colLoc = PRIME_ColIndex(headers, "Место")
    colContour = PRIME_ColIndex(headers, "Контур")
    colDiff = PRIME_ColIndex(headers, "Разница")
    colUnit = PRIME_ColIndex(headers, "Ед. изм.")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim conflicts As Long
    conflicts = 0

    Dim plan As PrimeDocPlan
    Dim sessionsInBatch As String
    sessionsInBatch = ""
    Dim rowForLine(999) As Long

    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(SH_INVENTORY) To lastRow
        Dim diffStr As String
        diffStr = Trim(oSheet.getCellByPosition(colDiff, r).getString())
        If diffStr <> "" And IsNumeric(diffStr) And CDbl(diffStr) <> 0 Then
            Dim sessionId As String
            sessionId = oSheet.getCellByPosition(colSession, r).getString()
            Dim snapshotAt As String
            snapshotAt = PRIME_MetaGet("INV_SNAPSHOT_" & sessionId)
            Dim code As String
            code = oSheet.getCellByPosition(colCode, r).getString()

            If PRIME_HasMovementsSince(code, snapshotAt) Then
                conflicts = conflicts + 1
            Else
                If plan.LineCount = 0 Then PRIME_InitPlan(plan, DOC_ADJUSTMENT, SH_INVENTORY, "")
                sessionsInBatch = sessionsInBatch & sessionId & "|" & code & "|" & oSheet.getCellByPosition(colLoc, r).getString() & ","

                Dim docLine As PrimeDocLine
                docLine.ProductCode = code
                docLine.QtyInput = CDbl(diffStr)
                docLine.UnitInput = oSheet.getCellByPosition(colUnit, r).getString()
                docLine.LocationTo = oSheet.getCellByPosition(colLoc, r).getString()
                docLine.Contour = oSheet.getCellByPosition(colContour, r).getString()
                docLine.Comment = "Инвентаризация " & sessionId
                rowForLine(plan.LineCount) = r
                PRIME_PlanAddLine(plan, docLine)
            End If
        End If
    Next r

    If plan.LineCount = 0 Then
        MsgBox "Нет строк с разницей для проведения. Конфликтов (обновите остатки): " & conflicts
        Exit Sub
    End If
    plan.SourceKey = "BATCH:" & sessionsInBatch

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Корректировка не проведена: " & PRIME_LastPostError()
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        oSheet.getCellByPosition(colDiff, rowForLine(i)).setString("проведено: " & docId)
    Next i

    MsgBox "Проведено корректировок: " & plan.LineCount & " (документ " & docId & "). Конфликтов (обновите остатки): " & conflicts
End Sub

Private Function PRIME_HasMovementsSince(ByVal productCode As String, ByVal sinceTs As String) As Boolean
    If sinceTs = "" Or Not PRIME_SheetExists(SH_DB_MOVEMENTS) Then
        PRIME_HasMovementsSince = False
        Exit Function
    End If
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_MOVEMENTS)
    Dim colProduct As Long, colDate As Long
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colDate = PRIME_ColIndex(headers, "MOVE_DATE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_MOVEMENTS)
    If UBound(table) < 1 Then
        PRIME_HasMovementsSince = False
        Exit Function
    End If
    Dim i As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(colProduct)) = productCode Then
            If CStr(table(i)(colDate)) > sinceTs Then
                PRIME_HasMovementsSince = True
                Exit Function
            End If
        End If
    Next i
    PRIME_HasMovementsSince = False
End Function
