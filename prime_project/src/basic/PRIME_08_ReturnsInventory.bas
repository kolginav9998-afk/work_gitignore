Option Explicit

' Дублируем определение типов из PRIME_04_Posting - StarBasic не поддерживает совместное
' использование Type между модулями (Type виден только в модуле, где объявлен), поэтому
' идентичное объявление нужно в каждом модуле, который строит PrimeDocPlan/PrimeDocLine.
Type PrimeDocLine
    ProductCode As String
    ProductName As String
    QtyInput As Double
    UnitInput As String
    LocationFrom As String
    LocationTo As String
    DestinationProject As String
    Recipient As String
    Comment As String
    Price As Double
    OriginalDocLineId As String
    QtyBase As Double
    LotId As String
End Type

Type PrimeDocPlan
    DocType As String
    DocDate As String
    SourceSheet As String
    SourceKey As String
    OrderId As String
    Lines(99) As PrimeDocLine
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

    Dim outRows() As Variant
    ReDim outRows(200)
    Dim n As Long
    n = 0

    Dim i As Long
    If UBound(lineTable) >= 1 Then
        For i = 1 To UBound(lineTable)
            Dim docId As String
            docId = CStr(lineTable(i)(colDocId))
            Dim docIdx As Long
            docIdx = PRIME_FindRowByKey(docTable, colDocDocId, docId)
            If docIdx >= 0 Then
                If CStr(docTable(docIdx)(colDocType)) = DOC_ISSUE Then
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
                            Dim row(UBound(headers)) As Variant
                            row(PRIME_ColIndex(headers, "Дата выдачи")) = CStr(docTable(docIdx)(colDocDate))
                            row(PRIME_ColIndex(headers, "Код")) = productCode
                            row(PRIME_ColIndex(headers, "Наименование")) = PRIME_GetProductField(productCode, "PRODUCT_NAME")
                            row(PRIME_ColIndex(headers, "Выдано")) = issuedQty
                            row(PRIME_ColIndex(headers, "Уже возвращено")) = returnedQty
                            row(PRIME_ColIndex(headers, "Осталось к возврату")) = remaining
                            row(PRIME_ColIndex(headers, "Ед. изм.")) = PRIME_GetProductField(productCode, "BASE_UNIT")
                            row(PRIME_ColIndex(headers, "Кому")) = CStr(lineTable(i)(colRecipient))
                            row(PRIME_ColIndex(headers, "_PRIME_OriginalLineId")) = lineId
                            If n > UBound(outRows) Then ReDim Preserve outRows(UBound(outRows) + 200)
                            outRows(n) = row
                            n = n + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If n > 0 Then
        ReDim Preserve outRows(n - 1)
        PRIME_AppendRowsBatch(SH_RETURNS, outRows)
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
    Dim posted As Long
    posted = 0

    Dim r As Long
    For r = PRIME_FormSchemaFirstDataRow(SH_RETURNS) To lastRow
        Dim qtyStr As String
        qtyStr = Trim(oSheet.getCellByPosition(colReturnNow, r).getString())
        If qtyStr <> "" And IsNumeric(qtyStr) And CDbl(qtyStr) > 0 Then
            Dim state As String
            state = oSheet.getCellByPosition(colState, r).getString()
            If Left(state, 4) <> "RET:" Then
                state = "RET:" & Format(Now, "YYYYMMDDHHMMSS") & Int(Rnd * 9999) & "-" & r
                oSheet.getCellByPosition(colState, r).setString(state)
            End If

            Dim originalLineId As String
            originalLineId = oSheet.getCellByPosition(colOriginalLine, r).getString()

            Dim plan As PrimeDocPlan
            PRIME_InitPlan(plan, DOC_RETURN, SH_RETURNS, state)

            Dim docLine As PrimeDocLine
            docLine.ProductCode = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код"), r).getString()
            docLine.QtyInput = CDbl(qtyStr)
            docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), r).getString()
            docLine.LocationTo = PRIME_GetProductField(docLine.ProductCode, "DEFAULT_LOCATION")
            docLine.OriginalDocLineId = originalLineId
            docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Комментарий"), r).getString()
            PRIME_PlanAddLine(plan, docLine)

            Dim docId As String
            docId = PRIME_PostDocument(plan)
            If docId = "" Then
                MsgBox "Строка " & (r + 1) & ": возврат не проведён - " & PRIME_LastPostError()
            Else
                oSheet.getCellByPosition(colReturnNow, r).setString("")
                oSheet.getCellByPosition(colState, r).setString("")
                posted = posted + 1
            End If
        End If
    Next r

    MsgBox "Проведено возвратов: " & posted
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

    Dim outRows() As Variant
    ReDim outRows(200)
    Dim n As Long
    n = 0

    If UBound(prodTable) >= 1 Then
        Dim p As Long
        For p = 1 To UBound(prodTable)
            Dim code As String
            code = CStr(prodTable(p)(colCode))
            Dim locations() As String
            Dim quantities() As Double
            PRIME_StockByLocation(code, locations, quantities)
            If UBound(locations) >= LBound(locations) Then
                Dim l As Long
                For l = LBound(locations) To UBound(locations)
                    If quantities(l) <> 0 Then
                        Dim row(UBound(headers)) As Variant
                        row(PRIME_ColIndex(headers, "Сессия")) = sessionId
                        row(PRIME_ColIndex(headers, "Код")) = code
                        row(PRIME_ColIndex(headers, "Наименование")) = CStr(prodTable(p)(colName))
                        row(PRIME_ColIndex(headers, "Место")) = locations(l)
                        row(PRIME_ColIndex(headers, "Категория")) = CStr(prodTable(p)(colCat))
                        row(PRIME_ColIndex(headers, "Подкатегория")) = CStr(prodTable(p)(colSub))
                        row(PRIME_ColIndex(headers, "Учёт")) = quantities(l)
                        row(PRIME_ColIndex(headers, "Ед. изм.")) = PRIME_GetProductField(code, "BASE_UNIT")
                        If n > UBound(outRows) Then ReDim Preserve outRows(UBound(outRows) + 200)
                        outRows(n) = row
                        n = n + 1
                    End If
                Next l
            End If
        Next p
    End If

    If n > 0 Then
        ReDim Preserve outRows(n - 1)
        PRIME_AppendRowsBatch(SH_INVENTORY, outRows)
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
Public Sub PRIME_Inventory_ConductButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_INVENTORY)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_INVENTORY)
    Dim colSession As Long, colCode As Long, colLoc As Long, colDiff As Long, colUnit As Long
    colSession = PRIME_ColIndex(headers, "Сессия")
    colCode = PRIME_ColIndex(headers, "Код")
    colLoc = PRIME_ColIndex(headers, "Место")
    colDiff = PRIME_ColIndex(headers, "Разница")
    colUnit = PRIME_ColIndex(headers, "Ед. изм.")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim posted As Long, conflicts As Long
    posted = 0 : conflicts = 0

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
                Dim plan As PrimeDocPlan
                PRIME_InitPlan(plan, DOC_ADJUSTMENT, SH_INVENTORY, sessionId & "|" & code & "|" & oSheet.getCellByPosition(colLoc, r).getString())
                Dim docLine As PrimeDocLine
                docLine.ProductCode = code
                docLine.QtyInput = CDbl(diffStr)
                docLine.UnitInput = oSheet.getCellByPosition(colUnit, r).getString()
                docLine.LocationTo = oSheet.getCellByPosition(colLoc, r).getString()
                docLine.Comment = "Инвентаризация " & sessionId
                PRIME_PlanAddLine(plan, docLine)

                Dim docId As String
                docId = PRIME_PostDocument(plan)
                If docId <> "" Then
                    posted = posted + 1
                    oSheet.getCellByPosition(colDiff, r).setString("проведено: " & docId)
                End If
            End If
        End If
    Next r

    MsgBox "Проведено корректировок: " & posted & ". Конфликтов (обновите остатки): " & conflicts
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
