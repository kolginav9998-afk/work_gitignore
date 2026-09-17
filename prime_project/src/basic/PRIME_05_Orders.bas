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


' PRIME_05_Orders
' Лист "Заказы": 25 бизнес-колонок 1.4.1 без изменений + 4 расширенных + 3 скрытых helper-колонки
' (_PRIME_OrderID, _PRIME_LineID, _PRIME_State - avoid_many_hidden_columns, было 10 в 1.4.1).
' Единственный обработчик событий - PRIME_OnContentChanged_Orders, точечный (одна строка,
' одна колонка), без пересчёта всего диапазона (закрывает Orders_RebuildGroupContexts из 1.4.1).

' === Событие листа (лёгкое, events.forbidden_actions соблюдены) ==============================
Public Sub PRIME_OnContentChanged_Orders(ByVal oRangeAddr As Variant)
    If Not PRIME_EventEnter() Then Exit Sub
    On Error Goto CleanExit

    If IsNull(oRangeAddr) Then GoTo CleanExit ' сигнал "весь лист"/структурное изменение - не обрабатываем

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colCode As Long, colFact As Long, colState As Long, colExpected As Long, colOrdered As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colFact = PRIME_ColIndex(headers, "Факт. количество")
    colState = PRIME_ColIndex(headers, "_PRIME_State")
    colExpected = PRIME_ColIndex(headers, "Ожидаемая дата")
    colOrdered = PRIME_ColIndex(headers, "Количество")

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= 1 Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Orders_AutofillByCode(oSheet, headers, r)
                ElseIf c = colFact And colState >= 0 Then
                    PRIME_Orders_TrackPendingDelivery(oSheet, headers, r)
                ElseIf c = colExpected Then
                    ' dates support (recommendation §38): нормализуем гибкий пользовательский
                    ' ввод даты в "YYYY-MM-DD" - нераспознанный формат оставляем как есть.
                    Dim rawExpected As String
                    rawExpected = Trim(oSheet.getCellByPosition(colExpected, r).getString())
                    If rawExpected <> "" Then
                        Dim parsed As String
                        parsed = PRIME_ParseFlexibleDate(rawExpected)
                        If parsed <> "" And parsed <> rawExpected Then
                            oSheet.getCellByPosition(colExpected, r).setString(parsed)
                        End If
                    End If
                    PRIME_Orders_RecomputeStatus(oSheet, headers, r)
                ElseIf c = colOrdered Then
                    PRIME_Orders_RecomputeStatus(oSheet, headers, r)
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
End Sub

' R22 (2.1.0): "В наличии сейчас" - представление того же COMMITTED-ledger, что и лист
' "Наличие", всегда для контура SC_GENERAL (Заказы не работают с цеховым/офисным контуром) и
' места, указанного в строке заказа. Пересчитывается тем же событием, что и статус - отдельной
' кнопки не требуется.
Public Sub PRIME_Orders_RefreshAvailability(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colAvail As Long, colCode As Long, colLoc As Long
    colAvail = PRIME_ColIndex(headers, "В наличии сейчас")
    colCode = PRIME_ColIndex(headers, "Код товара")
    If colAvail < 0 Or colCode < 0 Then Exit Sub
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub
    colLoc = PRIME_ColIndex(headers, "Место хранения")
    Dim loc As String
    loc = ""
    If colLoc >= 0 Then loc = Trim(oSheet.getCellByPosition(colLoc, row).getString())
    oSheet.getCellByPosition(colAvail, row).setValue(PRIME_LocationContourBalance(code, loc, SC_GENERAL))
End Sub

' order_status_logic (recommendation §14/§38): пересчитывает "Статус" по остатку/датам.
' "Отменено" - единственный статус, который эта функция никогда не перезаписывает (ручное
' решение пользователя, аналогично do_not_overwrite_non_empty_user_fields).
Public Sub PRIME_Orders_RecomputeStatus(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    PRIME_Orders_RefreshAvailability(oSheet, headers, row)
    Dim colStatus As Long
    colStatus = PRIME_ColIndex(headers, "Статус")
    If colStatus < 0 Then Exit Sub

    Dim currentStatus As String
    currentStatus = Trim(oSheet.getCellByPosition(colStatus, row).getString())
    If currentStatus = ORDER_STATUS_CANCELLED Then Exit Sub

    Dim colOrdered As Long, colReceived As Long, colExpected As Long
    colOrdered = PRIME_ColIndex(headers, "Количество")
    colReceived = PRIME_ColIndex(headers, "Получено всего")
    colExpected = PRIME_ColIndex(headers, "Ожидаемая дата")

    Dim orderedStr As String
    orderedStr = Trim(oSheet.getCellByPosition(colOrdered, row).getString())
    If orderedStr = "" Or Not IsNumeric(orderedStr) Then
        If currentStatus = "" Then oSheet.getCellByPosition(colStatus, row).setString(ORDER_STATUS_DRAFT)
        Exit Sub
    End If

    Dim receivedStr As String
    receivedStr = ""
    If colReceived >= 0 Then receivedStr = Trim(oSheet.getCellByPosition(colReceived, row).getString())
    Dim ordered As Double, received As Double
    ordered = CDbl(orderedStr)
    received = 0
    If receivedStr <> "" And IsNumeric(receivedStr) Then received = CDbl(receivedStr)

    Dim newStatus As String
    If received > 0 And received >= ordered - 0.0000005 Then
        newStatus = ORDER_STATUS_RECEIVED
    ElseIf received > 0 Then
        newStatus = ORDER_STATUS_PARTIAL
    Else
        newStatus = ORDER_STATUS_EXPECTED
    End If

    ' Полностью полученный заказ никогда не становится просроченным; пустая ожидаемая дата
    ' не создаёт просрочку (обе явно оговорены в order_status_logic.rules).
    If newStatus <> ORDER_STATUS_RECEIVED And colExpected >= 0 Then
        Dim expectedStr As String
        expectedStr = Trim(oSheet.getCellByPosition(colExpected, row).getString())
        If expectedStr <> "" And expectedStr < Format(Now, "YYYY-MM-DD") Then
            newStatus = ORDER_STATUS_OVERDUE
        End If
    End If

    oSheet.getCellByPosition(colStatus, row).setString(newStatus)
End Sub

' Ввод внутреннего кода - точечный lookup по memory index, без сканирования всего листа
' (live_code_lookup.Заказы). do_not_overwrite_non_empty_user_fields соблюдается PRIME_SetCellIfEmpty.
Private Sub PRIME_Orders_AutofillByCode(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    PRIME_SetCellIfEmpty(oSheet, headers, row, "Полное наименование товара", PRIME_GetProductField(code, "PRODUCT_NAME"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Ед. изм.", PRIME_GetProductField(code, "BASE_UNIT"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Категория", PRIME_GetProductField(code, "CATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Подкатегория", PRIME_GetProductField(code, "SUBCATEGORY"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Место хранения", PRIME_GetProductField(code, "DEFAULT_LOCATION"))
End Sub

Public Sub PRIME_SetCellIfEmpty(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, ByVal colName As String, ByVal value As String)
    Dim col As Long
    col = PRIME_ColIndex(headers, colName)
    If col = -1 Or value = "" Then Exit Sub
    If oSheet.getCellByPosition(col, row).getString() = "" Then
        oSheet.getCellByPosition(col, row).setString(value)
    End If
End Sub

' Клеймит незавершённую поставку идентификатором DELIVERY_ID в _PRIME_State при первом вводе
' "Факт. количество" (и сохраняет его при повторном редактировании того же значения до
' проведения) - это и есть SOURCE_KEY-компонент, обеспечивающий, что двойной клик "Провести"
' не создаёт вторую поставку, а два РАЗНЫХ ввода Факт.количества - создают.
Private Sub PRIME_Orders_TrackPendingDelivery(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colFact As Long, colState As Long
    colFact = PRIME_ColIndex(headers, "Факт. количество")
    colState = PRIME_ColIndex(headers, "_PRIME_State")

    Dim factStr As String
    factStr = Trim(oSheet.getCellByPosition(colFact, row).getString())
    Dim state As String
    state = oSheet.getCellByPosition(colState, row).getString()

    If factStr = "" Then
        If Left(state, 4) = "DLV:" Then oSheet.getCellByPosition(colState, row).setString("")
        Exit Sub
    End If

    If Left(state, 4) <> "DLV:" Then
        Dim newId As String
        newId = "DLV:" & Format(Now, "YYYYMMDDHHMMSS") & Int(Rnd * 9999)
        oSheet.getCellByPosition(colState, row).setString(newId)
    End If
End Sub

' === Кнопки =====================================================================================

Public Sub PRIME_Orders_NewOrder()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)

    Dim newRow As Long
    newRow = PRIME_FindLastRow(oSheet) + 1
    If newRow < 1 Then newRow = 1

    Dim orderId As String, lineId As String
    orderId = "ORD-" & Format(PRIME_SequenceNext("ORDER_ID"), "00000000")
    lineId = "OL-" & Format(PRIME_SequenceNext("ORDER_LINE_ID"), "00000000")

    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Номер"), newRow).setValue(1)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_OrderID"), newRow).setString(orderId)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_LineID"), newRow).setString(lineId)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_State"), newRow).setString("")

    Dim oCursor As Object
    oCursor = ThisComponent.CurrentController
    oCursor.setActiveSheet(oSheet)
    oCursor.select(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Полное наименование товара"), newRow))
End Sub

' add_order_line: новая позиция в том же заказе - копирует общие поля заказа (source/поставщик/
' площадка/документ), но НЕ трогает склад и не пересканирует историю (must_not).
Public Sub PRIME_Orders_AddOrderLine()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)

    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim srcRow As Long
    srcRow = oSel.RangeAddress.StartRow
    If srcRow < 1 Then Exit Sub

    Dim orderId As String
    orderId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_OrderID"), srcRow).getString()
    If orderId = "" Then Exit Sub

    Dim newRow As Long
    newRow = PRIME_FindLastRow(oSheet) + 1

    Dim commonCols As Variant
    commonCols = Array("От кого / площадка", "Продавец", "Источник прихода", "№ документа", _
        "Номер счета", "Дата заказа", "Дата документа", "Покупатель")
    Dim i As Long
    For i = LBound(commonCols) To UBound(commonCols)
        Dim col As Long
        col = PRIME_ColIndex(headers, commonCols(i))
        If col >= 0 Then
            oSheet.getCellByPosition(col, newRow).setString(oSheet.getCellByPosition(col, srcRow).getString())
        End If
    Next i

    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Номер"), newRow).setValue(1)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_OrderID"), newRow).setString(orderId)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_LineID"), newRow).setString("OL-" & Format(PRIME_SequenceNext("ORDER_LINE_ID"), "00000000"))
End Sub

Public Sub PRIME_Orders_ConductSelectedButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < 1 Then Exit Sub
    PRIME_Orders_ConductRow(oSheet, row)
End Sub

' R09 (2.1.0): все строки с заполненным "Факт. количество" проводятся ОДНИМ RECEIPT-документом
' (один DOC_ID/OP_ID/один store()) - раньше N готовых строк означали N отдельных документов и N
' file store() (например, 126 строк одной поставки = 126 сохранений файла, без атомарности "весь
' приход целиком или никак"). PRIME_PostReceiptLines и так строит партию/движение на КАЖДУЮ
' строку плана - здесь просто перестаём вызывать PostDocument в цикле по одной строке.
Public Sub PRIME_Orders_ConductAllReadyButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colFact As Long
    colFact = PRIME_ColIndex(headers, "Факт. количество")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim factQtyForLine(999) As Double
    Dim newProductRowForLine(999) As Boolean
    Dim batchKey As String
    batchKey = ""
    Dim commonOrderId As String
    Dim orderIdsDiffer As Boolean
    orderIdsDiffer = False

    Dim r As Long
    For r = 1 To lastRow
        If Trim(oSheet.getCellByPosition(colFact, r).getString()) <> "" Then
            Dim docLine As PrimeDocLine
            Dim rowOrderId As String, rowDeliveryKey As String
            If PRIME_Orders_BuildReceiptLine(oSheet, headers, r, docLine, rowOrderId, rowDeliveryKey) Then
                If plan.LineCount = 0 Then
                    PRIME_InitPlan(plan, DOC_RECEIPT, SH_ORDERS, "")
                    commonOrderId = rowOrderId
                ElseIf rowOrderId <> commonOrderId Then
                    orderIdsDiffer = True
                End If
                batchKey = batchKey & rowDeliveryKey & ","
                rowForLine(plan.LineCount) = r
                factQtyForLine(plan.LineCount) = docLine.QtyInput
                newProductRowForLine(plan.LineCount) = (docLine.ProductCode = "")
                PRIME_PlanAddLine(plan, docLine)
            End If
        End If
    Next r

    If plan.LineCount = 0 Then
        MsgBox "Нет строк с заполненным ""Факт. количество""."
        Exit Sub
    End If
    plan.SourceKey = "BATCH:" & batchKey
    ' DB_PRIME_DOCUMENTS.ORDER_ID - документ-уровневое поле; если батч охватывает НЕСКОЛЬКО
    ' разных заказов сразу, честнее оставить его пустым, чем ошибочно приписать весь документ
    ' одному из них - у КАЖДОЙ строки snapshot всё равно свой правильный ORDER_ID/ORDER_LINE_ID
    ' (см. PRIME_04_Posting.PRIME_BuildOrderSnapshotRow - берёт его из самой строки "Заказы").
    If Not orderIdsDiffer Then plan.OrderId = commonOrderId

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        If newProductRowForLine(i) Then
            oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), rowForLine(i)).setString(plan.Lines(i).ProductCode)
        End If
        PRIME_Orders_ApplyReceiptResult(oSheet, headers, rowForLine(i), factQtyForLine(i), docId)
    Next i

    MsgBox "Проведено строк: " & plan.LineCount & " (документ " & docId & ")"
End Sub

' Частичная поставка (partial_receipts): каждый клик по одной строке - отдельный документ/партия
' (одна строка = один DocLine, PostDocument сам создаёт для неё ОДИН DOC_ID/OP_ID); после успеха
' очищается "Факт. количество", обновляются "Получено всего"/"Осталось получить".
Public Sub PRIME_Orders_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)

    Dim docLine As PrimeDocLine
    Dim orderId As String, deliveryKey As String
    If Not PRIME_Orders_BuildReceiptLine(oSheet, headers, row, docLine, orderId, deliveryKey) Then Exit Sub
    Dim factQty As Double
    factQty = docLine.QtyInput

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_RECEIPT, SH_ORDERS, deliveryKey)
    plan.OrderId = orderId
    PRIME_PlanAddLine(plan, docLine)

    Dim docId As String
    docId = PRIME_PostDocument(plan)

    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    ' Если код товара был пуст - PRIME_PostDocument (ByRef plan) заполнил его сгенерированным
    ' ЕИ-кодом в plan.Lines(0) - подставляем обратно в лист.
    If plan.Lines(0).ProductCode <> "" And Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), row).getString()) = "" Then
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), row).setString(plan.Lines(0).ProductCode)
    End If

    PRIME_Orders_ApplyReceiptResult(oSheet, headers, row, factQty, docId)
    MsgBox "Проведено: " & docId
End Sub

' Читает строку "Заказы" и строит PrimeDocLine для RECEIPT - используется и одиночным
' проведением (ConductRow), и батчем (ConductAllReadyButton), чтобы логика не расходилась.
' Возвращает False (с MsgBox) для строк, которые не готовы к проведению - вызывающий обязан
' пропустить такую строку, а не прерывать весь батч.
Private Function PRIME_Orders_BuildReceiptLine(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, _
        ByRef docLine As PrimeDocLine, ByRef orderId As String, ByRef deliveryKey As String) As Boolean
    Dim factStr As String
    factStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Факт. количество"), row).getString())
    If factStr = "" Or Not IsNumeric(factStr) Then
        MsgBox "Строка " & (row + 1) & ": заполните ""Факт. количество"" перед проведением."
        PRIME_Orders_BuildReceiptLine = False
        Exit Function
    End If
    Dim factQty As Double
    factQty = CDbl(factStr)
    If factQty <= 0 Then
        MsgBox "Строка " & (row + 1) & ": количество должно быть больше нуля."
        PRIME_Orders_BuildReceiptLine = False
        Exit Function
    End If

    Dim lineId As String, state As String
    orderId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_OrderID"), row).getString()
    lineId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_LineID"), row).getString()
    state = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_State"), row).getString()
    Dim deliveryId As String
    If Left(state, 4) = "DLV:" Then
        deliveryId = state
    Else
        deliveryId = "DLV:" & Format(Now, "YYYYMMDDHHMMSS") & Int(Rnd * 9999) & "-" & row
    End If
    deliveryKey = orderId & "|" & lineId & "|" & deliveryId

    docLine.ProductCode = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), row).getString())
    docLine.ProductName = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Полное наименование товара"), row).getString()
    docLine.QtyInput = factQty
    docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    docLine.LocationTo = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
    docLine.Contour = SC_GENERAL
    docLine.OrderLineId = lineId
    On Error Resume Next
    docLine.Price = CDbl(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Цена"), row).getString())
    On Error Goto 0
    PRIME_Orders_BuildReceiptLine = True
End Function

' Обновляет "Получено всего"/"Осталось получить"/"Последний приход"/"Дата последнего прихода",
' очищает "Факт. количество" и pending state, пересчитывает статус и "В наличии сейчас" (R16/R17/R22).
Private Sub PRIME_Orders_ApplyReceiptResult(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, ByVal factQty As Double, ByVal docId As String)
    Dim colReceived As Long, colRemaining As Long, colOrdered As Long
    colReceived = PRIME_ColIndex(headers, "Получено всего")
    colRemaining = PRIME_ColIndex(headers, "Осталось получить")
    colOrdered = PRIME_ColIndex(headers, "Количество")

    Dim receivedSoFar As Double
    On Error Resume Next
    receivedSoFar = CDbl(oSheet.getCellByPosition(colReceived, row).getString())
    On Error Goto 0
    receivedSoFar = receivedSoFar + factQty

    Dim orderedQty As Double
    On Error Resume Next
    orderedQty = CDbl(oSheet.getCellByPosition(colOrdered, row).getString())
    On Error Goto 0

    If colReceived >= 0 Then oSheet.getCellByPosition(colReceived, row).setValue(receivedSoFar)
    If colRemaining >= 0 Then oSheet.getCellByPosition(colRemaining, row).setValue(orderedQty - receivedSoFar)

    Dim colLastDoc As Long, colLastDate As Long
    colLastDoc = PRIME_ColIndex(headers, "Последний приход")
    colLastDate = PRIME_ColIndex(headers, "Дата последнего прихода")
    If colLastDoc >= 0 Then oSheet.getCellByPosition(colLastDoc, row).setString(docId)
    If colLastDate >= 0 Then oSheet.getCellByPosition(colLastDate, row).setString(Format(Now, "YYYY-MM-DD"))

    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Факт. количество"), row).setString("")
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_State"), row).setString("")
    PRIME_Orders_RecomputeStatus(oSheet, headers, row)
End Sub

' "Распознать по артикулам": для строк без "Код товара", но с заполненным "Код поставщика"/
' "Артикул поставщика", подставляет код через product_aliases (ambiguous_match_behavior:
' не угадывать при нескольких совпадениях - оставляет строку как есть).
Public Sub PRIME_Orders_ResolveByArticleButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colCode As Long, colPlatform As Long, colSeller As Long, colSupCode As Long, colSupArt As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colPlatform = PRIME_ColIndex(headers, "От кого / площадка")
    colSeller = PRIME_ColIndex(headers, "Продавец")
    colSupCode = PRIME_ColIndex(headers, "Код поставщика")
    colSupArt = PRIME_ColIndex(headers, "Артикул поставщика")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim resolved As Long, ambiguous As Long
    resolved = 0 : ambiguous = 0

    Dim r As Long
    For r = 1 To lastRow
        If oSheet.getCellByPosition(colCode, r).getString() = "" Then
            Dim match As String
            match = PRIME_FindProductByAlias( _
                oSheet.getCellByPosition(colPlatform, r).getString(), _
                oSheet.getCellByPosition(colSeller, r).getString(), _
                oSheet.getCellByPosition(colSupCode, r).getString(), _
                oSheet.getCellByPosition(colSupArt, r).getString())
            If match = PRIME_ALIAS_AMBIGUOUS Then
                ambiguous = ambiguous + 1
            ElseIf match <> PRIME_ALIAS_NOT_FOUND Then
                oSheet.getCellByPosition(colCode, r).setString(match)
                PRIME_Orders_AutofillByCode(oSheet, headers, r)
                resolved = resolved + 1
            End If
        End If
    Next r

    MsgBox "Распознано по артикулу: " & resolved & ". Неоднозначных (требуют ручного выбора): " & ambiguous & "."
End Sub

Public Sub PRIME_Orders_FillAllByCodeButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        PRIME_Orders_AutofillByCode(oSheet, headers, r)
    Next r
    MsgBox "Автозаполнение по кодам выполнено для " & lastRow & " строк."
End Sub

' Черновик можно удалить, только если по нему ничего ещё не проведено (Получено всего = 0/пусто) -
' удаление уже проведённых данных как обычная операция запрещено (historical_corrections).
Public Sub PRIME_Orders_DeleteDraftButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < 1 Then Exit Sub

    Dim colReceived As Long
    colReceived = PRIME_ColIndex(headers, "Получено всего")
    Dim receivedStr As String
    receivedStr = Trim(oSheet.getCellByPosition(colReceived, row).getString())
    ' 2.0.1: было "receivedStr <> "" And CDbl(receivedStr) > 0" - StarBasic не короткозамыкает
    ' And, поэтому CDbl(receivedStr) выполнялся ВСЕГДА, включая случай пустой строки, вызывая
    ' крах ("Type mismatch") при удалении черновика с ещё не заполненным "Получено всего".
    If receivedStr <> "" Then
        If IsNumeric(receivedStr) Then
            If CDbl(receivedStr) > 0 Then
                MsgBox "По этой позиции уже есть проведённые поставки. Удаление черновика запрещено - используйте корректировку."
                Exit Sub
            End If
        End If
    End If

    oSheet.getRows().removeByIndex(row, 1)
End Sub
