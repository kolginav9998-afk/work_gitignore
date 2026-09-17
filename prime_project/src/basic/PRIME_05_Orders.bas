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
    Dim colRowType As Long
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        ' orders_must_show_receipt_positions_directly (2.1.2): дочерняя receipt-position строка
        ' не является самостоятельной order-line - редактирование её ячеек (в т.ч. случайное) не
        ' должно запускать автозаполнение по коду/пересчёт статуса заказа.
        If r >= 1 And (colRowType < 0 Or oSheet.getCellByPosition(colRowType, r).getString() <> "CHILD") Then
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
' batch_invalid_line fix (2.1.1): заполненная, но невалидная строка (PRIME_Orders_BuildReceiptLine
' вернула False - например, некорректное количество) теперь останавливает построение всего
' батча, а не молча исключается из него (см. тот же паттерн у PRIME_Issues_ConductAllButton).
Public Sub PRIME_Orders_ConductAllReadyButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colFact As Long
    colFact = PRIME_ColIndex(headers, "Факт. количество")
    Dim colRowType As Long
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim factQtyForLine(999) As Double
    Dim batchKey As String
    batchKey = ""
    Dim commonOrderId As String
    Dim orderIdsDiffer As Boolean
    orderIdsDiffer = False

    Dim r As Long
    For r = 1 To lastRow
        ' orders_must_show_receipt_positions_directly (2.1.2): дочерние receipt-position строки
        ' никогда не участвуют как самостоятельные order-lines в проведении.
        If colRowType >= 0 And oSheet.getCellByPosition(colRowType, r).getString() = "CHILD" Then GoTo NextRow
        If Trim(oSheet.getCellByPosition(colFact, r).getString()) <> "" Then
            Dim docLine As PrimeDocLine
            Dim rowOrderId As String, rowDeliveryKey As String
            If Not PRIME_Orders_BuildReceiptLine(oSheet, headers, r, docLine, rowOrderId, rowDeliveryKey) Then
                Exit Sub
            End If
            If plan.LineCount = 0 Then
                PRIME_InitPlan(plan, DOC_RECEIPT, SH_ORDERS, "")
                commonOrderId = rowOrderId
            ElseIf rowOrderId <> commonOrderId Then
                orderIdsDiffer = True
            End If
            batchKey = batchKey & rowDeliveryKey & ","
            rowForLine(plan.LineCount) = r
            factQtyForLine(plan.LineCount) = docLine.QtyInput
            PRIME_PlanAddLine(plan, docLine)
        End If
NextRow:
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

    ' orders_partial_receipts_show_wrong_ei fix (2.1.2): обрабатываем строки СТРОГО в обратном
    ' порядке (снизу вверх) - вставка дочерней строки под rowForLine(i) физически сдвигает вниз
    ' ВСЕ строки листа ниже неё. Если бы мы шли сверху вниз, rowForLine(i+1) (вычисленный ДО
    ' любых вставок) указывал бы уже не туда после первой же вставки выше него. Вставка ниже
    ' необработанных строк их не задевает, поэтому "снизу вверх" безопасно.
    Dim i As Long
    For i = plan.LineCount - 1 To 0 Step -1
        PRIME_Orders_InsertChildReceiptRow oSheet, headers, rowForLine(i), plan.Lines(i).ProductCode, plan.DocDate, _
            factQtyForLine(i), plan.Lines(i).Contour, plan.Lines(i).LocationTo, docId
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

    ' orders_partial_receipts_show_wrong_ei fix (2.1.2): каждый приход по этой order-line
    ' получает СВОЙ новый ЕИ-код (plan.Lines(0).ProductCode, решённый внутри PRIME_PostDocument) -
    ' он больше НЕ записывается в "Код товара" родительской строки (это и вызывало старый баг,
    ' когда второй частичный приход "терял" код первого/показывал устаревший). Вместо этого для
    ' него создаётся отдельная дочерняя receipt-position строка прямо под родительской.
    PRIME_Orders_InsertChildReceiptRow oSheet, headers, row, plan.Lines(0).ProductCode, plan.DocDate, _
        factQty, plan.Lines(0).Contour, plan.Lines(0).LocationTo, docId

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

' orders_must_show_receipt_positions_directly (2.1.2): физически вставляет дочернюю
' receipt-position строку СРАЗУ под родительской order-line (после уже существующих дочерних
' строк ЭТОЙ ЖЕ строки заказа - так несколько частичных приходов стопкой идут по порядку
' получения). Каждая такая строка - представление ОДНОГО конкретного ЕИ-кода/прихода, никогда
' не участвует в повторном проведении как самостоятельная order-line (см. guard'ы по
' "_PRIME_RowType" = "CHILD" выше и в ResolveByArticle/FillAllByCode ниже).
Public Sub PRIME_Orders_InsertChildReceiptRow(ByVal oSheet As Object, ByVal headers As Variant, ByVal parentRow As Long, _
        ByVal eiCode As String, ByVal receiptDate As String, ByVal qty As Double, ByVal contour As String, _
        ByVal location As String, ByVal docId As String)
    Dim colRowType As Long, colLineId As Long, colOrderId As Long
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")
    If colRowType < 0 Then Exit Sub ' файл не проходил апгрейд схемы 2.1.2 - дочерние строки не создаём

    colLineId = PRIME_ColIndex(headers, "_PRIME_LineID")
    colOrderId = PRIME_ColIndex(headers, "_PRIME_OrderID")
    Dim parentLineId As String, parentOrderId As String
    parentLineId = ""
    parentOrderId = ""
    If colLineId >= 0 Then parentLineId = oSheet.getCellByPosition(colLineId, parentRow).getString()
    If colOrderId >= 0 Then parentOrderId = oSheet.getCellByPosition(colOrderId, parentRow).getString()

    ' Ищем место ПОСЛЕ уже существующих дочерних строк этой же order-line, чтобы вторая/третья
    ' частичная поставка добавлялась ниже первой, а не между родителем и первой дочерней строкой.
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim insertAt As Long
    insertAt = parentRow + 1
    Do While insertAt <= lastRow
        If oSheet.getCellByPosition(colRowType, insertAt).getString() <> "CHILD" Then Exit Do
        If colLineId >= 0 Then
            If oSheet.getCellByPosition(colLineId, insertAt).getString() <> parentLineId Then Exit Do
        End If
        insertAt = insertAt + 1
    Loop

    oSheet.Rows.insertByIndex(insertAt, 1)

    oSheet.getCellByPosition(colRowType, insertAt).setString("CHILD")
    If colLineId >= 0 Then oSheet.getCellByPosition(colLineId, insertAt).setString(parentLineId)
    If colOrderId >= 0 Then oSheet.getCellByPosition(colOrderId, insertAt).setString(parentOrderId)

    Dim colName As Long, colCode As Long, colRecvDate As Long, colFact As Long, colLoc As Long
    Dim colContour As Long, colDoc As Long, colIssued As Long, colReturned As Long, colAvail As Long
    colName = PRIME_ColIndex(headers, "Полное наименование товара")
    colCode = PRIME_ColIndex(headers, "Код товара")
    colRecvDate = PRIME_ColIndex(headers, "Дата поступления")
    colFact = PRIME_ColIndex(headers, "Факт. количество")
    colLoc = PRIME_ColIndex(headers, "Место хранения")
    colContour = PRIME_ColIndex(headers, "Контур")
    colDoc = PRIME_ColIndex(headers, "№ документа")
    colIssued = PRIME_ColIndex(headers, "Выдано")
    colReturned = PRIME_ColIndex(headers, "Возвращено")
    colAvail = PRIME_ColIndex(headers, "В наличии сейчас")

    If colName >= 0 Then oSheet.getCellByPosition(colName, insertAt).setString("↳ приход")
    If colCode >= 0 Then oSheet.getCellByPosition(colCode, insertAt).setString(eiCode)
    If colRecvDate >= 0 Then oSheet.getCellByPosition(colRecvDate, insertAt).setString(receiptDate)
    If colFact >= 0 Then oSheet.getCellByPosition(colFact, insertAt).setValue(qty)
    If colLoc >= 0 Then oSheet.getCellByPosition(colLoc, insertAt).setString(location)
    If colContour >= 0 Then oSheet.getCellByPosition(colContour, insertAt).setString(PRIME_ContourDisplayName(contour))
    If colDoc >= 0 Then oSheet.getCellByPosition(colDoc, insertAt).setString(docId)
    ' Изначально ничего не выдано/не возвращено, остаток = только что полученное количество -
    ' PRIME_Orders_RefreshChildRow пересчитает эти три поля точно по ledger при следующей операции.
    If colIssued >= 0 Then oSheet.getCellByPosition(colIssued, insertAt).setValue(0)
    If colReturned >= 0 Then oSheet.getCellByPosition(colReturned, insertAt).setValue(0)
    If colAvail >= 0 Then oSheet.getCellByPosition(colAvail, insertAt).setValue(qty)

    PRIME_Orders_StyleChildRow(oSheet, headers, insertAt)
End Sub

' Визуально отличаем дочернюю receipt-position строку от обычной order-line (курсив + светло-
' серый фон на всю строку) - требование orders_must_show_receipt_positions_directly.
Public Sub PRIME_Orders_StyleChildRow(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(0, row, UBound(headers), row)
    oRange.CharPosture = com.sun.star.awt.FontSlant.ITALIC
    oRange.CellBackColor = RGB(240, 240, 240)
End Sub

' Пересчитывает "В наличии сейчас"/"Выдано"/"Возвращено" одной дочерней receipt-position строки
' по актуальному COMMITTED-ledger (R02 - остаток строго по товару/месту/контуру, контур для
' "Заказы" всегда SC_GENERAL - см. PRIME_Orders_RefreshAvailability). "Выдано"/"Возвращено" -
' по LOT_ID, связанному с ЕИ-кодом этой строки (2.1.1: 1 ЕИ-код = 1 партия).
Public Sub PRIME_Orders_RefreshChildRow(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colCode As Long, colLoc As Long, colAvail As Long, colIssued As Long, colReturned As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colLoc = PRIME_ColIndex(headers, "Место хранения")
    colAvail = PRIME_ColIndex(headers, "В наличии сейчас")
    colIssued = PRIME_ColIndex(headers, "Выдано")
    colReturned = PRIME_ColIndex(headers, "Возвращено")
    If colCode < 0 Then Exit Sub

    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Then Exit Sub

    Dim loc As String
    loc = ""
    If colLoc >= 0 Then loc = Trim(oSheet.getCellByPosition(colLoc, row).getString())

    If colAvail >= 0 Then
        oSheet.getCellByPosition(colAvail, row).setValue(PRIME_LocationContourBalance(code, loc, SC_GENERAL))
    End If

    Dim lotId As String
    lotId = PRIME_LotIdForProductCode(code)
    If colIssued >= 0 Then oSheet.getCellByPosition(colIssued, row).setValue(PRIME_IssuedQtyForLot(lotId))
    If colReturned >= 0 Then oSheet.getCellByPosition(colReturned, row).setValue(PRIME_ReturnedQtyForLot(lotId))
End Sub

' "Заказы" не получает событий с других листов (Выдачи/Возвраты/Перемещения/Инвентаризация) -
' вызывается из PRIME_04_Posting.PRIME_PostDocument ПОСЛЕ store(), чтобы дочерние строки, чей
' ЕИ-код затронут только что проведённым документом, отразили актуальный остаток/выдано/
' возвращено сразу же (manual_acceptance: "остаток на Заказы обновляется живьём после выдачи").
Public Sub PRIME_Orders_RefreshChildRowsForPlan(ByRef plan As PrimeDocPlan)
    If Not PRIME_SheetExists(SH_ORDERS) Then Exit Sub
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colRowType As Long, colCode As Long
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")
    colCode = PRIME_ColIndex(headers, "Код товара")
    If colRowType < 0 Or colCode < 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        If oSheet.getCellByPosition(colRowType, r).getString() = "CHILD" Then
            Dim rowCode As String
            rowCode = Trim(oSheet.getCellByPosition(colCode, r).getString())
            If rowCode <> "" Then
                Dim i As Long
                For i = 0 To plan.LineCount - 1
                    If plan.Lines(i).ProductCode = rowCode Then
                        PRIME_Orders_RefreshChildRow(oSheet, headers, r)
                        Exit For
                    End If
                Next i
            End If
        End If
    Next r
End Sub

' "Распознать по артикулам": для строк без "Код товара", но с заполненным "Код поставщика"/
' "Артикул поставщика", подставляет код через product_aliases (ambiguous_match_behavior:
' не угадывать при нескольких совпадениях - оставляет строку как есть).
Public Sub PRIME_Orders_ResolveByArticleButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colCode As Long, colPlatform As Long, colSeller As Long, colSupCode As Long, colSupArt As Long, colRowType As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colPlatform = PRIME_ColIndex(headers, "От кого / площадка")
    colSeller = PRIME_ColIndex(headers, "Продавец")
    colSupCode = PRIME_ColIndex(headers, "Код поставщика")
    colSupArt = PRIME_ColIndex(headers, "Артикул поставщика")
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim resolved As Long, ambiguous As Long
    resolved = 0 : ambiguous = 0

    Dim r As Long
    For r = 1 To lastRow
        ' orders_must_show_receipt_positions_directly (2.1.2): дочерняя receipt-position строка
        ' не является самостоятельной order-line - никогда не подставляем ей код по артикулу.
        If (colRowType < 0 Or oSheet.getCellByPosition(colRowType, r).getString() <> "CHILD") And _
                oSheet.getCellByPosition(colCode, r).getString() = "" Then
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
    Dim colRowType As Long
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        ' orders_must_show_receipt_positions_directly (2.1.2): дочерние receipt-position строки
        ' не автозаполняются как обычные order-lines.
        If colRowType < 0 Or oSheet.getCellByPosition(colRowType, r).getString() <> "CHILD" Then
            PRIME_Orders_AutofillByCode(oSheet, headers, r)
        End If
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
