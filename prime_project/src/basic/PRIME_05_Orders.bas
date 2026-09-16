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
    Dim colCode As Long, colFact As Long, colState As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colFact = PRIME_ColIndex(headers, "Факт. количество")
    colState = PRIME_ColIndex(headers, "_PRIME_State")

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= 1 Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Orders_AutofillByCode(oSheet, headers, r)
                ElseIf c = colFact And colState >= 0 Then
                    PRIME_Orders_TrackPendingDelivery(oSheet, headers, r)
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
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

Public Sub PRIME_Orders_ConductAllReadyButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colFact As Long
    colFact = PRIME_ColIndex(headers, "Факт. количество")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        If Trim(oSheet.getCellByPosition(colFact, r).getString()) <> "" Then
            PRIME_Orders_ConductRow(oSheet, r)
        End If
    Next r
End Sub

' Частичная поставка (partial_receipts): каждый клик - отдельный документ/партия, после успеха
' очищается Факт.количество, обновляются Получено всего/Осталось получить.
Public Sub PRIME_Orders_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)

    Dim factStr As String
    factStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Факт. количество"), row).getString())
    If factStr = "" Or Not IsNumeric(factStr) Then
        MsgBox "Строка " & (row + 1) & ": заполните ""Факт. количество"" перед проведением."
        Exit Sub
    End If
    Dim factQty As Double
    factQty = CDbl(factStr)
    If factQty <= 0 Then
        MsgBox "Строка " & (row + 1) & ": количество должно быть больше нуля."
        Exit Sub
    End If

    Dim orderId As String, lineId As String, state As String
    orderId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_OrderID"), row).getString()
    lineId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_LineID"), row).getString()
    state = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_State"), row).getString()
    Dim deliveryId As String
    If Left(state, 4) = "DLV:" Then
        deliveryId = state
    Else
        deliveryId = "DLV:" & Format(Now, "YYYYMMDDHHMMSS")
    End If

    Dim productCode As String
    productCode = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), row).getString())
    Dim productName As String
    productName = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Полное наименование товара"), row).getString()
    Dim unit As String
    unit = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    Dim location As String
    location = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Место хранения"), row).getString()
    Dim price As Double
    On Error Resume Next
    price = CDbl(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Цена"), row).getString())
    On Error Goto 0

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_RECEIPT, SH_ORDERS, orderId & "|" & lineId & "|" & deliveryId)
    plan.OrderId = orderId

    Dim docLine As PrimeDocLine
    docLine.ProductCode = productCode
    docLine.ProductName = productName
    docLine.QtyInput = factQty
    docLine.UnitInput = unit
    docLine.LocationTo = location
    docLine.Price = price
    PRIME_PlanAddLine(plan, docLine)

    Dim docId As String
    docId = PRIME_PostDocument(plan)

    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    ' Если код товара был пуст - подставляем сгенерированный ЕИ-код обратно в лист.
    If productCode = "" Then
        Dim createdCode As String
        createdCode = plan.Lines(0).ProductCode
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код товара"), row).setString(createdCode)
    End If

    ' Обновляем "Получено всего"/"Осталось получить", очищаем "Факт. количество" и pending state.
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

    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Факт. количество"), row).setString("")
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_State"), row).setString("")

    MsgBox "Проведено: " & docId
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
    If receivedStr <> "" And CDbl(receivedStr) > 0 Then
        MsgBox "По этой позиции уже есть проведённые поставки. Удаление черновика запрещено - используйте корректировку."
        Exit Sub
    End If

    oSheet.getRows().removeByIndex(row, 1)
End Sub
