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



' PRIME_06_Issues
' Лист "Выдачи": 12 исходных бизнес-колонок 1.4.1 без изменений + "Назначение / проект"
' (add_fields) + 1 скрытая helper-колонка _PRIME_IssueState (стабильный ISSUE_DRAFT_ID).
' Возврат НЕ обрабатывается здесь - единственный механизм возврата в PRIME_08_ReturnsInventory
' (закрывает конфликт старого/нового возврата 1.4.1, см. ARCHITECTURE §0).

Public Sub PRIME_OnContentChanged_Issues(ByVal oRangeAddr As Variant)
    If Not PRIME_EventEnter() Then Exit Sub
    On Error Goto CleanExit

    If IsNull(oRangeAddr) Then GoTo CleanExit

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Код")

    Dim colQty As Long
    colQty = PRIME_ColIndex(headers, "Кол-во")

    Dim r As Long, c As Long
    For r = oRangeAddr.StartRow To oRangeAddr.EndRow
        If r >= 1 Then
            For c = oRangeAddr.StartColumn To oRangeAddr.EndColumn
                If c = colCode Then
                    PRIME_Issues_AutofillByCode(oSheet, headers, r)
                    PRIME_Issues_RefreshInlineStock(oSheet, headers, r)
                ElseIf c = colQty Then
                    PRIME_Issues_RefreshInlineStock(oSheet, headers, r)
                End If
            Next c
        End If
    Next r

CleanExit:
    PRIME_EventLeave()
End Sub

' R23 (2.1.0): "В наличии" -> "Кол-во" -> "После выдачи" - представление того же
' COMMITTED-ledger, что и лист "Наличие". single_physical_warehouse (FINAL): один физический
' склад - "Выдачи" обязаны разрешать ЛЮБОЙ валидный EI_CODE независимо от того, пришёл он по
' Заказам, из Офиса, из Производства или как Деталь - контур больше не сужает физическую
' доступность (см. PRIME_04_Posting.PRIME_LocationContourBalance).
Private Sub PRIME_Issues_RefreshInlineStock(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colAvail As Long, colAfter As Long, colCode As Long, colFrom As Long, colQty As Long
    colAvail = PRIME_ColIndex(headers, "В наличии")
    colAfter = PRIME_ColIndex(headers, "После выдачи")
    colCode = PRIME_ColIndex(headers, "Код")
    colFrom = PRIME_ColIndex(headers, "Откуда")
    colQty = PRIME_ColIndex(headers, "Кол-во")
    If colAvail < 0 And colAfter < 0 Then Exit Sub

    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    Dim loc As String
    loc = ""
    If colFrom >= 0 Then loc = Trim(oSheet.getCellByPosition(colFrom, row).getString())
    Dim available As Double
    available = PRIME_LocationContourBalance(code, loc, "")
    If colAvail >= 0 Then oSheet.getCellByPosition(colAvail, row).setValue(available)

    If colAfter >= 0 Then
        Dim qtyStr As String
        qtyStr = Trim(oSheet.getCellByPosition(colQty, row).getString())
        If qtyStr <> "" And IsNumeric(qtyStr) Then
            oSheet.getCellByPosition(colAfter, row).setValue(available - CDbl(qtyStr))
        End If
    End If
End Sub

' live_code_lookup.Выдачи: Наименование, Ед.изм., Возвратный, Дата + location_rule для "Откуда".
Private Sub PRIME_Issues_AutofillByCode(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Код")
    Dim code As String
    code = Trim(oSheet.getCellByPosition(colCode, row).getString())
    If code = "" Or Not PRIME_ProductExists(code) Then Exit Sub

    PRIME_SetCellIfEmpty(oSheet, headers, row, "Наименование", PRIME_GetProductField(code, "PRODUCT_NAME"))
    PRIME_SetCellIfEmpty(oSheet, headers, row, "Ед. изм.", PRIME_GetProductField(code, "BASE_UNIT"))

    Dim colReturnable As Long
    colReturnable = PRIME_ColIndex(headers, "Возвратный")
    If colReturnable >= 0 And oSheet.getCellByPosition(colReturnable, row).getString() = "" Then
        Dim ret As String
        ret = PRIME_GetProductField(code, "RETURNABLE")
        If ret = "1" Then
            oSheet.getCellByPosition(colReturnable, row).setString("Да")
        ElseIf ret = "0" Then
            oSheet.getCellByPosition(colReturnable, row).setString("Нет")
        End If
    End If

    Dim colDate As Long
    colDate = PRIME_ColIndex(headers, "Дата")
    If colDate >= 0 And oSheet.getCellByPosition(colDate, row).getString() = "" Then
        oSheet.getCellByPosition(colDate, row).setString(Format(Now, "YYYY-MM-DD"))
    End If

    ' location_rule: подставить "Откуда" только если положительный остаток ровно в одном месте.
    Dim colFrom As Long
    colFrom = PRIME_ColIndex(headers, "Откуда")
    If colFrom >= 0 And oSheet.getCellByPosition(colFrom, row).getString() = "" Then
        Dim singleLoc As String
        singleLoc = PRIME_SingleLocationWithStock(code)
        If singleLoc <> "" Then
            oSheet.getCellByPosition(colFrom, row).setString(singleLoc)
        End If
    End If
End Sub

' === Кнопки =====================================================================================

Public Sub PRIME_Issues_NewIssueButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)

    Dim newRow As Long
    newRow = PRIME_FindLastRow(oSheet) + 1
    If newRow < 1 Then newRow = 1

    oSheet.getCellByPosition(PRIME_ColIndex(headers, "№"), newRow).setValue(newRow)
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_IssueState"), newRow).setString("ISS-" & Format(PRIME_SequenceNext("ISSUE_DRAFT_ID"), "00000000"))
    oSheet.getCellByPosition(PRIME_ColIndex(headers, "Дата"), newRow).setString(Format(Now, "YYYY-MM-DD"))

    ThisComponent.CurrentController.setActiveSheet(oSheet)
    ThisComponent.CurrentController.select(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код"), newRow))
End Sub

Public Sub PRIME_Issues_ConductSelectedButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    Dim row As Long
    row = oSel.RangeAddress.StartRow
    If row < 1 Then Exit Sub
    PRIME_Issues_ConductRow(oSheet, row)
End Sub

' R09/R10 (2.1.0): все строки с заполненным кодом проводятся ОДНИМ ISSUE-документом (один
' DOC_ID/OP_ID/один store()) - раньше N строк выдачи означали N отдельных документов. Строки
' одного товара из одного места теперь резервируют остаток друг у друга внутри одного плана
' (см. PRIME_04_Posting.PRIME_ValidateIssue) - две строки по 6 при остатке 10 отклоняются ВЕСЬ
' документ, а не проходят обе независимо.
' batch_invalid_line fix (2.1.1): раньше строка с заполненным "Код", но невалидная (например,
' пустое "Кол-во"), просто МОЛЧА (не считая одного MsgBox) пропускалась - весь батч без неё всё
' равно проводился, оставляя у пользователя ложное впечатление, что обработаны ВСЕ заполненные
' строки. Теперь любая заполненная, но не прошедшая BuildLine строка (кроме уже DONE - та
' пропускается штатно, это не ошибка) останавливает построение ВСЕГО батча целиком - ничего не
' проводится, пока пользователь не поправит или не очистит эту строку.
Public Sub PRIME_Issues_ConductAllButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)
    Dim colCode As Long
    colCode = PRIME_ColIndex(headers, "Код")
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_IssueState")
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)

    Dim plan As PrimeDocPlan
    Dim rowForLine(999) As Long
    Dim batchKey As String
    batchKey = ""

    Dim r As Long
    For r = 1 To lastRow
        If Trim(oSheet.getCellByPosition(colCode, r).getString()) <> "" Then
            If Left(oSheet.getCellByPosition(colState, r).getString(), 5) <> "DONE:" Then
                Dim docLine As PrimeDocLine
                Dim rowKey As String
                If Not PRIME_Issues_BuildLine(oSheet, headers, r, docLine, rowKey) Then
                    MsgBox "Проведение не выполнено: строка " & (r + 1) & " заполнена, но невалидна. Исправьте её или очистите перед проведением всего батча."
                    Exit Sub
                End If
                If plan.LineCount = 0 Then PRIME_InitPlan(plan, DOC_ISSUE, SH_ISSUES, "")
                batchKey = batchKey & rowKey & ","
                rowForLine(plan.LineCount) = r
                PRIME_PlanAddLine(plan, docLine)
            End If
        End If
    Next r

    If plan.LineCount = 0 Then Exit Sub
    plan.SourceKey = "BATCH:" & batchKey

    Dim docId As String
    docId = PRIME_PostDocument(plan)
    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To plan.LineCount - 1
        PRIME_Issues_RefreshInlineStock(oSheet, headers, rowForLine(i))
        oSheet.getCellByPosition(colState, rowForLine(i)).setString("DONE:" & _
            oSheet.getCellByPosition(colState, rowForLine(i)).getString())
    Next i
    MsgBox "Проведено строк выдачи: " & plan.LineCount & " (документ " & docId & ")"
End Sub

Public Sub PRIME_Issues_ConductRow(ByVal oSheet As Object, ByVal row As Long)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)

    Dim docLine As PrimeDocLine
    Dim draftId As String
    If Not PRIME_Issues_BuildLine(oSheet, headers, row, docLine, draftId) Then Exit Sub

    Dim plan As PrimeDocPlan
    PRIME_InitPlan(plan, DOC_ISSUE, SH_ISSUES, draftId)
    PRIME_PlanAddLine(plan, docLine)

    Dim docId As String
    docId = PRIME_PostDocument(plan)

    If docId = "" Then
        MsgBox "Проведение не выполнено: " & PRIME_LastPostError()
        Exit Sub
    End If

    PRIME_Issues_RefreshInlineStock(oSheet, headers, row)
    Dim colState As Long
    colState = PRIME_ColIndex(headers, "_PRIME_IssueState")
    oSheet.getCellByPosition(colState, row).setString("DONE:" & oSheet.getCellByPosition(colState, row).getString())
    MsgBox "Выдача проведена: " & docId
End Sub

' Читает строку "Выдачи" и строит PrimeDocLine - общий для одиночного и батч-проведения.
' Возвращает False (с MsgBox) для неготовых строк - вызывающий обязан пропустить такую строку,
' а не прерывать весь батч.
' batch_duplicate_posting fix (2.1.1): раньше _PRIME_IssueState никогда не помечался как
' "проведено" после успешного commit - строка, уже один раз успешно выданная, оставалась
' с непустым "Код" и стабильным draftId НАВСЕГДА, поэтому следующее "Провести все" (после
' добавления новых строк ниже) снова подхватывало её в НОВЫЙ батч. Комбинированный SOURCE_KEY
' такого батча ("BATCH:<старый_draftId>,<новый_draftId>,") ни разу раньше не встречался, поэтому
' PRIME_FindCommittedBySourceKey не находил совпадение - и вся строка проводилась ПОВТОРНО,
' списывая остаток дважды. Теперь успешно проведённая строка помечается префиксом "DONE:" и
' здесь же исключается из дальнейшего включения в любой план.
Private Function PRIME_Issues_BuildLine(ByVal oSheet As Object, ByVal headers As Variant, ByVal row As Long, _
        ByRef docLine As PrimeDocLine, ByRef draftId As String) As Boolean
    draftId = oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_IssueState"), row).getString()
    If Left(draftId, 5) = "DONE:" Then
        PRIME_Issues_BuildLine = False
        Exit Function
    End If
    If draftId = "" Then
        draftId = "ISS-" & Format(PRIME_SequenceNext("ISSUE_DRAFT_ID"), "00000000")
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_IssueState"), row).setString(draftId)
    End If

    Dim code As String
    code = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код"), row).getString())
    If code = "" Then
        MsgBox "Строка " & (row + 1) & ": не указан код товара."
        PRIME_Issues_BuildLine = False
        Exit Function
    End If

    Dim qtyStr As String
    qtyStr = Trim(oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кол-во"), row).getString())
    If qtyStr = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Строка " & (row + 1) & ": заполните корректное ""Кол-во""."
        PRIME_Issues_BuildLine = False
        Exit Function
    End If

    docLine.ProductCode = code
    docLine.QtyInput = CDbl(qtyStr)
    docLine.UnitInput = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), row).getString()
    docLine.LocationFrom = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Откуда"), row).getString()
    docLine.Contour = SC_GENERAL
    docLine.Recipient = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кто получил"), row).getString()
    docLine.DestinationProject = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Назначение / проект"), row).getString()
    docLine.Comment = oSheet.getCellByPosition(PRIME_ColIndex(headers, "Примечание"), row).getString()
    PRIME_Issues_BuildLine = True
End Function

Public Sub PRIME_Issues_FillAllByCodeButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)
    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        PRIME_Issues_AutofillByCode(oSheet, headers, r)
    Next r
    MsgBox "Автозаполнение по кодам выполнено для " & lastRow & " строк."
End Sub

' Пересчитывает "Возвращено"/"Дата возврата" по факту записей DB_PRIME_RETURNS - чисто отображение,
' не пишет в движения (диагностика/обновление, а не проведение).
Public Sub PRIME_Issues_RefreshReturnedColumnButton()
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)
    Dim colState As Long, colReturned As Long, colReturnDate As Long
    colState = PRIME_ColIndex(headers, "_PRIME_IssueState")
    colReturned = PRIME_ColIndex(headers, "Возвращено")
    colReturnDate = PRIME_ColIndex(headers, "Дата возврата")

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long
    For r = 1 To lastRow
        Dim draftId As String
        draftId = oSheet.getCellByPosition(colState, r).getString()
        If draftId <> "" Then
            Dim docId As String
            docId = PRIME_FindCommittedBySourceKey(draftId)
            If docId <> "" Then
                Dim lineId As String
                lineId = docId & "-L1"
                Dim returned As Double
                returned = PRIME_AlreadyReturnedQtyBase(lineId)
                If colReturned >= 0 Then oSheet.getCellByPosition(colReturned, r).setValue(returned)
            End If
        End If
    Next r
End Sub
