Option Explicit

' PRIME_11_Kits
' required_mode = "issue_bundle" для первого стабильного релиза: комплект - это просто набор
' обычных строк расхода компонентов, без виртуального остатка комплекта
' (virtual_kit_stock_created=false). Нехватка любого компонента отменяет добавление целиком
' (partial_issue_if_component_missing=false).
'
' Физическая сборка/разборка (ASSEMBLY/DISASSEMBLY) архитектурно поддержаны в PRIME_04_Posting
' (document_types включает оба типа), но конкретный UI для них НЕ реализован в этой версии -
' mandatory_for_first_stable_release=false. Это осознанное ограничение, а не недоделка: см.
' TEST_REPORT/финальный отчёт, known_limitations.

Public Sub PRIME_Kits_AddToIssuesButton()
    Dim kitId As String
    kitId = InputBox("KIT_ID комплекта для выдачи (см. лист ""Комплекты""):", "Добавить комплект")
    If Trim(kitId) = "" Then Exit Sub
    kitId = Trim(kitId)

    Dim qtyStr As String
    qtyStr = InputBox("Количество комплектов:", "Добавить комплект", "1")
    If Trim(qtyStr) = "" Or Not IsNumeric(qtyStr) Or CDbl(qtyStr) <= 0 Then
        MsgBox "Некорректное количество комплектов."
        Exit Sub
    End If
    Dim kitQty As Double
    kitQty = CDbl(qtyStr)

    Dim componentCodes() As String
    Dim componentQtyPerKit() As Double
    Dim componentUnits() As String
    Dim n As Long
    n = PRIME_Kits_ReadLines(kitId, componentCodes, componentUnits, componentQtyPerKit)

    If n = 0 Then
        MsgBox "Комплект " & kitId & " не найден, не активен или не содержит компонентов."
        Exit Sub
    End If

    ' Проверка наличия ВСЕХ компонентов ДО добавления хотя бы одной строки (partial_issue_if_component_missing=false).
    Dim i As Long
    Dim shortages As String
    shortages = ""
    For i = 0 To n - 1
        Dim needQty As Variant
        needQty = PRIME_ConvertQtyToBase(componentCodes(i), componentUnits(i), componentQtyPerKit(i) * kitQty)
        If IsEmpty(needQty) Then
            shortages = shortages & componentCodes(i) & " (нет коэффициента единицы " & componentUnits(i) & ")" & Chr(10)
        ElseIf PRIME_TotalLotBalance(componentCodes(i)) < CDbl(needQty) Then
            shortages = shortages & componentCodes(i) & " (нужно " & CDbl(needQty) & ", есть " & PRIME_TotalLotBalance(componentCodes(i)) & ")" & Chr(10)
        End If
    Next i

    If shortages <> "" Then
        MsgBox "Комплект не добавлен - недостаточно компонентов:" & Chr(10) & shortages
        Exit Sub
    End If

    ' Все компоненты в наличии - добавляем обычные строки расхода на лист "Выдачи".
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ISSUES)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ISSUES)

    For i = 0 To n - 1
        Dim newRow As Long
        newRow = PRIME_FindLastRow(oSheet) + 1
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "№"), newRow).setValue(newRow)
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Код"), newRow).setString(componentCodes(i))
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Наименование"), newRow).setString(PRIME_GetProductField(componentCodes(i), "PRODUCT_NAME"))
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Кол-во"), newRow).setValue(componentQtyPerKit(i) * kitQty)
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Ед. изм."), newRow).setString(componentUnits(i))
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Дата"), newRow).setString(Format(Now, "YYYY-MM-DD"))
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "Примечание"), newRow).setString("Комплект " & kitId & " x" & kitQty)
        oSheet.getCellByPosition(PRIME_ColIndex(headers, "_PRIME_IssueState"), newRow).setString("ISS-" & Format(PRIME_SequenceNext("ISSUE_DRAFT_ID"), "00000000"))
    Next i

    MsgBox "Добавлено строк расхода: " & n & ". Проверьте ""Откуда""/""Кто получил"" и проведите как обычную выдачу."
End Sub

' Читает компоненты комплекта из плоского листа "Комплекты" (одна строка = один компонент).
Private Function PRIME_Kits_ReadLines(ByVal kitId As String, ByRef codes() As String, ByRef units() As String, ByRef qtyPerKit() As Double) As Long
    If Not PRIME_SheetExists(SH_KITS) Then
        ReDim codes(-1) : ReDim units(-1) : ReDim qtyPerKit(-1)
        PRIME_Kits_ReadLines = 0
        Exit Function
    End If

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_KITS)
    Dim colKitId As Long, colProduct As Long, colQty As Long, colUnit As Long, colActive As Long
    colKitId = PRIME_ColIndex(headers, "KIT_ID")
    colProduct = PRIME_ColIndex(headers, "PRODUCT_CODE")
    colQty = PRIME_ColIndex(headers, "Количество на 1 комплект")
    colUnit = PRIME_ColIndex(headers, "Единица")
    colActive = PRIME_ColIndex(headers, "Активен")

    Dim table As Variant
    table = PRIME_ReadTable(SH_KITS)
    If UBound(table) < 1 Then
        ReDim codes(-1) : ReDim units(-1) : ReDim qtyPerKit(-1)
        PRIME_Kits_ReadLines = 0
        Exit Function
    End If

    ReDim codes(UBound(table))
    ReDim units(UBound(table))
    ReDim qtyPerKit(UBound(table))
    Dim n As Long
    n = 0
    Dim i As Long
    For i = 1 To UBound(table)
        If CStr(table(i)(colKitId)) = kitId Then
            Dim activeVal As String
            activeVal = CStr(table(i)(colActive))
            If activeVal = "1" Or LCase(activeVal) = "да" Or LCase(activeVal) = "true" Then
                codes(n) = CStr(table(i)(colProduct))
                units(n) = CStr(table(i)(colUnit))
                qtyPerKit(n) = CDbl(table(i)(colQty))
                n = n + 1
            End If
        End If
    Next i

    If n = 0 Then
        ReDim codes(-1) : ReDim units(-1) : ReDim qtyPerKit(-1)
    Else
        ReDim Preserve codes(n - 1)
        ReDim Preserve units(n - 1)
        ReDim Preserve qtyPerKit(n - 1)
    End If
    PRIME_Kits_ReadLines = n
End Function
