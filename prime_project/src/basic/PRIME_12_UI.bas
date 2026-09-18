Option Explicit

' PRIME_12_UI
' Единственный движок раскладки интерфейса (в 1.4.1 их было три разных - модули 19/31/37
' с разными списками листов). Команда "Восстановить интерфейс PRIME" НЕ вызывается во время
' обычного проведения (rebuild_interface_during_normal_posting=false) - только вручную.

' Легаси-архивный лист ("Расход — Производство/Детали" с 2.1.2 - см. авторитетный список
' видимых листов master task 2.1.2: у приходов на эти контуры уже есть активные листы
' "Приход — Производство"/"Приход — Детали", у расходов активного двойника пока нет) -
' только история, кнопки неактивны.
Public Sub PRIME_Legacy_ArchiveStub()
    MsgBox "Этот лист - архив истории версии 1.4.1, доступен только для чтения." & Chr(10) & _
        "Приход по этому контуру ведите на листе ""Приход — Производство"" или ""Приход — Детали""."
End Sub

' --- Навигация (кнопки-ярлыки на Инфо/Дашборде/Отчёте, просто переключают активный лист) -----
Public Sub PRIME_Nav_Dashboard()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_DASHBOARD))
End Sub

Public Sub PRIME_Nav_Report()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_REPORT_FINAL))
End Sub

Public Sub PRIME_Nav_ReportInput()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_REPORT_INPUT))
End Sub

Public Sub PRIME_Nav_Stock()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_STOCK))
End Sub

Public Sub PRIME_Nav_Orders()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_ORDERS))
End Sub

Public Sub PRIME_Nav_Search()
    ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_SEARCH))
End Sub

Public Sub PRIME_UI_RestoreInterfaceButton()
    PRIME_UI_RestoreInterfaceSilent()
    MsgBox "Интерфейс PRIME восстановлен."
End Sub

' Без MsgBox - для сборщика (tools/build_ods.py) и вызова из PRIME_Build_RunFullSetup.
Public Sub PRIME_UI_RestoreInterfaceSilent()
    PRIME_UI_ApplySheetVisibility()
    PRIME_UI_ApplyFreezeAndFilters()
    PRIME_UI_FixInfoPanelText()
    PRIME_Home_RefreshButton()
End Sub

' remove_text_references (PRIME 2.0.1): лист "Инфо" унаследован от шаблона 1.4.1 со статическим
' текстом ("Режим: PRODUCTION", "Релиз: 2.0.0", "Firebird: OK - embedded Firebird"), который
' никогда не обновлялся ни одним PRIME-модулем - после первой сборки 2.0.0 он навсегда оставался
' неверным (в частности, прямо утверждал зависимость от Firebird, которой в PRIME нет). Правим
' по совпадению подписи в колонке A, чтобы не сломаться, если раскладка панели когда-то изменится.
' Optional-с-значением-по-умолчанию (VBA-style) не компилируется в этой версии StarBasic -
' молча ломает компиляцию ВСЕЙ библиотеки Standard без видимого исключения (тот же класс
' проблемы, что и другие недокументированные ограничения StarBasic, см. историю отладки) -
' поэтому здесь просто константа, а не Optional-параметр.
Private Const PRIME_INFO_PANEL_MAX_ROW As Long = 20

Private Sub PRIME_UI_FixInfoPanelText()
    If Not PRIME_SheetExists(SH_INFO) Then Exit Sub
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_INFO)

    Dim r As Long
    For r = 0 To PRIME_INFO_PANEL_MAX_ROW
        Dim label As String
        label = Trim(oSheet.getCellByPosition(0, r).getString())
        Select Case label
            Case "Релиз"
                oSheet.getCellByPosition(1, r).setString(PRIME_SCHEMA_VERSION)
            Case "Firebird"
                oSheet.getCellByPosition(0, r).setString("Хранилище")
                oSheet.getCellByPosition(1, r).setString("Calc-only, без Firebird/Base")
        End Select
    Next r
End Sub

' 2.1.2 (авторитетный список видимых листов, master task): переписано с "перечисли, что скрыть"
' на "перечисли, что показать, скрой всё остальное" - и надёжнее (не пропустит какой-нибудь
' унаследованный от 1.4.1 лист, о котором забыли явно написать HIDE), и буквально соответствует
' формулировке задания "must be hidden... unless one of the exact sheets above is the active
' replacement". Единственные пользовательские листы, которые должны остаться видимыми:
Private Function PRIME_UI_VisibleSheetNames() As Variant
    PRIME_UI_VisibleSheetNames = Array(SH_HOME, SH_ORDERS, SH_RECEIPT_OFFICE, SH_RECEIPT_PRODUCTION, _
        SH_RECEIPT_DETAILS, SH_ISSUES, SH_RETURNS, SH_TRANSFERS, SH_INVENTORY, SH_STOCK, SH_SEARCH, _
        SH_JOURNAL, SH_KITS)
End Function

Private Function PRIME_UI_IsAllowedVisible(ByVal sheetName As String, ByVal allowList As Variant) As Boolean
    Dim i As Long
    For i = LBound(allowList) To UBound(allowList)
        If allowList(i) = sheetName Then
            PRIME_UI_IsAllowedVisible = True
            Exit Function
        End If
    Next i
    PRIME_UI_IsAllowedVisible = False
End Function

Private Sub PRIME_UI_ApplySheetVisibility()
    Dim allowList As Variant
    allowList = PRIME_UI_VisibleSheetNames()

    ' Calc не разрешает скрыть текущий активный лист - переключаемся на заведомо видимый ("Главная")
    ' ДО скрытия остальных, чтобы порядок перебора листов ниже не наткнулся на активный.
    If PRIME_SheetExists(SH_HOME) Then
        ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(SH_HOME))
    End If

    Dim oSheets As Object
    oSheets = ThisComponent.Sheets
    Dim i As Long
    For i = 0 To oSheets.Count - 1
        Dim oSheet As Object
        oSheet = oSheets.getByIndex(i)
        If PRIME_UI_IsAllowedVisible(oSheet.Name, allowList) Then
            oSheet.IsVisible = True
        Else
            On Error Resume Next ' активный лист сменить некому, если "Главная" ещё не создана - не падаем
            oSheet.IsVisible = False
            On Error Goto 0
        End If
    Next i

    ' Легаси-архивные листы (Расход — Производство/Детали) - только история, read-only даже
    ' если случайно вручную раскрыты (legacy_parallel_forms: hide/read-only, не активная форма).
    Dim legacySheets As Variant
    legacySheets = Array(SH_LEGACY_ISSUE_PROD, SH_LEGACY_ISSUE_PARTS)
    For i = LBound(legacySheets) To UBound(legacySheets)
        If PRIME_SheetExists(legacySheets(i)) Then
            PRIME_GetSheet(legacySheets(i)).IsProtected = True
        End If
    Next i
End Sub

Private Sub PRIME_UI_ApplyFreezeAndFilters()
    Dim dataSheets As Variant
    dataSheets = Array(SH_ORDERS, SH_ISSUES, SH_RECEIPT_SHOP, SH_ISSUE_SHOP, SH_RECEIPT_OFFICE, _
        SH_ISSUE_OFFICE, SH_RECEIPT_PRODUCTION, SH_RECEIPT_DETAILS, SH_RETURNS, SH_INVENTORY, SH_KITS)
    Dim oController As Object
    oController = ThisComponent.CurrentController

    Dim i As Long
    For i = LBound(dataSheets) To UBound(dataSheets)
        If PRIME_SheetExists(dataSheets(i)) Then
            Dim oSh As Object
            oSh = PRIME_GetSheet(dataSheets(i))

            Dim headerRow As Long, firstDataRow As Long
            headerRow = PRIME_FormSchemaHeaderRow(dataSheets(i))
            firstDataRow = headerRow + 1

            oController.setActiveSheet(oSh)
            ' Заморозить всё до и включая строку заголовка (freeze_headers) - для листов с
            ' декоративной панелью над таблицей (header_schema_registry) это не всегда строка 1.
            oController.freezeAtPosition(0, firstDataRow)

            Dim lastRow As Long, lastCol As Long
            lastRow = PRIME_FindLastRow(oSh)
            lastCol = PRIME_FindLastCol(oSh)
            If lastRow < headerRow Then lastRow = headerRow

            Dim oRange As Object
            oRange = oSh.getCellRangeByPosition(0, headerRow, lastCol, lastRow)

            ' Автофильтр (autofilter=true) через именованный диапазон базы данных.
            Dim rangeName As String
            rangeName = "PRIME_FILTER_" & i
            Dim oDBRanges As Object
            oDBRanges = ThisComponent.DatabaseRanges
            If oDBRanges.hasByName(rangeName) Then
                oDBRanges.removeByName(rangeName)
            End If
            oDBRanges.addNewByName(rangeName, oRange.RangeAddress)
            oDBRanges.getByName(rangeName).AutoFilter = True

            ' Видимые границы (visible_borders) сеткой по всему диапазону данных.
            Dim oBorder As New com.sun.star.table.BorderLine2
            oBorder.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
            oBorder.LineWidth = 18
            Dim oTableBorder As New com.sun.star.table.TableBorder2
            oTableBorder.TopLine = oBorder : oTableBorder.IsTopLineValid = True
            oTableBorder.BottomLine = oBorder : oTableBorder.IsBottomLineValid = True
            oTableBorder.LeftLine = oBorder : oTableBorder.IsLeftLineValid = True
            oTableBorder.RightLine = oBorder : oTableBorder.IsRightLineValid = True
            oTableBorder.HorizontalLine = oBorder : oTableBorder.IsHorizontalLineValid = True
            oTableBorder.VerticalLine = oBorder : oTableBorder.IsVerticalLineValid = True
            oRange.TableBorder2 = oTableBorder
        End If
    Next i
End Sub
