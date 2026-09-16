Option Explicit

' PRIME_12_UI
' Единственный движок раскладки интерфейса (в 1.4.1 их было три разных - модули 19/31/37
' с разными списками листов). Команда "Восстановить интерфейс PRIME" НЕ вызывается во время
' обычного проведения (rebuild_interface_during_normal_posting=false) - только вручную.

' Заглушка для кнопок 1.4.1, чья функциональность сознательно не перенесена в первый релиз
' PRIME 2.0.0 (см. known_limitations в финальном отчёте) - вместо тихого отказа или ошибки
' обращения к несуществующей Firebird-базе, пользователь получает понятное объяснение.
Public Sub PRIME_UI_NotImplementedStub()
    MsgBox "Эта функция версии 1.4.1 не перенесена в первый стабильный релиз PRIME 2.0.0." & Chr(10) & _
        "См. CHANGELOG_PRIME.md / TEST_REPORT.md, раздел ""Известные ограничения""."
End Sub

' Старая кнопка возврата на листе "Выдачи" удалена как отдельный механизм (конфликт двух
' реализаций возврата в 1.4.1, см. ARCHITECTURE §0) - редирект вместо тихого/ошибочного действия.
Public Sub PRIME_Issues_ReturnRedirectStub()
    MsgBox "Возврат теперь оформляется только на листе ""Возвраты"" (кнопки ""Обновить"" и ""Провести возвраты"")."
End Sub

' Легаси-архивные листы (Производство/Детали) - только история, кнопки неактивны.
Public Sub PRIME_Legacy_ArchiveStub()
    MsgBox "Этот лист - архив истории версии 1.4.1, доступен только для чтения." & Chr(10) & _
        "Текущий ввод ведите на листах ""Приход/Расход - Цех"" или ""Приход/Расход - Офис""."
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
    PRIME_UI_ApplySheetVisibility()
    PRIME_UI_ApplyFreezeAndFilters()
    MsgBox "Интерфейс PRIME восстановлен."
End Sub

Private Sub PRIME_UI_ApplySheetVisibility()
    Dim hiddenSheets As Variant
    hiddenSheets = Array(SH_SYS_META, SH_SYS_SEQ, SH_SYS_TX, SH_DB_PRODUCTS, SH_DB_ALIASES, SH_DB_PRODUCT_UNITS, _
        SH_DB_DOCUMENTS, SH_DB_DOC_LINES, SH_DB_MOVEMENTS, SH_DB_LOTS, SH_DB_ALLOCATIONS, SH_DB_RETURNS, _
        SH_DB_ORDER_SNAPSHOT, SH_DB_KITS, SH_DB_KIT_LINES, SH_DB_ACTS, SH_DB_AUDIT)
    Dim i As Long
    For i = LBound(hiddenSheets) To UBound(hiddenSheets)
        If PRIME_SheetExists(hiddenSheets(i)) Then
            PRIME_GetSheet(hiddenSheets(i)).IsVisible = False
        End If
    Next i

    ' Легаси-архивные листы (Производство/Детали) - только история, скрыты если пусты
    ' (legacy_parallel_forms: hide/read-only/redirect, а не активные формы ввода).
    Dim legacySheets As Variant
    legacySheets = Array(SH_LEGACY_RECEIPT_PROD, SH_LEGACY_ISSUE_PROD, SH_LEGACY_RECEIPT_PARTS, SH_LEGACY_ISSUE_PARTS)
    For i = LBound(legacySheets) To UBound(legacySheets)
        If PRIME_SheetExists(legacySheets(i)) Then
            Dim oLegacy As Object
            oLegacy = PRIME_GetSheet(legacySheets(i))
            oLegacy.IsVisible = (PRIME_FindLastRow(oLegacy) >= 1)
            oLegacy.IsProtected = True ' read-only - не активная форма ввода
        End If
    Next i
End Sub

Private Sub PRIME_UI_ApplyFreezeAndFilters()
    Dim dataSheets As Variant
    dataSheets = Array(SH_ORDERS, SH_ISSUES, SH_RECEIPT_SHOP, SH_ISSUE_SHOP, SH_RECEIPT_OFFICE, _
        SH_ISSUE_OFFICE, SH_RETURNS, SH_INVENTORY, SH_KITS)
    Dim oController As Object
    oController = ThisComponent.CurrentController

    Dim i As Long
    For i = LBound(dataSheets) To UBound(dataSheets)
        If PRIME_SheetExists(dataSheets(i)) Then
            Dim oSh As Object
            oSh = PRIME_GetSheet(dataSheets(i))

            oController.setActiveSheet(oSh)
            oController.freezeAtPosition(0, 1) ' заморозить строку заголовка (freeze_headers)

            Dim lastRow As Long, lastCol As Long
            lastRow = PRIME_FindLastRow(oSh)
            lastCol = PRIME_FindLastCol(oSh)
            If lastRow < 1 Then lastRow = 1

            Dim oRange As Object
            oRange = oSh.getCellRangeByPosition(0, 0, lastCol, lastRow)

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
