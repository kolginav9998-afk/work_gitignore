Option Explicit

' PRIME_17_Home (новый в 2.1.2)
' Лист "Главная" - простой landing/навигационный лист (НЕ дашборд с техническими данными):
' заголовок, версия, статус диагностики, крупные кнопки-переходы на все активные пользовательские
' листы, маленький блок сводки (незакрытые/просроченные заказы, последние операции). См.
' авторитетный список видимых листов и требования к "Главная" в master task 2.1.2 - никаких
' технических DB-контролов, легаси-действий, кнопок починки, отчётов или множества мелких
' контролов здесь быть не должно.

' Кнопки-переходы (13 штук, сетка 4 в ряд - см. tools/build_ods.py NEW_SHEET_BUTTON_COLS) -
' floating-controls поверх листа, занимают строки 0..9 (4 ряда по ~10мм + запас) - текстовое
' содержимое листа начинается НИЖЕ этого блока, чтобы кнопки его не перекрывали.
Private Const HOME_ROW_TITLE As Long = 10
Private Const HOME_ROW_VERSION As Long = 11
Private Const HOME_ROW_STATUS As Long = 12
Private Const HOME_ROW_SUMMARY_HEADER As Long = 14
Private Const HOME_ROW_OPEN_ORDERS As Long = 15
Private Const HOME_ROW_OVERDUE_ORDERS As Long = 16
Private Const HOME_ROW_RECENT_HEADER As Long = 18
Private Const HOME_ROW_RECENT_FIRST As Long = 19
Private Const HOME_RECENT_COUNT As Long = 5

' Создаёт лист "Главная" при первой сборке (идемпотентно - если лист уже существует, только
' обновляет статический текст, не трогая уже нарисованные кнопки/данные сводки).
Public Sub PRIME_Home_EnsureSheet()
    Dim oSheets As Object
    oSheets = ThisComponent.Sheets
    If Not oSheets.hasByName(SH_HOME) Then
        oSheets.insertNewByName(SH_HOME, 0) ' первым листом книги - именно landing page
    End If
    Dim oSheet As Object
    oSheet = oSheets.getByName(SH_HOME)

    oSheet.getCellByPosition(0, HOME_ROW_TITLE).setString("ПОКАТАК PRIME")
    oSheet.getCellByPosition(0, HOME_ROW_TITLE).CharHeight = 20
    oSheet.getCellByPosition(0, HOME_ROW_TITLE).CharWeight = com.sun.star.awt.FontWeight.BOLD

    oSheet.getCellByPosition(0, HOME_ROW_VERSION).setString("Версия: " & PRIME_SCHEMA_VERSION)
    oSheet.getCellByPosition(0, HOME_ROW_STATUS).setString("Статус: —")

    oSheet.getCellByPosition(0, HOME_ROW_SUMMARY_HEADER).setString("Сводка")
    oSheet.getCellByPosition(0, HOME_ROW_SUMMARY_HEADER).CharWeight = com.sun.star.awt.FontWeight.BOLD
    oSheet.getCellByPosition(0, HOME_ROW_OPEN_ORDERS).setString("Незакрытые заказы: —")
    oSheet.getCellByPosition(0, HOME_ROW_OVERDUE_ORDERS).setString("Просроченные заказы: —")

    oSheet.getCellByPosition(0, HOME_ROW_RECENT_HEADER).setString("Последние проведённые операции")
    oSheet.getCellByPosition(0, HOME_ROW_RECENT_HEADER).CharWeight = com.sun.star.awt.FontWeight.BOLD

    oSheet.Columns.getByIndex(0).Width = 12000
End Sub

' "Обновить" - единственная кнопка на "Главная" помимо навигации; пересчитывает версию, статус
' диагностики и блок сводки. Не пишет ни в один DB_PRIME_/SYS_PRIME_ лист (read-only view, как и
' Журнал/Наличие/Поиск).
Public Sub PRIME_Home_RefreshButton()
    If Not PRIME_SheetExists(SH_HOME) Then Exit Sub
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_HOME)

    oSheet.getCellByPosition(0, HOME_ROW_VERSION).setString("Версия: " & PRIME_SCHEMA_VERSION)

    Dim problems As Long
    problems = PRIME_Diagnostics_ProblemCount()
    If problems = 0 Then
        oSheet.getCellByPosition(0, HOME_ROW_STATUS).setString("Статус: PRIME OK")
    Else
        oSheet.getCellByPosition(0, HOME_ROW_STATUS).setString("Статус: есть ошибка диагностики (" & problems & ")")
    End If

    Dim openCount As Long, overdueCount As Long
    PRIME_Home_CountOrders(openCount, overdueCount)
    oSheet.getCellByPosition(0, HOME_ROW_OPEN_ORDERS).setString("Незакрытые заказы: " & openCount)
    oSheet.getCellByPosition(0, HOME_ROW_OVERDUE_ORDERS).setString("Просроченные заказы: " & overdueCount)

    PRIME_Home_RefreshRecentOperations(oSheet)
End Sub

' Незакрытые = не "Получено" и не "Отменено" (родительские строки заказа, без CHILD-строк
' receipt-позиций - см. PRIME_05_Orders._PRIME_RowType).
Private Sub PRIME_Home_CountOrders(ByRef openCount As Long, ByRef overdueCount As Long)
    openCount = 0
    overdueCount = 0
    If Not PRIME_SheetExists(SH_ORDERS) Then Exit Sub

    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colStatus As Long, colRowType As Long
    colStatus = PRIME_ColIndex(headers, "Статус")
    colRowType = PRIME_ColIndex(headers, "_PRIME_RowType")
    If colStatus < 0 Then Exit Sub

    Dim table As Variant
    table = PRIME_ReadTable(SH_ORDERS)
    If UBound(table) < 1 Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(table)
        If colRowType >= 0 Then
            If CStr(table(i)(colRowType)) = "CHILD" Then GoTo ContinueOrder
        End If
        Dim status As String
        status = CStr(table(i)(colStatus))
        If status = "" Then GoTo ContinueOrder
        If status <> ORDER_STATUS_RECEIVED And status <> ORDER_STATUS_CANCELLED Then
            openCount = openCount + 1
            If status = ORDER_STATUS_OVERDUE Then overdueCount = overdueCount + 1
        End If
ContinueOrder:
    Next i
End Sub

' Последние HOME_RECENT_COUNT COMMITTED документов (committed_only_everywhere) - DOC_ID
' присваивается последовательно, поэтому "последние" = последние COMMITTED строки таблицы,
' без отдельной сортировки по дате.
Private Sub PRIME_Home_RefreshRecentOperations(ByVal oSheet As Object)
    Dim r As Long
    For r = 0 To HOME_RECENT_COUNT - 1
        oSheet.getCellByPosition(0, HOME_ROW_RECENT_FIRST + r).setString("")
    Next r

    If Not PRIME_SheetExists(SH_DB_DOCUMENTS) Then Exit Sub
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_DOCUMENTS)
    Dim colDocId As Long, colDocType As Long, colDocDate As Long
    colDocId = PRIME_ColIndex(headers, "DOC_ID")
    colDocType = PRIME_ColIndex(headers, "DOC_TYPE")
    colDocDate = PRIME_ColIndex(headers, "DOC_DATE")

    Dim table As Variant
    table = PRIME_ReadTable(SH_DB_DOCUMENTS)
    If UBound(table) < 1 Then Exit Sub

    Dim found As Long
    found = 0
    Dim i As Long
    For i = UBound(table) To 1 Step -1
        If found >= HOME_RECENT_COUNT Then Exit For
        Dim docId As String
        docId = CStr(table(i)(colDocId))
        If PRIME_IsDocIdCommitted(docId) Then
            oSheet.getCellByPosition(0, HOME_ROW_RECENT_FIRST + found).setString( _
                CStr(table(i)(colDocDate)) & "  " & CStr(table(i)(colDocType)) & "  " & docId)
            found = found + 1
        End If
    Next i
End Sub

' === Навигация - большие кнопки-переходы на все активные пользовательские листы =================
Private Sub PRIME_Home_GoTo(ByVal sheetName As String)
    If PRIME_SheetExists(sheetName) Then
        ThisComponent.CurrentController.setActiveSheet(PRIME_GetSheet(sheetName))
    End If
End Sub

Public Sub PRIME_Home_GoOrders()
    PRIME_Home_GoTo(SH_ORDERS)
End Sub

Public Sub PRIME_Home_GoReceiptOffice()
    PRIME_Home_GoTo(SH_RECEIPT_OFFICE)
End Sub

Public Sub PRIME_Home_GoReceiptProduction()
    PRIME_Home_GoTo(SH_RECEIPT_PRODUCTION)
End Sub

Public Sub PRIME_Home_GoReceiptDetails()
    PRIME_Home_GoTo(SH_RECEIPT_DETAILS)
End Sub

Public Sub PRIME_Home_GoIssues()
    PRIME_Home_GoTo(SH_ISSUES)
End Sub

Public Sub PRIME_Home_GoStock()
    PRIME_Home_GoTo(SH_STOCK)
End Sub

Public Sub PRIME_Home_GoReturns()
    PRIME_Home_GoTo(SH_RETURNS)
End Sub

Public Sub PRIME_Home_GoTransfers()
    PRIME_Home_GoTo(SH_TRANSFERS)
End Sub

Public Sub PRIME_Home_GoInventory()
    PRIME_Home_GoTo(SH_INVENTORY)
End Sub

Public Sub PRIME_Home_GoSearch()
    PRIME_Home_GoTo(SH_SEARCH)
End Sub

Public Sub PRIME_Home_GoJournal()
    PRIME_Home_GoTo(SH_JOURNAL)
End Sub

Public Sub PRIME_Home_GoKits()
    PRIME_Home_GoTo(SH_KITS)
End Sub
