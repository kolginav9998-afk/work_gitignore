Option Explicit

' PRIME_14_MigrationInstaller
' Установка схемы (EnsureSchema, разово, НЕ на каждое проведение - в отличие от 1.4.1, где
' WMSDBST_EnsureSchema/WMSARCH_EnsureSchema вызывались на каждую строку) и миграция 1.4.1 -> 2.0.0.
' disable_legacy_events_first: перед миграцией отключаем старые обработчики WMS_*.

' === Установка / EnsureSchema ===================================================================
Public Sub PRIME_Install_EnsureSchema()
    PRIME_Install_EnsureSheetWithHeaders(SH_SYS_META, PRIME_SysMetaColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_SYS_SEQ, PRIME_SysSeqColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_SYS_TX, PRIME_SysTxColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_PRODUCTS, PRIME_DbProductsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_ALIASES, PRIME_DbAliasesColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_PRODUCT_UNITS, PRIME_DbProductUnitsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_DOCUMENTS, PRIME_DbDocumentsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_DOC_LINES, PRIME_DbDocLinesColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_MOVEMENTS, PRIME_DbMovementsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_LOTS, PRIME_DbLotsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_ALLOCATIONS, PRIME_DbAllocationsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_RETURNS, PRIME_DbReturnsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_ORDER_SNAPSHOT, PRIME_DbOrderSnapshotColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_KITS, PRIME_DbKitsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_KIT_LINES, PRIME_DbKitLinesColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_ACTS, PRIME_DbActsColumns())
    PRIME_Install_EnsureSheetWithHeaders(SH_DB_AUDIT, PRIME_DbAuditColumns())

    If PRIME_MetaGet("SCHEMA_VERSION") = "" Then
        PRIME_MetaSet("SCHEMA_VERSION", PRIME_SCHEMA_VERSION)
        PRIME_MetaSet("INSTALLED_AT", Format(Now, "YYYY-MM-DD HH:MM:SS"))
    End If

    PRIME_InvalidateProductIndex()
    PRIME_RebuildCommittedKeyCache()
End Sub

Private Sub PRIME_Install_EnsureSheetWithHeaders(ByVal sheetName As String, ByVal columns As Variant)
    Dim oSheets As Object
    oSheets = ThisComponent.Sheets
    Dim isNew As Boolean
    isNew = Not oSheets.hasByName(sheetName)
    If isNew Then
        oSheets.insertNewByName(sheetName, oSheets.Count)
    End If

    Dim oSheet As Object
    oSheet = oSheets.getByName(sheetName)

    If isNew Then
        Dim n As Long
        n = UBound(columns) - LBound(columns) + 1
        Dim headerRow(0) As Variant
        Dim rowData(n - 1) As Variant
        Dim i As Long
        For i = 0 To n - 1
            rowData(i) = columns(LBound(columns) + i)
        Next i
        headerRow(0) = rowData
        oSheet.getCellRangeByPosition(0, 0, n - 1, 0).setDataArray(headerRow)
        oSheet.IsVisible = False
        PRIME_InvalidateHeaderCache(sheetName)
    End If
End Sub

' Создаёт бизнес-лист с заголовком, ТОЛЬКО если он ещё не существует - не трогает уже
' присутствующие пользовательские листы (replacing_user_data_with_empty_templates запрещено).
Public Sub PRIME_Install_EnsureBusinessSheet(ByVal sheetName As String, ByVal businessColumns As Variant, ByVal hiddenColumns As Variant)
    Dim oSheets As Object
    oSheets = ThisComponent.Sheets
    If oSheets.hasByName(sheetName) Then Exit Sub

    oSheets.insertNewByName(sheetName, oSheets.Count)
    Dim oSheet As Object
    oSheet = oSheets.getByName(sheetName)

    Dim bn As Long
    bn = UBound(businessColumns) - LBound(businessColumns) + 1
    Dim hn As Long
    hn = 0
    If Not IsMissing(hiddenColumns) Then
        If Not IsNull(hiddenColumns) Then hn = UBound(hiddenColumns) - LBound(hiddenColumns) + 1
    End If

    Dim total As Long
    total = bn + hn
    Dim rowData(total - 1) As Variant
    Dim i As Long
    For i = 0 To bn - 1
        rowData(i) = businessColumns(LBound(businessColumns) + i)
    Next i
    For i = 0 To hn - 1
        rowData(bn + i) = hiddenColumns(LBound(hiddenColumns) + i)
    Next i

    Dim headerRow(0) As Variant
    headerRow(0) = rowData
    oSheet.getCellRangeByPosition(0, 0, total - 1, 0).setDataArray(headerRow)
    PRIME_InvalidateHeaderCache(sheetName)
End Sub

Public Sub PRIME_Install_EnsureAllBusinessSheetsButton()
    PRIME_Install_EnsureAllBusinessSheetsSilent()
    MsgBox "Все листы PRIME проверены/созданы."
End Sub

' Вариант без MsgBox - вызывается сборщиком (tools/build_ods.py) в headless-режиме, где
' любой диалог означал бы зависание процесса сборки.
Public Sub PRIME_Install_EnsureAllBusinessSheetsSilent()
    PRIME_Install_EnsureSchema()
    PRIME_Install_EnsureBusinessSheet(SH_ORDERS, PRIME_ArrayConcat(PRIME_OrdersBusinessColumns(), PRIME_OrdersExtraColumns()), PRIME_OrdersHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_ISSUES, PRIME_IssuesColumns(), PRIME_IssuesHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_RECEIPT_SHOP, PRIME_WorkflowReceiptColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_ISSUE_SHOP, PRIME_WorkflowIssueColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_RECEIPT_OFFICE, PRIME_WorkflowReceiptColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_ISSUE_OFFICE, PRIME_WorkflowIssueColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_RETURNS, PRIME_ReturnsColumns(), PRIME_ReturnsHiddenColumns())
    PRIME_Install_EnsureBusinessSheet(SH_INVENTORY, PRIME_InventoryColumns(), Array())
    PRIME_Install_EnsureBusinessSheet(SH_KITS, PRIME_KitsColumns(), Array())
    PRIME_Install_EnsureBusinessSheet(SH_STOCK, PRIME_StockColumns(), Array())
    PRIME_Install_EnsureBusinessSheet(SH_STOCK_ORDERS, PRIME_StockOrdersColumns(), Array())
    PRIME_Install_EnsureBusinessSheet(SH_SEARCH, PRIME_SearchColumns(), Array())
    PRIME_Install_EnsureBusinessSheet(SH_DIAGNOSTICS, Array("Диагностика"), Array())
    PRIME_Install_EnsureBusinessSheet(SH_DASHBOARD, Array("Показатель", "Значение"), Array())
    PRIME_Install_EnsureBusinessSheet(SH_REPORT_INPUT, Array("Показатель", "Значение"), Array())
    PRIME_Install_EnsureBusinessSheet(SH_REPORT_FINAL, Array("Отчёт"), Array())
End Sub

Private Function PRIME_ArrayConcat(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim na As Long, nb As Long
    na = UBound(a) - LBound(a) + 1
    nb = UBound(b) - LBound(b) + 1
    Dim result(na + nb - 1) As Variant
    Dim i As Long
    For i = 0 To na - 1
        result(i) = a(LBound(a) + i)
    Next i
    For i = 0 To nb - 1
        result(na + i) = b(LBound(b) + i)
    Next i
    PRIME_ArrayConcat = result
End Function

' === Резервная копия ============================================================================
Public Sub PRIME_Backup_CreateButton()
    PRIME_Backup_Create("manual")
End Sub

Public Function PRIME_Backup_Create(ByVal reason As String) As String
    On Error Goto Fail
    Dim dir As String
    dir = PRIME_EnsureDir(PRIME_DIR_BACKUPS)
    Dim fileName As String
    fileName = "POKATAK_backup_" & Format(Now, "YYYYMMDD_HHMMSS") & ".ods"
    Dim fullPath As String
    fullPath = dir & fileName

    Dim saveArgs(0) As New com.sun.star.beans.PropertyValue
    saveArgs(0).Name = "FilterName"
    saveArgs(0).Value = "calc8"
    ThisComponent.storeToURL(ConvertToURL(fullPath), saveArgs())

    PRIME_Backup_Create = fullPath
    Exit Function
Fail:
    MsgBox "Резервная копия НЕ создана: " & Error$
    PRIME_Backup_Create = ""
End Function

' === Отключение легаси-событий ===================================================================
' disable_legacy_events_first: перед миграцией снимаем старые обработчики WMS_* с листов,
' чтобы они не срабатывали параллельно с новыми PRIME_OnContentChanged_*.
Public Sub PRIME_Install_DisableLegacyEvents()
    Dim sheetNames As Variant
    sheetNames = Array(SH_ORDERS, SH_ISSUES, SH_RECEIPT_SHOP, SH_ISSUE_SHOP, SH_RECEIPT_OFFICE, SH_ISSUE_OFFICE, _
        SH_RETURNS, "Приход — Производство", "Расход — Производство", "Приход — Детали", "Расход — Детали")
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        If PRIME_SheetExists(sheetNames(i)) Then
            On Error Resume Next
            Dim oEvents As Object
            oEvents = PRIME_GetSheet(sheetNames(i)).Events
            oEvents.revokeByName("OnChange") ' ключ события листа "Contents changed" в UNO - OnChange, не OnContentChanged
            On Error Goto 0
        End If
    Next i
End Sub

' === Миграция колонок бизнес-листов, унаследованных от 1.4.1 (issues_migration, PRIME 2.0.1) ===
' Дефект 2.0.0: PRIME_Install_EnsureBusinessSheet создаёт целевую раскладку колонок ТОЛЬКО для
' листа, которого ещё не существует - но почти все бизнес-листы (Выдачи и все 4 листа
' Приход/Расход - Цех/Офис, Возвраты, Инвентаризация, Остаток, Остаток - Заказы, База - Поиск)
' УЖЕ существуют в шаблоне 1.4.1 (с декоративной "шапкой" и/или легаси _WMS_*-колонками), из-за
' чего EnsureBusinessSheet молча ничего для них не делает - все эти листы оставались на исходной
' раскладке 1.4.1, а не на PRIME_*Columns() схеме, которую ожидает остальной код. Единственный
' лист, для которого это было явно исправлено в 2.0.0 - "Заказы" (PRIME_Migration_MigrateOrdersSheet).
' Эта функция - универсальная замена: приводит УЖЕ существующий лист к целевому набору колонок,
' перенося данные ПО ИМЕНИ колонки (колонки без соответствия в целевой схеме, включая любые
' легаси _WMS_*-поля, не переносятся - non_negotiable_architecture.legacy_code_allowed_in_production_ods=false).
' Идемпотентна: если текущий заголовок уже точно совпадает с целевым - не трогает лист.
Public Function PRIME_Migration_MigrateBusinessSheetColumns(ByVal sheetName As String, ByVal targetColumns As Variant) As Boolean
    If Not PRIME_SheetExists(sheetName) Then
        PRIME_Migration_MigrateBusinessSheetColumns = False
        Exit Function
    End If

    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim headerRow As Long
    headerRow = PRIME_FormSchemaHeaderRow(sheetName)

    Dim targetCount As Long
    targetCount = UBound(targetColumns) - LBound(targetColumns) + 1

    Dim lastCol As Long
    lastCol = PRIME_FindLastCol(oSheet)
    Dim oldHeaders() As String
    If lastCol >= 0 Then
        Dim headerData As Variant
        headerData = oSheet.getCellRangeByPosition(0, headerRow, lastCol, headerRow).getDataArray()
        ReDim oldHeaders(lastCol)
        Dim hc As Long
        For hc = 0 To lastCol
            oldHeaders(hc) = CStr(headerData(0)(hc))
        Next hc
    Else
        ReDim oldHeaders(-1)
    End If

    ' Уже мигрировано (точное совпадение имени и порядка колонок) - ничего не делаем, чтобы
    ' повторный запуск сборки/миграции не тёр уже введённые пользователем данные.
    Dim alreadyMigrated As Boolean
    alreadyMigrated = (UBound(oldHeaders) - LBound(oldHeaders) + 1 = targetCount)
    If alreadyMigrated Then
        Dim tc As Long
        For tc = 0 To targetCount - 1
            If oldHeaders(tc) <> CStr(targetColumns(LBound(targetColumns) + tc)) Then
                alreadyMigrated = False
                Exit For
            End If
        Next tc
    End If
    If alreadyMigrated Then
        PRIME_Migration_MigrateBusinessSheetColumns = False
        Exit Function
    End If

    Dim oldLastRow As Long
    oldLastRow = PRIME_FindLastRow(oSheet) ' по ТЕКУЩЕМУ headerRow - вызывается до перезаписи
    Dim oldDataRowCount As Long
    oldDataRowCount = oldLastRow - headerRow ' может быть <= 0, если данных ещё нет

    Dim totalNewRows As Long
    totalNewRows = 1 ' заголовок
    If oldDataRowCount > 0 Then totalNewRows = totalNewRows + oldDataRowCount

    Dim newTable(totalNewRows - 1) As Variant
    Dim headerOut(targetCount - 1) As Variant
    Dim i As Long
    For i = 0 To targetCount - 1
        headerOut(i) = targetColumns(LBound(targetColumns) + i)
    Next i
    newTable(0) = headerOut

    If oldDataRowCount > 0 Then
        Dim oldDataRange As Variant
        oldDataRange = oSheet.getCellRangeByPosition(0, headerRow, lastCol, oldLastRow).getDataArray()
        Dim r As Long
        For r = 1 To oldDataRowCount
            Dim newRow(targetCount - 1) As Variant
            Dim c As Long
            For c = 0 To targetCount - 1
                Dim colName As String
                colName = CStr(targetColumns(LBound(targetColumns) + c))
                Dim oldIdx As Long
                oldIdx = PRIME_Migration_IndexOfName(oldHeaders, colName)
                If oldIdx >= 0 Then
                    newRow(c) = oldDataRange(r)(oldIdx)
                Else
                    newRow(c) = ""
                End If
            Next c
            newTable(r) = newRow
        Next r
    End If

    ' Очищаем максимум старой/новой области, затем пишем целевую раскладку одним батчем.
    Dim clearLastCol As Long
    clearLastCol = lastCol
    If targetCount - 1 > clearLastCol Then clearLastCol = targetCount - 1
    Dim clearLastRow As Long
    clearLastRow = oldLastRow
    If headerRow + totalNewRows - 1 > clearLastRow Then clearLastRow = headerRow + totalNewRows - 1
    If clearLastCol >= 0 And clearLastRow >= headerRow Then
        oSheet.getCellRangeByPosition(0, headerRow, clearLastCol, clearLastRow).clearContents(1023)
    End If
    oSheet.getCellRangeByPosition(0, headerRow, targetCount - 1, headerRow + totalNewRows - 1).setDataArray(newTable)

    PRIME_InvalidateHeaderCache(sheetName)
    PRIME_Migration_MigrateBusinessSheetColumns = True
End Function

Private Function PRIME_Migration_IndexOfName(ByVal names() As String, ByVal target As String) As Long
    If UBound(names) < LBound(names) Then
        PRIME_Migration_IndexOfName = -1
        Exit Function
    End If
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If names(i) = target Then
            PRIME_Migration_IndexOfName = i
            Exit Function
        End If
    Next i
    PRIME_Migration_IndexOfName = -1
End Function

' Вызывает миграцию колонок для КАЖДОГО бизнес-листа, для которого PRIME_Install_EnsureBusinessSheet
' не гарантирует целевую раскладку (см. комментарий выше) - "Заказы" сюда не входит, у неё
' отдельная, более сложная миграция (PRIME_Migration_MigrateOrdersSheet, генерирует ЕИ-коды).
Public Sub PRIME_Migration_MigrateAllBusinessSheets()
    PRIME_Migration_MigrateBusinessSheetColumns SH_ISSUES, PRIME_ArrayConcat(PRIME_IssuesColumns(), PRIME_IssuesHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_RECEIPT_SHOP, PRIME_ArrayConcat(PRIME_WorkflowReceiptColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_ISSUE_SHOP, PRIME_ArrayConcat(PRIME_WorkflowIssueColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_RECEIPT_OFFICE, PRIME_ArrayConcat(PRIME_WorkflowReceiptColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_ISSUE_OFFICE, PRIME_ArrayConcat(PRIME_WorkflowIssueColumns(), PRIME_WorkflowHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_RETURNS, PRIME_ArrayConcat(PRIME_ReturnsColumns(), PRIME_ReturnsHiddenColumns())
    PRIME_Migration_MigrateBusinessSheetColumns SH_INVENTORY, PRIME_InventoryColumns()
    PRIME_Migration_MigrateBusinessSheetColumns SH_STOCK, PRIME_StockColumns()
    PRIME_Migration_MigrateBusinessSheetColumns SH_STOCK_ORDERS, PRIME_StockOrdersColumns()
    PRIME_Migration_MigrateBusinessSheetColumns SH_SEARCH, PRIME_SearchColumns()
End Sub

' === Миграция 1.4.1 -> 2.0.0 =====================================================================
' Zero-arg, no-dialog вариант миграции для сборщика (tools/build_ods.py) и автоматических тестов -
' UNO script provider не умеет удобно принимать ByRef-массивы через invoke() извне, поэтому
' здесь всё держится в локальных переменных одного вызова, без диалогов.
'
' PRIME_Build_RunFullSetup - ЕДИНСТВЕННАЯ функция, которую должен вызывать внешний сборщик
' (tools/build_ods.py) одним invoke(). Эмпирически подтверждено (см. историю отладки в git log):
' повторный вызов PRIME_Install_EnsureSchema непосредственно перед PRIME_Migration_MigrateOrdersSheet
' в рамках одного и того же invoke() ломает выполнение ("Object variable not set", без видимого
' исключения наружу) - поэтому PRIME_Build_MigrateSilent (в отличие от интерактивной
' PRIME_Migration_RunButton, у которой это её единственный вызов EnsureSchema) НЕ повторяет
' EnsureSchema, полагаясь на то, что PRIME_Install_EnsureAllBusinessSheetsSilent уже вызвала её.
Public Sub PRIME_Build_RunFullSetup(ByVal runMigration As Boolean)
    PRIME_Install_EnsureAllBusinessSheetsSilent()
    If runMigration Then
        PRIME_Build_MigrateSilent()
    Else
        PRIME_Install_DisableLegacyEvents()
    End If
    PRIME_UI_RestoreInterfaceSilent()
End Sub

' Вызывать ТОЛЬКО после того, как PRIME_Install_EnsureSchema уже был вызван в этом же сеансе
' (через PRIME_Install_EnsureAllBusinessSheetsSilent) - см. предупреждение выше.
Public Sub PRIME_Build_MigrateSilent()
    PRIME_Install_DisableLegacyEvents()
    ' Приводит Выдачи/workflow-листы/Возвраты/Инвентаризацию/Остаток/Поиск к целевой раскладке
    ' ДО миграции "Заказы" - порядок не важен для корректности (независимые листы), но так
    ' ошибки в одной миграции не маскируют результат другой при чтении лога.
    PRIME_Migration_MigrateAllBusinessSheets()
    Dim oldCodes() As String
    Dim newCodes() As String
    Dim mapCount As Long
    mapCount = PRIME_Migration_MigrateOrdersSheet(oldCodes, newCodes)
    If mapCount > 0 Then
        PRIME_Migration_RemapCodes(SH_ISSUES, "Код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_RECEIPT_SHOP, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_ISSUE_SHOP, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_RECEIPT_OFFICE, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_ISSUE_OFFICE, "Внутренний код", oldCodes, newCodes)
    End If
    PRIME_MetaSet("SCHEMA_VERSION", PRIME_SCHEMA_VERSION)
    PRIME_MetaSet("MIGRATED_AT", Format(Now, "YYYY-MM-DD HH:MM:SS"))
    PRIME_MetaSet("MIGRATED_FROM", "1.4.1")
End Sub

Public Sub PRIME_Migration_RunButton()
    Dim answer As Integer
    answer = MsgBox("Выполнить миграцию 1.4.1 -> PRIME " & PRIME_SCHEMA_VERSION & "?" & Chr(10) & _
        "Будет создана резервная копия перед началом. Действие затрагивает структуру листа ""Заказы"".", _
        MB_YESNO + MB_ICONQUESTION, "Миграция PRIME")
    If answer <> IDYES Then Exit Sub

    Dim backupPath As String
    backupPath = PRIME_Backup_Create("before_migration")
    If backupPath = "" Then
        MsgBox "Миграция остановлена: не удалось создать резервную копию."
        Exit Sub
    End If

    PRIME_Install_DisableLegacyEvents()
    PRIME_Install_EnsureSchema()
    PRIME_Migration_MigrateAllBusinessSheets()

    Dim renameMap As Object
    Dim oldCodes() As String
    Dim newCodes() As String
    Dim mapCount As Long
    mapCount = PRIME_Migration_MigrateOrdersSheet(oldCodes, newCodes)

    If mapCount > 0 Then
        PRIME_Migration_RemapCodes(SH_ISSUES, "Код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_RECEIPT_SHOP, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_ISSUE_SHOP, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_RECEIPT_OFFICE, "Внутренний код", oldCodes, newCodes)
        PRIME_Migration_RemapCodes(SH_ISSUE_OFFICE, "Внутренний код", oldCodes, newCodes)
    End If

    PRIME_MetaSet("SCHEMA_VERSION", PRIME_SCHEMA_VERSION)
    PRIME_MetaSet("MIGRATED_AT", Format(Now, "YYYY-MM-DD HH:MM:SS"))
    PRIME_MetaSet("MIGRATED_FROM", "1.4.1")

    MsgBox "Миграция завершена. Товаров создано/сопоставлено: " & mapCount & Chr(10) & _
        "Резервная копия исходного файла: " & backupPath & Chr(10) & Chr(10) & _
        "ВАЖНО: история частичных поставок и складские остатки из старой Firebird-базы НЕ " & _
        "переносятся автоматически (см. MIGRATION_1.4.1_TO_2.0.md). ""Получено всего"" по старым " & _
        "заказам установлено в 0 - при необходимости выполните инвентаризацию после миграции."
End Sub

' Приводит лист "Заказы" к целевой раскладке колонок (preserve_all_existing_25_order_columns,
' reorder_columns, add_new_columns, do_not_recreate_dozen_legacy_helper_columns) и генерирует
' ЕИ-коды для товаров, ещё не имеющих кода в этом формате. Возвращает число сопоставленных кодов
' и заполняет oldCodes()/newCodes() для переноса замены в другие листы.
Public Function PRIME_Migration_MigrateOrdersSheet(ByRef oldCodes() As String, ByRef newCodes() As String) As Long
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim oldHeaders As Variant
    oldHeaders = PRIME_HeaderMap(SH_ORDERS)

    If PRIME_ColIndex(oldHeaders, "_PRIME_OrderID") >= 0 Then
        ' Уже мигрировано ранее - ничего не делаем со структурой, только досеиваем товары.
        PRIME_Migration_MigrateOrdersSheet = PRIME_Migration_SeedProductsFromOrders(oldCodes, newCodes)
        Exit Function
    End If

    Dim targetHeaders As Variant
    targetHeaders = PRIME_ArrayConcat(PRIME_ArrayConcat(PRIME_OrdersBusinessColumns(), PRIME_OrdersExtraColumns()), PRIME_OrdersHiddenColumns())

    Dim oldTable As Variant
    oldTable = PRIME_ReadTable(SH_ORDERS)
    Dim lastRow As Long
    lastRow = UBound(oldTable)
    If lastRow < 0 Then lastRow = 0

    Dim targetCount As Long
    targetCount = UBound(targetHeaders) - LBound(targetHeaders) + 1

    Dim newTable(lastRow) As Variant
    Dim headerRow(targetCount - 1) As Variant
    Dim tc As Long
    For tc = 0 To targetCount - 1
        headerRow(tc) = targetHeaders(LBound(targetHeaders) + tc)
    Next tc
    newTable(0) = headerRow

    Dim r As Long
    For r = 1 To lastRow
        Dim newRow(targetCount - 1) As Variant
        For tc = 0 To targetCount - 1
            Dim colName As String
            colName = targetHeaders(LBound(targetHeaders) + tc)
            Dim oldIdx As Long
            oldIdx = PRIME_ColIndex(oldHeaders, colName)
            If oldIdx >= 0 And r <= UBound(oldTable) Then
                newRow(tc) = oldTable(r)(oldIdx)
            Else
                newRow(tc) = ""
            End If
        Next tc
        newTable(r) = newRow
    Next r

    ' Очищаем весь текущий диапазон листа (может быть шире/уже нового) и пишем заново одним батчем.
    Dim oldLastCol As Long
    oldLastCol = PRIME_FindLastCol(oSheet)
    Dim clearCols As Long
    clearCols = oldLastCol
    If targetCount - 1 > clearCols Then clearCols = targetCount - 1
    oSheet.getCellRangeByPosition(0, 0, clearCols, lastRow).clearContents(1023)
    oSheet.getCellRangeByPosition(0, 0, targetCount - 1, lastRow).setDataArray(newTable)
    PRIME_InvalidateHeaderCache(SH_ORDERS)

    ' Присваиваем _PRIME_OrderID/_PRIME_LineID существующим строкам, где их ещё нет.
    Dim headers2 As Variant
    headers2 = PRIME_HeaderMap(SH_ORDERS)
    Dim colOrderId As Long, colLineId As Long, colReceived As Long, colOrdered As Long, colRemaining As Long
    colOrderId = PRIME_ColIndex(headers2, "_PRIME_OrderID")
    colLineId = PRIME_ColIndex(headers2, "_PRIME_LineID")
    colReceived = PRIME_ColIndex(headers2, "Получено всего")
    colOrdered = PRIME_ColIndex(headers2, "Количество")
    colRemaining = PRIME_ColIndex(headers2, "Осталось получить")

    For r = 1 To lastRow
        If oSheet.getCellByPosition(colOrderId, r).getString() = "" Then
            oSheet.getCellByPosition(colOrderId, r).setString("ORD-" & Format(PRIME_SequenceNext("ORDER_ID"), "00000000"))
        End If
        If oSheet.getCellByPosition(colLineId, r).getString() = "" Then
            oSheet.getCellByPosition(colLineId, r).setString("OL-" & Format(PRIME_SequenceNext("ORDER_LINE_ID"), "00000000"))
        End If
        If oSheet.getCellByPosition(colReceived, r).getString() = "" Then
            oSheet.getCellByPosition(colReceived, r).setValue(0)
        End If
        If oSheet.getCellByPosition(colRemaining, r).getString() = "" Then
            Dim orderedStr As String
            orderedStr = oSheet.getCellByPosition(colOrdered, r).getString()
            If IsNumeric(orderedStr) Then
                oSheet.getCellByPosition(colRemaining, r).setValue(CDbl(orderedStr))
            End If
        End If
        PRIME_Orders_RecomputeStatus(oSheet, headers2, r)
    Next r

    PRIME_Migration_MigrateOrdersSheet = PRIME_Migration_SeedProductsFromOrders(oldCodes, newCodes)
End Function

' Заводит DB_PRIME_PRODUCTS по уникальным старым кодам товара из "Заказы" (identity = старый код,
' НЕ название - name_is_unique_key=false) и генерирует новые ЕИ-коды, если старый код ещё не в
' формате "ЕИ-########" (rename_existing_codes_automatically=false - уже-ЕИ-коды не переименовываем).
Public Function PRIME_Migration_SeedProductsFromOrders(ByRef oldCodes() As String, ByRef newCodes() As String) As Long
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(SH_ORDERS)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_ORDERS)
    Dim colCode As Long, colName As Long, colUnit As Long, colLoc As Long, colCat As Long, colSub As Long
    colCode = PRIME_ColIndex(headers, "Код товара")
    colName = PRIME_ColIndex(headers, "Полное наименование товара")
    colUnit = PRIME_ColIndex(headers, "Ед. изм.")
    colLoc = PRIME_ColIndex(headers, "Место хранения")
    colCat = PRIME_ColIndex(headers, "Категория")
    colSub = PRIME_ColIndex(headers, "Подкатегория")

    Dim table As Variant
    table = PRIME_ReadTable(SH_ORDERS)

    ReDim oldCodes(UBound(table))
    ReDim newCodes(UBound(table))
    Dim n As Long
    n = 0

    If UBound(table) >= 1 Then
        Dim r As Long
        For r = 1 To UBound(table)
            Dim code As String
            code = Trim(CStr(table(r)(colCode)))
            If code <> "" And Not PRIME_ProductExists(code) And Not PRIME_Migration_AlreadyMapped(oldCodes, n, code) Then
                Dim finalCode As String
                If PRIME_Migration_IsEICode(code) Then
                    finalCode = code
                    PRIME_Migration_InsertProductWithCode(finalCode, CStr(table(r)(colName)), CStr(table(r)(colUnit)), _
                        CStr(table(r)(colLoc)), CStr(table(r)(colCat)), CStr(table(r)(colSub)))
                Else
                    finalCode = PRIME_CreateProduct(CStr(table(r)(colName)), CStr(table(r)(colUnit)), _
                        CStr(table(r)(colLoc)), CStr(table(r)(colCat)), CStr(table(r)(colSub)), False, "")
                End If
                oldCodes(n) = code
                newCodes(n) = finalCode
                n = n + 1
            End If
        Next r
    End If

    If n = 0 Then
        ReDim oldCodes(-1)
        ReDim newCodes(-1)
    Else
        ReDim Preserve oldCodes(n - 1)
        ReDim Preserve newCodes(n - 1)
    End If
    PRIME_Migration_SeedProductsFromOrders = n
End Function

Private Function PRIME_Migration_AlreadyMapped(ByVal oldCodes() As String, ByVal n As Long, ByVal code As String) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If oldCodes(i) = code Then
            PRIME_Migration_AlreadyMapped = True
            Exit Function
        End If
    Next i
    PRIME_Migration_AlreadyMapped = False
End Function

Private Function PRIME_Migration_IsEICode(ByVal code As String) As Boolean
    If Left(code, Len(PRODUCT_CODE_PREFIX)) <> PRODUCT_CODE_PREFIX Then
        PRIME_Migration_IsEICode = False
        Exit Function
    End If
    Dim suffix As String
    suffix = Mid(code, Len(PRODUCT_CODE_PREFIX) + 1)
    PRIME_Migration_IsEICode = (Len(suffix) > 0 And IsNumeric(suffix))
End Function

Private Sub PRIME_Migration_InsertProductWithCode(ByVal code As String, ByVal name As String, ByVal unit As String, _
        ByVal location As String, ByVal category As String, ByVal subcategory As String)
    Dim headers As Variant
    headers = PRIME_HeaderMap(SH_DB_PRODUCTS)
    Dim row(UBound(headers)) As Variant
    Dim now As String
    now = Format(Now, "YYYY-MM-DD HH:MM:SS")
    row(PRIME_ColIndex(headers, "PRODUCT_CODE")) = code
    row(PRIME_ColIndex(headers, "PRODUCT_NAME")) = name
    row(PRIME_ColIndex(headers, "BASE_UNIT")) = unit
    row(PRIME_ColIndex(headers, "DEFAULT_LOCATION")) = location
    row(PRIME_ColIndex(headers, "CATEGORY")) = category
    row(PRIME_ColIndex(headers, "SUBCATEGORY")) = subcategory
    row(PRIME_ColIndex(headers, "ACTIVE")) = 1
    row(PRIME_ColIndex(headers, "CREATED_AT")) = now
    row(PRIME_ColIndex(headers, "UPDATED_AT")) = now
    Dim rows(0) As Variant
    rows(0) = row
    PRIME_AppendRowsBatch(SH_DB_PRODUCTS, rows)
    PRIME_InvalidateProductIndex()

    ' Обновляем максимальный суффикс последовательности ЕИ-кодов, чтобы PRIME_NextProductCode
    ' не выдал уже занятый номер (migration_rule).
    If PRIME_Migration_IsEICode(code) Then
        Dim suffix As Long
        suffix = CLng(Mid(code, Len(PRODUCT_CODE_PREFIX) + 1))
        Dim seqHeaders As Variant
        seqHeaders = PRIME_HeaderMap(SH_SYS_SEQ)
        Dim seqTable As Variant
        seqTable = PRIME_ReadTable(SH_SYS_SEQ)
        Dim idx As Long
        idx = PRIME_FindRowByKey(seqTable, PRIME_ColIndex(seqHeaders, "SEQ_NAME"), "PRODUCT_CODE")
        Dim current As Long
        current = 0
        If idx >= 0 Then current = CLng(seqTable(idx)(PRIME_ColIndex(seqHeaders, "NEXT_VALUE")))
        If suffix >= current Then
            ' Прокручиваем последовательность вперёд, пока она не превысит уже занятый суффикс -
            ' PRIME_NextProductCode() после этого гарантированно не выдаст занятый номер.
            Do While PRIME_SequenceNext("PRODUCT_CODE") <= suffix
            Loop
        End If
    End If
End Sub

Private Sub PRIME_Migration_RemapCodes(ByVal sheetName As String, ByVal codeColumnName As String, ByVal oldCodes() As String, ByVal newCodes() As String)
    If Not PRIME_SheetExists(sheetName) Then Exit Sub
    Dim oSheet As Object
    oSheet = PRIME_GetSheet(sheetName)
    Dim headers As Variant
    headers = PRIME_HeaderMap(sheetName)
    Dim col As Long
    col = PRIME_ColIndex(headers, codeColumnName)
    If col < 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = PRIME_FindLastRow(oSheet)
    Dim r As Long, i As Long
    For r = 1 To lastRow
        Dim v As String
        v = oSheet.getCellByPosition(col, r).getString()
        If v <> "" Then
            For i = LBound(oldCodes) To UBound(oldCodes)
                If oldCodes(i) = v Then
                    oSheet.getCellByPosition(col, r).setString(newCodes(i))
                    Exit For
                End If
            Next i
        End If
    Next r
End Sub
