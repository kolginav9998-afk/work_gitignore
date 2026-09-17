Option Explicit

' PRIME_00_Config
' Единая точка констант: имена листов, версии, лимиты, справочные значения.
' Никакой логики I/O здесь быть не должно (см. forbidden: очень крупные монолитные функции,
' циклические зависимости) - только константы и простые справочные функции.

Public Const PRIME_SCHEMA_VERSION As String = "2.0.1"
Public Const PRIME_BUILD_DATE As String = "2026-09-17"

' --- Системные (скрытые) листы ---
Public Const SH_SYS_META As String = "SYS_PRIME_META"
Public Const SH_SYS_SEQ As String = "SYS_PRIME_SEQ"
Public Const SH_SYS_TX As String = "SYS_PRIME_TX"
Public Const SH_DB_PRODUCTS As String = "DB_PRIME_PRODUCTS"
Public Const SH_DB_ALIASES As String = "DB_PRIME_ALIASES"
Public Const SH_DB_PRODUCT_UNITS As String = "DB_PRIME_PRODUCT_UNITS"
Public Const SH_DB_DOCUMENTS As String = "DB_PRIME_DOCUMENTS"
Public Const SH_DB_DOC_LINES As String = "DB_PRIME_DOC_LINES"
Public Const SH_DB_MOVEMENTS As String = "DB_PRIME_MOVEMENTS"
Public Const SH_DB_LOTS As String = "DB_PRIME_LOTS"
Public Const SH_DB_ALLOCATIONS As String = "DB_PRIME_ALLOCATIONS"
Public Const SH_DB_RETURNS As String = "DB_PRIME_RETURNS"
Public Const SH_DB_ORDER_SNAPSHOT As String = "DB_PRIME_ORDER_SNAPSHOT"
Public Const SH_DB_KITS As String = "DB_PRIME_KITS"
Public Const SH_DB_KIT_LINES As String = "DB_PRIME_KIT_LINES"
Public Const SH_DB_ACTS As String = "DB_PRIME_ACTS"
Public Const SH_DB_AUDIT As String = "DB_PRIME_AUDIT"

' --- Бизнес-листы (пользовательский UI, сохраняем привычные имена 1.4.1) ---
Public Const SH_ORDERS As String = "Заказы"
Public Const SH_ISSUES As String = "Выдачи"
Public Const SH_STOCK As String = "Остаток"
Public Const SH_STOCK_ORDERS As String = "Остаток — Заказы"
Public Const SH_SEARCH As String = "База - Поиск"
Public Const SH_RETURNS As String = "Возвраты"
Public Const SH_INVENTORY As String = "Инвентаризация"
Public Const SH_KITS As String = "Комплекты"
Public Const SH_DASHBOARD As String = "Дашборд"
Public Const SH_REPORT_INPUT As String = "Отчет — ввод"
Public Const SH_REPORT_FINAL As String = "Отчет руководителю"
Public Const SH_INFO As String = "Инфо"
Public Const SH_DIAGNOSTICS As String = "Диагностика PRIME"

Public Const SH_RECEIPT_SHOP As String = "Приход — Цех"
Public Const SH_ISSUE_SHOP As String = "Расход — Цех"
Public Const SH_RECEIPT_OFFICE As String = "Приход — Офис"
Public Const SH_ISSUE_OFFICE As String = "Расход — Офис"

' Легаси-формы 1.4.1 - только как архив истории, не активные формы ввода (см. ARCHITECTURE §5)
Public Const SH_LEGACY_RECEIPT_PROD As String = "Приход — Производство"
Public Const SH_LEGACY_ISSUE_PROD As String = "Расход — Производство"
Public Const SH_LEGACY_RECEIPT_PARTS As String = "Приход — Детали"
Public Const SH_LEGACY_ISSUE_PARTS As String = "Расход — Детали"

' --- Реестр схемы форм (header_schema_registry, PRIME 2.0.1) ---------------------------------
' Устраняет дефект 2.0.0: код молча предполагал, что заголовок таблицы всегда в строке 0
' (0-based). На деле большинство бизнес-листов унаследовали от шаблона 1.4.1 декоративную
' "шапку" (название + описание + строка статуса) ПЕРЕД реальной строкой заголовка колонок -
' у "Приход/Расход - Цех/Офис", "Возвраты", "Инвентаризация" заголовок физически в строке 4,
' у "Остаток" и "База - Поиск" - в строке 5. PRIME_Install_EnsureBusinessSheet при этом молча
' ничего не делает для уже существующих листов (сборка НЕ создаёт новый лист поверх старого),
' поэтому весь низкоуровневый доступ (PRIME_02_Store) обязан спрашивать здесь, где на самом
' деле искать заголовок конкретного листа, а не считать, что это всегда строка 0.
' Изменение HeaderRow для листа, у которого уже есть реальные пользовательские данные,
' требует соответствующей миграции (PRIME_14_MigrationInstaller) - само по себе изменение
' этой функции данные не переносит.
Public Function PRIME_FormSchemaHeaderRow(ByVal sheetName As String) As Long
    Select Case sheetName
        Case SH_RECEIPT_SHOP, SH_ISSUE_SHOP, SH_RECEIPT_OFFICE, SH_ISSUE_OFFICE, SH_RETURNS, SH_INVENTORY, SH_STOCK_ORDERS
            PRIME_FormSchemaHeaderRow = 4
        Case SH_STOCK, SH_SEARCH
            PRIME_FormSchemaHeaderRow = 5
        Case Else
            ' Заказы, Выдачи, Комплекты, Остаток — Заказы, все SYS_PRIME_*/DB_PRIME_* -
            ' заголовок в строке 0 (либо унаследовано от 1.4.1 без декоративной шапки,
            ' либо лист создаётся PRIME с нуля).
            PRIME_FormSchemaHeaderRow = 0
    End Select
End Function

' Первая строка данных = сразу после строки заголовка. Единая точка, используемая вместо
' разбросанных по коду "магических" 1/2 (header_schema_registry.rules).
Public Function PRIME_FormSchemaFirstDataRow(ByVal sheetName As String) As Long
    PRIME_FormSchemaFirstDataRow = PRIME_FormSchemaHeaderRow(sheetName) + 1
End Function

' --- Статус заказа (order_status_logic, PRIME 2.0.1) ---
Public Const ORDER_STATUS_DRAFT As String = "Черновик"
Public Const ORDER_STATUS_EXPECTED As String = "Ожидается"
Public Const ORDER_STATUS_PARTIAL As String = "Частично получено"
Public Const ORDER_STATUS_RECEIVED As String = "Получено"
Public Const ORDER_STATUS_OVERDUE As String = "Просрочено"
Public Const ORDER_STATUS_CANCELLED As String = "Отменено"

' Разбор пользовательского ввода даты (PRIME 2.0.1, recommendation dates support):
' "24.08", "24/08", "24-08", "24.08.2026", "24/08/2026", "24-08-2026". День/месяц идут первыми
' (формат заказчика, не ISO); год по умолчанию - текущий, если не указан. Возвращает
' нормализованную "YYYY-MM-DD" либо "" при нераспознанном вводе - лучше оставить ввод как есть
' (вызывающий не перезапишет ячейку), чем молча исказить дату по неверной догадке о формате.
Public Function PRIME_ParseFlexibleDate(ByVal inputText As String) As String
    Dim s As String
    s = Trim(inputText)
    If s = "" Then
        PRIME_ParseFlexibleDate = ""
        Exit Function
    End If

    Dim sep As String
    If InStr(s, ".") > 0 Then
        sep = "."
    ElseIf InStr(s, "/") > 0 Then
        sep = "/"
    ElseIf InStr(s, "-") > 0 Then
        sep = "-"
    Else
        PRIME_ParseFlexibleDate = ""
        Exit Function
    End If

    Dim parts() As String
    parts = Split(s, sep)
    If UBound(parts) < 1 Then
        PRIME_ParseFlexibleDate = ""
        Exit Function
    End If
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
        PRIME_ParseFlexibleDate = ""
        Exit Function
    End If

    Dim dayNum As Long, monthNum As Long, yearNum As Long
    dayNum = CLng(parts(0))
    monthNum = CLng(parts(1))
    ' StarBasic "And"/"Or" не короткозамкнуты (в отличие от AndAlso/OrElse) - оба операнда
    ' вычисляются ВСЕГДА, поэтому "UBound(parts) >= 2 And IsNumeric(parts(2))" обращался бы к
    ' parts(2) даже когда массив короче (только день+месяц, UBound=1), вызывая "Subscript out
    ' of range" - в этой среде такая ошибка внутри вызванной извне функции не долетает как
    ' видимое исключение, а молча обрывает вызывающий Sub целиком. Поэтому - вложенный If.
    yearNum = Year(Now)
    If UBound(parts) >= 2 Then
        If IsNumeric(parts(2)) Then
            yearNum = CLng(parts(2))
            If yearNum < 100 Then yearNum = yearNum + 2000
        End If
    End If

    If dayNum < 1 Or dayNum > 31 Or monthNum < 1 Or monthNum > 12 Then
        PRIME_ParseFlexibleDate = ""
        Exit Function
    End If

    PRIME_ParseFlexibleDate = Format(yearNum, "0000") & "-" & Format(monthNum, "00") & "-" & Format(dayNum, "00")
End Function

' --- Идентичность товара ---
Public Const PRODUCT_CODE_PREFIX As String = "ЕИ-"
Public Const PRODUCT_CODE_DIGITS As Integer = 8

' --- Типы документов (posting_engine.document_types) ---
Public Const DOC_RECEIPT As String = "RECEIPT"
Public Const DOC_ISSUE As String = "ISSUE"
Public Const DOC_RETURN As String = "RETURN"
Public Const DOC_TRANSFER As String = "TRANSFER"
Public Const DOC_ADJUSTMENT As String = "ADJUSTMENT"
Public Const DOC_ASSEMBLY As String = "ASSEMBLY"
Public Const DOC_DISASSEMBLY As String = "DISASSEMBLY"

' --- Состояния транзакции (transaction_model.states) ---
Public Const TX_PREPARED As String = "PREPARED"
Public Const TX_COMMITTED As String = "COMMITTED"
Public Const TX_FAILED As String = "FAILED"
Public Const TX_CANCELLED As String = "CANCELLED"

' --- Лимиты производительности / устойчивости ---
Public Const PRIME_QTY_PRECISION_DIGITS As Integer = 6
' Должно совпадать с "Lines(99) As PrimeDocLine" в каждой копии Type PrimeDocPlan
' (PRIME_04_Posting и модули 05/06/07/08, дублирующие Type - см. их комментарии).
Public Const PRIME_DOC_PLAN_MAX_LINE_INDEX As Long = 99
Public Const PRIME_IMPORT_TARGET_ROWS As Long = 20000
Public Const PRIME_MAX_EVENT_ROWS As Long = 1 ' PRIME_OnContentChanged всегда обрабатывает одну строку

' --- Пути (относительно каталога документа) ---
Public Const PRIME_DIR_BACKUPS As String = "Backups"
Public Const PRIME_DIR_ACTS As String = "Acts"
Public Const PRIME_DIR_EXPORTS As String = "Exports"
Public Const PRIME_DIR_DIAGNOSTICS As String = "Diagnostics"
Public Const PRIME_LOG_FILE As String = "prime_runtime.log"

' --- Этапы логирования (logging.stages) ---
Public Const STAGE_BUTTON_ENTER As String = "BUTTON_ENTER"
Public Const STAGE_VALIDATION_START As String = "VALIDATION_START"
Public Const STAGE_VALIDATION_OK As String = "VALIDATION_OK"
Public Const STAGE_PLAN_READY As String = "PLAN_READY"
Public Const STAGE_TX_PREPARED As String = "TX_PREPARED"
Public Const STAGE_DOCS_WRITTEN As String = "DOCS_WRITTEN"
Public Const STAGE_LINES_WRITTEN As String = "LINES_WRITTEN"
Public Const STAGE_LOTS_WRITTEN As String = "LOTS_WRITTEN"
Public Const STAGE_MOVEMENTS_WRITTEN As String = "MOVEMENTS_WRITTEN"
Public Const STAGE_TX_COMMITTED As String = "TX_COMMITTED"
Public Const STAGE_STORE_START As String = "STORE_START"
Public Const STAGE_STORE_OK As String = "STORE_OK"
Public Const STAGE_UI_FINALIZED As String = "UI_FINALIZED"
Public Const STAGE_ERROR As String = "ERROR"

' Заголовки листа "Заказы" - 25 бизнес-колонок 1.4.1, порядок сохранён 1:1 (ARCHITECTURE §5).
' "От кого / площадка" и "Продавец" соседствуют (orders_sheet.column_rule).
Public Function PRIME_OrdersBusinessColumns() As Variant
    Dim cols(24) As String
    cols(0)  = "Номер"
    cols(1)  = "Полное наименование товара"
    cols(2)  = "Код товара"
    cols(3)  = "Код поставщика"
    cols(4)  = "Артикул поставщика"
    cols(5)  = "От кого / площадка"
    cols(6)  = "Продавец"
    cols(7)  = "Источник прихода"
    cols(8)  = "№ документа"
    cols(9)  = "Номер счета"
    cols(10) = "Дата заказа"
    cols(11) = "Дата документа"
    cols(12) = "Дата поступления"
    cols(13) = "Количество"
    cols(14) = "Факт. количество"
    cols(15) = "Ед. изм."
    cols(16) = "Цена"
    cols(17) = "Сумма"
    cols(18) = "Покупатель"
    cols(19) = "Категория"
    cols(20) = "Подкатегория"
    cols(21) = "Место хранения"
    cols(22) = "Статус"
    cols(23) = "Контроль"
    cols(24) = "Комментарий"
    PRIME_OrdersBusinessColumns = cols
End Function

' Расширенный хвост (Ожидаемая дата, Назначение/проект, Получено всего, Осталось получить)
Public Function PRIME_OrdersExtraColumns() As Variant
    Dim cols(3) As String
    cols(0) = "Ожидаемая дата"
    cols(1) = "Назначение / проект"
    cols(2) = "Получено всего"
    cols(3) = "Осталось получить"
    PRIME_OrdersExtraColumns = cols
End Function

' Скрытые технические helper-колонки листа "Заказы" (avoid_many_hidden_columns => только 3, не 10 как в 1.4.1)
Public Function PRIME_OrdersHiddenColumns() As Variant
    Dim cols(2) As String
    cols(0) = "_PRIME_OrderID"
    cols(1) = "_PRIME_LineID"
    cols(2) = "_PRIME_State"
    PRIME_OrdersHiddenColumns = cols
End Function

' Единственная скрытая helper-колонка листа "Выдачи": стабильный ISSUE_DRAFT_ID (idempotency),
' присваивается один раз при создании строки, не меняется при повторных кликах "Провести".
Public Function PRIME_IssuesHiddenColumns() As Variant
    Dim cols(0) As String
    cols(0) = "_PRIME_IssueState"
    PRIME_IssuesHiddenColumns = cols
End Function

Public Function PRIME_IssuesColumns() As Variant
    Dim cols(12) As String
    cols(0)  = "№"
    cols(1)  = "Код"
    cols(2)  = "Наименование"
    cols(3)  = "Кол-во"
    cols(4)  = "Ед. изм."
    cols(5)  = "Кто получил"
    cols(6)  = "Дата"
    cols(7)  = "Откуда"
    cols(8)  = "Возвратный"
    cols(9)  = "Возвращено"
    cols(10) = "Дата возврата"
    cols(11) = "Назначение / проект"
    cols(12) = "Примечание"
    PRIME_IssuesColumns = cols
End Function

' Приход — Цех / Приход — Офис (единый workflow вместо разных "бухгалтерий" 1.4.1).
' "Внутренний код" - обязательное поле, по нему живой lookup (live_code_lookup."Приход — Цех"/"Офис").
Public Function PRIME_WorkflowReceiptColumns() As Variant
    Dim cols(11) As String
    cols(0)  = "Дата"
    cols(1)  = "Внутренний код"
    cols(2)  = "Наименование"
    cols(3)  = "Артикул"
    cols(4)  = "Кол-во"
    cols(5)  = "Ед. изм."
    cols(6)  = "Кто сдал"
    cols(7)  = "Место хранения"
    cols(8)  = "Категория"
    cols(9)  = "Подкатегория"
    cols(10) = "Документ"
    cols(11) = "Комментарий"
    PRIME_WorkflowReceiptColumns = cols
End Function

Public Function PRIME_WorkflowIssueColumns() As Variant
    Dim cols(11) As String
    cols(0)  = "Дата"
    cols(1)  = "Внутренний код"
    cols(2)  = "Наименование"
    cols(3)  = "Артикул"
    cols(4)  = "Кол-во"
    cols(5)  = "Ед. изм."
    cols(6)  = "Кому"
    cols(7)  = "Место хранения"
    cols(8)  = "Возвратный"
    cols(9)  = "Назначение / проект"
    cols(10) = "Документ"
    cols(11) = "Комментарий"
    PRIME_WorkflowIssueColumns = cols
End Function

' === Схемы скрытых системных листов (используются PRIME_14_MigrationInstaller.PRIME_Install_EnsureSchema
' и должны буквально совпадать с полями, которые читают/пишут PRIME_02_Store/03_Catalog/04_Posting -
' единственное место, где заводится новое поле, это одновременно и здесь, и в коде, который его использует).

Public Function PRIME_SysMetaColumns() As Variant
    PRIME_SysMetaColumns = Array("KEY", "VALUE")
End Function

Public Function PRIME_SysSeqColumns() As Variant
    PRIME_SysSeqColumns = Array("SEQ_NAME", "NEXT_VALUE")
End Function

Public Function PRIME_SysTxColumns() As Variant
    PRIME_SysTxColumns = Array("OP_ID", "SOURCE_KEY", "DOC_ID", "STATE", "STARTED_AT", "COMMITTED_AT", "HASH", "ERROR")
End Function

Public Function PRIME_DbProductsColumns() As Variant
    PRIME_DbProductsColumns = Array("PRODUCT_CODE", "PRODUCT_NAME", "BASE_UNIT", "DEFAULT_LOCATION", _
        "CATEGORY", "SUBCATEGORY", "RETURNABLE", "ACCOUNT_TYPE", "ACTIVE", "CREATED_AT", "UPDATED_AT")
End Function

Public Function PRIME_DbAliasesColumns() As Variant
    PRIME_DbAliasesColumns = Array("ALIAS_ID", "PRODUCT_CODE", "PLATFORM", "SELLER", "SUPPLIER_CODE", _
        "SUPPLIER_ARTICLE", "COMMENT", "ACTIVE")
End Function

Public Function PRIME_DbProductUnitsColumns() As Variant
    PRIME_DbProductUnitsColumns = Array("PRODUCT_CODE", "UNIT_NAME", "FACTOR_TO_BASE", "ACTIVE")
End Function

Public Function PRIME_DbDocumentsColumns() As Variant
    PRIME_DbDocumentsColumns = Array("DOC_ID", "DOC_TYPE", "DOC_DATE", "SOURCE_SHEET", "SOURCE_KEY", "ORDER_ID", "STATUS")
End Function

Public Function PRIME_DbDocLinesColumns() As Variant
    PRIME_DbDocLinesColumns = Array("DOC_LINE_ID", "DOC_ID", "PRODUCT_CODE", "QTY_BASE", "UNIT", "PRICE", _
        "LOCATION_FROM", "LOCATION_TO", "DESTINATION_PROJECT", "RECIPIENT", "COMMENT")
End Function

Public Function PRIME_DbMovementsColumns() As Variant
    PRIME_DbMovementsColumns = Array("MOVE_ID", "DOC_ID", "DOC_LINE_ID", "PRODUCT_CODE", "LOT_ID", "QTY_BASE", _
        "LOCATION", "MOVE_DATE", "OP_ID")
End Function

Public Function PRIME_DbLotsColumns() As Variant
    PRIME_DbLotsColumns = Array("LOT_ID", "PRODUCT_CODE", "RECEIPT_DOC_ID", "RECEIPT_LINE_ID", "RECEIPT_DATE", _
        "LOCATION", "ORIGINAL_QTY_BASE", "BASE_UNIT", "ORIGIN", "ORDER_ID")
End Function

Public Function PRIME_DbAllocationsColumns() As Variant
    PRIME_DbAllocationsColumns = Array("ALLOC_ID", "DOC_LINE_ID", "LOT_ID", "QTY_BASE")
End Function

Public Function PRIME_DbReturnsColumns() As Variant
    PRIME_DbReturnsColumns = Array("RETURN_ID", "ORIGINAL_ISSUE_DOC_LINE_ID", "RETURN_DOC_ID", "QTY_BASE", "RETURN_DATE")
End Function

Public Function PRIME_DbOrderSnapshotColumns() As Variant
    PRIME_DbOrderSnapshotColumns = Array("ORDER_ID", "RECEIPT_DOC_ID", "LOT_ID", "DELIVERY_QTY", "PRODUCT_CODE", "DOC_DATE")
End Function

' DB_PRIME_KITS/KIT_LINES заведены структурно (system_sheets), но в первом релизе фактическим
' источником данных комплектов служит видимый плоский лист "Комплекты" (PRIME_KitsColumns) -
' см. PRIME_11_Kits и известные ограничения в финальном отчёте.
Public Function PRIME_DbKitsColumns() As Variant
    PRIME_DbKitsColumns = Array("KIT_ID", "NAME", "VERSION", "ACTIVE")
End Function

Public Function PRIME_DbKitLinesColumns() As Variant
    PRIME_DbKitLinesColumns = Array("KIT_ID", "PRODUCT_CODE", "QTY_PER_KIT", "UNIT")
End Function

Public Function PRIME_DbActsColumns() As Variant
    PRIME_DbActsColumns = Array("ACT_ID", "DOC_ID", "FILE_PATH", "GENERATED_AT", "STATUS")
End Function

Public Function PRIME_DbAuditColumns() As Variant
    PRIME_DbAuditColumns = Array("TS", "OP_ID", "STAGE", "DURATION_MS", "SHEET", "ROW", "ERROR_NO", "ERROR_TEXT")
End Function

' Единая скрытая helper-колонка для всех 4 workflow-листов (стабильный SOURCE_KEY, по аналогии с Issues).
Public Function PRIME_WorkflowHiddenColumns() As Variant
    Dim cols(0) As String
    cols(0) = "_PRIME_WFState"
    PRIME_WorkflowHiddenColumns = cols
End Function

' Лист "Возвраты" - единственный механизм возврата (single_return_mechanism).
' "Вернуть сейчас" - редактируемое пользователем поле (не техническое), стоит последним
' перед скрытыми helper-колонками.
Public Function PRIME_ReturnsColumns() As Variant
    Dim cols(9) As String
    cols(0) = "Дата выдачи"
    cols(1) = "Код"
    cols(2) = "Наименование"
    cols(3) = "Выдано"
    cols(4) = "Уже возвращено"
    cols(5) = "Осталось к возврату"
    cols(6) = "Ед. изм."
    cols(7) = "Кому"
    cols(8) = "Вернуть сейчас"
    cols(9) = "Комментарий"
    PRIME_ReturnsColumns = cols
End Function

Public Function PRIME_ReturnsHiddenColumns() As Variant
    Dim cols(1) As String
    cols(0) = "_PRIME_OriginalLineId"
    cols(1) = "_PRIME_ReturnState"
    PRIME_ReturnsHiddenColumns = cols
End Function

' Лист "Остаток" - view поверх COMMITTED-движений, batch_output, без скрытого 2000-лимита.
Public Function PRIME_StockColumns() As Variant
    Dim cols(5) As String
    cols(0) = "Код"
    cols(1) = "Наименование"
    cols(2) = "Ед. изм."
    cols(3) = "Место хранения"
    cols(4) = "Остаток"
    cols(5) = "Последняя операция"
    PRIME_StockColumns = cols
End Function

' Лист "Остаток — Заказы" - снимки заказа + остаток соответствующей партии (include_zero_lot_balance).
Public Function PRIME_StockOrdersColumns() As Variant
    Dim cols(6) As String
    cols(0) = "ORDER_ID"
    cols(1) = "Дата"
    cols(2) = "Код товара"
    cols(3) = "Наименование"
    cols(4) = "Количество поставки"
    cols(5) = "LOT_ID"
    cols(6) = "Текущий остаток партии"
    PRIME_StockOrdersColumns = cols
End Function

' Лист "База - Поиск" - batch_search по документам/строкам/партиям.
Public Function PRIME_SearchColumns() As Variant
    Dim cols(6) As String
    cols(0) = "Тип"
    cols(1) = "Дата"
    cols(2) = "DOC_ID"
    cols(3) = "Код"
    cols(4) = "Наименование"
    cols(5) = "Количество"
    cols(6) = "Подробности"
    PRIME_SearchColumns = cols
End Function

' Лист "Инвентаризация".
' Лист "Комплекты" - плоская таблица (одна строка на компонент), как в ТЗ kits.fields.
' Первый релиз: только режим "набор для выдачи" (issue_bundle), без виртуального остатка комплекта.
Public Function PRIME_KitsColumns() As Variant
    Dim cols(6) As String
    cols(0) = "KIT_ID"
    cols(1) = "Название"
    cols(2) = "Версия"
    cols(3) = "PRODUCT_CODE"
    cols(4) = "Количество на 1 комплект"
    cols(5) = "Единица"
    cols(6) = "Активен"
    PRIME_KitsColumns = cols
End Function

Public Function PRIME_InventoryColumns() As Variant
    Dim cols(9) As String
    cols(0) = "Сессия"
    cols(1) = "Код"
    cols(2) = "Наименование"
    cols(3) = "Место"
    cols(4) = "Категория"
    cols(5) = "Подкатегория"
    cols(6) = "Учёт"
    cols(7) = "Факт"
    cols(8) = "Разница"
    cols(9) = "Ед. изм."
    PRIME_InventoryColumns = cols
End Function
