Option Explicit

' PRIME_00_Config
' Единая точка констант: имена листов, версии, лимиты, справочные значения.
' Никакой логики I/O здесь быть не должно (см. forbidden: очень крупные монолитные функции,
' циклические зависимости) - только константы и простые справочные функции.

Public Const PRIME_SCHEMA_VERSION As String = "2.1.0"
Public Const PRIME_BUILD_DATE As String = "2026-09-17"

' --- Контуры остатка (stock_architecture, PRIME 2.1.0) ---
' Один и тот же товар может физически лежать в трёх независимо учитываемых контурах: общий
' склад (Заказы/Выдачи/Возвраты/Перемещения), детали цеха (Приход/Расход - Цех) и офис
' (Приход/Расход - Офис). Контур - часть идентичности партии/движения (наравне с местом
' хранения), а не отдельная параллельная база - FIFO/остаток обязаны фильтровать по нему.
Public Const SC_GENERAL As String = "GENERAL"
Public Const SC_WORKSHOP_DETAILS As String = "WORKSHOP_DETAILS"
Public Const SC_OFFICE As String = "OFFICE"

Public Function PRIME_ContourForSheet(ByVal sheetName As String) As String
    Select Case sheetName
        Case SH_RECEIPT_SHOP, SH_ISSUE_SHOP
            PRIME_ContourForSheet = SC_WORKSHOP_DETAILS
        Case SH_RECEIPT_OFFICE, SH_ISSUE_OFFICE
            PRIME_ContourForSheet = SC_OFFICE
        Case Else
            PRIME_ContourForSheet = SC_GENERAL
    End Select
End Function

Public Function PRIME_ContourDisplayName(ByVal contour As String) As String
    Select Case contour
        Case SC_WORKSHOP_DETAILS
            PRIME_ContourDisplayName = "Детали цеха"
        Case SC_OFFICE
            PRIME_ContourDisplayName = "Офис"
        Case Else
            PRIME_ContourDisplayName = "Склад"
    End Select
End Function

' Обратное преобразование для ручного ввода на листе "Перемещения" (пользователь вводит
' привычное название контура, не технический ключ) - нераспознанное/пустое значение
' консервативно трактуется как общий склад (SC_GENERAL), а не как ошибка ввода.
Public Function PRIME_ContourFromDisplayName(ByVal displayName As String) As String
    Select Case LCase(Trim(displayName))
        Case "детали цеха", "цех", "workshop_details"
            PRIME_ContourFromDisplayName = SC_WORKSHOP_DETAILS
        Case "офис", "office"
            PRIME_ContourFromDisplayName = SC_OFFICE
        Case Else
            PRIME_ContourFromDisplayName = SC_GENERAL
    End Select
End Function

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
Public Const SH_DB_RETURN_ALLOCATIONS As String = "DB_PRIME_RETURN_ALLOCATIONS"

' --- Бизнес-листы (пользовательский UI, сохраняем привычные имена 1.4.1) ---
Public Const SH_ORDERS As String = "Заказы"
Public Const SH_ISSUES As String = "Выдачи"
' 2.1.0: "Остаток" переименован в "Наличие" (stock_architecture) - это сводный обзор, а не
' обязательный ежедневный экран (см. inline stock на рабочих листах). Старое имя листа 1.4.1/
' 2.0.x переносится сюда только через явную миграцию-переименование (см. MigrationInstaller),
' само по себе изменение этой константы существующий физический лист не переименовывает.
Public Const SH_STOCK As String = "Наличие"
Public Const SH_STOCK_LEGACY_NAME As String = "Остаток"
Public Const SH_STOCK_ORDERS As String = "Остаток — Заказы"
Public Const SH_SEARCH As String = "База - Поиск"
Public Const SH_RETURNS As String = "Возвраты"
Public Const SH_INVENTORY As String = "Инвентаризация"
Public Const SH_KITS As String = "Комплекты"
Public Const SH_TRANSFERS As String = "Перемещения"
Public Const SH_JOURNAL As String = "Журнал"
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
' Должно совпадать с "Lines(999) As PrimeDocLine" в каждой копии Type PrimeDocPlan
' (PRIME_04_Posting и модули 05/06/07/08/15, дублирующие Type - см. их комментарии).
' 2.1.0: поднято с 99 до 999 (минимум 1000 строк в одном документе, mandatory_document_size) -
' одна поставка/выдача/возврат/перемещение из многих строк проводится ОДНИМ документом
' (DOC_ID/OP_ID/store()), а не циклом отдельных PostDocument-вызовов.
Public Const PRIME_DOC_PLAN_MAX_LINE_INDEX As Long = 999
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
Public Const STAGE_PRODUCTS_WRITTEN As String = "PRODUCTS_WRITTEN"
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

' Расширенный хвост (Ожидаемая дата, Назначение/проект, Получено всего, Осталось получить +
' 2.1.0 inline stock/traceability: В наличии сейчас, Последний приход, Дата последнего прихода).
Public Function PRIME_OrdersExtraColumns() As Variant
    Dim cols(6) As String
    cols(0) = "Ожидаемая дата"
    cols(1) = "Назначение / проект"
    cols(2) = "Получено всего"
    cols(3) = "Осталось получить"
    cols(4) = "В наличии сейчас"
    cols(5) = "Последний приход"
    cols(6) = "Дата последнего прихода"
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

' 2.1.0 inline stock (R23): "В наличии" -> "Кол-во" -> "После выдачи" - обе вычисляемые колонки
' обновляются автозаполнением по коду/вводу количества (см. PRIME_06_Issues) и не хранят
' независимый остаток - это представление того же COMMITTED-ledger, что и лист "Наличие".
Public Function PRIME_IssuesColumns() As Variant
    Dim cols(14) As String
    cols(0)  = "№"
    cols(1)  = "Код"
    cols(2)  = "Наименование"
    cols(3)  = "В наличии"
    cols(4)  = "Кол-во"
    cols(5)  = "После выдачи"
    cols(6)  = "Ед. изм."
    cols(7)  = "Кто получил"
    cols(8)  = "Дата"
    cols(9)  = "Откуда"
    cols(10) = "Возвратный"
    cols(11) = "Возвращено"
    cols(12) = "Дата возврата"
    cols(13) = "Назначение / проект"
    cols(14) = "Примечание"
    PRIME_IssuesColumns = cols
End Function

' Приход — Цех / Приход — Офис (единый workflow вместо разных "бухгалтерий" 1.4.1).
' "Внутренний код" - обязательное поле, по нему живой lookup (live_code_lookup."Приход — Цех"/"Офис").
' 2.1.0 inline stock (R19/R21): "Остаток .../Приход/Будет ..." - подпись зависит от контура
' листа (Цех -> "деталей", Офис -> "офиса"/"в офисе"), но логика (PRIME_07_Workflows) одна и та
' же для обоих контуров - контур определяет только WHERE считается остаток, не КАК.
Public Function PRIME_WorkflowReceiptColumns(ByVal sheetName As String) As Variant
    Dim beforeLabel As String, afterLabel As String
    If sheetName = SH_RECEIPT_OFFICE Then
        beforeLabel = "Остаток офиса"
        afterLabel = "Будет в офисе"
    Else
        beforeLabel = "Остаток деталей"
        afterLabel = "Будет деталей"
    End If
    Dim cols(13) As String
    cols(0)  = "Дата"
    cols(1)  = "Внутренний код"
    cols(2)  = "Наименование"
    cols(3)  = "Артикул"
    cols(4)  = beforeLabel
    cols(5)  = "Кол-во"
    cols(6)  = afterLabel
    cols(7)  = "Ед. изм."
    cols(8)  = "Кто сдал"
    cols(9)  = "Место хранения"
    cols(10) = "Категория"
    cols(11) = "Подкатегория"
    cols(12) = "Документ"
    cols(13) = "Комментарий"
    PRIME_WorkflowReceiptColumns = cols
End Function

Public Function PRIME_WorkflowIssueColumns(ByVal sheetName As String) As Variant
    Dim beforeLabel As String, afterLabel As String
    If sheetName = SH_ISSUE_OFFICE Then
        beforeLabel = "Остаток офиса"
        afterLabel = "Будет в офисе"
    Else
        beforeLabel = "Остаток деталей"
        afterLabel = "Будет деталей"
    End If
    Dim cols(13) As String
    cols(0)  = "Дата"
    cols(1)  = "Внутренний код"
    cols(2)  = "Наименование"
    cols(3)  = "Артикул"
    cols(4)  = beforeLabel
    cols(5)  = "Кол-во"
    cols(6)  = afterLabel
    cols(7)  = "Ед. изм."
    cols(8)  = "Кому"
    cols(9)  = "Место хранения"
    cols(10) = "Возвратный"
    cols(11) = "Назначение / проект"
    cols(12) = "Документ"
    cols(13) = "Комментарий"
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
    ' OP_ID: пусто = видим всегда (миграция/легаси/прямое создание); непусто = видим только когда
    ' PRIME_IsOpIdCommitted(OP_ID)=True (транзакционное создание нового товара внутри проведения,
    ' см. PRIME_04_Posting.PRIME_WriteNewProductsIfAny и PRIME_03_Catalog.PRIME_BuildProductIndex).
    PRIME_DbProductsColumns = Array("PRODUCT_CODE", "PRODUCT_NAME", "BASE_UNIT", "DEFAULT_LOCATION", _
        "CATEGORY", "SUBCATEGORY", "RETURNABLE", "ACCOUNT_TYPE", "ACTIVE", "CREATED_AT", "UPDATED_AT", "OP_ID")
End Function

Public Function PRIME_DbAliasesColumns() As Variant
    PRIME_DbAliasesColumns = Array("ALIAS_ID", "PRODUCT_CODE", "PLATFORM", "SELLER", "SUPPLIER_CODE", _
        "SUPPLIER_ARTICLE", "COMMENT", "ACTIVE")
End Function

Public Function PRIME_DbProductUnitsColumns() As Variant
    PRIME_DbProductUnitsColumns = Array("PRODUCT_CODE", "UNIT_NAME", "FACTOR_TO_BASE", "ACTIVE")
End Function

' OP_ID добавлен в 2.1.0 (committed_only_everywhere): позволяет проверить, что владеющая
' транзакция реально COMMITTED, не читая отдельный SYS_PRIME_TX по SOURCE_KEY - см.
' PRIME_02_Store.PRIME_IsDocIdCommitted.
Public Function PRIME_DbDocumentsColumns() As Variant
    PRIME_DbDocumentsColumns = Array("DOC_ID", "DOC_TYPE", "DOC_DATE", "SOURCE_SHEET", "SOURCE_KEY", "ORDER_ID", "STATUS", "OP_ID")
End Function

Public Function PRIME_DbDocLinesColumns() As Variant
    PRIME_DbDocLinesColumns = Array("DOC_LINE_ID", "DOC_ID", "PRODUCT_CODE", "QTY_BASE", "UNIT", "PRICE", _
        "LOCATION_FROM", "LOCATION_TO", "DESTINATION_PROJECT", "RECIPIENT", "COMMENT")
End Function

' STOCK_CONTOUR добавлен в 2.1.0 (stock_architecture) - см. PRIME_ContourForSheet.
Public Function PRIME_DbMovementsColumns() As Variant
    PRIME_DbMovementsColumns = Array("MOVE_ID", "DOC_ID", "DOC_LINE_ID", "PRODUCT_CODE", "LOT_ID", "QTY_BASE", _
        "LOCATION", "MOVE_DATE", "OP_ID", "STOCK_CONTOUR")
End Function

' STOCK_CONTOUR/PARENT_LOT_ID добавлены в 2.1.0: партия физически привязана к одному
' месту+контуру и никогда не "телепортируется" - перемещение создаёт НОВУЮ партию с
' PARENT_LOT_ID = исходная (lot lineage, critical_fixes.transfer), инвентаризация консервативно
' расходует/создаёт партии так же, как обычное списание/приход (critical_fixes.inventory).
Public Function PRIME_DbLotsColumns() As Variant
    PRIME_DbLotsColumns = Array("LOT_ID", "PRODUCT_CODE", "RECEIPT_DOC_ID", "RECEIPT_LINE_ID", "RECEIPT_DATE", _
        "LOCATION", "ORIGINAL_QTY_BASE", "BASE_UNIT", "ORIGIN", "ORDER_ID", "STOCK_CONTOUR", "PARENT_LOT_ID")
End Function

Public Function PRIME_DbAllocationsColumns() As Variant
    PRIME_DbAllocationsColumns = Array("ALLOC_ID", "DOC_LINE_ID", "LOT_ID", "QTY_BASE")
End Function

' OP_ID добавлен в 2.1.0 - без него неудавшийся (FAILED) возврат, чьи строки уже физически
' попали в лист до отката, ошибочно уменьшал бы "Осталось к возврату" навсегда (см. критику
' в REQUIREMENTS_MATRIX R07/R08).
Public Function PRIME_DbReturnsColumns() As Variant
    PRIME_DbReturnsColumns = Array("RETURN_ID", "ORIGINAL_ISSUE_DOC_LINE_ID", "RETURN_DOC_ID", "QTY_BASE", "RETURN_DATE", "OP_ID")
End Function

' Новая таблица 2.1.0 (R08): какая ЧАСТЬ какой конкретной allocation исходной выдачи уже была
' возвращена - без неё второй частичный возврат снова "с нуля" проходит allocations по FIFO и
' может повторно вернуть в ту же самую партию, из которой уже был засчитан первый возврат,
' вместо следующей по очереди (см. REQUIREMENTS_MATRIX R08).
Public Function PRIME_DbReturnAllocationsColumns() As Variant
    PRIME_DbReturnAllocationsColumns = Array("RETURN_ID", "ALLOC_ID", "LOT_ID", "QTY_BASE", "OP_ID")
End Function

' Полный immutable snapshot заказа (2.1.0, R13) - копия всех значимых реквизитов заказа И
' поставки на момент проведения прихода, не только количества/партии, как в 2.0.x. Реквизиты
' поставщика/счёта/цены и т.п. не восстановимы из мутируемого листа "Заказы" впоследствии,
' поэтому должны быть скопированы сюда в момент commit, а не вычисляться на лету при просмотре.
Public Function PRIME_DbOrderSnapshotColumns() As Variant
    PRIME_DbOrderSnapshotColumns = Array( _
        "ORDER_ID", "ORDER_LINE_ID", "RECEIPT_DOC_ID", "RECEIPT_LINE_ID", "LOT_ID", _
        "PRODUCT_CODE", "PRODUCT_NAME", "SUPPLIER_OR_PLATFORM", "SELLER", "SUPPLIER_CODE", _
        "SUPPLIER_ARTICLE", "INVOICE_NUMBER", "DOCUMENT_NUMBER", "ORDER_DATE", "EXPECTED_DATE", _
        "DOCUMENT_DATE", "RECEIPT_DATE", "ORDERED_QTY", "RECEIVED_QTY", "UNIT", "PRICE", _
        "AMOUNT", "BUYER", "CATEGORY", "SUBCATEGORY", "LOCATION", "STOCK_CONTOUR", _
        "DESTINATION_PROJECT", "COMMENT")
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

' Лист "Наличие" (2.1.0, было "Остаток") - сводный view поверх COMMITTED-движений всех контуров,
' batch_output, без скрытого 2000-лимита. Не обязательный ежедневный экран - остаток по каждому
' контуру дублируется inline на рабочих листах (см. PRIME_OrdersExtraColumns/WorkflowReceipt/
' IssueColumns/IssuesColumns), здесь - только общий обзор с фильтром по контуру.
Public Function PRIME_StockColumns() As Variant
    Dim cols(6) As String
    cols(0) = "Код"
    cols(1) = "Наименование"
    cols(2) = "Контур"
    cols(3) = "Место хранения"
    cols(4) = "Ед. изм."
    cols(5) = "Остаток"
    cols(6) = "Последняя операция"
    PRIME_StockColumns = cols
End Function

' Лист "Перемещения" (2.1.0, новый) - TRANSFER между контурами/местами с сохранением партийности
' (lot lineage) - см. PRIME_15_Transfers/PRIME_04_Posting.PRIME_PostTransferLines.
Public Function PRIME_TransfersColumns() As Variant
    Dim cols(9) As String
    cols(0) = "Дата"
    cols(1) = "Внутренний код"
    cols(2) = "Наименование"
    cols(3) = "Кол-во"
    cols(4) = "Ед. изм."
    cols(5) = "Контур — откуда"
    cols(6) = "Место — откуда"
    cols(7) = "Контур — куда"
    cols(8) = "Место — куда"
    cols(9) = "Комментарий"
    PRIME_TransfersColumns = cols
End Function

Public Function PRIME_TransfersHiddenColumns() As Variant
    Dim cols(0) As String
    cols(0) = "_PRIME_TransferState"
    PRIME_TransfersHiddenColumns = cols
End Function

' Лист "Журнал" (2.1.0, новый) - только COMMITTED документы, read-only, по кнопке "Обновить"
' (batch_output, не построчный пересчёт).
Public Function PRIME_JournalColumns() As Variant
    Dim cols(9) As String
    cols(0) = "DOC_ID"
    cols(1) = "OP_ID"
    cols(2) = "Тип"
    cols(3) = "Дата/время"
    cols(4) = "Статус"
    cols(5) = "Источник"
    cols(6) = "Контур"
    cols(7) = "Поставщик/Получатель"
    cols(8) = "Количество строк"
    cols(9) = "Комментарий"
    PRIME_JournalColumns = cols
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

' 2.1.0: добавлена "Контур" - инвентаризация теперь снимает остаток по каждому (месту, контуру)
' отдельно (R06 lot-consistency), а не только по месту, иначе один и тот же код на одном месте,
' но в разных контурах (склад/детали цеха), задваивал бы разницу.
Public Function PRIME_InventoryColumns() As Variant
    Dim cols(10) As String
    cols(0) = "Сессия"
    cols(1) = "Код"
    cols(2) = "Наименование"
    cols(3) = "Место"
    cols(4) = "Контур"
    cols(5) = "Категория"
    cols(6) = "Подкатегория"
    cols(7) = "Учёт"
    cols(8) = "Факт"
    cols(9) = "Разница"
    cols(10) = "Ед. изм."
    PRIME_InventoryColumns = cols
End Function
