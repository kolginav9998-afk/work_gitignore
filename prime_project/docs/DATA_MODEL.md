# PRIME 2.0.1 — модель данных

Описывает РЕАЛЬНО реализованную схему (соответствует `src/basic/PRIME_00_Config.bas`
`PRIME_Db*Columns()`/`PRIME_Sys*Columns()` функциям и `PRIME_14_MigrationInstaller.PRIME_Install_EnsureSchema`,
которые физически создают эти листы) — не намерение на будущее.

Все системные листы: скрыты, без ручного редактирования, только пакетный I/O через
`PRIME_02_Store`. Остаток — VIEW поверх `DB_PRIME_MOVEMENTS` (только `COMMITTED`-движения,
проверено фильтром `PRIME_IsOpIdCommitted` с 2.0.1 — см. `POSTING_PROTOCOL.md`), не хранимое поле.

## Реестр строки заголовка (header_schema_registry, PRIME 2.0.1)

Не все бизнес-листы имеют заголовок колонок в строке 0. Листы, унаследованные от шаблона 1.4.1
с декоративной "шапкой" (название + описание + строка статуса) перед реальной таблицей, имеют
физический заголовок в строке 4 или 5 — см. `PRIME_00_Config.PRIME_FormSchemaHeaderRow`:

| Строка заголовка | Листы |
|---|---|
| 0 | Заказы, Выдачи, Комплекты, все `SYS_PRIME_*`/`DB_PRIME_*` |
| 4 | Приход/Расход — Цех/Офис, Возвраты, Инвентаризация, Остаток — Заказы |
| 5 | Остаток, База - Поиск |

Весь низкоуровневый доступ (`PRIME_02_Store`: `PRIME_HeaderMap`, `PRIME_ReadTable`,
`PRIME_AppendRowsBatch`, `PRIME_UpdateRowsBatch`, `PRIME_FindLastRow`, `PRIME_ClearDataRows`)
берёт номер строки заголовка отсюда, а не предполагает строку 0 — до 2.0.1 это предположение
было жёстко зашито, из-за чего почти все бизнес-листы, кроме "Заказы", фактически работали на
исходной раскладке колонок 1.4.1, а не на схеме PRIME (см. `CHANGELOG.md`, `KNOWN_ISSUES.md` №16).

## Системные листы (SYS_PRIME_*)

### SYS_PRIME_META
| Поле | Назначение |
|---|---|
| KEY | версия схемы, дата установки/миграции (`SCHEMA_VERSION`, `INSTALLED_AT`, `MIGRATED_AT`, `MIGRATED_FROM`, `CURRENT_INVENTORY_SESSION`, `INV_SNAPSHOT_<session>`) |
| VALUE | значение |

### SYS_PRIME_SEQ
| Поле | Назначение |
|---|---|
| SEQ_NAME | PRODUCT_CODE, ORDER_ID, ORDER_LINE_ID, DOC_ID, LOT_ID, MOVE_ID, ISSUE_DRAFT_ID, WF_DRAFT_ID, INVENTORY_SESSION_ID, ACT_ID и т.д. — заводится автоматически при первом использовании |
| NEXT_VALUE | последнее выданное значение |

`PRIME_SequenceNext(seqName)` — единственная точка выдачи новых ID/кодов.

### SYS_PRIME_TX
| Поле | Назначение |
|---|---|
| OP_ID | уникальный идентификатор операции (timestamp+random) |
| SOURCE_KEY | идемпотентность (см. POSTING_PROTOCOL.md) |
| DOC_ID | ссылка на DB_PRIME_DOCUMENTS |
| STATE | PREPARED / COMMITTED / FAILED |
| STARTED_AT / COMMITTED_AT | текстовые timestamp |
| HASH, ERROR | зарезервировано / текст ошибки |

## Товарный справочник

### DB_PRIME_PRODUCTS
`PRODUCT_CODE (PK, "ЕИ-00000001…"), PRODUCT_NAME, BASE_UNIT, DEFAULT_LOCATION, CATEGORY, SUBCATEGORY, RETURNABLE, ACCOUNT_TYPE, ACTIVE, CREATED_AT, UPDATED_AT`

### DB_PRIME_ALIASES
`ALIAS_ID, PRODUCT_CODE, PLATFORM, SELLER, SUPPLIER_CODE, SUPPLIER_ARTICLE, COMMENT, ACTIVE`.
Составной контекст = «От кого/площадка» + «Продавец» + «Код поставщика» + «Артикул поставщика».
При неоднозначном совпадении — не угадывать (`PRIME_ALIAS_AMBIGUOUS`), требовать выбор пользователя.
**Известное ограничение:** нет UI для добавления/редактирования алиасов в первом релизе (см. KNOWN_ISSUES.md) -
заполняется прямым редактированием скрытого листа или будущей миграцией.

### DB_PRIME_PRODUCT_UNITS
`PRODUCT_CODE, UNIT_NAME, FACTOR_TO_BASE, ACTIVE`. Базовая единица товара всегда имеет фактор 1
неявно. Отсутствие фактора для указанной единицы блокирует **весь** документ.

## Документы и движения

### DB_PRIME_DOCUMENTS
`DOC_ID, DOC_TYPE, DOC_DATE, SOURCE_SHEET, SOURCE_KEY, ORDER_ID, STATUS`.
DOC_TYPE ∈ {RECEIPT, ISSUE, RETURN, TRANSFER, ADJUSTMENT, ASSEMBLY, DISASSEMBLY} (последние два —
архитектурно готовы в `PRIME_04_Posting`, без UI в первом релизе).

### DB_PRIME_DOC_LINES
`DOC_LINE_ID, DOC_ID, PRODUCT_CODE, QTY_BASE, UNIT, PRICE, LOCATION_FROM, LOCATION_TO,
DESTINATION_PROJECT, RECIPIENT, COMMENT`.
`DESTINATION_PROJECT` явно присутствует и всегда записывается — закрывает дефект 1.4.1
(`WMSWF_IssueRowCon` читал «Назначение», но не сохранял).

### DB_PRIME_MOVEMENTS
`MOVE_ID, DOC_ID, DOC_LINE_ID, PRODUCT_CODE, LOT_ID, QTY_BASE (+/-), LOCATION, MOVE_DATE, OP_ID`.
**Единственный источник истины по складу.** Остаток товара = сумма QTY_BASE по всем движениям
этого товара (`PRIME_TotalLotBalance`/`PRIME_StockByLocation` в `PRIME_03_Catalog`/`PRIME_04_Posting`).

### DB_PRIME_LOTS
`LOT_ID, PRODUCT_CODE, RECEIPT_DOC_ID, RECEIPT_LINE_ID, RECEIPT_DATE, LOCATION,
ORIGINAL_QTY_BASE, BASE_UNIT, ORIGIN, ORDER_ID`.
Текущий остаток партии = сумма связанных COMMITTED-движений с этим LOT_ID (не хранимое поле).

### DB_PRIME_ALLOCATIONS
`ALLOC_ID, DOC_LINE_ID (строка выдачи), LOT_ID, QTY_BASE` — результат FIFO-разбиения одной
строки выдачи на несколько партий; используется для симметричного возврата в партии.

### DB_PRIME_RETURNS
`RETURN_ID, ORIGINAL_ISSUE_DOC_LINE_ID, RETURN_DOC_ID, QTY_BASE, RETURN_DATE`.
Лимит на возврат = `issued_qty (из DOC_LINES) - SUM(уже возвращённого по этому ORIGINAL_ISSUE_DOC_LINE_ID)`.

### DB_PRIME_ORDER_SNAPSHOT
`ORDER_ID, RECEIPT_DOC_ID, LOT_ID, DELIVERY_QTY, PRODUCT_CODE, DOC_DATE` — неизменяемый снимок
каждой поставки по заказу; источник для листа «Остаток — Заказы».

### DB_PRIME_KITS / DB_PRIME_KIT_LINES
`KIT_ID, NAME, VERSION, ACTIVE` / `KIT_ID, PRODUCT_CODE, QTY_PER_KIT, UNIT`.
**Известное ограничение:** структурно созданы (`PRIME_Install_EnsureSchema`), но фактический
источник данных комплектов в первом релизе — видимый плоский лист «Комплекты»
(`PRIME_KitsColumns`), не эти скрытые таблицы (см. `PRIME_11_Kits.bas` и KNOWN_ISSUES.md).

### DB_PRIME_ACTS
`ACT_ID, DOC_ID, FILE_PATH, GENERATED_AT, STATUS` — акт всегда строится по DOC_ID из журнала,
не из текущего выделения ячеек.

### DB_PRIME_AUDIT
`TS, OP_ID, STAGE, DURATION_MS, SHEET, ROW, ERROR_NO, ERROR_TEXT` — этапы см. POSTING_PROTOCOL.md.
Запись в этот лист обёрнута в `On Error Resume Next` (единственное осознанное исключение из
общего правила «без подавления ошибок» — сбой логирования не должен ронять проведение).

## Бизнес-листы (видимые, пользовательские)

| Лист | Бизнес-колонок | Скрытые helper-колонки |
|---|---|---|
| Заказы | 25 исходных 1.4.1 + 4 новых (Ожидаемая дата, Назначение/проект, Получено всего, Осталось получить) | `_PRIME_OrderID`, `_PRIME_LineID`, `_PRIME_State` (3, было 10 `_WMS_*` в 1.4.1) |
| Выдачи | 12 исходных + Назначение/проект | `_PRIME_IssueState` |
| Приход/Расход — Цех, Офис (4 листа) | Дата, Внутренний код, Наименование, Артикул, Кол-во, Ед.изм., контрагент, Место, категория/подкатегория (приход) или Возвратный/Назначение (расход), Документ, Комментарий | `_PRIME_WFState` |
| Возвраты | Дата выдачи, Код, Наименование, Выдано, Уже возвращено, Осталось к возврату, Ед.изм., Кому, Вернуть сейчас, Комментарий | `_PRIME_OriginalLineId`, `_PRIME_ReturnState` |
| Инвентаризация | Сессия, Код, Наименование, Место, Категория, Подкатегория, Учёт, Факт, Разница, Ед.изм. | — |
| Комплекты | KIT_ID, Название, Версия, PRODUCT_CODE, Количество на 1 комплект, Единица, Активен | — |
| Остаток | Код, Наименование, Ед.изм., Место хранения, Остаток, Последняя операция | — |
| Остаток — Заказы | ORDER_ID, Дата, Код товара, Наименование, Количество поставки, LOT_ID, Текущий остаток партии | — |
| База - Поиск | Тип, Дата, DOC_ID, Код, Наименование, Количество, Подробности | — |

Точный порядок и полный список колонок — единственный источник истины в
`src/basic/PRIME_00_Config.bas` (функции `PRIME_*Columns()`); эта таблица — обзор, не копия.
