# PRIME 2.0.1 — протокол проведения документов

Описывает РЕАЛЬНО реализованный протокол в `src/basic/PRIME_04_Posting.bas`
(`PRIME_PostDocument`) — единственную функцию, которая пишет в `DB_PRIME_*`/`SYS_PRIME_TX`.
Формы (Заказы, Выдачи, цеховые/офисные листы, Возвраты, Инвентаризация) только строят
`PrimeDocPlan` в памяти и вызывают `PRIME_PostDocument`; сами они никогда не пишут в
системные листы напрямую.

## 1. Структура плана документа (`PrimeDocPlan`/`PrimeDocLine`)

`PrimeDocPlan`: `DocType, DocDate, SourceSheet, SourceKey, OrderId, Lines(99), LineCount`.
`PrimeDocLine`: `ProductCode, ProductName, QtyInput, UnitInput, LocationFrom, LocationTo,
DestinationProject, Recipient, Comment, Price, OriginalDocLineId, QtyBase, LotId`.

`DocType` ∈ `RECEIPT | ISSUE | RETURN | TRANSFER | ADJUSTMENT` (константы `DOC_*` в
`PRIME_00_Config`). `ASSEMBLY`/`DISASSEMBLY` упомянуты в `DATA_MODEL.md` как архитектурно
предусмотренные в схеме документов, но `PRIME_PostDocument` для них ветки `Select Case` не
реализует — попытка провести такой `DocType` завершится ошибкой 1011 "Неизвестный тип документа".

Лимит строк одной операции — 100 (`Lines(99)`, константа `PRIME_DOC_PLAN_MAX_LINE_INDEX = 99`
в `PRIME_00_Config`, не `UBound()` — см. `KNOWN_ISSUES.md`/комментарий в коде про ограничение
StarBasic на `UBound` поля-массива в `Type`). Документ с 101-й строкой завершается ошибкой 1010
до попытки проведения.

## 2. `PRIME_PostDocument` — пошаговый алгоритм (как реализовано)

```
Если Not PRIME_TryEnter() Then                      ' 2.0.1: реальный boolean-мьютекс, не декорация
    вернуть "" немедленно, ничего не писать, LastPostError = "операция уже выполняется"
AuditLog(BUTTON_ENTER)
On Error Goto PostFailed
  1. existingDoc = FindCommittedBySourceKey(SourceKey)
     Если найден -> Leave(), вернуть existingDoc (НЕ создавать новых движений)      [идемпотентность]
  2. ValidateAndExpandPlan(plan)                     ' по DocType, см. §3
     Если ошибка -> Leave(), вернуть "", LastPostError = текст
  3. opId = NewOpId() ; docId = "DOC-" & SequenceNext("DOC_ID")
     WriteTxRow(opId, SourceKey, docId, PREPARED)                                  [TX_PREPARED]
  4. WriteDocumentHeader(docId, plan)  -> DB_PRIME_DOCUMENTS (1 строка)
     WriteDocLines(docId, plan)        -> DB_PRIME_DOC_LINES (LineCount строк)
     Select Case DocType:
        RECEIPT     -> PostReceiptLines     (создаёт LOT на строку + движение прихода [+ снимок заказа])
        ISSUE       -> PostIssueLines       (FIFO по партиям: allocations + движения списания)
        RETURN      -> PostReturnLines      (движения по allocations исходной выдачи + DB_PRIME_RETURNS)
        TRANSFER    -> PostTransferLines    (2 движения на строку: -QTY на LocationFrom, +QTY на LocationTo)
        ADJUSTMENT  -> PostAdjustmentLines  (1 движение со знаком QtyBase)
  5. UpdateTxState(opId, COMMITTED)                                                [TX_COMMITTED]
     RegisterCommittedKey(SourceKey, docId)         ' в памяти, кэш SOURCE_KEY->DOC_ID; committedKeyRegistered=True
  6. ThisComponent.store()                          ' один store() на весь документ
  7. Leave() ; вернуть docId
PostFailed:
  Если TX-строка уже была записана -> UpdateTxState(opId, FAILED, errText)  (On Error Resume Next)
  Если committedKeyRegistered -> UnregisterCommittedKey(SourceKey)    ' 2.0.1: см. §4
  AuditLog(ERROR) ; Leave() ; вернуть ""
```

**Важно (2.0.1):** проверка `PRIME_TryEnter()` — это первое, что делает функция, и если лок уже
занят (операция уже выполняется), функция **немедленно возвращается**, не доходя даже до
`AuditLog(BUTTON_ENTER)` и не выполняя `PRIME_Leave()` (лок не был захвачен этим вызовом — снимать
нечего). В 2.0.0 результат `PRIME_TryEnter()` не проверялся вовсе: функция всегда продолжала
выполнение независимо от того, что вернул лок — то есть двойной клик/повторный вызов реально не
блокировался. Это было найдено и исправлено при разборе для 2.0.1 (см. `CHANGELOG.md`).

Каждый этап логируется в `DB_PRIME_AUDIT` через `PRIME_AuditLog` со стадией из
`PRIME_00_Config` (`BUTTON_ENTER, VALIDATION_START, VALIDATION_OK, TX_PREPARED, DOCS_WRITTEN,
LINES_WRITTEN, LOTS_WRITTEN, MOVEMENTS_WRITTEN, TX_COMMITTED, STORE_START, STORE_OK,
UI_FINALIZED, ERROR`). `PRIME_AuditLog` сам обёрнут в `On Error Resume Next` — единственное
осознанное исключение из общего правила «без подавления ошибок»: сбой логирования не должен
ронять проведение (см. `golden_invariants`, KNOWN_ISSUES.md).

`PRIME_Leave()` вызывается на КАЖДОМ пути выхода функции (успех, идемпотентный повтор, ошибка
валидации, ошибка исполнения через `PostFailed`) — не в отдельном `On Error Resume Next` вокруг
записи, а явно на каждой ветке.

## 3. Валидация по типу документа (`PRIME_ValidateAndExpandPlan`)

Общая часть для всех типов: каждая строка должна иметь либо `ProductCode`, либо
`ProductName` (для создания нового товара); `QtyInput <> 0`; `QtyInput < 0` разрешён
только для `ADJUSTMENT` (недостача инвентаризации).

- **RECEIPT**: если `ProductCode` пуст — создаётся новая карточка товара (`PRIME_CreateProduct`,
  только когда явно нет кода). Конвертация количества в базовую единицу — обязательна;
  отсутствие коэффициента для указанной единицы блокирует **весь** документ. `LotId`
  предварительно помечается `"LOT-"`, финальный номер партии присваивается в `PostReceiptLines`.
- **ISSUE**: конвертация в базовую единицу + проверка `PRIME_TotalLotBalance(ProductCode) >=
  QtyBase` по КАЖДОЙ строке. Недостаточность остатка хотя бы в одной строке отклоняет весь
  документ ещё до записи (`shortage never leaves a partial issue`).
- **RETURN**: `QtyBase` не может превышать `issuedQty - alreadyReturnedQty` (допуск
  `0.0000005` на плавающую точку) для `OriginalDocLineId`, где `issuedQty` берётся из
  `DB_PRIME_DOC_LINES`, а `alreadyReturnedQty` — сумма всех строк `DB_PRIME_RETURNS` по этому же
  `OriginalDocLineId` (не только по одному предыдущему возврату).
- **TRANSFER**: остаток на конкретном `LocationFrom` (не суммарный остаток товара) должен быть
  `>= QtyBase`; берётся через `PRIME_StockByLocation` (парные массивы мест/количеств, см.
  `PRIME_03_Catalog` — Collection не может перечислить собственные ключи в StarBasic, поэтому
  используется этот, а не Collection-based, механизм).
- **ADJUSTMENT**: знак сохраняется через конвертацию (`Sgn(QtyInput)` применяется после
  конвертации `Abs(QtyInput)` в базовую единицу) — недостача инвентаризации даёт отрицательную
  `QtyBase`.

Любая ошибка валидации (включая отсутствие коэффициента пересчёта единицы) отклоняет **весь**
документ ДО записи TX-строки — на этом этапе в системных листах ничего ещё не изменено.

## 4. Идемпотентность (SOURCE_KEY)

`PRIME_FindCommittedBySourceKey`/`PRIME_RegisterCommittedKey` (`PRIME_02_Store`) держат
in-memory кэш `SOURCE_KEY -> DOC_ID`, построенный один раз за сеанс из `SYS_PRIME_TX` (только
строки со `STATE = COMMITTED`; `PREPARED`/`FAILED` не попадают в кэш и не блокируют повтор).
Проверка идемпотентности — **первый** содержательный шаг `PRIME_PostDocument`, до любого
изменения `DB_PRIME_PRODUCTS`/`DB_PRIME_LOTS` (закрывает дефект 1.4.1, где `EnsureProductCon`
выполнялся до проверки повтора).

Повторный вызов с уже COMMITTED `SOURCE_KEY`: новых строк/движений не создаётся, возвращается
существующий `DOC_ID`; вызывающая форма показывает "уже проведено" вместо повторного эффекта.

**Согласованность кэша с реальным состоянием TX (2.0.1).** `PRIME_RegisterCommittedKey`
вызывается на шаге 5 — ДО `ThisComponent.store()` на шаге 6. Если `store()` всё же провалится
(диск, права доступа, файл занят), выполнение попадает в `PostFailed`, где TX корректно
откатывается на `FAILED` — но до 2.0.1 запись в кэше `SOURCE_KEY -> DOC_ID` оставалась, как
будто операция COMMITTED, и повторная попытка того же `SOURCE_KEY` получала "уже проведено"
вместо шанса на повторную попытку. Теперь `PostFailed` вызывает `PRIME_UnregisterCommittedKey`,
если регистрация успела произойти — кэш и `SYS_PRIME_TX` больше не расходятся.

| Операция | SOURCE_KEY (формируется вызывающей формой) |
|---|---|
| Приход по заказу | ORDER_ID + ORDER_LINE_ID + номер поставки |
| Выдача | идентификатор черновика выдачи |
| Возврат | OriginalDocLineId + идентификатор возврата |
| Перемещение | идентификатор документа перемещения |
| Инвентаризация (ADJUSTMENT) | INVENTORY_SESSION_ID + строка |

## 5. `SYS_PRIME_TX` — состояния

`PREPARED` записывается один раз, сразу после успешной валидации, ДО записи в
`DB_PRIME_DOCUMENTS`/`DOC_LINES`/движения. `COMMITTED` записывается ПОСЛЕ того, как все
строки документа/партий/движений уже записаны — последний шаг перед `store()`. `FAILED`
записывается из блока `PostFailed` только если `PREPARED`-строка уже существовала (т.е. если
исключение произошло до шага 3, TX-строки вообще нет).

**Известное ограничение (честно, не заявляем как решённое):** строка `PREPARED` без
последующего `COMMITTED`/`FAILED` (например, аварийное завершение LibreOffice между шагом 3 и
5) не имеет автоматического механизма докрутки при следующем открытии — она остаётся в
`SYS_PRIME_TX` как видимый диагностический след. С 2.0.1 то, что она **гарантированно не
участвует в остатке**, — это не только формулировка в документации, а реально проверяемое
поведение кода: `PRIME_LotBalance`/`PRIME_StockByLocation`/лист "Остаток" фильтруют движения по
`PRIME_IsOpIdCommitted(OP_ID)` (кэш COMMITTED `OP_ID` в `PRIME_02_Store`, перестраивается из
`SYS_PRIME_TX`). В 2.0.0 этот фильтр отсутствовал в коде — все движения суммировались независимо
от состояния транзакции, то есть "не участвует в остатке" было декларацией, а не фактом; это и
было исправлено. Ручной или автоматический разбор зависшей `PREPARED`-записи ("crash recovery")
по-прежнему не реализован. См. `KNOWN_ISSUES.md`.

## 6. Один `store()`

`ThisComponent.store()` вызывается один раз, после того как TX-строка уже переведена в
`COMMITTED`. Нет отдельного «закрытия документа/БД»: закрывает дефект 1.4.1 (`WMSDB_Close`:
commit → store → close → store → close, неявно подтверждавший недоопределённую операцию).

## 7. FIFO (ISSUE) и симметричный возврат (RETURN)

`PRIME_FifoLotsForProduct` возвращает партии товара, отсортированные по `(RECEIPT_DATE,
LOT_ID)` (сортировка вставками — число партий одного товара ожидается небольшим, не
рассчитано на тысячи партий на один товар), с текущим балансом = суммой `QTY_BASE` COMMITTED-
движений по каждой партии (`PRIME_LotBalance`).

`PostIssueLines` разбирает нужное количество по партиям в этом порядке, на каждую партию
пишет строку `DB_PRIME_ALLOCATIONS` (`DOC_LINE_ID, LOT_ID, QTY_BASE`) и отрицательное движение.
Если после разбора остаётся неудовлетворённый остаток (не должно происходить после
`PRIME_ValidateIssue`, но проверяется отдельно как защита от рассинхронизации) — весь документ
уже начатый **обрывается ошибкой 1020** уже ПОСЛЕ записи `DOC_ID`/`DOC_LINES`, т.е. в этом
единственном защитном сценарии возможен документ без движений, зафиксированный в
`SYS_PRIME_TX` как `FAILED` (не `COMMITTED` — на остаток не влияет).

`PostReturnLines` для каждой возвращаемой строки читает `DB_PRIME_ALLOCATIONS` исходной строки
выдачи (`PRIME_AllocationsForDocLine`) и распределяет возвращаемое количество по тем же
партиям в том же порядке (FIFO по allocations, не по датам заново) — так партия, из которой
физически выдали товар, получает его обратно. **Известное ограничение:** если у исходной
строки выдачи нет записей в `DB_PRIME_ALLOCATIONS` (легаси-данные, смигрированные без
allocations), возврат всё равно проводится единым "безлотовым" движением (`LOT_ID=""`) на
оставшееся количество, чтобы не заблокировать документ — это ухудшает трассируемость партии
для таких строк (см. `KNOWN_ISSUES.md`/`MIGRATION.md`).

## 8. TRANSFER и ADJUSTMENT

`TRANSFER`: на каждую строку — ровно 2 движения без `LOT_ID` (`-QTY_BASE` на
`LocationFrom`, `+QTY_BASE` на `LocationTo`), сумма остаётся неизменной по построению
(`total_stock_change: 0`), не отдельной проверкой после факта.

`ADJUSTMENT`: одно движение на строку со знаком `QtyBase`, установленным на этапе валидации;
`LOT_ID` не проставляется (инвентаризационная разница не привязана к конкретной партии).

## 9. Что НЕ реализовано в этом протоколе (честно)

- `ASSEMBLY`/`DISASSEMBLY` — есть в схеме `DB_PRIME_DOCUMENTS.DOC_TYPE`/`DATA_MODEL.md`, нет
  ветки в `PRIME_PostDocument`.
- Автоматическое восстановление зависшего `PREPARED` при следующем открытии книги — нет.
- Хэш плана документа (`SYS_PRIME_TX.HASH`) — колонка существует, но `PRIME_WriteTxRow` всегда
  пишет в неё пустую строку; проверка целостности по хэшу не реализована.
