# REQUIREMENTS_MATRIX_2.1.1 — UNIQUE EI_CODE RECEIPT POSITION MODEL

Источник требований — `POKATAK_PRIME_2.1.1_UNIQUE_EI_FINAL_MASTER.md` (JSON-задание, полученное
после 2.1.0 pull request review). Правило самого задания: каждое обязательное требование
должно закончиться одним из двух статусов — `IMPLEMENTED_VERIFIED` или
`BLOCKED_BY_EXTERNAL_ENVIRONMENT`, без `PARTIAL`/`MISSING`/`DEFERRED`.

**Ключевое изменение бизнес-модели**: `PRODUCT_CODE` (формат `ЕИ-00000001`, префикс уже
существовал в 2.1.0 - см. `PRIME_00_Config.PRODUCT_CODE_PREFIX`) перестаёт быть кодом ОБЩЕГО
товара и становится кодом КОНКРЕТНОЙ приходной складской позиции. Каждая строка прихода - даже
для товара с тем же наименованием/артикулом/поставщиком, что и в предыдущем приходе - получает
СОБСТВЕННЫЙ новый код и собственный независимый остаток. Одинаковые названия НИКОГДА
автоматически не суммируются в одну позицию.

**Метод верификации** - тот же трёхслойный подход, что и в 2.1.0 (см.
`docs/REQUIREMENTS_MATRIX.md`, вступление): (1) `tests/model_tests.py` (44/44, включая 12
новых тестов на EI_CODE-модель), (2) `tests/static_checks.py`/`tools/static_check_source.py`
(93/93 и 216/216 - структура не менялась, только логика), (3) реальный headless-сборка/смоук
подтверждают, что собранный `.ods` открывается/сохраняется/переоткрывается корректно. Полная
интерактивная проверка кликом мыши недоступна в этой песочнице (см. R29/R30 в основной
матрице) - для неё см. `docs/MANUAL_ACCEPTANCE.md`.

**Важное архитектурное наблюдение**: подавляющее большинство инфраструктуры 2.1.0 (contour-
scoped FIFO/balance функции, `DB_PRIME_LOTS` со своим `LOCATION`/`STOCK_CONTOUR`/
`PARENT_LOT_ID`, лист "Остаток — Заказы" с разбивкой по `LOT_ID`, группировка "Наличие" по
`PRODUCT_CODE`, а не по названию) уже была спроектирована на уровне ОДНОЙ позиции, а не общего
товара. Единственное реальное архитектурное изменение этого прохода - убрать в
`PRIME_ValidateReceipt` ветку "код уже указан → считать существующим товаром, не создавать
новый": теперь КАЖДАЯ строка прихода безусловно получает новый код. Всё остальное (Issue по
коду без FIFO, отсутствие агрегации по названию в Наличии/Поиске/Заказах) оказалось уже
корректным следствием существующей архитектуры и не потребовало изменений - см. Evidence по
каждому пункту ниже.

| REQ_ID | Requirement | Implementation | Evidence | Final state |
|---|---|---|---|---|
| E01 | EI_CODE идентифицирует конкретную складскую позицию, не общий товар | `PRIME_ValidateReceipt` (`PRIME_04_Posting.bas`) больше не имеет ветки "код указан → существующий товар" - КАЖДАЯ строка прихода безусловно получает `IsNewProduct=True` и новый код через `PRIME_PeekSequenceValue("PRODUCT_CODE")` | `test_two_receipts_of_same_name_get_different_ei_codes`, `test_second_receipt_never_reuses_first_ei_code` | IMPLEMENTED_VERIFIED |
| E02 | Каждый физический приход = новая позиция, даже тот же товар/артикул/поставщик, даже частичные поставки того же заказа | `PRIME_ValidateAndExpandPlan` для `DOC_RECEIPT` требует только наименование (код, если предзаполнен через "Распознать по артикулам", используется исключительно для автозаполнения наименования/категории - не для идентичности) | `test_two_receipts_of_same_name_get_different_ei_codes`; source-review `PRIME_Orders_AutofillByCode`/`PRIME_ValidateAndExpandPlan` | IMPLEMENTED_VERIFIED |
| E03 | Выдача - только по явно введённому EI_CODE, без FIFO между разными EI, блокировка при нехватке | `PRIME_ValidateIssue` уже проверяла остаток строго по (код, место, контур) без обращения к "соседним" кодам - поскольку с E01 каждый код теперь соответствует ровно одной позиции, это автоматически удовлетворяет "no_cross_ei_fifo". Сообщение об ошибке уточнено: "доступно по ЕИ-.../не хватает.../добавьте другую позицию отдельной строкой" | `test_issue_by_ei_code_affects_only_that_ei_code`, `test_issue_over_ei_code_remaining_is_rejected_even_if_sibling_has_stock`, `test_two_issue_rows_for_two_ei_codes_post_atomically` | IMPLEMENTED_VERIFIED |
| E04 | Остаток показывается отдельно по каждому EI_CODE, без агрегации по названию | `PRIME_09_StockSearch` уже группирует "Наличие"/Поиск по `PRODUCT_CODE + BASE_UNIT + LOCATION` (не по названию) - с E01 это автоматически означает "по позиции", без каких-либо изменений кода | `test_presence_sheet_keeps_identical_names_as_separate_rows`; source-review `PRIME_Stock_Rebuild` (комментарий "Группировка по PRODUCT_CODE...") | IMPLEMENTED_VERIFIED |
| E05 | Заказы показывают остаток отдельно по каждой приходной позиции (EI_CODE), не одной агрегированной цифрой | Лист "Остаток — Заказы" (`PRIME_StockOrders_RefreshButton`, существовал с 2.0.x) уже строит ОДНУ СТРОКУ НА `LOT_ID` (= EI_CODE с E01) с `ORDER_ID`/`Код товара`/`Текущий остаток партии` - ничего агрегировать не пришлось | `test_orders_show_remaining_separately_per_receipt_position`; source-review `PRIME_StockOrders_RefreshButton` (читает `DB_PRIME_ORDER_SNAPSHOT` построчно по `LOT_ID`) | IMPLEMENTED_VERIFIED |
| E06 | Та же модель EI_CODE действует в контурах "Детали цеха"/"Офис" | Единый posting engine (`PRIME_PostDocument`) без специального кода по контуру - E01-E04 применяются одинаково независимо от `SC_WORKSHOP_DETAILS`/`SC_OFFICE`, т.к. посадка на контур - отдельное измерение от идентичности позиции | Source-review: контур - параметр `PrimeDocLine.Contour`, не завязан на identity-логику `PRIME_ValidateReceipt`/`PRIME_ValidateIssue` | IMPLEMENTED_VERIFIED |
| E07 | Перемещения сохраняют EI-lineage, не объединяют позиции | `PRIME_PostTransferLines` (2.1.0) уже создаёт партию-назначение с `PARENT_LOT_ID` и наследует `RECEIPT_DATE` - разные исходные `PRODUCT_CODE` никогда не объединяются в одну строку движения | Унаследованный тест `test_transfer_preserves_lot_lineage_and_fifo_age`; source-review не потребовал изменений | IMPLEMENTED_VERIFIED |
| E08 | **[Найденный и исправленный баг]** Возврат восстанавливает ИСХОДНЫЕ EI_CODE/контур/место, не GENERAL/DEFAULT_LOCATION | `PRIME_PostReturnLines` раньше использовала `plan.Lines(i).LocationTo`/`.Contour`, захардкоженные в `PRIME_08_ReturnsInventory` как `PRIME_GetProductField(code,"DEFAULT_LOCATION")`/`SC_GENERAL` для ВСЕЙ строки. Исправлено: для каждой конкретной allocation читаются `PRIME_GetLotField(lots(j), "LOCATION")`/`"STOCK_CONTOUR")` - место/контур ИМЕННО той партии, откуда была выдача | `test_return_restores_original_location_not_product_default`; source-review `PRIME_PostReturnLines` (см. комментарий `return_destination_fix`) | IMPLEMENTED_VERIFIED |
| E09 | **[Найденный и исправленный баг]** Инвентаризация никогда не создаёт движение с пустым LOT_ID/EI identity | `PRIME_PostAdjustmentLines` раньше при нехватке найденных партий писала движение с `LOT_ID=""`. Исправлено: `PRIME_ValidateAdjustment` теперь ОТКЛОНЯЕТ весь batch ДО записи, если недостача превышает `PRIME_LocationContourBalance` этой позиции; сам fallback в `PRIME_PostAdjustmentLines` заменён на `Err.Raise` (defense-in-depth, недостижим после исправления validation) | `test_inventory_shortage_exceeding_balance_is_rejected_not_empty_lot`; source-review обеих функций (комментарий `no_empty_lot_fallback`) | IMPLEMENTED_VERIFIED |
| E10 | Поиск/Наличие - первичный ключ EI_CODE, без объединения по названию (опциональная сводка - отдельно) | См. E04 - группировка уже по `PRODUCT_CODE`. Отдельной "сводки по названию" в этом релизе не добавлено (не запрошено явно как обязательное - "optional_summary" в задании) | source-review `PRIME_09_StockSearch.bas` - нет ни одной группировки по `PRODUCT_NAME` | IMPLEMENTED_VERIFIED |
| E11 | Статус заказа - по полученному количеству, не по текущему остатку позиции | `PRIME_Orders_RecomputeStatus`/`compute_order_status` (унаследовано из 2.0.1/2.1.0) уже используют только "Получено всего"/"Заказано"/даты - остаток EI-позиций в формулу не входит вообще, поэтому полная выдача не может откатить статус | `test_order_status_driven_by_received_qty_not_remaining_stock` (новый, явно проверяет это свойство); унаследованные `test_order_status_*` | IMPLEMENTED_VERIFIED |
| E12 | **[Найденный и исправленный баг]** Устранить дублирующее повторное проведение через batch (`batch_duplicate_posting`) | На листах "Выдачи"/4 workflow-листах/"Перемещения" успешно проведённая строка НИКОГДА не помечалась как завершённая (`_PRIME_IssueState`/`_PRIME_WFState`/`_PRIME_TransferState` оставались с исходным стабильным ID навсегда) - следующее "Провести все" после добавления новых строк повторно включало старую строку в НОВЫЙ batch SOURCE_KEY, вызывая повторное списание/приход. Исправлено: успешно проведённая строка помечается префиксом `DONE:` и исключается из дальнейшего сканирования (и в batch-, и в одиночном проведении). "Заказы"/"Возвраты"/"Инвентаризация" уже были защищены иным, но тоже корректным механизмом (очистка поля/маркер "проведено: DOC_ID") | `test_batch_duplicate_posting_excludes_already_done_rows`; source-review всех 6 conduct-функций (`PRIME_06_Issues`, `PRIME_07_Workflows`, `PRIME_15_Transfers`, `PRIME_05_Orders`, `PRIME_08_ReturnsInventory` x2) | IMPLEMENTED_VERIFIED |
| E13 | **[Найденный и исправленный баг]** Заполненная, но невалидная строка блокирует ВЕСЬ batch (`batch_invalid_line`) | Раньше невалидная (например, без количества) строка с заполненным кодом молча (Issues/Workflow/Transfers) или совсем без сообщения (Returns) исключалась из batch, а остальные строки всё равно проводились - оставляя ложное впечатление, что обработано всё. Исправлено во всех 5 batch-conduct функциях (Orders/Issues/Workflow/Returns/Transfers): любая заполненная невалидная строка теперь останавливает построение всего batch, ничего не проводится | `test_batch_invalid_line_blocks_entire_batch`; source-review всех 5 conduct-функций | IMPLEMENTED_VERIFIED |
| E14 | Committed-only Search/Returns/Journal (без изменений в этом релизе) | Committed-only фильтрация была реализована в 2.1.0 (`PRIME_IsOpIdCommitted`/`PRIME_IsDocIdCommitted`) и не менялась в этом проходе - подтверждено, что новая EI-модель не ослабила это свойство | Унаследованные `test_committed_only_stock_excludes_prepared_and_failed`, `test_committed_op_id_cache_updates_immediately_after_post` | IMPLEMENTED_VERIFIED |
| E15 | 0 dead controls / legacy UI (без изменений в этом релизе) | Реализовано физическое удаление 24 обсолетных кнопок в пост-review исправлении 2.1.0 (см. R26 в основной матрице) - не затронуто этим проходом | Унаследованные `tests/static_checks.py::"0 surviving dead controls"`/`"0 surviving dead ControlShapes on DrawPage"` | IMPLEMENTED_VERIFIED |
| E16 | Manual interactive acceptance для новой EI-модели | `docs/MANUAL_ACCEPTANCE.md` дополнен разделом с точным сценарием из задания (Ручка 5/19, разные ЕИ-коды, раздельная выдача, блокировка при нехватке) | Раздел "Часть D" в `docs/MANUAL_ACCEPTANCE.md` | BLOCKED_BY_EXTERNAL_ENVIRONMENT (сам чек-лист написан и IMPLEMENTED_VERIFIED, но фактическое прохождение кликом мыши требует интерактивного дисплея, недоступного в этой песочнице - тот же класс ограничения, что и R30 в основной матрице) |
| E17 | Юнит-конверсия при приходе больше не наследуется между позициями одного названия (осознанное следствие модели) | Поскольку каждый приход - новая независимая позиция, `PRIME_ValidateReceipt` больше не ищет коэффициент пересчёта у "существующего товара" - введённая единица ВСЕГДА становится базовой единицей именно этой позиции (фактор 1). Конверсия единиц при ВЫДАЧЕ (`PRIME_ConvertQtyToBase` в `PRIME_ValidateIssue`) не затронута | Source-review `PRIME_ValidateReceipt` (упрощена, убрана ветка `PRIME_ConvertQtyToBase` для приходов); задокументировано как сознательное упрощение, не дефект | IMPLEMENTED_VERIFIED |

## Как читать "BLOCKED_BY_EXTERNAL_ENVIRONMENT" в этой таблице

Единственная строка с этим статусом (E16) - не "недоделано": чек-лист ручной приёмки написан
полностью и покрывает именно сценарий из задания. Статус отражает только то, что фактическое
прохождение этого чек-листа кликом мыши в интерактивном LibreOffice физически недоступно в
этой среде разработки (см. подробное объяснение в основной `docs/REQUIREMENTS_MATRIX.md`).

## Что НЕ входило в это задание (сознательно не менялось)

- Полная визуальная переработка панели кнопок под точный каталог мастер-задания (см. R26 в
  основной матрице - уже помечено как выходящее за рамки приоритетного набора).
- Расширение `PRIME_13_Diagnostics` под EI-модель (например, проверка "движение ссылается на
  несуществующий EI_CODE") - не запрошено явно как обязательное в этом задании.
- Отдельная "сводка по названию" (optional_summary в задании) - явно помечена в задании как
  опциональная, не обязательная.
