```json
{
  "task": {
    "name": "POKATAK PRIME 2.0.0",
    "type": "full_project_rebuild",
    "role": [
      "senior software architect",
      "LibreOffice Calc engineer",
      "LibreOffice Basic engineer",
      "UNO engineer",
      "WMS architect",
      "reliability engineer",
      "test engineer"
    ],
    "input_files": {
      "current_project": "POKATAK_PRIME_LINUX_1.4.1_CLEAN(1).zip",
      "audit": "POKATAK_AUDIT_2026-09-15(1).md"
    },
    "required_output": "ПОКАТАК_PRIME_2.0.0.ods",
    "objective": "Создать финальную стабильную версию ПОКАТАК на основе существующего проекта. Сохранить привычный пользовательский интерфейс и полезный функционал, но заменить внутреннее ядро на простую, устойчивую Calc-only архитектуру без обязательной LibreOffice Base/Firebird зависимости."
  },

  "priority": [
    "не допускать зависаний и аварийных закрытий LibreOffice",
    "не допускать двойного проведения",
    "не допускать частично проведённых документов",
    "правильно считать остатки",
    "гарантировать сохранность проведённых данных",
    "обеспечить высокую скорость работы",
    "автоматизировать ручные действия",
    "добавлять дополнительные функции только после стабилизации ядра"
  ],

  "core_principle": {
    "rule": "Если приходится выбирать между сложной автоматизацией и более простой надёжной реализацией, выбирать надёжность.",
    "do_not_patch_old_architecture_blindly": true,
    "preserve_ui_not_internal_bugs": true,
    "single_source_of_truth": true,
    "single_posting_engine": true
  },

  "phase_0_project_analysis": {
    "required": true,
    "actions": [
      "Распаковать ZIP проекта.",
      "Распаковать ODS как ODF ZIP.",
      "Извлечь все LibreOffice Basic модули.",
      "Полностью прочитать аудит.",
      "Проанализировать content.xml.",
      "Проанализировать события листов.",
      "Проанализировать кнопки и ScriptCode.",
      "Построить карту листов.",
      "Построить карту button -> macro.",
      "Построить карту event -> macro.",
      "Найти все реальные пути изменения складских данных.",
      "Найти старые и новые конкурирующие реализации.",
      "Найти все зависимости от Base, Firebird, SDBC и ODB.",
      "Найти все commit, rollback, save, store и EnsureSchema.",
      "Найти тяжёлые Contents changed handlers.",
      "Изучить заказы, приходы, выдачи, возвраты, цех, офис, партии, FIFO, поиск, остатки, акты, импорт, отчёты, диагностику и комплекты."
    ]
  },

  "architecture": {
    "runtime_database": "Calc-only",
    "main_file": "ПОКАТАК_PRIME_2.0.0.ods",
    "base_firebird": {
      "required_in_runtime": false,
      "odb_required": false,
      "sdbc_required": false,
      "sql_required_for_daily_work": false,
      "allowed_usage": [
        "legacy archive",
        "optional one-time read-only import"
      ]
    },
    "runtime_rule": "Если WMS_DATA_PORTABLE.odb отсутствует, PRIME должен полностью работать."
  },

  "system_sheets": [
    "SYS_PRIME_META",
    "SYS_PRIME_SEQ",
    "SYS_PRIME_TX",
    "DB_PRIME_PRODUCTS",
    "DB_PRIME_ALIASES",
    "DB_PRIME_PRODUCT_UNITS",
    "DB_PRIME_DOCUMENTS",
    "DB_PRIME_DOC_LINES",
    "DB_PRIME_MOVEMENTS",
    "DB_PRIME_LOTS",
    "DB_PRIME_ALLOCATIONS",
    "DB_PRIME_RETURNS",
    "DB_PRIME_ORDER_SNAPSHOT",
    "DB_PRIME_KITS",
    "DB_PRIME_KIT_LINES",
    "DB_PRIME_ACTS",
    "DB_PRIME_AUDIT"
  ],

  "system_sheets_rules": {
    "hidden": true,
    "manual_user_editing": false,
    "batch_io": true,
    "source_of_truth": "committed movements",
    "stock_sheet_is_view_only": true,
    "do_not_store_editable_stock_balance": true
  },

  "transaction_model": {
    "fields": [
      "OP_ID",
      "SOURCE_KEY",
      "DOC_ID",
      "STATE",
      "STARTED_AT",
      "COMMITTED_AT",
      "HASH",
      "ERROR"
    ],
    "states": [
      "PREPARED",
      "COMMITTED",
      "FAILED",
      "CANCELLED"
    ],
    "posting_algorithm": [
      "Получить PRIME runtime lock.",
      "Включить event guard.",
      "Считать весь документ пользователя в память.",
      "Проверить все строки до записи.",
      "Создать полный план документа в памяти.",
      "Создать header.",
      "Создать lines.",
      "Определить products.",
      "Рассчитать lots.",
      "Рассчитать FIFO allocations.",
      "Сформировать movements.",
      "Сформировать returns и links.",
      "Сформировать snapshots.",
      "Проверить остаток по всему документу.",
      "Проверить единицы и конвертации.",
      "Проверить SOURCE_KEY на повтор.",
      "Записать PREPARED transaction.",
      "Пакетно записать системные таблицы.",
      "Последним логическим шагом выставить COMMITTED.",
      "Обновить интерфейс.",
      "Вызвать ThisComponent.store() один раз.",
      "Только после успешного store очищать разрешённые поля.",
      "Всегда снимать event guard и runtime lock через cleanup."
    ],
    "stock_considers_states": [
      "COMMITTED"
    ],
    "prepared_affects_stock": false
  },

  "crash_recovery": {
    "on_open": [
      "Найти PREPARED без COMMITTED.",
      "Не учитывать их в остатке.",
      "Показывать их в диагностике.",
      "Не допроводить автоматически.",
      "Предоставить безопасное удаление или отмену незавершённого staging."
    ],
    "automatic_legacy_repair": false
  },

  "idempotency": {
    "required": true,
    "rule": "Повторный клик не создаёт повторного движения.",
    "source_key_examples": {
      "order_receipt": "ORDER_ID + ORDER_LINE_ID + DELIVERY_ID",
      "issue": "ISSUE_DRAFT_ID",
      "return": "ORIGINAL_ISSUE_LINE_ID + RETURN_ID",
      "inventory": "INVENTORY_SESSION_ID + LINE_ID",
      "transfer": "TRANSFER_DOC_ID"
    },
    "when_existing_committed_source_key_found": [
      "Не писать новые движения.",
      "Вернуть существующий DOC_ID.",
      "Показать статус Уже проведено."
    ]
  },

  "events": {
    "content_changed_must_be_lightweight": true,
    "handler": "PRIME_OnContentChanged",
    "forbidden_actions": [
      "проведение склада",
      "полное сканирование истории",
      "пересчёт всего листа",
      "перестройка всех заказов",
      "обновление дашборда",
      "обновление всего остатка",
      "сохранение ODS",
      "открытие Base",
      "SQL",
      "создание партий",
      "миграции",
      "перестройка кнопок",
      "массовое форматирование"
    ],
    "allowed_actions": [
      "определить изменившуюся строку",
      "проверить колонку внутреннего кода",
      "найти товар в memory index",
      "подставить несколько пустых полей текущей строки"
    ]
  },

  "runtime": {
    "required_functions": [
      "PRIME_TryEnter",
      "PRIME_Leave",
      "PRIME_EventEnter",
      "PRIME_EventLeave",
      "PRIME_ResetRuntimeLock"
    ],
    "event_guard_type": "nested_counter",
    "boolean_busy_flag_only": false,
    "double_click_protection": true,
    "cleanup_required": true,
    "lock_controllers": {
      "allowed": true,
      "only_for_short_batch_ui_operations": true,
      "always_unlock_in_cleanup": true
    }
  },

  "performance": {
    "preferred_range_api": [
      "getDataArray",
      "setDataArray"
    ],
    "forbidden_patterns": [
      "читать 10000 строк по одной ячейке через UNO",
      "писать тысячи строк по одной ячейке",
      "ReDim Preserve на каждой итерации",
      "полный UsedArea scan на каждый ввод",
      "перестраивать стили при проведении",
      "запускать migration при обычной операции",
      "сохранять ODS внутри вложенных helper-функций"
    ],
    "caches": [
      "header -> column index",
      "product_code -> product record",
      "aliases",
      "units",
      "committed source keys"
    ],
    "cache_invalidation_required": true
  },

  "product_identity": {
    "primary_key": "internal_product_code",
    "new_code_format": "ЕИ-00000001",
    "prefix": "ЕИ-",
    "prefix_alphabet": "Cyrillic",
    "sequence_storage": "SYS_PRIME_SEQ",
    "first_code_if_empty": "ЕИ-00000001",
    "migration_rule": "Следующий номер = max существующего числового суффикса ЕИ-кодов + 1.",
    "rename_existing_codes_automatically": false,
    "name_is_unique_key": false,
    "supplier_article_is_global_unique_key": false
  },

  "product_aliases": {
    "matching_context": [
      "От кого / площадка",
      "Продавец",
      "Код поставщика",
      "Артикул поставщика"
    ],
    "result": "Код товара",
    "ambiguous_match_behavior": "Не угадывать. Требовать выбор пользователя.",
    "fuzzy_search": "suggestion_only"
  },

  "product_table": {
    "sheet": "DB_PRIME_PRODUCTS",
    "fields": [
      "PRODUCT_CODE",
      "PRODUCT_NAME",
      "BASE_UNIT",
      "DEFAULT_LOCATION",
      "CATEGORY",
      "SUBCATEGORY",
      "RETURNABLE",
      "ACCOUNT_TYPE",
      "ACTIVE",
      "CREATED_AT",
      "UPDATED_AT"
    ]
  },

  "product_units": {
    "sheet": "DB_PRIME_PRODUCT_UNITS",
    "fields": [
      "PRODUCT_CODE",
      "UNIT_NAME",
      "FACTOR_TO_BASE",
      "ACTIVE"
    ],
    "supported_examples": [
      "шт",
      "упак",
      "кг",
      "м"
    ],
    "product_specific_conversion": true,
    "example": {
      "unit": "упак",
      "base_unit": "шт",
      "factor": 1000
    },
    "missing_conversion_behavior": "Block entire posting.",
    "quantity_precision_digits": 6
  },

  "live_code_lookup": {
    "required": true,
    "database_connection_allowed": false,
    "lookup_source": "memory index from DB_PRIME_PRODUCTS",
    "do_not_overwrite_non_empty_user_fields": true,

    "Выдачи": {
      "trigger": "ввод внутреннего кода",
      "autofill": [
        "Наименование",
        "Ед. изм.",
        "Возвратный",
        "Дата"
      ],
      "location_rule": "Подставить Откуда только если положительный остаток находится ровно в одном месте.",
      "multiple_locations": "Оставить Откуда пустым и показать подсказку."
    },

    "Заказы": {
      "trigger": "ввод существующего внутреннего кода",
      "autofill": [
        "Полное наименование товара",
        "Ед. изм.",
        "Категория",
        "Подкатегория",
        "Место хранения"
      ]
    },

    "Приход — Цех": {
      "autofill": [
        "Наименование",
        "Ед. изм.",
        "Категория",
        "Подкатегория",
        "Место хранения"
      ]
    },

    "Расход — Цех": {
      "autofill": [
        "Наименование",
        "Ед. изм.",
        "Категория",
        "Подкатегория",
        "Место хранения"
      ]
    },

    "Приход — Офис": {
      "autofill": [
        "Наименование",
        "Ед. изм.",
        "Категория",
        "Подкатегория",
        "Место хранения"
      ]
    },

    "Расход — Офис": {
      "autofill": [
        "Наименование",
        "Ед. изм.",
        "Категория",
        "Подкатегория",
        "Место хранения"
      ]
    },

    "buttons": [
      "Подтянуть по коду",
      "Подтянуть все коды"
    ]
  },

  "orders_sheet": {
    "sheet": "Заказы",
    "delete_existing_user_columns": false,

    "columns": [
      "Номер",
      "Полное наименование товара",
      "Код товара",
      "Код поставщика",
      "Артикул поставщика",
      "От кого / площадка",
      "Продавец",
      "Источник прихода",
      "№ документа",
      "Номер счета",
      "Дата заказа",
      "Дата документа",
      "Дата поступления",
      "Количество",
      "Факт. количество",
      "Ед. изм.",
      "Цена",
      "Сумма",
      "Покупатель",
      "Категория",
      "Подкатегория",
      "Место хранения",
      "Статус",
      "Контроль",
      "Комментарий",
      "Ожидаемая дата",
      "Назначение / проект",
      "Получено всего",
      "Осталось получить"
    ],

    "column_rule": "От кого / площадка и Продавец должны находиться рядом.",

    "allowed_hidden_helpers": [
      "_PRIME_OrderID",
      "_PRIME_LineID",
      "_PRIME_State"
    ],

    "avoid_many_hidden_columns": true
  },

  "order_identity": {
    "order_key": "ORDER_ID",
    "line_key": "ORDER_LINE_ID",
    "visual_line_number_is_global_key": false
  },

  "new_order": {
    "actions": [
      "Создать новую строку.",
      "Создать ORDER_ID.",
      "Создать ORDER_LINE_ID.",
      "Установить визуальный номер позиции 1.",
      "Переместить курсор в поле товара."
    ],
    "must_not": [
      "изменять склад",
      "сканировать всю историю",
      "перестраивать все заказы",
      "открывать Base",
      "строить dashboard",
      "делать тяжёлый save"
    ]
  },

  "add_order_line": {
    "same_order_id": true,
    "new_order_line_id": true,
    "copy_common_order_fields": true,
    "full_history_rebuild": false
  },

  "partial_receipts": {
    "ordered_quantity_field": "Количество",
    "current_delivery_quantity_field": "Факт. количество",
    "each_delivery_creates": [
      "separate receipt document",
      "separate lot",
      "historical snapshot"
    ],
    "after_success": [
      "Очистить Факт. количество.",
      "Обновить Получено всего.",
      "Обновить Осталось получить."
    ],
    "example": {
      "ordered": 10,
      "delivery_1": 6,
      "after_delivery_1": {
        "received_total": 6,
        "remaining": 4
      },
      "delivery_2": 4,
      "after_delivery_2": {
        "received_total": 10,
        "remaining": 0,
        "receipt_documents": 2,
        "lots": 2
      }
    }
  },

  "multiple_supplier_documents": {
    "multiple_upd_per_invoice": true,
    "each_receipt_snapshot_preserves": [
      "document number",
      "document date",
      "receipt date",
      "supplier",
      "seller",
      "invoice",
      "quantity",
      "ORDER_ID"
    ]
  },

  "posting_engine": {
    "function": "PRIME_PostDocument",
    "single_engine_required": true,
    "document_types": [
      "RECEIPT",
      "ISSUE",
      "RETURN",
      "TRANSFER",
      "ADJUSTMENT",
      "ASSEMBLY",
      "DISASSEMBLY"
    ],
    "ui_forms_only_build_document_model": true,
    "ui_forms_must_not_write_movements_directly": true
  },

  "lots": {
    "required": true,
    "fields": [
      "LOT_ID",
      "PRODUCT_CODE",
      "RECEIPT_DOC_ID",
      "RECEIPT_LINE_ID",
      "RECEIPT_DATE",
      "LOCATION",
      "ORIGINAL_QTY_BASE",
      "BASE_UNIT",
      "ORIGIN",
      "ORDER_ID"
    ],
    "current_lot_balance": "Calculated from COMMITTED movements."
  },

  "fifo": {
    "automatic": true,
    "manual_lot_required": false,
    "one_issue_line_can_use_multiple_lots": true,
    "sort_order": [
      "receipt date",
      "LOT_ID"
    ],
    "shortage_behavior": "Reject entire document.",
    "example": {
      "lot_1": 10,
      "lot_2": 20,
      "issue": 15,
      "allocation": [
        10,
        5
      ]
    }
  },

  "issues_sheet": {
    "sheet": "Выдачи",
    "preserve_fields": [
      "№",
      "Код",
      "Наименование",
      "Кол-во",
      "Ед. изм.",
      "Кто получил",
      "Дата",
      "Откуда",
      "Возвратный",
      "Возвращено",
      "Дата возврата",
      "Примечание"
    ],
    "add_fields": [
      "Назначение / проект"
    ],
    "buttons": [
      "Новая выдача",
      "Провести выбранное",
      "Провести заполненные",
      "Подтянуть по коду",
      "Подтянуть все коды",
      "Добавить комплект",
      "Возвраты",
      "Остаток",
      "Поиск"
    ],
    "repeat_posting_behavior": "Do not create duplicate issue."
  },

  "workflows": {
    "active_ui_sheets": [
      "Приход — Цех",
      "Расход — Цех",
      "Приход — Офис",
      "Расход — Офис"
    ],
    "same_posting_engine": true,
    "internal_code_required": true,
    "issue_destination_must_be_saved": true,

    "legacy_parallel_forms": [
      "Приход — Производство",
      "Расход — Производство",
      "Приход — Детали",
      "Расход — Детали"
    ],

    "legacy_behavior": [
      "hide",
      "read-only",
      "redirect to new workflows"
    ],

    "legacy_parallel_posting_allowed": false
  },

  "returns": {
    "single_return_mechanism": true,
    "linked_to_original_issue_line": true,
    "maximum_return_formula": "issued_qty - already_returned_qty",
    "duplicate_safe": true,
    "creates": [
      "RETURN document",
      "IN movement",
      "link to original ISSUE",
      "traceable return lot or equivalent link"
    ],
    "example": {
      "issued": 5,
      "returned": 2,
      "remaining_returnable": 3,
      "repeat_same_return": "no additional stock increase"
    },
    "disable_conflicting_legacy_return_buttons": true,
    "automatic_legacy_return_repair": false
  },

  "transfers": {
    "single_document": true,
    "movements": [
      "OUT from location A",
      "IN to location B"
    ],
    "same_quantity": true,
    "total_stock_change": 0,
    "atomic_logical_operation": true
  },

  "inventory": {
    "workflow": [
      "Create inventory session.",
      "Capture snapshot timestamp.",
      "Read expected stock.",
      "User enters physical count.",
      "Before posting check movements after snapshot.",
      "On conflict require refresh or explicit handling.",
      "Create adjustment only for difference.",
      "Never rewrite historical movements."
    ]
  },

  "stock_view": {
    "sheet": "Остаток",
    "data_source": "COMMITTED movements",
    "group_by": [
      "PRODUCT_CODE",
      "BASE_UNIT",
      "LOCATION"
    ],
    "optional_detail": "LOT_ID",
    "columns": [
      "Код",
      "Наименование",
      "Артикул",
      "Ед. изм.",
      "Место хранения",
      "Остаток",
      "Последняя операция"
    ],
    "filters": [
      "Источник",
      "Категория",
      "Подкатегория",
      "Место",
      "Только ненулевые",
      "Только отрицательные",
      "Код",
      "Название"
    ],
    "silent_2000_row_limit": false,
    "batch_output": true
  },

  "order_stock_view": {
    "sheet": "Остаток — Заказы",
    "source": "saved receipt/order snapshots",
    "columns": [
      "ORDER_ID",
      "Документ",
      "Дата",
      "Код товара",
      "Наименование",
      "Количество поставки",
      "Ед. изм.",
      "LOT_ID",
      "Текущий остаток партии"
    ],
    "include_zero_lot_balance": true,
    "do_not_repeat_global_product_stock_as_lot_stock": true
  },

  "search": {
    "sheet": "База - Поиск",
    "uses_base": false,
    "search_fields": [
      "Внутренний код",
      "Наименование",
      "Артикул",
      "Поставщик",
      "Продавец",
      "№ документа",
      "Счёт",
      "ORDER_ID",
      "DOC_ID",
      "LOT_ID",
      "Получатель",
      "Назначение"
    ],
    "batch_search": true
  },

  "acts": {
    "writer_odt_supported": true,
    "source": [
      "DB_PRIME_DOCUMENTS",
      "DB_PRIME_DOC_LINES"
    ],
    "source_is_current_selection": false,
    "internal_code_required_in_act": true,
    "can_regenerate_by_doc_id": true,
    "writer_failure_must_not_rollback_stock_document": true,
    "false_success_message_allowed": false,
    "primary_format": "ODT",
    "pdf": "separate command"
  },

  "kits": {
    "required_mode": "issue_bundle",
    "sheet": "Комплекты",
    "fields": [
      "KIT_ID",
      "Название",
      "Версия",
      "PRODUCT_CODE",
      "Количество на 1 комплект",
      "Единица",
      "Активен"
    ],
    "issue_button": "Добавить комплект",
    "workflow": [
      "Выбрать комплект.",
      "Ввести количество комплектов.",
      "Рассчитать количество каждого компонента.",
      "Проверить наличие всех компонентов.",
      "Добавить обычные строки расхода.",
      "Сохранить KIT_ID для трассировки."
    ],
    "virtual_kit_stock_created": false,
    "partial_issue_if_component_missing": false,
    "example": {
      "kit_quantity": 3,
      "components_per_kit": {
        "ЕИ-00000010": 2,
        "ЕИ-00000011": 4
      },
      "result_issue": {
        "ЕИ-00000010": 6,
        "ЕИ-00000011": 12
      }
    },
    "physical_assembly": {
      "architecture_ready": true,
      "mandatory_for_first_stable_release": false
    }
  },

  "import_orders": {
    "preserve_feature": true,
    "preview_required": true,
    "header_detection": true,
    "date_normalization": true,
    "merged_cell_propagation_within_order_only": true,
    "silent_2500_limit": false,
    "target_supported_rows": 20000,
    "report_errors_and_duplicates": true,
    "identical_product_row_is_not_document_duplicate": true,
    "duplicate_document_detection": "document key or hash",
    "batch_write": true,
    "automatic_stock_posting_after_import": false
  },

  "dates": {
    "accepted_formats": [
      "24.08",
      "24/08",
      "24-08",
      "24.08.2026",
      "24/08/2026",
      "24-08-2026"
    ],
    "missing_year": "use current year",
    "full_sheet_rebuild_on_date_normalization": false
  },

  "ui": {
    "preserve_general_style": true,
    "merged_cells_in_work_tables": false,
    "freeze_headers": true,
    "autofilter": true,
    "visible_borders": true,
    "clear_errors": true,
    "work_actions_via_sheet_buttons": true,
    "user_should_not_open_macro_dialog_for_daily_work": true,
    "technical_sheets_hidden": true,
    "technical_columns_minimal": true,
    "interface_rebuild_command": "Восстановить интерфейс PRIME",
    "rebuild_interface_during_normal_posting": false
  },

  "error_handling": {
    "on_error_resume_next_allowed_only_for": [
      "cleanup",
      "optional UNO property detection",
      "object close"
    ],
    "forbidden_for": [
      "document writes",
      "movement writes",
      "quantities",
      "transaction state",
      "store",
      "stock calculations",
      "act creation"
    ]
  },

  "logging": {
    "path": "Diagnostics/prime_runtime.log",
    "must_not_break_operation_if_unavailable": true,
    "commercial_document_content": false,
    "fields": [
      "timestamp",
      "OP_ID",
      "stage",
      "duration",
      "sheet",
      "row",
      "error number",
      "error text"
    ],
    "stages": [
      "BUTTON_ENTER",
      "VALIDATION_START",
      "VALIDATION_OK",
      "PLAN_READY",
      "TX_PREPARED",
      "DOCS_WRITTEN",
      "LINES_WRITTEN",
      "LOTS_WRITTEN",
      "MOVEMENTS_WRITTEN",
      "TX_COMMITTED",
      "STORE_START",
      "STORE_OK",
      "UI_FINALIZED",
      "ERROR"
    ]
  },

  "diagnostics": {
    "read_only": true,
    "store_allowed": false,
    "modify_data_allowed": false,
    "automatic_repairs_allowed": false,
    "checks": [
      "duplicate PRODUCT_CODE",
      "duplicate DOC_ID",
      "duplicate MOVE_ID",
      "duplicate committed SOURCE_KEY",
      "movement without product",
      "lot without product",
      "allocation without issue",
      "allocation without lot",
      "negative stock",
      "return greater than issue",
      "orphan document line",
      "committed document without required movement",
      "PREPARED transactions",
      "unit mismatches",
      "sequence validity",
      "kit integrity"
    ],
    "aggregate_quantities_separately_by_unit": true
  },

  "backup": {
    "before_migration": true,
    "before_update": true,
    "manual_button": "Создать резервную копию",
    "optional_first_successful_launch_of_day": true,
    "after_every_line": false,
    "directory": "Backups",
    "filename_pattern": "POKATAK_backup_YYYYMMDD_HHMMSS.ods"
  },

  "migration": {
    "source_version": "1.4.1",
    "target_version": "2.0.0",
    "disable_legacy_events_first": true,
    "create_prime_tables": true,
    "orders_mapping_by_headers": true,
    "preserve_all_existing_25_order_columns": true,
    "reorder_columns": true,
    "add_new_columns": true,
    "auto_post_existing_rows": false,
    "preserve_legacy_data_snapshot_if_needed": true,
    "do_not_recreate_dozen_legacy_helper_columns": true
  },

  "legacy_odb_import": {
    "optional": true,
    "read_only": true,
    "preview_before_import": true,
    "write_back_to_odb": false,
    "prime_runtime_dependency_after_import": false,
    "if_driver_unavailable": "Continue PRIME without legacy import."
  },

  "basic_modules": [
    "PRIME_00_Config",
    "PRIME_01_Runtime",
    "PRIME_02_Store",
    "PRIME_03_Catalog",
    "PRIME_04_Posting",
    "PRIME_05_Orders",
    "PRIME_06_Issues",
    "PRIME_07_Workflows",
    "PRIME_08_ReturnsInventory",
    "PRIME_09_StockSearch",
    "PRIME_10_ActsReports",
    "PRIME_11_Kits",
    "PRIME_12_UI",
    "PRIME_13_Diagnostics",
    "PRIME_14_MigrationInstaller"
  ],

  "basic_rules": {
    "option_explicit": true,
    "circular_dependency_chaos": false,
    "very_large_monolithic_functions": false,
    "legacy_modules_can_remain_as_archived_source": true,
    "legacy_modules_bound_to_runtime_ui": false
  },

  "data_access_layer": {
    "required": true,
    "only_store_layer_directly_accesses_hidden_db_sheets": true,
    "helpers": [
      "HeaderMap",
      "ReadTable",
      "FindLastRow",
      "AppendRowsBatch",
      "UpdateRowsBatch",
      "FindByKey",
      "SequenceNext"
    ]
  },

  "historical_corrections": {
    "delete_posted_movements_as_normal_workflow": false,
    "method": [
      "reversal document",
      "adjustment document"
    ],
    "require_original_doc_reference": true,
    "require_reason": true,
    "full_reset": {
      "admin_only": true,
      "backup_required": true,
      "explicit_confirmation_phrase_required": true
    }
  },

  "dashboard": {
    "auto_refresh_on_every_cell_change": false,
    "explicit_refresh": true,
    "show_snapshot_timestamp": true,
    "failed_new_snapshot_must_not_replace_previous_successful_snapshot": true
  },

  "manager_report": {
    "preserve_feature": true,
    "metrics": [
      "documents",
      "document lines",
      "unique products",
      "receipts",
      "issues",
      "returns",
      "recipients",
      "destinations",
      "open orders",
      "overdue orders",
      "discrepancies",
      "inventory adjustments",
      "kits",
      "problems"
    ],
    "rows_are_not_documents": true,
    "saved_report_is_snapshot": true
  },

  "static_tests": {
    "required": true,
    "checks": [
      "ODS is valid ZIP",
      "XML parses successfully",
      "PRIME modules embedded",
      "script library manifest valid",
      "sheet events reference existing PRIME macros",
      "buttons reference existing PRIME macros",
      "working buttons do not reference WMSDB modules",
      "normal runtime has no required ODB dependency",
      "all original order columns remain",
      "new PRIME order columns exist",
      "ЕИ code generator exists",
      "hidden PRIME tables exist",
      "legacy heavy events disabled"
    ]
  },

  "model_tests": {
    "required": true,
    "cases": [
      "partial receipt 10 -> 6 -> 4",
      "FIFO across multiple lots",
      "shortage blocks entire document",
      "return after issue",
      "duplicate click",
      "transfer conserves total stock",
      "inventory adjustment",
      "product code uniqueness",
      "ambiguous supplier alias",
      "unit conversion",
      "kit availability"
    ]
  },

  "libreoffice_integration_tests": {
    "required_when_environment_allows": true,
    "use_copy_and_separate_profile": true,
    "steps": [
      "Open final ODS.",
      "Run installation or migration if possible.",
      "Save.",
      "Close.",
      "Reopen.",
      "Verify structure and persistence."
    ],
    "if_basic_macro_execution_not_possible": [
      "Perform maximum available ODS open/save smoke test.",
      "Create TEST_CHECKLIST.md.",
      "Do not claim Basic runtime was dynamically verified."
    ]
  },

  "acceptance_tests": [
    {
      "name": "new_order",
      "expected": [
        "1 row created",
        "fast operation",
        "stock unchanged"
      ]
    },
    {
      "name": "lookup_by_internal_code",
      "expected": [
        "enter ЕИ code in Выдачи",
        "name populated",
        "unit populated",
        "no Base connection"
      ]
    },
    {
      "name": "partial_receipt",
      "input": "ordered 10, receive 6 then 4",
      "expected": [
        "2 receipt docs",
        "2 lots",
        "received total 10",
        "remaining 0"
      ]
    },
    {
      "name": "126_line_receipt",
      "expected": [
        "all 126 lines processed",
        "quantities correct",
        "no per-cell DB dependency"
      ]
    },
    {
      "name": "1000_lines",
      "expected": "batch processing works"
    },
    {
      "name": "10000_history_rows",
      "expected": "ordinary cell input does not scan all history"
    },
    {
      "name": "fifo",
      "input": "lot 10 + lot 20, issue 15",
      "expected": "allocation 10 + 5"
    },
    {
      "name": "shortage",
      "input": "available 10, issue 11",
      "expected": "zero new movements"
    },
    {
      "name": "return",
      "input": "issue 5, return 2",
      "expected": [
        "stock increases exactly 2",
        "repeating same return adds nothing"
      ]
    },
    {
      "name": "transfer",
      "expected": "total product stock unchanged"
    },
    {
      "name": "double_click",
      "expected": "movement count unchanged after second click"
    },
    {
      "name": "close_reopen",
      "expected": [
        "committed documents remain",
        "stock remains",
        "pending does not affect stock"
      ]
    },
    {
      "name": "act_after_form_clear",
      "expected": [
        "act generated by DOC_ID",
        "internal code present"
      ]
    },
    {
      "name": "kit",
      "input": "3 kits containing 2 A + 4 B each",
      "expected": [
        "issue A = 6",
        "issue B = 12",
        "shortage of any component cancels complete posting"
      ]
    }
  ],

  "fault_injection": {
    "development_only": true,
    "production_default": false,
    "points": [
      "BEFORE_PREPARED",
      "AFTER_PREPARED",
      "AFTER_DOCS",
      "AFTER_LINES",
      "AFTER_LOTS",
      "AFTER_MOVEMENTS",
      "BEFORE_COMMIT_MARKER",
      "AFTER_COMMIT_MARKER",
      "BEFORE_STORE",
      "AFTER_STORE"
    ],
    "required_assertions": [
      "failure before commit marker does not affect stock",
      "unfinished operations are visible after reopen",
      "retry does not duplicate committed movement"
    ]
  },

  "forbidden": [
    "two active posting mechanisms for the same operation",
    "stock posting from Contents changed",
    "schema migration during normal posting",
    "save from nested helper functions",
    "multiple unnecessary store operations per document",
    "clearing user input before successful store",
    "automatic product merge by equal names",
    "supplier article as global product identity",
    "fuzzy automatic merge",
    "creating a new product code when existing internal code is supplied",
    "hiding critical write failures with On Error Resume Next",
    "full workbook rebuild after simple cell edit",
    "claiming runtime validation based only on source-text tests",
    "removing existing user columns from Заказы",
    "replacing user data with empty templates",
    "requiring additional third-party software"
  ],

  "source_tree": {
    "root": "prime_project",
    "directories": [
      "src/basic",
      "src/templates",
      "tools",
      "tests",
      "docs",
      "build",
      "dist"
    ],
    "store_new_basic_as_bas_files": true,
    "preserve_legacy_sources_for_reference": true
  },

  "builder": {
    "reproducible": true,
    "actions": [
      "Load template ODS.",
      "Inject PRIME Basic modules.",
      "Update script library manifests.",
      "Update events.",
      "Update buttons.",
      "Validate XML.",
      "Validate ZIP.",
      "Run static tests.",
      "Create final ODS."
    ]
  },

  "dist": {
    "required_files": [
      "ПОКАТАК_PRIME_2.0.0.ods",
      "НАЧАТЬ_PRIME.txt",
      "CHANGELOG_PRIME.md",
      "MIGRATION_1.4.1_TO_2.0.md",
      "TEST_REPORT.md"
    ],
    "runtime_directories_created_automatically": [
      "Backups",
      "Acts",
      "Exports",
      "Diagnostics"
    ],
    "odb_required_for_daily_user": false
  },

  "start_prime_manual": {
    "filename": "НАЧАТЬ_PRIME.txt",
    "must_explain": [
      "how to open PRIME",
      "how to enable macros",
      "how to create first backup",
      "how to run Проверить PRIME",
      "how to create an order",
      "how to post receipt",
      "how to issue by internal code",
      "how to return an item",
      "how to update stock",
      "how to create an act",
      "where backups are stored",
      "what to send when an error occurs"
    ],
    "error_report_items": [
      "OP_ID",
      "Diagnostics/prime_runtime.log"
    ]
  },

  "definition_of_done": {
    "code_only_is_not_enough": true,
    "required_process": [
      "Build final ODS.",
      "Validate project structure.",
      "Validate events.",
      "Validate buttons.",
      "Run available tests.",
      "Fix discovered problems.",
      "Rebuild.",
      "Repeat validation."
    ],
    "main_deliverable": "dist/ПОКАТАК_PRIME_2.0.0.ods"
  },

  "final_report": {
    "required": true,
    "include": [
      "exact final ODS path",
      "final file size",
      "list of PRIME modules",
      "disabled legacy modules",
      "disabled legacy events",
      "active PRIME events",
      "data storage architecture",
      "ЕИ code generation logic",
      "transaction protocol",
      "FIFO implementation",
      "code lookup implementation",
      "return implementation",
      "kit implementation",
      "tests actually executed",
      "tests not executable in current environment",
      "known limitations"
    ],
    "false_claims_of_full_verification_forbidden": true
  },

  "conflict_resolution_order": [
    "data integrity",
    "audit findings",
    "this PRIME specification",
    "preservation of familiar UI",
    "legacy internal behavior"
  ],

  "final_instruction": "Не пытайся просто исправлять старый Firebird-код. Цель — построить простое однопользовательское офлайн WMS-ядро внутри Calc: один ODS, один каталог товаров, один внутренний код, один posting engine, один журнал движений, один FIFO, один механизм возврата и один источник истины. Сначала доведи до устойчивости товары, документы, партии, движения, FIFO, возвраты, остатки, сохранение и идемпотентность. Затем подключай акты, импорт, отчёты, dashboard и комплекты. Итогом должен быть реально собранный ПОКАТАК_PRIME_2.0.0.ods."
}
```