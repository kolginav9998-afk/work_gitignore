#!/usr/bin/env python3
"""
PRIME 2.0.1 builder.

Берёт шаблон ODS 1.4.1 (src/templates/POKATAK_WMS_1.4.1_template.ods), внедряет 15 модулей
PRIME_*.bas в библиотеку Basic "Standard", затем УДАЛЯЕТ все 38 legacy WMS_*-модулей из
собранного .ods (legacy_removal, PRIME 2.0.1: legacy_code_allowed_in_production_ods=false -
архивная копия исходников хранится в git под legacy_reference/, а не в самом файле; в 2.0.0
эти модули ошибочно оставались встроенными как "архив", только отвязанные от кнопок/событий),
создаёт системные/бизнес-листы, привязывает лёгкие
обработчики PRIME_OnContentChanged_* к событию листа "OnChange" (реальный ключ события в UNO,
подтверждено через headless-исследование шаблона - НЕ "OnContentChanged"), перепривязывает
кнопки на новые PRIME-макросы по явной карте (builder.actions: Update events, Update buttons),
запускает миграцию структуры листа "Заказы" и сохраняет результат.

Требует запущенный headless LibreOffice с UNO-сокетом (см. start_soffice()/main()).
"""
import argparse
import shutil
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.script import ScriptEventDescriptor

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_BASIC_DIR = REPO_ROOT / "src" / "basic"
DEFAULT_TEMPLATE = REPO_ROOT / "src" / "templates" / "POKATAK_WMS_1.4.1_template.ods"

# legacy_removal (PRIME 2.0.1): все 38 WMS_* модулей 1.4.1, физически встроенные в шаблон -
# в 2.0.0 они оставались в собранном .ods как "архивный исходный код", только отвязанные от
# кнопок/событий (legacy_modules_bound_to_runtime_ui=false). Мастер-задание 2.0.1 требует
# полностью убрать их из production .ods (legacy_code_allowed_in_production_ods=false) -
# архивная копия исходников остаётся в репозитории под legacy_reference/ (git), не в самом
# файле WMS. Список подтверждён построчной выгрузкой библиотеки "Standard" из шаблона
# (см. tools/list_legacy_modules.py в истории отладки) - должен совпадать 1:1 с реально
# встроенными модулями, иначе static_check_source.py/static_checks.py укажут на расхождение.
LEGACY_MODULES_TO_REMOVE = [
    "WMS_02_Orders_FINAL", "WMS_03_Issues_FINAL", "WMS_04_DB_Connection", "WMS_05_DB_Install",
    "WMS_06_DB_Orders", "WMS_07_DB_Issues", "WMS_08_DB_Returns", "WMS_09_DB_Search",
    "WMS_10_DB_Stock", "WMS_11_DB_UnitsLots", "WMS_12_DB_SmartReceipt", "WMS_13_SystemCenter",
    "WMS_14_UniversalReceipt", "WMS_15_SafetyCore", "WMS_16_Acts", "WMS_17_ActIntegration",
    "WMS_18_ActsRegistry", "WMS_19_ProductionUI", "WMS_20_ManualOperations", "WMS_21_Architecture",
    "WMS_22_References", "WMS_23_ReturnsInventory", "WMS_24_GlobalSearch", "WMS_25_WorkflowActs",
    "WMS_26_StabilityDiagnostics", "WMS_27_RuntimeCore", "WMS_28_DBEngine", "WMS_30_Integrity",
    "WMS_31_UIEngine", "WMS_32_AdminReset", "WMS_33_OfflineExports", "WMS_34_ManagerReport",
    "WMS_35_OrderExtras", "WMS_36_OrderImport", "WMS_37_PrimeUI", "WMS_38_Dashboard",
    "WMS_99_Installer", "WMS_CORE_Common_FINAL",
]

PRIME_MODULES = [
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
    "PRIME_14_MigrationInstaller",
    "PRIME_15_Transfers",
    "PRIME_16_Journal",
    "PRIME_17_Home",
]

# 2.1.0 (R25/R26): "Перемещения"/"Журнал" - совершенно новые листы без единого пре-существующего
# form-контрола в шаблоне 1.4.1 (нечего перепривязывать - BUTTON_MAP работает только поверх уже
# нарисованных кнопок), а у "Комплекты" в 2.0.0/2.0.1 не было привязано вообще ни одной кнопки
# (только одна InputBox-функция, вызываемая исключительно через Сервис -> Макросы). Вместо того
# чтобы оставить эти три листа без единой видимой кнопки (тот же класс проблемы, что и "0 visible
# stubs", только наоборот - "0 visible buttons"), здесь программно создаются реальные
# push-button контролы поверх декоративной шапки листа (строка 1, под заголовком/описанием) и
# сразу привязываются к соответствующему PRIME-макросу - см. create_new_sheet_buttons().
NEW_SHEET_BUTTONS = {
    # FINAL mega-task main_sheet.navigation_buttons: все 3 "Расход — ..." добавлены рядом со
    # своими "Приход — ..." (single_physical_warehouse - расход такой же активный экран, как
    # приход, никаких отдельных "складов"). 16 кнопок всего - ровно 4 полных ряда по
    # NEW_SHEET_BUTTON_COLS["Главная"]=4.
    "Главная": [
        ("Обновить", "Standard.PRIME_17_Home.PRIME_Home_RefreshButton"),
        ("Заказы", "Standard.PRIME_17_Home.PRIME_Home_GoOrders"),
        ("Приход — Офис", "Standard.PRIME_17_Home.PRIME_Home_GoReceiptOffice"),
        ("Расход — Офис", "Standard.PRIME_17_Home.PRIME_Home_GoIssueOffice"),
        ("Приход — Производство", "Standard.PRIME_17_Home.PRIME_Home_GoReceiptProduction"),
        ("Расход — Производство", "Standard.PRIME_17_Home.PRIME_Home_GoIssueProduction"),
        ("Приход — Детали", "Standard.PRIME_17_Home.PRIME_Home_GoReceiptDetails"),
        ("Расход — Детали", "Standard.PRIME_17_Home.PRIME_Home_GoIssueDetails"),
        ("Выдачи", "Standard.PRIME_17_Home.PRIME_Home_GoIssues"),
        ("Наличие", "Standard.PRIME_17_Home.PRIME_Home_GoStock"),
        ("Возвраты", "Standard.PRIME_17_Home.PRIME_Home_GoReturns"),
        ("Перемещения", "Standard.PRIME_17_Home.PRIME_Home_GoTransfers"),
        ("Инвентаризация", "Standard.PRIME_17_Home.PRIME_Home_GoInventory"),
        ("Поиск", "Standard.PRIME_17_Home.PRIME_Home_GoSearch"),
        ("Журнал", "Standard.PRIME_17_Home.PRIME_Home_GoJournal"),
        ("Комплекты", "Standard.PRIME_17_Home.PRIME_Home_GoKits"),
    ],
    "Перемещения": [
        ("Новое перемещение", "Standard.PRIME_15_Transfers.PRIME_Transfers_NewButton"),
        ("Провести выбранное", "Standard.PRIME_15_Transfers.PRIME_Transfers_ConductSelectedButton"),
        ("Провести все", "Standard.PRIME_15_Transfers.PRIME_Transfers_ConductAllButton"),
    ],
    "Журнал": [
        ("Обновить", "Standard.PRIME_16_Journal.PRIME_Journal_RefreshButton"),
    ],
    "Комплекты": [
        ("Добавить комплект в выдачу", "Standard.PRIME_11_Kits.PRIME_Kits_AddToIssuesButton"),
    ],
}

# Листы с лёгким обработчиком PRIME_OnContentChanged_* (ARCHITECTURE §4): Заказы, Выдачи,
# 4 активных цеховых/офисных листа, Возвраты. Ключ события в UNO - "OnChange".
SHEET_EVENT_HANDLERS = {
    "Заказы": "Standard.PRIME_05_Orders.PRIME_OnContentChanged_Orders",
    "Выдачи": "Standard.PRIME_06_Issues.PRIME_OnContentChanged_Issues",
    "Приход — Цех": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Цех": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Приход — Офис": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Офис": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Приход — Производство": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Приход — Детали": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Производство": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Детали": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Возвраты": "Standard.PRIME_08_ReturnsInventory.PRIME_OnContentChanged_Returns",
    "Перемещения": "Standard.PRIME_15_Transfers.PRIME_OnContentChanged_Transfers",
}

# Карта (лист, старое имя контрола) -> новый PRIME-макрос, либо сентинел HIDE.
# Построена по фактической выгрузке кнопок из шаблона (dump_all_buttons.py) - см. коммит с
# картой кнопок в PRIME_ARCHITECTURE обсуждении.
#
# 2.1.0 (R26, "0 visible stubs"): раньше нереализованные функции 1.4.1 привязывались к
# PRIME_UI_NotImplementedStub - кнопка оставалась ВИДИМОЙ и кликабельной, только показывала
# сообщение "не перенесено". Аудит внешней проверки справедливо указал, что 52 такие видимые
# кнопки-заглушки (23 NotImplementedStub + 28 ArchiveStub на легаси-листах + 1 redirect) не
# соответствуют требованию "0 dead buttons". Реальных новых фич под них в этом проходе не
# добавлено (см. KNOWN_ISSUES - осознанно отложенные).
# Пост-review исправление: HIDE изначально означал только EnableVisible=False - контрол
# физически оставался в форме/DrawPage. Повторный аудит справедливо указал, что "production
# ODS should contain no dead user controls" - невидимый, но всё ещё существующий control -
# это тоже dead control. Теперь HIDE означает физическое удаление: и модели контрола из
# form, и его ControlShape с DrawPage листа (см. rebind_buttons ниже) - после сборки такой
# контрол не существует в документе вообще, а не просто скрыт.
# ArchiveStub на 4 легаси-архивных листах (Производство/Детали) - НЕ HIDE: это не "недоделанная
# фича", а осознанно read-only архив истории (см. ARCHITECTURE §5) - сами эти листы уже
# скрываются целиком, если в них нет исторических данных (PRIME_UI_ApplySheetVisibility), а
# если данные есть, кнопка-подсказка "это архив, только для чтения" - корректное, честное
# поведение защищённого read-only листа, а не незакрытый долг.
HIDE = None
ARCHIVE_STUB = "Standard.PRIME_12_UI.PRIME_Legacy_ArchiveStub"

# (sheet, template control name) -> new visible caption. Only for pre-existing 1.4.1 template
# controls that are being repurposed for a different meaning than their original 1.4.1 label -
# see the "Из заказов" comment next to WMS_STOCK_SEARCH in BUTTON_MAP above.
RELABEL_MAP = {
    ("Наличие", "WMS_STOCK_SEARCH"): "Из заказов",
}

BUTTON_MAP = {
    # Инфо
    ("Инфо", "WMSC_REFRESH"): "Standard.PRIME_10_ActsReports.PRIME_Dashboard_RefreshButton",
    ("Инфо", "WMSC_SELF"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    ("Инфо", "WMSC_STOCK"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Инфо", "WMSC_SEARCH"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Инфо", "WMSC_INCOMPLETE"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    ("Инфо", "WMSC_AUDIT"): HIDE,
    ("Инфо", "WMSC_ACTS"): HIDE,
    ("Инфо", "WMSC_ACTOPEN"): "Standard.PRIME_10_ActsReports.PRIME_Acts_OpenByDocIdButton",
    ("Инфо", "WMSC_UPDATE"): "Standard.PRIME_14_MigrationInstaller.PRIME_Migration_RunButton",
    ("Инфо", "WMS_CLEAN_PRODUCT"): HIDE,
    ("Инфо", "WMS_CLEAN_ORDER"): HIDE,
    ("Инфо", "WMS_CLEAN_DOC"): HIDE,
    ("Инфо", "WMS_CLEAN_LOT"): HIDE,
    ("Инфо", "WMS_CLEAN_OPER"): HIDE,
    ("Инфо", "WMS_CLEAN_ALL"): HIDE,
    ("Инфо", "WMS_EXPORT_REPORT_DATA"): HIDE,
    ("Инфо", "pa_control_111"): "Standard.PRIME_12_UI.PRIME_Nav_Dashboard",
    ("Инфо", "pa_control_112"): "Standard.PRIME_12_UI.PRIME_Nav_Report",
    # Заказы
    ("Заказы", "WMS_ORD_BTN_NEW"): "Standard.PRIME_05_Orders.PRIME_Orders_NewOrder",
    ("Заказы", "WMS_ORD_BTN_UNIVERSAL"): "Standard.PRIME_05_Orders.PRIME_Orders_NewOrder",
    ("Заказы", "WMS_ORD_BTN_CONDUCT_POS"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductSelectedButton",
    ("Заказы", "WMS_ORD_BTN_CONDUCT"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductSelectedButton",
    ("Заказы", "WMS_ORD_BTN_CONDUCTALL"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductAllReadyButton",
    ("Заказы", "WMS_ORD_BTN_LOTUNITS"): HIDE,
    ("Заказы", "WMS_ORD_BTN_BULKFILL"): "Standard.PRIME_05_Orders.PRIME_Orders_FillAllByCodeButton",
    ("Заказы", "WMS_ORD_BTN_ARTICLE"): "Standard.PRIME_05_Orders.PRIME_Orders_ResolveByArticleButton",
    ("Заказы", "WMS_ORD_BTN_UPD"): HIDE,
    ("Заказы", "WMS_ORD_BTN_DELDRAFT"): "Standard.PRIME_05_Orders.PRIME_Orders_DeleteDraftButton",
    # Выдачи
    ("Выдачи", "WMS_ISS_BTN_NEW"): "Standard.PRIME_06_Issues.PRIME_Issues_NewIssueButton",
    ("Выдачи", "WMS_ISS_BTN_CONDUCT"): "Standard.PRIME_06_Issues.PRIME_Issues_ConductSelectedButton",
    ("Выдачи", "WMS_ISS_BTN_CONDUCTALL"): "Standard.PRIME_06_Issues.PRIME_Issues_ConductAllButton",
    # single_return_mechanism: возврат теперь оформляется только на листе "Возвраты" -
    # старая кнопка возврата на "Выдачи" скрыта, а не оставлена с redirect-сообщением (R26).
    ("Выдачи", "WMS_ISS_BTN_RETURN"): HIDE,
    ("Выдачи", "WMS_ISS_BTN_LOOKUP"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Выдачи", "WMS_ISS_BTN_BULKFILL"): "Standard.PRIME_06_Issues.PRIME_Issues_FillAllByCodeButton",
    # Наличие (2.1.0, было "Остаток" - см. PRIME_Migration_RenameStockToNalichie). FINAL
    # mega-task (critical_fix, confirmed bug #3): "Из производства"/"Из офиса"/"Детали" были
    # привязаны к контуру (WMS_STOCK_PROD реально показывал SC_GENERAL, а не производство) -
    # переведены на фильтр по происхождению EI (PRIME_Stock_ShowOrigin*Button, см.
    # PRIME_09_StockSearch). Labels в самом шаблоне 1.4.1 уже "Из офиса"/"Из производства"/
    # "Детали" - совпадают со stock_and_presence.quick_filters один в один, меняется только
    # обработчик. WMS_STOCK_SEARCH дублировал WMS_STOCK_HISTORY (тот же PRIME_Search_RunButton,
    # ui_cleanup: remove_duplicate_buttons) - переиспользован под недостающий фильтр
    # "Из заказов" (см. RELABEL_MAP ниже, меняет подпись кнопки на "Из заказов").
    ("Наличие", "WMS_STOCK_ALL"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Наличие", "WMS_STOCK_OFFICE"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowOriginOfficeButton",
    ("Наличие", "WMS_STOCK_PROD"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowOriginProductionButton",
    ("Наличие", "WMS_STOCK_PARTS"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowOriginDetailsButton",
    ("Наличие", "WMS_STOCK_NEG"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowNegativeButton",
    ("Наличие", "WMS_STOCK_HISTORY"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Наличие", "WMS_STOCK_LOTS"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowLotsByCodeButton",
    ("Наличие", "WMS_STOCK_SEARCH"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowOriginOrdersButton",
    ("Наличие", "WMS_STOCK_DIAG"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    # Поиск (2.1.2: переименован из "База - Поиск" ДО этой стадии сборки - см.
    # PRIME_14_MigrationInstaller.PRIME_Migration_RenameSearchToPoisk, вызывается раньше в
    # конвейере, поэтому ключ здесь - уже новое имя, как и у "Наличие" выше)
    ("Поиск", "WMS_S2_FIND"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Поиск", "WMS_S2_ALL"): "Standard.PRIME_09_StockSearch.PRIME_Search_ShowAllButton",
    ("Поиск", "WMS_S2_CLEAR"): "Standard.PRIME_09_StockSearch.PRIME_Search_ClearButton",
    # Справочники
    ("Справочники", "WMS_REF_SAVE"): HIDE,
    ("Справочники", "WMS_REF_REFRESH"): HIDE,
    # Возвраты
    ("Возвраты", "WMS_RET_0"): "Standard.PRIME_08_ReturnsInventory.PRIME_Returns_RefreshButton",
    ("Возвраты", "WMS_RET_1"): "Standard.PRIME_08_ReturnsInventory.PRIME_Returns_ConductButton",
    # Инвентаризация
    ("Инвентаризация", "WMS_INV_0"): "Standard.PRIME_08_ReturnsInventory.PRIME_Inventory_LoadButton",
    ("Инвентаризация", "WMS_INV_1"): "Standard.PRIME_08_ReturnsInventory.PRIME_Inventory_RecalcButton",
    ("Инвентаризация", "WMS_INV_2"): "Standard.PRIME_08_ReturnsInventory.PRIME_Inventory_ConductButton",
    # Дашборд
    ("Дашборд", "pa_control_94"): "Standard.PRIME_10_ActsReports.PRIME_Dashboard_RefreshButton",
    ("Дашборд", "pa_control_95"): "Standard.PRIME_12_UI.PRIME_Nav_Orders",
    ("Дашборд", "pa_control_96"): "Standard.PRIME_12_UI.PRIME_Nav_Stock",
    ("Дашборд", "pa_control_97"): "Standard.PRIME_12_UI.PRIME_Nav_Report",
    # Отчет — ввод (быстрое автозаполнение из истории не реализовано в этом релизе - HIDE, не
    # visible-stub; BTN_7/8/9 реализованы полноценно и остаются активными)
    ("Отчет — ввод", "WMSRPT_BTN_0"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_1"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_2"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_3"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_4"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_5"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_6"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_7"): "Standard.PRIME_10_ActsReports.PRIME_Report_RefreshFactsButton",
    ("Отчет — ввод", "WMSRPT_BTN_8"): "Standard.PRIME_10_ActsReports.PRIME_Report_BuildButton",
    ("Отчет — ввод", "WMSRPT_BTN_9"): "Standard.PRIME_10_ActsReports.PRIME_Report_SaveButton",
    ("Отчет — ввод", "WMSRPT_BTN_10"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_11"): HIDE,
    ("Отчет — ввод", "WMSRPT_BTN_12"): HIDE,
    # Отчет руководителю
    ("Отчет руководителю", "pa_control_108"): "Standard.PRIME_12_UI.PRIME_Nav_ReportInput",
    ("Отчет руководителю", "pa_control_109"): "Standard.PRIME_10_ActsReports.PRIME_Report_SaveButton",
    ("Отчет руководителю", "pa_control_110"): "Standard.PRIME_12_UI.PRIME_Nav_Dashboard",
    # Остаток — Заказы
    ("Остаток — Заказы", "WMS_OXS_0"): "Standard.PRIME_09_StockSearch.PRIME_StockOrders_RefreshButton",
    ("Остаток — Заказы", "WMS_OXS_1"): "Standard.PRIME_09_StockSearch.PRIME_StockOrders_RefreshButton",
}

# Активные workflow-листы получают единый обработчик. FINAL mega-task: single_physical_warehouse -
# "Расход — Производство"/"Расход — Детали" перестают быть архивом истории 1.4.1 (как и
# "Приход — Производство"/"Приход — Детали" стали активными в 2.1.2) и становятся полноценными
# активными issue-листами того же engine, что и "Расход — Офис"/"Выдачи" - см.
# PRIME_00_Config.SH_ISSUE_PRODUCTION/SH_ISSUE_DETAILS и PRIME_07_Workflows.
# "Приход/Расход — Цех" остаются в списке ради работающих кнопок на случай исторических данных на
# скрытом архивном листе (сам лист исключён из авторитетного видимого списка - см.
# PRIME_12_UI.PRIME_UI_VisibleSheetNames), но НЕ являются активным пользовательским экраном.
ACTIVE_WORKFLOW_SHEETS = [
    "Приход — Цех", "Расход — Цех", "Приход — Офис", "Расход — Офис",
    "Приход — Производство", "Приход — Детали",
    "Расход — Производство", "Расход — Детали",
]
WORKFLOW_BUTTON_NAMES = {
    "WMS_WF_ROW": "Standard.PRIME_07_Workflows.PRIME_Workflow_ConductRowButton",
    "WMS_WF_ALL": "Standard.PRIME_07_Workflows.PRIME_Workflow_ConductAllButton",
    "WMS_WF_ACT": "Standard.PRIME_10_ActsReports.PRIME_Acts_CreateFromDocButton",
    "WMS_WF_STOCK": "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    "WMS_WF_SEARCH": "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    "WMS_WF_FILLART": "Standard.PRIME_07_Workflows.PRIME_Workflow_FillArticleButton",
    "WMS_WF_REPEAT": "Standard.PRIME_07_Workflows.PRIME_Workflow_RepeatFieldsButton",
}
for _sheet in ACTIVE_WORKFLOW_SHEETS:
    for _ctrl, _macro in WORKFLOW_BUTTON_NAMES.items():
        BUTTON_MAP[(_sheet, _ctrl)] = _macro

# confirmed bug #5 (forbidden_wording): every active workflow sheet's WMS_WF_SEARCH button is
# labelled "Поиск в базе" in the 1.4.1 template - same forbidden "база" wording as the old
# "База - Поиск" sheet name (already fixed - see PRIME_14_MigrationInstaller.
# PRIME_Migration_RenameSearchToPoisk). The macro binding was already correct
# (PRIME_Search_RunButton); only the caption was stale.
for _sheet in ACTIVE_WORKFLOW_SHEETS:
    RELABEL_MAP[(_sheet, "WMS_WF_SEARCH")] = "Поиск"


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def script_uri(module_dot_sub: str) -> str:
    return f"vnd.sun.star.script:{module_dot_sub}?language=Basic&location=document"


def kill_stale_soffice(profile_dir: Path):
    # xvfb-run wraps soffice.bin in a shell, so proc.terminate() (which only signals the
    # xvfb-run wrapper PID) does not reliably kill the actual soffice.bin child - it can be
    # left running indefinitely, still holding the UserInstallation profile directory. A
    # second build reusing the same fixed profile path then races that orphan and can fail
    # opaquely at doc.store() with SfxBaseModel::storeSelf (observed empirically while
    # developing 2.0.1 - not a code defect in the macros, but real enough to guard against
    # in CI, where a retried step could hit the exact same collision).
    subprocess.run(["pkill", "-9", "-f", f"soffice.bin.*{profile_dir}"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def start_soffice(profile_dir: Path, port: int) -> subprocess.Popen:
    kill_stale_soffice(profile_dir)
    if profile_dir.exists():
        shutil.rmtree(profile_dir)
    cmd = [
        "xvfb-run", "-a", "soffice",
        "--headless", "--invisible", "--nocrashreport", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard",
        f"-env:UserInstallation=file://{profile_dir}",
        f"--accept=socket,host=localhost,port={port};urp;StarOffice.ComponentContext",
    ]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return proc


def connect(port: int, retries: int = 30):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    last_err = None
    for _ in range(retries):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
            return ctx
        except Exception as e:  # noqa: BLE001 - retry loop, re-raised below if exhausted
            last_err = e
            time.sleep(1)
    raise RuntimeError(f"Could not connect to soffice UNO socket: {last_err}")


def inject_basic_modules(doc):
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    if not libs.isLibraryLoaded("Standard"):
        libs.loadLibrary("Standard")
    lib = libs.getByName("Standard")

    for module_name in PRIME_MODULES:
        source_path = SRC_BASIC_DIR / f"{module_name}.bas"
        source = source_path.read_text(encoding="utf-8")
        if lib.hasByName(module_name):
            lib.replaceByName(module_name, source)
        else:
            lib.insertByName(module_name, source)
        print(f"  injected module {module_name} ({len(source)} bytes)")


def remove_legacy_modules(doc):
    libs = doc.BasicLibraries
    lib = libs.getByName("Standard")
    removed = 0
    missing = []
    for module_name in LEGACY_MODULES_TO_REMOVE:
        if lib.hasByName(module_name):
            lib.removeByName(module_name)
            removed += 1
        else:
            missing.append(module_name)
    print(f"  removed {removed}/{len(LEGACY_MODULES_TO_REMOVE)} legacy WMS_* modules")
    if missing:
        print(f"  NOTE: not found (already absent?): {missing}")


def invoke_macro(doc, module_dot_sub: str, args=()):
    script = doc.getScriptProvider().getScript(script_uri(module_dot_sub))
    return script.invoke(args, (), ())


def bind_sheet_events(doc):
    for sheet_name, module_dot_sub in SHEET_EVENT_HANDLERS.items():
        if not doc.Sheets.hasByName(sheet_name):
            print(f"  WARNING: sheet '{sheet_name}' not found, skipping event binding")
            continue
        sheet = doc.Sheets.getByName(sheet_name)
        props_any = uno.Any(
            "[]com.sun.star.beans.PropertyValue",
            (make_prop("EventType", "Script"), make_prop("Script", script_uri(module_dot_sub))),
        )
        uno.invoke(sheet.Events, "replaceByName", ("OnChange", props_any))
        print(f"  bound OnChange on '{sheet_name}' -> {module_dot_sub}")


def rebind_buttons(doc):
    rebound = 0
    # (sheet, form, ctrl_name) of every HIDE-mapped control found - removed in a SEPARATE pass
    # below, after the scan finishes. Removing a control while iterating form.getByIndex(i)/
    # form.Count by index would shift every later index in the same form and skip controls.
    to_remove = []
    for sheet_idx in range(doc.Sheets.Count):
        sheet = doc.Sheets.getByIndex(sheet_idx)
        forms = sheet.DrawPage.Forms
        for form_idx in range(forms.Count):
            form = forms.getByIndex(form_idx)
            for ctrl_idx in range(form.Count):
                ctrl = form.getByIndex(ctrl_idx)
                key = (sheet.Name, ctrl.Name)
                if key in RELABEL_MAP:
                    ctrl.Label = RELABEL_MAP[key]
                if key not in BUTTON_MAP:
                    continue
                target = BUTTON_MAP[key]
                if target is HIDE:
                    to_remove.append((sheet, form, ctrl.Name))
                    continue
                desc = ScriptEventDescriptor()
                desc.ListenerType = "XActionListener"
                desc.EventMethod = "actionPerformed"
                desc.ScriptType = "Script"
                desc.ScriptCode = script_uri(target)
                try:
                    form.revokeScriptEvent(ctrl_idx, "XActionListener", "actionPerformed", "")
                except Exception:
                    pass
                form.registerScriptEvent(ctrl_idx, desc)
                rebound += 1

    # Production ODS should contain no dead user controls: obsolete controls are physically
    # removed (both the ControlShape on the sheet's DrawPage and the control model in the
    # form), not merely hidden with EnableVisible=False - a hidden-but-present control is
    # still a dead control.
    removed = 0
    for sheet, form, ctrl_name in to_remove:
        draw_page = sheet.DrawPage
        shape_found = False
        for shape_idx in range(draw_page.Count):
            shape = draw_page.getByIndex(shape_idx)
            # Match by the control model's Name property, not shape identity - pyuno hands out
            # a fresh wrapper object on each getByIndex() call, so "shape.Control is ctrl" (or
            # even "==") is unreliable across separate lookups of what is the same UNO object.
            try:
                if shape.Control.Name == ctrl_name:
                    draw_page.remove(shape)
                    shape_found = True
                    break
            except Exception:
                continue
        try:
            form.removeByName(ctrl_name)
        except Exception:
            pass
        # Verify against the document's actual state rather than trusting a clean return -
        # empirically, form.removeByName() can raise (likely a disposal-notification artifact
        # of the pyuno bridge) even though the control was in fact removed; checking
        # form.hasByName() afterwards reports what really happened, not what the call claimed.
        if not form.hasByName(ctrl_name):
            removed += 1
        else:
            print(f"  WARNING: failed to remove dead control '{sheet.Name}'!'{ctrl_name}' from its form")
        if not shape_found:
            print(f"  NOTE: no ControlShape found for removed control '{sheet.Name}'!'{ctrl_name}' (model removed anyway)")

    mapped_actions = sum(1 for v in BUTTON_MAP.values() if v is not HIDE)
    mapped_hides = sum(1 for v in BUTTON_MAP.values() if v is HIDE)
    print(f"  rebound {rebound} buttons (of {mapped_actions} mapped actions), removed {removed} dead controls (of {mapped_hides} mapped for removal)")
    if rebound < mapped_actions or removed < mapped_hides:
        print(f"  NOTE: {mapped_actions - rebound + mapped_hides - removed} mapped (sheet, control) pairs were not found in the template")


# 2.1.2: "Главная" needs a real button GRID (13 nav/refresh buttons - a single row would run off
# the screen), unlike the 1-3 buttons on Перемещения/Журнал/Комплекты (single row is fine there).
# Keyed by sheet name; sheets not listed here keep the original single-row layout (cols=None).
NEW_SHEET_BUTTON_COLS = {
    "Главная": 4,
}


def create_new_sheet_buttons(doc):
    """See NEW_SHEET_BUTTONS above: draws real push-button controls on sheets that have no
    pre-existing template control to rebind (brand-new sheets, or a sheet that already existed
    but never had a single bound button)."""
    created = 0
    for sheet_name, buttons in NEW_SHEET_BUTTONS.items():
        if not doc.Sheets.hasByName(sheet_name):
            print(f"  WARNING: sheet '{sheet_name}' not found, skipping button creation")
            continue
        sheet = doc.Sheets.getByName(sheet_name)
        draw_page = sheet.DrawPage
        forms = draw_page.Forms
        if forms.Count == 0:
            form = doc.createInstance("com.sun.star.form.component.Form")
            form.Name = "PRIME_Form"
            forms.insertByIndex(0, form)
        else:
            form = forms.getByIndex(0)

        # One button per ~35mm of width (1 mm = 100 1/100mm units), 8mm tall, just under the top
        # of the sheet - same decorative-header row every business sheet already uses for its
        # title/description band, so this does not collide with the real data header. Sheets in
        # NEW_SHEET_BUTTON_COLS wrap to a new row after that many buttons instead of one long row.
        cols_per_row = NEW_SHEET_BUTTON_COLS.get(sheet_name)
        x_margin = 200   # 2mm left margin
        y_pos = 100      # 1mm from top
        width = 3500     # 35mm
        height = 800     # 8mm
        gap = 200        # 2mm
        x_cursor = x_margin
        for idx, (caption, macro) in enumerate(buttons):
            if cols_per_row and idx > 0 and idx % cols_per_row == 0:
                x_cursor = x_margin
                y_pos += height + gap

            ctrl_name = f"PRIME_BTN_{sheet_name}_{idx}"
            model = doc.createInstance("com.sun.star.form.component.CommandButton")
            model.Name = ctrl_name
            model.Label = caption
            form.insertByName(ctrl_name, model)
            ctrl_idx = form.Count - 1

            shape = doc.createInstance("com.sun.star.drawing.ControlShape")
            shape.Control = model
            shape.Size = uno.createUnoStruct("com.sun.star.awt.Size", width, height)
            shape.Position = uno.createUnoStruct("com.sun.star.awt.Point", x_cursor, y_pos)
            draw_page.add(shape)
            desc = ScriptEventDescriptor()
            desc.ListenerType = "XActionListener"
            desc.EventMethod = "actionPerformed"
            desc.ScriptType = "Script"
            desc.ScriptCode = script_uri(macro)
            form.registerScriptEvent(ctrl_idx, desc)
            created += 1
            x_cursor += width + gap
    print(f"  created {created} new button controls on {len(NEW_SHEET_BUTTONS)} sheets")


def build(template: Path, output: Path, port: int, profile_dir: Path, run_migration: bool):
    template = template.resolve()
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, output)

    print("Starting headless LibreOffice ...")
    proc = start_soffice(profile_dir, port)
    try:
        time.sleep(3)
        ctx = connect(port)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        print(f"Opening {output} ...")
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(output)), "_blank", 0, (make_prop("Hidden", True),))

        print("Injecting PRIME Basic modules ...")
        inject_basic_modules(doc)

        # ВАЖНО: EnsureSchema и любая последующая структурная операция (EnsureBusinessSheet,
        # Migrate) должны идти в ОТДЕЛЬНЫХ top-level invoke()-вызовах от внешнего скрипта, а не в
        # одной цепочке вызовов внутри одного invoke(). Эмпирически подтверждено (см. историю
        # отладки, включая отдельную bisection-сессию 2.1.0): "EnsureSchema, затем сразу
        # EnsureBusinessSheet(нового листа) или MigrateOrdersSheet в рамках ОДНОГО invoke()"
        # молча ничего не создаёт/не мигрирует для этого второго шага (без видимого исключения),
        # а те же операции, разнесённые по отдельным invoke(), работают штатно. Это тот самый
        # механизм, из-за которого "Комплекты"/"Перемещения"/"Журнал" реально отсутствовали в
        # ранее собранных .ods, хотя код, создающий их, формально вызывался. Поэтому здесь ТРИ
        # отдельных вызова (EnsureSchema, затем EnsureAllBusinessSheetsOnly, затем Migrate/
        # DisableLegacyEvents), а не один PRIME_Build_RunFullSetup.
        print("Running PRIME_Install_EnsureSchema ...")
        invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_EnsureSchema")

        print("Running PRIME_Install_EnsureAllBusinessSheetsOnly (separate invoke) ...")
        invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_EnsureAllBusinessSheetsOnly")

        if run_migration:
            print("Running PRIME_Build_MigrateSilent (separate invoke) ...")
            invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Build_MigrateSilent")
        else:
            print("Running PRIME_Install_DisableLegacyEvents (separate invoke) ...")
            invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_DisableLegacyEvents")

        print("Applying PRIME UI layout (separate invoke) ...")
        invoke_macro(doc, "Standard.PRIME_12_UI.PRIME_UI_RestoreInterfaceSilent")

        # confirmed bug #4: called as its OWN top-level invoke() - see the empirically-documented
        # "chained structural ops within one invoke() silently no-op past the first" gotcha
        # explained in the big comment above (EnsureSchema/EnsureBusinessSheet history).
        print("Fixing 'Наличие' panel text (separate invoke) ...")
        invoke_macro(doc, "Standard.PRIME_12_UI.PRIME_UI_FixStockPanelTextButton")

        print("Binding PRIME_OnContentChanged_* sheet events (OnChange) ...")
        bind_sheet_events(doc)

        print("Rebinding buttons to PRIME macros ...")
        rebind_buttons(doc)

        print("Creating buttons on new/previously-buttonless sheets ...")
        create_new_sheet_buttons(doc)

        print("Removing legacy WMS_* modules from production ODS ...")
        remove_legacy_modules(doc)

        print("Saving ...")
        doc.store()
        doc.close(False)
        print(f"Done: {output}")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
        kill_stale_soffice(profile_dir)  # see kill_stale_soffice() - proc.terminate() alone is not enough


def main():
    parser = argparse.ArgumentParser(description="Build ПОКАТАК_PRIME_2.0.1.ods from the 1.4.1 template")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--port", type=int, default=2002)
    parser.add_argument("--profile-dir", type=Path, default=Path("/tmp/prime_build_profile"))
    parser.add_argument("--no-migration", action="store_true",
                         help="Skip the Orders-sheet migration pass (template has no legacy data by default, so it is a no-op reorder+dedup, but can be skipped for speed)")
    args = parser.parse_args()

    build(args.template, args.output, args.port, args.profile_dir, run_migration=not args.no_migration)


if __name__ == "__main__":
    sys.exit(main())
