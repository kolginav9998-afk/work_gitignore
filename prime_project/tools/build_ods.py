#!/usr/bin/env python3
"""
PRIME 2.0.0 builder.

Берёт шаблон ODS 1.4.1 (src/templates/POKATAK_WMS_1.4.1_template.ods), внедряет 15 модулей
PRIME_*.bas в библиотеку Basic "Standard" (старые WMS_* модули остаются как архивный исходный
код - legacy_modules_can_remain_as_archived_source, но отвязываются от кнопок/событий -
legacy_modules_bound_to_runtime_ui=false), создаёт системные/бизнес-листы, привязывает лёгкие
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
]

# Листы с лёгким обработчиком PRIME_OnContentChanged_* (ARCHITECTURE §4): Заказы, Выдачи,
# 4 активных цеховых/офисных листа, Возвраты. Ключ события в UNO - "OnChange".
SHEET_EVENT_HANDLERS = {
    "Заказы": "Standard.PRIME_05_Orders.PRIME_OnContentChanged_Orders",
    "Выдачи": "Standard.PRIME_06_Issues.PRIME_OnContentChanged_Issues",
    "Приход — Цех": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Цех": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Приход — Офис": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Расход — Офис": "Standard.PRIME_07_Workflows.PRIME_OnContentChanged_Workflow",
    "Возвраты": "Standard.PRIME_08_ReturnsInventory.PRIME_OnContentChanged_Returns",
}

# Карта (лист, старое имя контрола) -> новый PRIME-макрос. Построена по фактической выгрузке
# кнопок из шаблона (dump_all_buttons.py) - см. коммит с картой кнопок в PRIME_ARCHITECTURE
# обсуждении. Всё, что не перечислено ниже, остаётся привязанным к старому WMS_*-макросу
# (архивный код, см. NotImplementedStub для явных отказов там, где это важно проговорить).
STUB = "Standard.PRIME_12_UI.PRIME_UI_NotImplementedStub"
ARCHIVE_STUB = "Standard.PRIME_12_UI.PRIME_Legacy_ArchiveStub"

BUTTON_MAP = {
    # Инфо
    ("Инфо", "WMSC_REFRESH"): "Standard.PRIME_10_ActsReports.PRIME_Dashboard_RefreshButton",
    ("Инфо", "WMSC_SELF"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    ("Инфо", "WMSC_STOCK"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Инфо", "WMSC_SEARCH"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Инфо", "WMSC_INCOMPLETE"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    ("Инфо", "WMSC_AUDIT"): STUB,
    ("Инфо", "WMSC_ACTS"): STUB,
    ("Инфо", "WMSC_ACTOPEN"): "Standard.PRIME_10_ActsReports.PRIME_Acts_OpenByDocIdButton",
    ("Инфо", "WMSC_UPDATE"): "Standard.PRIME_14_MigrationInstaller.PRIME_Migration_RunButton",
    ("Инфо", "WMS_CLEAN_PRODUCT"): STUB,
    ("Инфо", "WMS_CLEAN_ORDER"): STUB,
    ("Инфо", "WMS_CLEAN_DOC"): STUB,
    ("Инфо", "WMS_CLEAN_LOT"): STUB,
    ("Инфо", "WMS_CLEAN_OPER"): STUB,
    ("Инфо", "WMS_CLEAN_ALL"): STUB,
    ("Инфо", "WMS_EXPORT_REPORT_DATA"): STUB,
    ("Инфо", "pa_control_111"): "Standard.PRIME_12_UI.PRIME_Nav_Dashboard",
    ("Инфо", "pa_control_112"): "Standard.PRIME_12_UI.PRIME_Nav_Report",
    # Заказы
    ("Заказы", "WMS_ORD_BTN_NEW"): "Standard.PRIME_05_Orders.PRIME_Orders_NewOrder",
    ("Заказы", "WMS_ORD_BTN_UNIVERSAL"): "Standard.PRIME_05_Orders.PRIME_Orders_NewOrder",
    ("Заказы", "WMS_ORD_BTN_CONDUCT_POS"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductSelectedButton",
    ("Заказы", "WMS_ORD_BTN_CONDUCT"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductSelectedButton",
    ("Заказы", "WMS_ORD_BTN_CONDUCTALL"): "Standard.PRIME_05_Orders.PRIME_Orders_ConductAllReadyButton",
    ("Заказы", "WMS_ORD_BTN_LOTUNITS"): STUB,
    ("Заказы", "WMS_ORD_BTN_BULKFILL"): "Standard.PRIME_05_Orders.PRIME_Orders_FillAllByCodeButton",
    ("Заказы", "WMS_ORD_BTN_ARTICLE"): "Standard.PRIME_05_Orders.PRIME_Orders_ResolveByArticleButton",
    ("Заказы", "WMS_ORD_BTN_UPD"): STUB,
    ("Заказы", "WMS_ORD_BTN_DELDRAFT"): "Standard.PRIME_05_Orders.PRIME_Orders_DeleteDraftButton",
    # Выдачи
    ("Выдачи", "WMS_ISS_BTN_NEW"): "Standard.PRIME_06_Issues.PRIME_Issues_NewIssueButton",
    ("Выдачи", "WMS_ISS_BTN_CONDUCT"): "Standard.PRIME_06_Issues.PRIME_Issues_ConductSelectedButton",
    ("Выдачи", "WMS_ISS_BTN_CONDUCTALL"): "Standard.PRIME_06_Issues.PRIME_Issues_ConductAllButton",
    ("Выдачи", "WMS_ISS_BTN_RETURN"): "Standard.PRIME_12_UI.PRIME_Issues_ReturnRedirectStub",
    ("Выдачи", "WMS_ISS_BTN_LOOKUP"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Выдачи", "WMS_ISS_BTN_BULKFILL"): "Standard.PRIME_06_Issues.PRIME_Issues_FillAllByCodeButton",
    # Остаток
    ("Остаток", "WMS_STOCK_ALL"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Остаток", "WMS_STOCK_OFFICE"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Остаток", "WMS_STOCK_PROD"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Остаток", "WMS_STOCK_PARTS"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowAllButton",
    ("Остаток", "WMS_STOCK_NEG"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowNegativeButton",
    ("Остаток", "WMS_STOCK_HISTORY"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Остаток", "WMS_STOCK_LOTS"): "Standard.PRIME_09_StockSearch.PRIME_Stock_ShowLotsByCodeButton",
    ("Остаток", "WMS_STOCK_SEARCH"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("Остаток", "WMS_STOCK_DIAG"): "Standard.PRIME_13_Diagnostics.PRIME_Diagnostics_RunButton",
    # База - Поиск
    ("База - Поиск", "WMS_S2_FIND"): "Standard.PRIME_09_StockSearch.PRIME_Search_RunButton",
    ("База - Поиск", "WMS_S2_ALL"): "Standard.PRIME_09_StockSearch.PRIME_Search_ShowAllButton",
    ("База - Поиск", "WMS_S2_CLEAR"): "Standard.PRIME_09_StockSearch.PRIME_Search_ClearButton",
    # Справочники
    ("Справочники", "WMS_REF_SAVE"): STUB,
    ("Справочники", "WMS_REF_REFRESH"): STUB,
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
    # Отчет — ввод
    ("Отчет — ввод", "WMSRPT_BTN_0"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_1"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_2"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_3"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_4"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_5"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_6"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_7"): "Standard.PRIME_10_ActsReports.PRIME_Report_RefreshFactsButton",
    ("Отчет — ввод", "WMSRPT_BTN_8"): "Standard.PRIME_10_ActsReports.PRIME_Report_BuildButton",
    ("Отчет — ввод", "WMSRPT_BTN_9"): "Standard.PRIME_10_ActsReports.PRIME_Report_SaveButton",
    ("Отчет — ввод", "WMSRPT_BTN_10"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_11"): STUB,
    ("Отчет — ввод", "WMSRPT_BTN_12"): STUB,
    # Отчет руководителю
    ("Отчет руководителю", "pa_control_108"): "Standard.PRIME_12_UI.PRIME_Nav_ReportInput",
    ("Отчет руководителю", "pa_control_109"): "Standard.PRIME_10_ActsReports.PRIME_Report_SaveButton",
    ("Отчет руководителю", "pa_control_110"): "Standard.PRIME_12_UI.PRIME_Nav_Dashboard",
    # Остаток — Заказы
    ("Остаток — Заказы", "WMS_OXS_0"): "Standard.PRIME_09_StockSearch.PRIME_StockOrders_RefreshButton",
    ("Остаток — Заказы", "WMS_OXS_1"): "Standard.PRIME_09_StockSearch.PRIME_StockOrders_RefreshButton",
}

# Активные workflow-листы получают единый обработчик; легаси-архивные (Производство/Детали)
# получают явный "архив истории" стаб на каждой кнопке, а не тихо оставленный старый макрос,
# который может попытаться обратиться к отсутствующему Firebird.
ACTIVE_WORKFLOW_SHEETS = ["Приход — Цех", "Расход — Цех", "Приход — Офис", "Расход — Офис"]
LEGACY_WORKFLOW_SHEETS = ["Приход — Производство", "Расход — Производство", "Приход — Детали", "Расход — Детали"]
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
for _sheet in LEGACY_WORKFLOW_SHEETS:
    for _ctrl in WORKFLOW_BUTTON_NAMES:
        BUTTON_MAP[(_sheet, _ctrl)] = ARCHIVE_STUB


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def script_uri(module_dot_sub: str) -> str:
    return f"vnd.sun.star.script:{module_dot_sub}?language=Basic&location=document"


def start_soffice(profile_dir: Path, port: int) -> subprocess.Popen:
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
    missing_sheets = set()
    for sheet_idx in range(doc.Sheets.Count):
        sheet = doc.Sheets.getByIndex(sheet_idx)
        forms = sheet.DrawPage.Forms
        for form_idx in range(forms.Count):
            form = forms.getByIndex(form_idx)
            for ctrl_idx in range(form.Count):
                ctrl = form.getByIndex(ctrl_idx)
                key = (sheet.Name, ctrl.Name)
                if key not in BUTTON_MAP:
                    continue
                target = BUTTON_MAP[key]
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
    print(f"  rebound {rebound} buttons (of {len(BUTTON_MAP)} mapped)")
    if rebound < len(BUTTON_MAP):
        print(f"  NOTE: {len(BUTTON_MAP) - rebound} mapped (sheet, control) pairs were not found in the template")


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

        print("Running PRIME_Install_EnsureAllBusinessSheetsSilent ...")
        invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_EnsureAllBusinessSheetsSilent")

        if run_migration:
            print("Running PRIME_Build_MigrateSilent (Orders column migration + product seeding) ...")
            invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Build_MigrateSilent")
        else:
            print("Running PRIME_Install_DisableLegacyEvents ...")
            invoke_macro(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_DisableLegacyEvents")

        print("Binding PRIME_OnContentChanged_* sheet events (OnChange) ...")
        bind_sheet_events(doc)

        print("Rebinding buttons to PRIME macros ...")
        rebind_buttons(doc)

        print("Applying PRIME UI layout (PRIME_UI_RestoreInterfaceButton) ...")
        invoke_macro(doc, "Standard.PRIME_12_UI.PRIME_UI_RestoreInterfaceButton")

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


def main():
    parser = argparse.ArgumentParser(description="Build ПОКАТАК_PRIME_2.0.0.ods from the 1.4.1 template")
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
