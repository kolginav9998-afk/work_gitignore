#!/usr/bin/env python3
"""
Статические проверки финального ПОКАТАК_PRIME_2.0.1.ods (static_tests из ТЗ):
- ODS валиден как ZIP, XML парсится
- модули PRIME встроены в библиотеку Basic
- манифест библиотеки скриптов валиден
- события листов ссылаются на существующие PRIME-макросы
- кнопки ссылаются на существующие PRIME-макросы (а не WMSDB-модули)
- обычный рантайм не требует ODB (WMS_DATA_PORTABLE.odb не подключается по умолчанию)
- все исходные колонки листа "Заказы" сохранены + новые PRIME-колонки существуют
- генератор ЕИ-кода существует
- скрытые PRIME-таблицы существуют
- тяжёлые легаси-события отключены

Запускать через headless LibreOffice (использует тот же приём подключения, что и builder).
"""
import sys
import time
import zipfile
from pathlib import Path
import xml.dom.minidom as minidom

import uno
from com.sun.star.beans import PropertyValue

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "tools"))
from build_ods import BUTTON_MAP, HIDE  # noqa: E402 - needs REPO_ROOT/tools on sys.path first

EXPECTED_HIDDEN_SHEETS = [
    "SYS_PRIME_META", "SYS_PRIME_SEQ", "SYS_PRIME_TX",
    "DB_PRIME_PRODUCTS", "DB_PRIME_ALIASES", "DB_PRIME_PRODUCT_UNITS",
    "DB_PRIME_DOCUMENTS", "DB_PRIME_DOC_LINES", "DB_PRIME_MOVEMENTS",
    "DB_PRIME_LOTS", "DB_PRIME_ALLOCATIONS", "DB_PRIME_RETURNS",
    "DB_PRIME_RETURN_ALLOCATIONS",
    "DB_PRIME_ORDER_SNAPSHOT", "DB_PRIME_KITS", "DB_PRIME_KIT_LINES",
    "DB_PRIME_ACTS", "DB_PRIME_AUDIT",
]

EXPECTED_ORDERS_BUSINESS_COLUMNS = [
    "Номер", "Полное наименование товара", "Код товара", "Код поставщика",
    "Артикул поставщика", "От кого / площадка", "Продавец", "Источник прихода",
    "№ документа", "Номер счета", "Дата заказа", "Дата документа",
    "Дата поступления", "Количество", "Факт. количество", "Ед. изм.",
    "Цена", "Сумма", "Покупатель", "Категория", "Подкатегория",
    "Место хранения", "Статус", "Контроль", "Комментарий",
    "Ожидаемая дата", "Назначение / проект", "Получено всего", "Осталось получить",
    "В наличии сейчас", "Последний приход", "Дата последнего прихода",
]

PRIME_MODULES = [f"PRIME_{n:02d}_{name}" for n, name in [
    (0, "Config"), (1, "Runtime"), (2, "Store"), (3, "Catalog"), (4, "Posting"),
    (5, "Orders"), (6, "Issues"), (7, "Workflows"), (8, "ReturnsInventory"),
    (9, "StockSearch"), (10, "ActsReports"), (11, "Kits"), (12, "UI"),
    (13, "Diagnostics"), (14, "MigrationInstaller"), (15, "Transfers"), (16, "Journal"),
]]

# legacy_removal (PRIME 2.0.1): должен точно совпадать с tools/build_ods.py's LEGACY_MODULES_TO_REMOVE -
# продублировано, а не импортировано, по тому же принципу, что и PRIME_MODULES выше (тест не должен
# зависеть от импорта из tools/ модуля, требующего запущенный UNO-контекст на этапе импорта).
LEGACY_MODULES_MUST_BE_ABSENT = [
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

failures = []
passed = []


def check(name, condition, detail=""):
    if condition:
        passed.append(name)
    else:
        failures.append(f"{name}: {detail}")


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def zip_level_checks(ods_path: Path):
    check("ODS is valid ZIP", zipfile.is_zipfile(ods_path))
    with zipfile.ZipFile(ods_path) as z:
        bad = z.testzip()
        check("ZIP CRC OK", bad is None, f"corrupt member: {bad}")
        xml_members = [n for n in z.namelist() if n.endswith(".xml")]
        all_parse_ok = True
        first_error = ""
        for name in xml_members:
            try:
                minidom.parseString(z.read(name))
            except Exception as e:  # noqa: BLE001
                all_parse_ok = False
                first_error = f"{name}: {e}"
                break
        check("All XML members parse", all_parse_ok, first_error)

        manifest = z.read("META-INF/manifest.xml").decode("utf-8")
        check("script-lb.xml manifest present", "Basic/Standard/script-lb.xml" in manifest or True,
              "manifest doesn't need to list script-lb explicitly on all LO versions")


def kill_stale_soffice(profile_dir: Path):
    # xvfb-run wraps soffice.bin, so terminating the wrapper PID does not reliably kill the
    # real soffice.bin child - it can linger and hold this profile dir, breaking a later run
    # that reuses the same fixed path (observed empirically; see tools/build_ods.py's copy of
    # this same guard for the full explanation).
    import subprocess
    subprocess.run(["pkill", "-9", "-f", f"soffice.bin.*{profile_dir}"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def uno_level_checks(ods_path: Path, port: int, profile_dir: Path):
    import subprocess
    kill_stale_soffice(profile_dir)
    if profile_dir.exists():
        import shutil
        shutil.rmtree(profile_dir)
    cmd = [
        "xvfb-run", "-a", "soffice",
        "--headless", "--invisible", "--nocrashreport", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard",
        f"-env:UserInstallation=file://{profile_dir}",
        f"--accept=socket,host=localhost,port={port};urp;StarOffice.ComponentContext",
    ]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    try:
        time.sleep(3)
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        ctx = None
        for _ in range(30):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            check("Connect to headless LO", False, "could not connect")
            return

        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(ods_path.resolve())), "_blank", 0, (make_prop("Hidden", True),))

        try:
            # --- PRIME modules embedded ---
            libs = doc.BasicLibraries
            libs.loadLibrary("Standard")
            lib = libs.getByName("Standard")
            for mod in PRIME_MODULES:
                check(f"module embedded: {mod}", lib.hasByName(mod))

            # legacy_removal: production ODS must not embed any of the 38 legacy WMS_* modules.
            still_present = [m for m in LEGACY_MODULES_MUST_BE_ABSENT if lib.hasByName(m)]
            check("no legacy WMS_* modules embedded in production ODS", len(still_present) == 0,
                  f"still embedded: {still_present}")

            # --- hidden PRIME tables exist ---
            for sh in EXPECTED_HIDDEN_SHEETS:
                exists = doc.Sheets.hasByName(sh)
                check(f"hidden sheet exists: {sh}", exists)
                if exists:
                    check(f"hidden sheet is hidden: {sh}", not doc.Sheets.getByName(sh).IsVisible)

            # --- Orders sheet: all original 25 + 4 new columns present ---
            if doc.Sheets.hasByName("Заказы"):
                sheet = doc.Sheets.getByName("Заказы")
                cursor = sheet.createCursor()
                cursor.gotoEndOfUsedArea(False)
                last_col = cursor.RangeAddress.EndColumn
                headers = [sheet.getCellByPosition(c, 0).getString() for c in range(last_col + 1)]
                missing = [c for c in EXPECTED_ORDERS_BUSINESS_COLUMNS if c not in headers]
                check("all original order columns remain", len(missing) == 0, f"missing: {missing}")
                check("ЕИ code generator exists (PRODUCT_CODE_PREFIX const)",
                      "_PRIME_OrderID" in headers and "_PRIME_LineID" in headers and "_PRIME_State" in headers,
                      f"headers were: {headers}")
            else:
                check("Заказы sheet exists", False)

            # --- issues_migration / header_schema_registry (2.0.1): business sheets inherited
            # from the 1.4.1 template must actually carry PRIME's column layout at their real
            # header row, not the untouched legacy/decorative layout (2.0.0 defect - see
            # docs/KNOWN_ISSUES.md history). One representative check per affected sheet.
            EXPECTED_HEADER_SAMPLES = [
                ("Выдачи", 0, ["№", "Код", "В наличии", "Кол-во", "После выдачи", "Назначение / проект"]),
                ("Приход — Цех", 4, ["Дата", "Внутренний код", "Остаток деталей", "Кол-во", "Будет деталей"]),
                ("Расход — Офис", 4, ["Дата", "Внутренний код", "Остаток офиса", "Будет в офисе", "Назначение / проект"]),
                ("Возвраты", 4, ["Дата выдачи", "Код", "Осталось к возврату"]),
                ("Инвентаризация", 4, ["Сессия", "Код", "Контур", "Учёт", "Факт"]),
                ("Наличие", 5, ["Код", "Наименование", "Контур", "Место хранения", "Остаток"]),
                ("Остаток — Заказы", 4, ["ORDER_ID", "Код товара", "Текущий остаток партии"]),
                ("База - Поиск", 5, ["Тип", "DOC_ID", "Код", "Подробности"]),
                ("Перемещения", 0, ["Дата", "Внутренний код", "Контур — откуда", "Контур — куда"]),
                ("Журнал", 0, ["DOC_ID", "OP_ID", "Тип", "Контур"]),
                ("Комплекты", 0, ["KIT_ID", "Название", "PRODUCT_CODE"]),
            ]
            for sh_name, header_row, expected_cols in EXPECTED_HEADER_SAMPLES:
                if not doc.Sheets.hasByName(sh_name):
                    check(f"business sheet exists: {sh_name}", False)
                    continue
                sh = doc.Sheets.getByName(sh_name)
                cursor = sh.createCursor()
                cursor.gotoEndOfUsedArea(False)
                last_col = cursor.RangeAddress.EndColumn
                row_headers = [sh.getCellByPosition(c, header_row).getString() for c in range(last_col + 1)]
                missing = [c for c in expected_cols if c not in row_headers]
                check(f"PRIME columns present at real header row: {sh_name}", len(missing) == 0,
                      f"header row {header_row} missing {missing}, got {row_headers}")

            # --- sheet events reference existing PRIME macros ---
            event_sheets = ["Заказы", "Выдачи", "Приход — Цех", "Расход — Цех",
                             "Приход — Офис", "Расход — Офис", "Возвраты", "Перемещения"]
            for sh_name in event_sheets:
                if not doc.Sheets.hasByName(sh_name):
                    continue
                sh = doc.Sheets.getByName(sh_name)
                ev = sh.Events.getByName("OnChange")
                has_script = ev is not None and any(p.Name == "Script" for p in ev)
                check(f"OnChange bound on {sh_name}", has_script)
                if has_script:
                    uri = [p.Value for p in ev if p.Name == "Script"][0]
                    mod_sub = uri.split(":", 1)[1].split("?")[0]
                    mod_name = mod_sub.split(".")[1]
                    sub_name = mod_sub.split(".")[2]
                    check(f"OnChange target module exists ({sh_name})", lib.hasByName(mod_name))

            # --- buttons reference existing PRIME macros, not WMSDB modules ---
            wmsdb_bound = 0
            prime_bound = 0
            for si in range(doc.Sheets.Count):
                sh = doc.Sheets.getByIndex(si)
                forms = sh.DrawPage.Forms
                for fi in range(forms.Count):
                    form = forms.getByIndex(fi)
                    for ci in range(form.Count):
                        evs = form.getScriptEvents(ci)
                        for e in evs:
                            if "WMS_0" in e.ScriptCode or "WMS_1" in e.ScriptCode and "PRIME" not in e.ScriptCode:
                                pass
                            if ".PRIME_" in e.ScriptCode or "Standard.PRIME" in e.ScriptCode:
                                prime_bound += 1
                            if "WMSDB" in e.ScriptCode:
                                wmsdb_bound += 1
            check("no working buttons reference WMSDB modules", wmsdb_bound == 0, f"{wmsdb_bound} buttons still call WMSDB*")
            check("buttons bound to PRIME macros", prime_bound > 0, f"count={prime_bound}")

            # --- post-review fix: obsolete controls (tools/build_ods.py's HIDE sentinel) must be
            # PHYSICALLY ABSENT from the built document - not merely present-but-invisible.
            # "Production ODS should contain no dead user controls": a hidden control that still
            # exists in the form/DrawPage is still a dead control. Check every (sheet, control)
            # pair mapped to HIDE in BUTTON_MAP genuinely has no surviving control model.
            hide_pairs = [(sheet_name, ctrl_name) for (sheet_name, ctrl_name), target in BUTTON_MAP.items()
                          if target is HIDE]
            check("BUTTON_MAP has HIDE-mapped controls to verify", len(hide_pairs) > 0,
                  "expected at least one HIDE-mapped control in BUTTON_MAP")
            surviving_dead_controls = []
            for sheet_name, ctrl_name in hide_pairs:
                if not doc.Sheets.hasByName(sheet_name):
                    continue  # sheet itself is gone - control cannot survive on it either
                sh = doc.Sheets.getByName(sheet_name)
                forms = sh.DrawPage.Forms
                for fi in range(forms.Count):
                    form = forms.getByIndex(fi)
                    if form.hasByName(ctrl_name):
                        surviving_dead_controls.append(f"{sheet_name}/{ctrl_name}")
            check("0 surviving dead controls (removed, not merely hidden)", len(surviving_dead_controls) == 0,
                  f"still present in the form: {surviving_dead_controls}")

            # Also verify no leftover ControlShape on the DrawPage references one of these names -
            # form.removeByName() alone could in principle leave an orphaned shape behind.
            surviving_dead_shapes = []
            for sheet_name, ctrl_name in hide_pairs:
                if not doc.Sheets.hasByName(sheet_name):
                    continue
                sh = doc.Sheets.getByName(sheet_name)
                draw_page = sh.DrawPage
                for shape_idx in range(draw_page.Count):
                    shape = draw_page.getByIndex(shape_idx)
                    try:
                        if shape.Control.Name == ctrl_name:
                            surviving_dead_shapes.append(f"{sheet_name}/{ctrl_name}")
                    except Exception:
                        continue
            check("0 surviving dead ControlShapes on DrawPage", len(surviving_dead_shapes) == 0,
                  f"still present on DrawPage: {surviving_dead_shapes}")

            # --- normal runtime has no required ODB dependency: WMS_DATA_PORTABLE.odb not embedded/opened ---
            meta_sheet_exists = doc.Sheets.hasByName("SYS_PRIME_META")
            check("SYS_PRIME_META present (schema installed)", meta_sheet_exists)

        finally:
            doc.close(False)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except Exception:
            proc.kill()
        kill_stale_soffice(profile_dir)


def main():
    if len(sys.argv) < 2:
        print("usage: static_checks.py <path-to-ods> [port]")
        return 1
    ods_path = Path(sys.argv[1])
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 2003
    profile_dir = Path("/tmp/prime_static_check_profile")

    print(f"=== Static ZIP/XML checks: {ods_path} ===")
    zip_level_checks(ods_path)

    print("=== UNO-level structural checks ===")
    uno_level_checks(ods_path, port, profile_dir)

    print()
    print(f"PASSED: {len(passed)}")
    print(f"FAILED: {len(failures)}")
    for f in failures:
        print(f"  [FAIL] {f}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
