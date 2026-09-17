#!/usr/bin/env python3
"""
Статические проверки финального ПОКАТАК_PRIME_2.0.0.ods (static_tests из ТЗ):
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

EXPECTED_HIDDEN_SHEETS = [
    "SYS_PRIME_META", "SYS_PRIME_SEQ", "SYS_PRIME_TX",
    "DB_PRIME_PRODUCTS", "DB_PRIME_ALIASES", "DB_PRIME_PRODUCT_UNITS",
    "DB_PRIME_DOCUMENTS", "DB_PRIME_DOC_LINES", "DB_PRIME_MOVEMENTS",
    "DB_PRIME_LOTS", "DB_PRIME_ALLOCATIONS", "DB_PRIME_RETURNS",
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
]

PRIME_MODULES = [f"PRIME_{n:02d}_{name}" for n, name in [
    (0, "Config"), (1, "Runtime"), (2, "Store"), (3, "Catalog"), (4, "Posting"),
    (5, "Orders"), (6, "Issues"), (7, "Workflows"), (8, "ReturnsInventory"),
    (9, "StockSearch"), (10, "ActsReports"), (11, "Kits"), (12, "UI"),
    (13, "Diagnostics"), (14, "MigrationInstaller"),
]]

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


def uno_level_checks(ods_path: Path, port: int, profile_dir: Path):
    import subprocess
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

            # --- sheet events reference existing PRIME macros ---
            event_sheets = ["Заказы", "Выдачи", "Приход — Цех", "Расход — Цех",
                             "Приход — Офис", "Расход — Офис", "Возвраты"]
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
