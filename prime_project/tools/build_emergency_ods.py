#!/usr/bin/env python3
"""
build_emergency_ods.py — ПОКАТАК PRIME EMERGENCY WORKING VERSION.

По прямому указанию пользователя: остановить дальнейшее усложнение основного движка PRIME
(src/basic/PRIME_00..17.bas — не тронуты ни строкой) и вместо этого за один проход собрать
маленькую, полностью независимую, формула-ориентированную рабочую версию:

    Приход (Заказы / Приход — Офис / Приход — Производство / Приход — Детали)
        -> каждая фактически пришедшая позиция получает новый уникальный ЕИ-код
    Выдача (один лист "Выдачи")
        -> списывает по явно введённому ЕИ-коду, остаток = живая формула SUMIFS,
           не отдельная таблица движений

Никаких скрытых DB_PRIME_*/SYS_PRIME_* листов, никакого транзакционного протокола
PREPARED/COMMITTED, никаких событий листа (OnChange) — только явные кнопки "Провести всё".
Это осознанный выбор простоты и надёжности для срочной версии, а не недосмотр.

Собирается с нуля (private:factory/scalc), а не из шаблона 1.4.1 — в этой версии нет ничего,
что нужно было бы наследовать из легаси-шаблона.
"""
import argparse
import shutil
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet import TableValidationVisibility
from com.sun.star.sheet.ValidationType import LIST as VALIDATION_LIST, DECIMAL as VALIDATION_DECIMAL
from com.sun.star.sheet.ValidationAlertStyle import STOP as ALERT_STOP
from com.sun.star.sheet.ConditionOperator import GREATER_EQUAL
from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.script import ScriptEventDescriptor

REPO_ROOT = Path(__file__).resolve().parents[1]
MODULE_SOURCE = REPO_ROOT / "src" / "basic_emergency" / "Emergency.bas"

SH_ORDERS = "Заказы"
SH_OFFICE = "Приход — Офис"
SH_PROD = "Приход — Производство"
SH_PARTS = "Приход — Детали"
SH_ISSUE = "Выдачи"
SH_STOCK = "Наличие"
SH_SEQ = "_SEQ"

MAX_ROW = 999  # предзаполненных строк с формулами (1..999, заголовок в строке 0)

UNIT_LIST = "шт\nкг\nм\nупак\nл\nкомпл"


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def script_uri(module_dot_sub: str) -> str:
    return f"vnd.sun.star.script:Standard.{module_dot_sub}?language=Basic&location=document"


def kill_stale_soffice(profile_dir: Path):
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
    return subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def connect(port: int, retries: int = 30):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    last_err = None
    for _ in range(retries):
        try:
            return resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(1)
    raise RuntimeError(f"Could not connect to soffice UNO socket: {last_err}")


def set_headers(sheet, headers):
    for c, text in enumerate(headers):
        cell = sheet.getCellByPosition(c, 0)
        cell.setString(text)
        cell.CharWeight = BOLD
    sheet.Columns.getByIndex(0).OptimalWidth = True


def apply_unit_validation(sheet, col):
    cell_range = sheet.getCellRangeByPosition(col, 1, col, MAX_ROW)
    validation = cell_range.Validation
    validation.Type = VALIDATION_LIST
    validation.setFormula1(UNIT_LIST)
    validation.ShowList = TableValidationVisibility.UNSORTED
    validation.IgnoreBlankCells = True
    validation.ShowErrorMessage = True
    validation.ErrorAlertStyle = ALERT_STOP
    validation.ErrorTitle = "Единица измерения"
    validation.ErrorMessage = "Выберите единицу измерения из списка (шт/кг/м/упак/л/компл)."
    cell_range.Validation = validation


def apply_nonnegative_number_validation(sheet, col, title):
    cell_range = sheet.getCellRangeByPosition(col, 1, col, MAX_ROW)
    validation = cell_range.Validation
    validation.Type = VALIDATION_DECIMAL
    validation.Operator = GREATER_EQUAL
    validation.setFormula1("0")
    validation.IgnoreBlankCells = True
    validation.ShowErrorMessage = True
    validation.ErrorAlertStyle = ALERT_STOP
    validation.ErrorTitle = title
    validation.ErrorMessage = "Только число, не меньше нуля."
    cell_range.Validation = validation


def fill_formula_column(sheet, col, formula_for_row):
    for r in range(1, MAX_ROW + 1):
        sheet.getCellByPosition(col, r).setFormula(formula_for_row(r + 1))  # +1: Calc-формулы 1-индексные


def build(output: Path, port: int, profile_dir: Path):
    print("Starting headless LibreOffice ...")
    proc = start_soffice(profile_dir, port)
    try:
        time.sleep(3)
        ctx = connect(port)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        print("Creating blank Calc document ...")
        doc = desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, (make_prop("Hidden", True),))

        sheets = doc.Sheets
        # Пустой документ Calc всегда создаётся ровно с одним листом в этой версии LO -
        # переиспользуем его под первый нужный лист, остальные 5 добавляем по имени.
        sheets.getByIndex(0).Name = SH_ORDERS
        for name in (SH_OFFICE, SH_PROD, SH_PARTS, SH_ISSUE, SH_STOCK, SH_SEQ):
            sheets.insertNewByName(name, sheets.Count)

        print("Writing headers, formulas, validations ...")

        # --- Заказы: 0 Номер,1 Наименование,2 Количество,3 Факт.количество,4 Ед.изм.,5 Место,
        # 6 Дата,7 Комментарий,8 ЕИ-код,9 Остаток(формула),10 Статус ---
        orders = sheets.getByName(SH_ORDERS)
        set_headers(orders, ["Номер заказа", "Наименование", "Количество", "Факт. количество",
                              "Ед. изм.", "Место", "Дата", "Комментарий", "ЕИ-код", "Остаток", "Статус"])
        apply_unit_validation(orders, 4)
        apply_nonnegative_number_validation(orders, 2, "Количество")
        apply_nonnegative_number_validation(orders, 3, "Факт. количество")
        # locale note: this environment's Calc formula argument separator is ";", not "," -
        # SUMIFS(...,...) parses as Err:508 ("pair missing") here; verified empirically (see
        # tools/build_emergency_ods.py commit message) before writing this the semicolon way.
        fill_formula_column(orders, 9, lambda r: (
            f'=IF(I{r}="";"";D{r}-SUMIFS({SH_ISSUE}.D:D;{SH_ISSUE}.B:B;I{r};{SH_ISSUE}.K:K;"Проведено"))'
        ))

        # --- Приход — Офис / Производство: 0 ЕИ-код,1 Наименование,2 Пришло,3 Ед.изм.,
        # 4 Остаток(формула),5 Место,6 Дата,7 Комментарий ---
        for name in (SH_OFFICE, SH_PROD):
            sh = sheets.getByName(name)
            set_headers(sh, ["ЕИ-код", "Наименование", "Пришло", "Ед. изм.", "Остаток",
                              "Место", "Дата", "Комментарий"])
            apply_unit_validation(sh, 3)
            apply_nonnegative_number_validation(sh, 2, "Пришло")
            fill_formula_column(sh, 4, lambda r: (
                f'=IF(A{r}="";"";C{r}-SUMIFS({SH_ISSUE}.D:D;{SH_ISSUE}.B:B;A{r};{SH_ISSUE}.K:K;"Проведено"))'
            ))

        # --- Приход — Детали: 0 ЕИ-код,1 Наименование,2 Артикул,3 Пришло,4 Ед.изм.,
        # 5 Остаток(формула),6 Место,7 Дата,8 Комментарий ---
        parts = sheets.getByName(SH_PARTS)
        set_headers(parts, ["ЕИ-код", "Наименование", "Артикул", "Пришло", "Ед. изм.", "Остаток",
                             "Место", "Дата", "Комментарий"])
        apply_unit_validation(parts, 4)
        apply_nonnegative_number_validation(parts, 3, "Пришло")
        fill_formula_column(parts, 5, lambda r: (
            f'=IF(A{r}="";"";D{r}-SUMIFS({SH_ISSUE}.D:D;{SH_ISSUE}.B:B;A{r};{SH_ISSUE}.K:K;"Проведено"))'
        ))

        # --- Выдачи: 0 Дата,1 ЕИ-код,2 Наименование,3 Количество,4 Ед.изм.,5 Доступно,
        # 6 Кто получил,7 Куда,8 Возвратный,9 Комментарий,10 Статус ---
        issues = sheets.getByName(SH_ISSUE)
        set_headers(issues, ["Дата", "ЕИ-код", "Наименование", "Количество", "Ед. изм.",
                              "Доступно", "Кто получил", "Куда", "Возвратный", "Комментарий", "Статус"])
        apply_unit_validation(issues, 4)
        apply_nonnegative_number_validation(issues, 3, "Количество")

        # --- Наличие: 0 ЕИ-код,1 Наименование,2 Откуда пришло,3 Пришло,4 Выдано,5 Остаток,
        # 6 Место (только просмотр, заполняется макросом "Обновить") ---
        stock = sheets.getByName(SH_STOCK)
        set_headers(stock, ["ЕИ-код", "Наименование", "Откуда пришло", "Пришло", "Выдано",
                             "Остаток", "Место"])

        # --- _SEQ: скрытый счётчик ЕИ-кодов ---
        seq = sheets.getByName(SH_SEQ)
        seq.getCellByPosition(0, 0).setValue(0)
        seq.IsVisible = False

        controller = doc.CurrentController
        for name in (SH_ORDERS, SH_OFFICE, SH_PROD, SH_PARTS, SH_ISSUE, SH_STOCK):
            controller.setActiveSheet(sheets.getByName(name))
            controller.freezeAtPosition(0, 1)

        print("Injecting Emergency Basic module ...")
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        if not libs.isLibraryLoaded("Standard"):
            libs.loadLibrary("Standard")
        lib = libs.getByName("Standard")
        source = MODULE_SOURCE.read_text(encoding="utf-8")
        if lib.hasByName("Emergency"):
            lib.replaceByName("Emergency", source)
        else:
            lib.insertByName("Emergency", source)

        print("Creating buttons ...")
        buttons_by_sheet = {
            SH_ORDERS: [("Провести всё", "Emergency.Orders_ConductAllButton")],
            SH_OFFICE: [("Провести всё", "Emergency.Office_ConductAllButton")],
            SH_PROD: [("Провести всё", "Emergency.Production_ConductAllButton")],
            SH_PARTS: [("Провести всё", "Emergency.Parts_ConductAllButton")],
            SH_ISSUE: [("Провести всё", "Emergency.Issues_ConductAllButton")],
            SH_STOCK: [("Обновить", "Emergency.Stock_RefreshButton")],
        }
        for sheet_name, buttons in buttons_by_sheet.items():
            sheet = sheets.getByName(sheet_name)
            draw_page = sheet.DrawPage
            form = doc.createInstance("com.sun.star.form.component.Form")
            form.Name = "EmergencyForm"
            draw_page.Forms.insertByIndex(0, form)

            x_cursor = 200  # 2mm
            for idx, (caption, macro) in enumerate(buttons):
                ctrl_name = f"EMG_BTN_{sheet_name}_{idx}"
                model = doc.createInstance("com.sun.star.form.component.CommandButton")
                model.Name = ctrl_name
                model.Label = caption
                form.insertByName(ctrl_name, model)
                ctrl_idx = form.Count - 1

                shape = doc.createInstance("com.sun.star.drawing.ControlShape")
                shape.Control = model
                shape.Size = uno.createUnoStruct("com.sun.star.awt.Size", 3500, 800)
                shape.Position = uno.createUnoStruct("com.sun.star.awt.Point", x_cursor, 3500)
                draw_page.add(shape)

                desc = ScriptEventDescriptor()
                desc.ListenerType = "XActionListener"
                desc.EventMethod = "actionPerformed"
                desc.ScriptType = "Script"
                desc.ScriptCode = script_uri(macro)
                form.registerScriptEvent(ctrl_idx, desc)

                x_cursor += 3500 + 200

        # Активный лист по открытию - Заказы, первый по порядку.
        doc.CurrentController.setActiveSheet(sheets.getByName(SH_ORDERS))

        print(f"Saving to {output} ...")
        output.parent.mkdir(parents=True, exist_ok=True)
        save_args = (make_prop("FilterName", "calc8"), make_prop("Overwrite", True))
        doc.storeToURL(uno.systemPathToFileUrl(str(output.resolve())), save_args)
        doc.close(False)
        print(f"Done: {output}")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
        kill_stale_soffice(profile_dir)


def main():
    parser = argparse.ArgumentParser(description="Build the PRIME EMERGENCY working version .ods")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--port", type=int, default=2810)
    parser.add_argument("--profile-dir", type=Path, default=Path("/tmp/prime_emergency_build_profile"))
    args = parser.parse_args()
    build(args.output, args.port, args.profile_dir)


if __name__ == "__main__":
    sys.exit(main())
