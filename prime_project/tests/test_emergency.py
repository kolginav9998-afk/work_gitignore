#!/usr/bin/env python3
"""
test_emergency.py — headless UNO acceptance test for the PRIME EMERGENCY WORKING VERSION
(tools/build_emergency_ods.py / src/basic_emergency/Emergency.bas), reproducing the user's own
literal 9-point acceptance scenario verbatim (see chat: "EMERGENCY WORKING VERSION — НУЖНА
СЕГОДНЯ", point 9), including the mandatory save/close/reopen check.
"""
import sys
import time
import subprocess
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import build_emergency_ods as builder  # noqa: E402
import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402

PORT = 2811
PROFILE_DIR = Path("/tmp/prime_emergency_test_profile")
OUTPUT = Path("/tmp/PRIME_EMERGENCY_functional_test.ods")

failures = []


def check(condition, message):
    if not condition:
        failures.append(message)
        print(f"  [FAIL] {message}")
    else:
        print(f"  [OK] {message}")


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect_with_retry(port, retries=30):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    for _ in range(retries):
        try:
            return resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(1)
    raise RuntimeError("connect failed")


def invoke_macro(doc, sub_name, args=()):
    script = doc.getScriptProvider().getScript(
        f"vnd.sun.star.script:Standard.Emergency.{sub_name}?language=Basic&location=document")
    return script.invoke(args, (), ())


def set_row(sheet, row, values):
    """values: dict col_index -> value (str or float)."""
    for col, val in values.items():
        cell = sheet.getCellByPosition(col, row)
        if isinstance(val, str):
            cell.setString(val)
        else:
            cell.setValue(val)


def main():
    builder.build(OUTPUT, PORT, PROFILE_DIR)

    proc = builder.start_soffice(PROFILE_DIR, PORT + 1)
    try:
        time.sleep(3)
        ctx = connect_with_retry(PORT + 1)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(OUTPUT.resolve())), "_blank", 0, (make_prop("Hidden", True),))

        orders = doc.Sheets.getByName("Заказы")
        office = doc.Sheets.getByName("Приход — Офис")
        prod = doc.Sheets.getByName("Приход — Производство")
        parts = doc.Sheets.getByName("Приход — Детали")
        issues = doc.Sheets.getByName("Выдачи")

        # === Заказы: Ручка 5 -> EI-A, Ручка 19 -> EI-B ===
        print("=== Заказы: два прихода одной ручки, разные ЕИ ===")
        set_row(orders, 1, {0: "З-1", 1: "Ручка", 3: 5.0, 4: "шт"})
        set_row(orders, 2, {0: "З-1", 1: "Ручка", 3: 19.0, 4: "шт"})
        invoke_macro(doc, "Orders_ConductAllButton")
        ei_a = orders.getCellByPosition(8, 1).getString()
        ei_b = orders.getCellByPosition(8, 2).getString()
        check(ei_a.startswith("ЕИ-") and ei_b.startswith("ЕИ-"), f"both rows got an EI code ({ei_a!r}, {ei_b!r})")
        check(ei_a != ei_b, f"EI-A != EI-B ({ei_a!r} vs {ei_b!r})")

        # === Выдачи EI-B 19 -> остаток EI-B = 0 ===
        print("=== Выдачи EI-B 19 -> остаток 0 ===")
        set_row(issues, 1, {0: "2026-01-01", 1: ei_b, 3: 19.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        status_b = issues.getCellByPosition(10, 1).getString()
        check(status_b == "Проведено", f"issue of EI-B fully posted (status={status_b!r})")
        bal_b = orders.getCellByPosition(9, 2).getValue()
        check(bal_b == 0, f"остаток EI-B стал 0 (got {bal_b})")

        # === Выдачи EI-A 1 -> остаток EI-A = 4 ===
        print("=== Выдачи EI-A 1 -> остаток 4 ===")
        set_row(issues, 2, {0: "2026-01-01", 1: ei_a, 3: 1.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        bal_a = orders.getCellByPosition(9, 1).getValue()
        check(bal_a == 4, f"остаток EI-A стал 4 (got {bal_a})")

        # === Приход — Офис 10 -> EI-C, выдача 3 -> остаток 7 ===
        print("=== Приход — Офис 10 -> EI-C, выдача 3 -> остаток 7 ===")
        set_row(office, 1, {1: "Бумага", 2: 10.0, 3: "упак", 5: "Склад-1"})
        invoke_macro(doc, "Office_ConductAllButton")
        ei_c = office.getCellByPosition(0, 1).getString()
        check(ei_c.startswith("ЕИ-") and ei_c not in (ei_a, ei_b), f"EI-C minted ({ei_c!r})")
        set_row(issues, 3, {0: "2026-01-01", 1: ei_c, 3: 3.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        bal_c = office.getCellByPosition(4, 1).getValue()
        check(bal_c == 7, f"Приход — Офис остаток стал 7 (got {bal_c})")

        # === Приход — Производство 20 -> EI-D, выдача 5 -> остаток 15 ===
        print("=== Приход — Производство 20 -> EI-D, выдача 5 -> остаток 15 ===")
        set_row(prod, 1, {1: "Краска", 2: 20.0, 3: "л", 5: "Цех-1"})
        invoke_macro(doc, "Production_ConductAllButton")
        ei_d = prod.getCellByPosition(0, 1).getString()
        check(ei_d.startswith("ЕИ-") and ei_d not in (ei_a, ei_b, ei_c), f"EI-D minted ({ei_d!r})")
        set_row(issues, 4, {0: "2026-01-01", 1: ei_d, 3: 5.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        bal_d = prod.getCellByPosition(4, 1).getValue()
        check(bal_d == 15, f"Приход — Производство остаток стал 15 (got {bal_d})")

        # === Приход — Детали 50 -> EI-E, выдача 8 -> остаток 42 ===
        print("=== Приход — Детали 50 -> EI-E, выдача 8 -> остаток 42 ===")
        set_row(parts, 1, {1: "Деталь", 2: "ART-1", 3: 50.0, 4: "шт", 6: "Цех-2"})
        invoke_macro(doc, "Parts_ConductAllButton")
        ei_e = parts.getCellByPosition(0, 1).getString()
        check(ei_e.startswith("ЕИ-") and ei_e not in (ei_a, ei_b, ei_c, ei_d), f"EI-E minted ({ei_e!r})")
        set_row(issues, 5, {0: "2026-01-01", 1: ei_e, 3: 8.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        bal_e = parts.getCellByPosition(5, 1).getValue()
        check(bal_e == 42, f"Приход — Детали остаток стал 42 (got {bal_e})")

        # === Попытка выдать больше остатка блокируется ===
        print("=== Попытка выдать больше остатка (EI-A: доступно 4, просим 5) ===")
        set_row(issues, 6, {0: "2026-01-01", 1: ei_a, 3: 5.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        status_over = issues.getCellByPosition(10, 6).getString()
        check("Недостаточно остатка" in status_over, f"over-issue blocked with a clear message (got {status_over!r})")
        bal_a_after = orders.getCellByPosition(9, 1).getValue()
        check(bal_a_after == 4, f"остаток EI-A НЕ изменился после блокировки (got {bal_a_after})")

        # === Идемпотентность: повторный клик "Провести всё" не портит уже проведённые строки ===
        print("=== Идемпотентность: повторный клик 'Провести всё' ===")
        invoke_macro(doc, "Orders_ConductAllButton")
        check(orders.getCellByPosition(8, 1).getString() == ei_a, "повторное 'Провести всё' (Заказы) не меняет уже присвоенный EI-A")
        check(orders.getCellByPosition(8, 2).getString() == ei_b, "повторное 'Провести всё' (Заказы) не меняет уже присвоенный EI-B")
        invoke_macro(doc, "Issues_ConductAllButton")
        check(orders.getCellByPosition(9, 1).getValue() == 4, "повторное 'Провести всё' (Выдачи) не меняет остаток EI-A (по-прежнему 4)")
        check(orders.getCellByPosition(9, 2).getValue() == 0, "повторное 'Провести всё' (Выдачи) не меняет остаток EI-B (по-прежнему 0)")

        # === Неизвестный ЕИ-код в "Выдачи" ===
        print("=== Неизвестный ЕИ-код блокируется с понятным сообщением ===")
        set_row(issues, 7, {0: "2026-01-01", 1: "ЕИ-99999999", 3: 1.0})
        invoke_macro(doc, "Issues_ConductAllButton")
        status_unknown = issues.getCellByPosition(10, 7).getString()
        check("не найден" in status_unknown, f"неизвестный ЕИ даёт понятное сообщение, не тихий сбой (got {status_unknown!r})")

        # === Проверка типов валидации данных (выпадающий список Ед.изм., числовое Количество) ===
        print("=== Валидация ячеек (Ед.изм. список, Количество только число) ===")
        unit_validation = issues.getCellRangeByPosition(4, 1, 4, 1).Validation
        check(unit_validation.Type == builder.VALIDATION_LIST, f"Ед.изм. на 'Выдачи' - выпадающий список (got type={unit_validation.Type})")
        qty_validation = issues.getCellRangeByPosition(3, 1, 3, 1).Validation
        check(qty_validation.Type == builder.VALIDATION_DECIMAL, f"Количество на 'Выдачи' - только число >= 0 (got type={qty_validation.Type})")

        # === Наличие: одна строка на каждый EI ===
        print("=== Наличие ===")
        stock = doc.Sheets.getByName("Наличие")
        invoke_macro(doc, "Stock_RefreshButton")
        stock_eis = {stock.getCellByPosition(0, r).getString() for r in range(1, 20)
                     if stock.getCellByPosition(0, r).getString() != ""}
        check({ei_a, ei_b, ei_c, ei_d, ei_e} <= stock_eis,
              f"Наличие содержит все 5 EI (got {sorted(stock_eis)})")

        doc.store()
        doc.close(False)

        # === Save -> close -> reopen: остатки сохраняются ===
        print("=== Save -> close -> reopen ===")
        doc2 = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(OUTPUT.resolve())), "_blank", 0, (make_prop("Hidden", True),))
        orders2 = doc2.Sheets.getByName("Заказы")
        office2 = doc2.Sheets.getByName("Приход — Офис")
        check(orders2.getCellByPosition(9, 1).getValue() == 4, "reopen: остаток EI-A сохранился (4)")
        check(orders2.getCellByPosition(9, 2).getValue() == 0, "reopen: остаток EI-B сохранился (0)")
        check(office2.getCellByPosition(4, 1).getValue() == 7, "reopen: остаток EI-C (Приход — Офис) сохранился (7)")
        doc2.close(False)

        print()
        print(f"PASSED: {'yes' if not failures else 'no'} ({len(failures)} failures)")
        return 1 if failures else 0
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
        builder.kill_stale_soffice(PROFILE_DIR)


if __name__ == "__main__":
    sys.exit(main())
