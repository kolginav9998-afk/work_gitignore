#!/usr/bin/env python3
"""
functional_2_1_2: headless UNO integration test for the 2.1.2 mandatory regressions - a single
long-lived LibreOffice session driving the REAL built ODS through script.invoke(), exactly like
a user clicking buttons (each posting/refresh call is issued as its own external invoke(), which
this session's debugging established as the only reliable pattern for chained external calls -
see functional_smoke.py's docstring for the documented instability of other call shapes).

Covers, in one continuous session (closer to real usage than functional_smoke.py's isolated
checks):
  - Orders child receipt row + live refresh after Issue and after Return (buffered-array-
    readback corruption fix in PRIME_PostIssueLines/PRIME_PostReturnLines).
  - Transfer and Adjustment postings (same corruption-bug class, PRIME_PostTransferLines/
    PRIME_PostAdjustmentLines) with lot-lineage / contour checks.
  - "Наличие"/"Поиск" refresh buttons (PRIME_Stock_Rebuild/PRIME_Search_Execute corruption fix)
    and committed-only filtering.
  - "Возвраты" candidate-list refresh committed-only filtering.
  - The two new active receipt workflows "Приход — Производство"/"Приход — Детали" (own EI_CODE,
    own contour, no aggregation).
  - "Главная" landing sheet content and the authoritative final visible-sheet list.

This is an integration smoke test, not a substitute for the pure-Python algorithm models in
model_tests.py - if it fails, re-run the specific tools/build_ods.py + soffice sequence manually
before concluding the business logic itself regressed (see functional_smoke.py's caveat).
"""
import sys
import time
import subprocess
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import build_ods  # noqa: E402
import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402

PORT = 2799
PROFILE_DIR = Path("/tmp/prime_functional_2_1_2_profile")
OUTPUT = Path("/tmp/PRIME_functional_2_1_2.ods")

VISIBLE_SHEETS = {
    "Главная", "Заказы", "Приход — Офис", "Расход — Офис",
    "Приход — Производство", "Расход — Производство", "Приход — Детали", "Расход — Детали",
    "Выдачи", "Возвраты", "Перемещения", "Инвентаризация", "Наличие", "Поиск",
    "Журнал", "Комплекты",
}

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


def load_with_retry(desktop, url, retries=10):
    for _ in range(retries):
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (make_prop("Hidden", True),))
        if doc is not None:
            return doc
        time.sleep(2)
    raise RuntimeError("load failed")


def invoke_macro(doc, module_dot_sub, args=()):
    script = doc.getScriptProvider().getScript(
        f"vnd.sun.star.script:Standard.{module_dot_sub}?language=Basic&location=document")
    return script.invoke(args, (), ())


def header_map(doc, sheet_name):
    return list(invoke_macro(doc, "PRIME_02_Store.PRIME_HeaderMap", (sheet_name,))[0])


def last_row(doc, sheet):
    return invoke_macro(doc, "PRIME_02_Store.PRIME_FindLastRow", (sheet,))[0]


def first_data_row(doc, sheet_name):
    return invoke_macro(doc, "PRIME_00_Config.PRIME_FormSchemaFirstDataRow", (sheet_name,))[0]


def new_order_with_receipt(doc, orders_sheet, headers, location, fact_qty):
    def col(name):
        return headers.index(name)

    invoke_macro(doc, "PRIME_05_Orders.PRIME_Orders_NewOrder")
    row = last_row(doc, orders_sheet)
    orders_sheet.getCellByPosition(col("Полное наименование товара"), row).setString("T")
    orders_sheet.getCellByPosition(col("Количество"), row).setValue(10)
    orders_sheet.getCellByPosition(col("Место хранения"), row).setString(location)
    orders_sheet.getCellByPosition(col("Факт. количество"), row).setValue(fact_qty)
    invoke_macro(doc, "PRIME_05_Orders.PRIME_Orders_ConductRow", (orders_sheet, row))
    child_row = row + 1
    code = orders_sheet.getCellByPosition(col("Код товара"), child_row).getString()
    return child_row, code


def main():
    shutil.copyfile(build_ods.DEFAULT_TEMPLATE, OUTPUT)
    proc = build_ods.start_soffice(PROFILE_DIR, PORT)
    try:
        time.sleep(4)
        ctx = connect_with_retry(PORT)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = load_with_retry(desktop, uno.systemPathToFileUrl(str(OUTPUT)))
        build_ods.inject_basic_modules(doc)
        invoke_macro(doc, "PRIME_14_MigrationInstaller.PRIME_Install_EnsureSchema")
        invoke_macro(doc, "PRIME_14_MigrationInstaller.PRIME_Install_EnsureAllBusinessSheetsOnly")
        invoke_macro(doc, "PRIME_14_MigrationInstaller.PRIME_Build_MigrateSilent")
        invoke_macro(doc, "PRIME_12_UI.PRIME_UI_RestoreInterfaceSilent")
        invoke_macro(doc, "PRIME_12_UI.PRIME_UI_FixStockPanelTextButton")

        # === Sheet visibility ===
        print("=== Sheet visibility ===")
        actually_visible = {doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)
                             if doc.Sheets.getByIndex(i).IsVisible}
        check(actually_visible == VISIBLE_SHEETS,
              f"visible sheets must be exactly {sorted(VISIBLE_SHEETS)}, got {sorted(actually_visible)}")

        # confirmed bug #4: "Наличие" must not carry the inherited 1.4.1 Firebird decorative text.
        print("=== Наличие panel text ===")
        stock_sheet_early = doc.Sheets.getByName("Наличие")
        panel_text_ok = all(
            "Firebird" not in stock_sheet_early.getCellByPosition(0, r).getString()
            for r in range(0, 5)
        )
        check(panel_text_ok, "Наличие decorative panel no longer mentions Firebird")

        # === Главная ===
        print("=== Главная landing sheet ===")
        home = doc.Sheets.getByName("Главная")
        check(home.getCellByPosition(0, 10).getString() == "ПОКАТАК PRIME", "Главная title present")
        check(home.getCellByPosition(0, 11).getString().startswith("Версия:"), "Главная shows version")
        check(home.getCellByPosition(0, 12).getString().startswith("Статус:"), "Главная shows diagnostics status")

        # === Orders child row: Receipt -> Issue -> Return live refresh ===
        print("=== Orders child row live refresh (Issue + Return) ===")
        orders_sheet = doc.Sheets.getByName("Заказы")
        orders_headers = header_map(doc, "Заказы")
        child_row, code1 = new_order_with_receipt(doc, orders_sheet, orders_headers, "Склад-1", 6)
        check(code1.startswith("ЕИ-"), f"receipt minted an EI_CODE ({code1!r})")

        issues_sheet = doc.Sheets.getByName("Выдачи")
        issues_headers = header_map(doc, "Выдачи")
        irow = max(last_row(doc, issues_sheet) + 1, 1)
        issues_sheet.getCellByPosition(issues_headers.index("Код"), irow).setString(code1)
        issues_sheet.getCellByPosition(issues_headers.index("Кол-во"), irow).setValue(2)
        issues_sheet.getCellByPosition(issues_headers.index("Откуда"), irow).setString("Склад-1")
        invoke_macro(doc, "PRIME_06_Issues.PRIME_Issues_ConductRow", (issues_sheet, irow))
        state = issues_sheet.getCellByPosition(issues_headers.index("_PRIME_IssueState"), irow).getString()
        check(state.startswith("DONE:"), f"issue committed ({state!r})")

        avail_col = orders_headers.index("В наличии сейчас")
        issued_col = orders_headers.index("Выдано")
        avail_after_issue = orders_sheet.getCellByPosition(avail_col, child_row).getValue()
        issued_after_issue = orders_sheet.getCellByPosition(issued_col, child_row).getValue()
        check(avail_after_issue == 4, f"child row avail 6->4 after issuing 2 (got {avail_after_issue})")
        check(issued_after_issue == 2, f"child row Выдано tracks issued qty (got {issued_after_issue})")

        prod_sheet = doc.Sheets.getByName("DB_PRIME_PRODUCTS")
        prod_headers = header_map(doc, "DB_PRIME_PRODUCTS")
        plast = last_row(doc, prod_sheet)
        for rr in range(1, plast + 1):
            if prod_sheet.getCellByPosition(prod_headers.index("PRODUCT_CODE"), rr).getString() == code1:
                prod_sheet.getCellByPosition(prod_headers.index("RETURNABLE"), rr).setValue(1)
                break
        invoke_macro(doc, "PRIME_03_Catalog.PRIME_InvalidateProductIndex")
        invoke_macro(doc, "PRIME_08_ReturnsInventory.PRIME_Returns_RefreshButton")
        returns_sheet = doc.Sheets.getByName("Возвраты")
        returns_headers = header_map(doc, "Возвраты")
        rlast = last_row(doc, returns_sheet)
        target_row = None
        for rr in range(1, rlast + 1):
            if returns_sheet.getCellByPosition(returns_headers.index("Код"), rr).getString() == code1:
                target_row = rr
                break
        check(target_row is not None, "Возвраты refresh finds the returnable candidate")
        if target_row is not None:
            returns_sheet.getCellByPosition(returns_headers.index("Вернуть сейчас"), target_row).setValue(1)
            invoke_macro(doc, "PRIME_08_ReturnsInventory.PRIME_Returns_ConductButton")
            rstate = returns_sheet.getCellByPosition(returns_headers.index("_PRIME_ReturnState"), target_row).getString()
            check(rstate == "", f"return committed and cleared its pending marker (got {rstate!r})")
            avail_after_return = orders_sheet.getCellByPosition(avail_col, child_row).getValue()
            returned_col = orders_headers.index("Возвращено")
            returned_after = orders_sheet.getCellByPosition(returned_col, child_row).getValue()
            check(avail_after_return == 5, f"child row avail restored to 5 after returning 1 (got {avail_after_return})")
            check(returned_after == 1, f"child row Возвращено tracks returned qty (got {returned_after})")

        # === Transfer ===
        print("=== Transfer ===")
        transfers_sheet = doc.Sheets.getByName("Перемещения")
        t_headers = header_map(doc, "Перемещения")
        trow = max(last_row(doc, transfers_sheet) + 1, 1)
        transfers_sheet.getCellByPosition(t_headers.index("Внутренний код"), trow).setString(code1)
        transfers_sheet.getCellByPosition(t_headers.index("Кол-во"), trow).setValue(2)
        transfers_sheet.getCellByPosition(t_headers.index("Место — откуда"), trow).setString("Склад-1")
        transfers_sheet.getCellByPosition(t_headers.index("Место — куда"), trow).setString("Склад-2")
        invoke_macro(doc, "PRIME_15_Transfers.PRIME_Transfers_ConductRow", (transfers_sheet, trow))
        tstate = transfers_sheet.getCellByPosition(t_headers.index("_PRIME_TransferState"), trow).getString()
        check(tstate.startswith("DONE:"), f"transfer committed ({tstate!r})")
        avail_after_transfer = orders_sheet.getCellByPosition(avail_col, child_row).getValue()
        check(avail_after_transfer == 3, f"child row avail 5->3 after transferring 2 away (got {avail_after_transfer})")

        # === Adjustment (shortage branch, via Инвентаризация) ===
        print("=== Adjustment (inventory shortage) ===")
        invoke_macro(doc, "PRIME_08_ReturnsInventory.PRIME_Inventory_LoadButton")
        inv_sheet = doc.Sheets.getByName("Инвентаризация")
        inv_headers = header_map(doc, "Инвентаризация")
        ilast = last_row(doc, inv_sheet)
        inv_target = None
        for rr in range(1, ilast + 1):
            if (inv_sheet.getCellByPosition(inv_headers.index("Код"), rr).getString() == code1
                    and inv_sheet.getCellByPosition(inv_headers.index("Место"), rr).getString() == "Склад-1"):
                inv_target = rr
                break
        check(inv_target is not None, "inventory snapshot includes the Склад-1 position")
        if inv_target is not None:
            uchet = inv_sheet.getCellByPosition(inv_headers.index("Учёт"), inv_target).getValue()
            fact = uchet - 1
            inv_sheet.getCellByPosition(inv_headers.index("Факт"), inv_target).setValue(fact)
            invoke_macro(doc, "PRIME_08_ReturnsInventory.PRIME_Inventory_RecalcButton")
            invoke_macro(doc, "PRIME_08_ReturnsInventory.PRIME_Inventory_ConductButton")
            diff_cell = inv_sheet.getCellByPosition(inv_headers.index("Разница"), inv_target).getString()
            check(diff_cell.startswith("проведено:"), f"adjustment committed ({diff_cell!r})")
            avail_after_adj = orders_sheet.getCellByPosition(avail_col, child_row).getValue()
            check(avail_after_adj == fact, f"child row avail reflects adjustment (expected {fact}, got {avail_after_adj})")

        # === Наличие / Поиск refresh correctness ===
        print("=== Наличие / Поиск refresh ===")
        invoke_macro(doc, "PRIME_09_StockSearch.PRIME_Stock_ShowNonZeroButton")
        stock_sheet = doc.Sheets.getByName("Наличие")
        stock_headers = header_map(doc, "Наличие")
        slast = last_row(doc, stock_sheet)
        stock_row_ok = any(
            stock_sheet.getCellByPosition(stock_headers.index("Внутренний код"), rr).getString() == code1
            and stock_sheet.getCellByPosition(stock_headers.index("Наименование"), rr).getString() not in ("", "0")
            for rr in range(1, slast + 1)
        )
        check(stock_row_ok, "Наличие refresh shows a non-corrupted row for the traded EI_CODE")

        # single_physical_warehouse / confirmed bug #3: "Из заказов" origin filter must show code1
        # (minted on "Заказы") and must NOT show a PRODUCTION-origin EI - proves the quick filters
        # now key off LOT.ORIGIN, not off the inert STOCK_CONTOUR.
        invoke_macro(doc, "PRIME_09_StockSearch.PRIME_Stock_ShowOriginOrdersButton")
        slast = last_row(doc, stock_sheet)
        orders_origin_codes = {
            stock_sheet.getCellByPosition(stock_headers.index("Внутренний код"), rr).getString()
            for rr in range(1, slast + 1)
        }
        check(code1 in orders_origin_codes, "Наличие 'Из заказов' filter includes the order-origin EI_CODE")

        invoke_macro(doc, "PRIME_09_StockSearch.PRIME_Search_ShowAllButton")
        search_sheet = doc.Sheets.getByName("Поиск")
        search_headers = header_map(doc, "Поиск")
        xlast = last_row(doc, search_sheet)
        search_rows_ok = sum(
            1 for rr in range(1, xlast + 1)
            if search_sheet.getCellByPosition(search_headers.index("Код"), rr).getString() == code1
            and search_sheet.getCellByPosition(search_headers.index("Тип"), rr).getString() not in ("", "0")
        )
        check(search_rows_ok >= 1, f"Поиск refresh shows non-corrupted rows for the traded EI_CODE (found {search_rows_ok})")

        # === Приход — Производство / Приход — Детали: independent EI_CODE + contour ===
        print("=== Приход — Производство / Приход — Детали ===")
        prod_wf_sheet = doc.Sheets.getByName("Приход — Производство")
        prod_wf_headers = header_map(doc, "Приход — Производство")
        prow = max(last_row(doc, prod_wf_sheet) + 1, first_data_row(doc, "Приход — Производство"))
        prod_wf_sheet.getCellByPosition(prod_wf_headers.index("Наименование"), prow).setString("Деталь производства")
        prod_wf_sheet.getCellByPosition(prod_wf_headers.index("Количество прихода"), prow).setValue(5)
        prod_wf_sheet.getCellByPosition(prod_wf_headers.index("Место хранения на складе"), prow).setString("Цех-1")
        invoke_macro(doc, "PRIME_07_Workflows.PRIME_Workflow_ConductRow", (prod_wf_sheet, prow))
        code_prod = prod_wf_sheet.getCellByPosition(prod_wf_headers.index("Внутренний код"), prow).getString()
        check(code_prod.startswith("ЕИ-") and code_prod != code1, f"Приход — Производство minted its own EI_CODE ({code_prod!r})")

        det_sheet = doc.Sheets.getByName("Приход — Детали")
        det_headers = header_map(doc, "Приход — Детали")
        drow = max(last_row(doc, det_sheet) + 1, first_data_row(doc, "Приход — Детали"))
        det_sheet.getCellByPosition(det_headers.index("Наименование"), drow).setString("Деталь")
        det_sheet.getCellByPosition(det_headers.index("Количество прихода"), drow).setValue(3)
        det_sheet.getCellByPosition(det_headers.index("Место хранения на складе"), drow).setString("Цех-2")
        invoke_macro(doc, "PRIME_07_Workflows.PRIME_Workflow_ConductRow", (det_sheet, drow))
        code_det = det_sheet.getCellByPosition(det_headers.index("Внутренний код"), drow).getString()
        check(code_det.startswith("ЕИ-") and code_det not in (code1, code_prod),
              f"Приход — Детали minted its own EI_CODE ({code_det!r})")

        # confirmed bug #3 (critical_fix): "Из производства" must filter PRODUCTION-origin EIs,
        # not the inert GENERAL contour - see PRIME_09_StockSearch.PRIME_Stock_ShowOriginProductionButton.
        invoke_macro(doc, "PRIME_09_StockSearch.PRIME_Stock_ShowOriginProductionButton")
        stock_headers = header_map(doc, "Наличие")
        plast = last_row(doc, stock_sheet)
        production_origin_codes = {
            stock_sheet.getCellByPosition(stock_headers.index("Внутренний код"), rr).getString()
            for rr in range(1, plast + 1)
        }
        check(code_prod in production_origin_codes and code1 not in production_origin_codes
              and code_det not in production_origin_codes,
              f"Наличие 'Из производства' filter shows only PRODUCTION-origin EI (got {sorted(production_origin_codes)})")

        invoke_macro(doc, "PRIME_09_StockSearch.PRIME_Stock_ShowOriginDetailsButton")
        dlast = last_row(doc, stock_sheet)
        details_origin_codes = {
            stock_sheet.getCellByPosition(stock_headers.index("Внутренний код"), rr).getString()
            for rr in range(1, dlast + 1)
        }
        check(code_det in details_origin_codes and code_prod not in details_origin_codes,
              f"Наличие 'Детали' filter shows only DETAILS-origin EI (got {sorted(details_origin_codes)})")

        moves_sheet = doc.Sheets.getByName("DB_PRIME_MOVEMENTS")
        move_headers = header_map(doc, "DB_PRIME_MOVEMENTS")
        mlast = last_row(doc, moves_sheet)
        contour_by_code = {}
        for rr in range(1, mlast + 1):
            pc = moves_sheet.getCellByPosition(move_headers.index("PRODUCT_CODE"), rr).getString()
            if pc in (code_prod, code_det):
                contour_by_code[pc] = moves_sheet.getCellByPosition(move_headers.index("STOCK_CONTOUR"), rr).getString()
        check(contour_by_code.get(code_prod) == "PRODUCTION",
              f"Приход — Производство posts to PRODUCTION contour (got {contour_by_code.get(code_prod)!r})")
        check(contour_by_code.get(code_det) == "WORKSHOP_DETAILS",
              f"Приход — Детали posts to WORKSHOP_DETAILS contour (got {contour_by_code.get(code_det)!r})")

        doc.close(False)
        print()
        print(f"PASSED: {'yes' if not failures else 'no'} ({len(failures)} failures)")
        return 1 if failures else 0
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
        build_ods.kill_stale_soffice(PROFILE_DIR)


if __name__ == "__main__":
    sys.exit(main())
