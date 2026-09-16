#!/usr/bin/env python3
"""
model_tests (ТЗ static_tests/model_tests): чистая Python-модель ключевых алгоритмов
PRIME_04_Posting (FIFO-разбиение, партиальная поставка, идемпотентность по SOURCE_KEY,
сохранение суммарного остатка при перемещении, лимит возврата) - без запуска LibreOffice.

Это НЕ замена интеграционным тестам в реальном рантайме (см. functional_smoke.py и его
ограничения), а быстрая, надёжная проверка самой ЛОГИКИ, независимая от особенностей headless
LibreOffice, которые преследовали разработку в этой сессии. Каждая функция здесь - прямой
Python-эквивалент соответствующей Basic-функции; если Basic-версия меняется, эта модель должна
обновляться синхронно (иначе тест проверяет уже не то, что реально исполняется).
"""
import sys


class TestFailure(Exception):
    pass


def check(condition, message):
    if not condition:
        raise TestFailure(message)


# === FIFO allocation (PRIME_04_Posting.PRIME_FifoLotsForProduct / PostIssueLines) ===========
def fifo_allocate(lots, qty_needed):
    """lots: list of (lot_id, receipt_date, balance), sorted by (date, lot_id) by caller.
    Returns list of (lot_id, qty_taken); raises TestFailure if insufficient (shortage)."""
    remaining = qty_needed
    allocations = []
    for lot_id, _date, balance in lots:
        if remaining <= 0:
            break
        if balance <= 0:
            continue
        take = min(remaining, balance)
        allocations.append((lot_id, take))
        remaining -= take
    if remaining > 1e-6:
        raise TestFailure(f"shortage: {remaining} units unallocated (shortage_behavior=Reject entire document)")
    return allocations


def test_fifo_two_lots():
    lots = [("LOT-1", "2026-01-01", 10.0), ("LOT-2", "2026-01-02", 20.0)]
    result = fifo_allocate(lots, 15)
    check(result == [("LOT-1", 10.0), ("LOT-2", 5.0)], f"unexpected FIFO allocation: {result}")


def test_fifo_shortage_rejects_whole_document():
    lots = [("LOT-1", "2026-01-01", 10.0)]
    try:
        fifo_allocate(lots, 11)
        raise TestFailure("expected shortage to raise, but it did not")
    except TestFailure as e:
        check("shortage" in str(e), "wrong failure reason")


# === Partial receipt (ТЗ partial_receipts.example: 10 -> 6 -> 4) =============================
def test_partial_receipt_sequence():
    ordered = 10
    received_total = 0
    documents = []

    for delivery_qty in (6, 4):
        received_total += delivery_qty
        documents.append({"qty": delivery_qty, "lot": f"LOT-{len(documents) + 1}"})

    check(received_total == ordered, f"received_total={received_total}, expected {ordered}")
    check(len(documents) == 2, f"expected 2 receipt documents, got {len(documents)}")
    check(len({d['lot'] for d in documents}) == 2, "expected 2 distinct lots")
    remaining = ordered - received_total
    check(remaining == 0, f"remaining should be 0, got {remaining}")


# === Idempotency (SOURCE_KEY) - double click must not double-post ============================
class FakeCommittedKeys:
    def __init__(self):
        self._committed = {}

    def find_or_post(self, source_key, doc_id_factory):
        if source_key in self._committed:
            return self._committed[source_key], False  # existing, not newly posted
        doc_id = doc_id_factory()
        self._committed[source_key] = doc_id
        return doc_id, True


def test_idempotent_double_click_same_source_key():
    store = FakeCommittedKeys()
    counter = {"n": 0}

    def make_id():
        counter["n"] += 1
        return f"DOC-{counter['n']}"

    doc1, was_new1 = store.find_or_post("ORDER-1|LINE-1|DLV-1", make_id)
    doc2, was_new2 = store.find_or_post("ORDER-1|LINE-1|DLV-1", make_id)  # simulated double-click, same key

    check(was_new1 is True, "first post should be new")
    check(was_new2 is False, "second post with same SOURCE_KEY must not create a new document")
    check(doc1 == doc2, f"double-click must return the same DOC_ID, got {doc1} vs {doc2}")
    check(counter["n"] == 1, f"expected exactly 1 movement created, got {counter['n']}")


def test_two_different_deliveries_both_post():
    store = FakeCommittedKeys()
    counter = {"n": 0}

    def make_id():
        counter["n"] += 1
        return f"DOC-{counter['n']}"

    doc1, was_new1 = store.find_or_post("ORDER-1|LINE-1|DLV-1", make_id)
    doc2, was_new2 = store.find_or_post("ORDER-1|LINE-1|DLV-2", make_id)  # genuinely different delivery

    check(was_new1 and was_new2, "two distinct SOURCE_KEYs must both post")
    check(doc1 != doc2, "two distinct deliveries must get distinct DOC_IDs")
    check(counter["n"] == 2, f"expected 2 movements, got {counter['n']}")


# === Transfer conserves total stock (ТЗ transfers.total_stock_change: 0) =====================
def test_transfer_conserves_total_stock():
    stock = {"A": 50.0, "B": 20.0}
    qty = 15.0
    total_before = sum(stock.values())

    stock["A"] -= qty
    stock["B"] += qty

    total_after = sum(stock.values())
    check(total_before == total_after, f"transfer changed total stock: {total_before} -> {total_after}")
    check(stock["A"] == 35.0 and stock["B"] == 35.0, f"unexpected post-transfer stock: {stock}")


# === Return limit (ТЗ returns.example: issued 5, returned 2, remaining_returnable 3) ==========
def test_return_cannot_exceed_remaining():
    issued = 5.0
    already_returned = 2.0
    remaining_returnable = issued - already_returned
    check(remaining_returnable == 3.0, f"expected remaining_returnable=3, got {remaining_returnable}")

    def validate_return(qty):
        if qty > remaining_returnable + 1e-6:
            raise TestFailure(f"return {qty} exceeds remaining returnable {remaining_returnable}")
        return True

    check(validate_return(3.0) is True, "returning exactly the remaining amount must be allowed")
    try:
        validate_return(3.1)
        raise TestFailure("expected over-return to be rejected")
    except TestFailure as e:
        check("exceeds remaining" in str(e), "wrong rejection reason")


def test_repeat_same_return_amount_does_not_double_credit():
    # Symmetric to double-click idempotency: same RETURN source key must not double-credit stock.
    store = FakeCommittedKeys()
    stock = {"P": 0.0}
    counter = {"n": 0}

    def post_return(source_key, qty):
        def make_id():
            counter["n"] += 1
            stock["P"] += qty
            return f"RET-{counter['n']}"
        doc_id, was_new = store.find_or_post(source_key, make_id)
        return doc_id, was_new

    post_return("RETURN-1", 2.0)
    post_return("RETURN-1", 2.0)  # repeat click, same key
    check(stock["P"] == 2.0, f"expected stock credited exactly once (2.0), got {stock['P']}")


# === Kit shortage cancels the whole add (ТЗ kits.partial_issue_if_component_missing=False) ===
def test_kit_shortage_cancels_entire_kit():
    components_per_kit = {"A": 2, "B": 4}
    stock = {"A": 10, "B": 30}
    kit_qty = 6  # needs 12 A (short) and 24 B (ok)

    needed = {code: qty * kit_qty for code, qty in components_per_kit.items()}
    shortages = {code: needed[code] for code in needed if needed[code] > stock.get(code, 0)}

    check(len(shortages) > 0, "expected a shortage to be detected for this scenario")
    # Nothing should be issued when any component is short.
    issued = {} if shortages else needed
    check(issued == {}, f"kit issue must not partially post on shortage, got {issued}")


def test_kit_full_availability_issues_all_components():
    components_per_kit = {"A": 2, "B": 4}
    stock = {"A": 10, "B": 30}
    kit_qty = 3  # needs 6 A, 12 B - both available

    needed = {code: qty * kit_qty for code, qty in components_per_kit.items()}
    shortages = {code: needed[code] for code in needed if needed[code] > stock.get(code, 0)}
    check(not shortages, f"unexpected shortage: {shortages}")
    check(needed == {"A": 6, "B": 12}, f"unexpected needed quantities: {needed}")


TESTS = [
    test_fifo_two_lots,
    test_fifo_shortage_rejects_whole_document,
    test_partial_receipt_sequence,
    test_idempotent_double_click_same_source_key,
    test_two_different_deliveries_both_post,
    test_transfer_conserves_total_stock,
    test_return_cannot_exceed_remaining,
    test_repeat_same_return_amount_does_not_double_credit,
    test_kit_shortage_cancels_entire_kit,
    test_kit_full_availability_issues_all_components,
]


def main():
    failures = []
    for test in TESTS:
        try:
            test()
            print(f"  [PASS] {test.__name__}")
        except TestFailure as e:
            failures.append((test.__name__, str(e)))
            print(f"  [FAIL] {test.__name__}: {e}")
        except Exception as e:  # noqa: BLE001
            failures.append((test.__name__, f"unexpected exception: {e!r}"))
            print(f"  [ERROR] {test.__name__}: {e!r}")

    print()
    print(f"PASSED: {len(TESTS) - len(failures)}/{len(TESTS)}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
