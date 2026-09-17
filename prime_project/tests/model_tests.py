#!/usr/bin/env python3
"""
model_tests (ТЗ static_tests/model_tests): чистая Python-модель ключевых алгоритмов
PRIME_04_Posting (FIFO-разбиение, партиальная поставка, идемпотентность по SOURCE_KEY,
сохранение суммарного остатка при перемещении, лимит возврата) - без запуска LibreOffice.
С версии 2.0.1 также покрывает конкретные дефекты 2.0.0, найденные и исправленные в
PRIME_01_Runtime/PRIME_02_Store/PRIME_04_Posting: операционный лок, event guard, фильтр
"только COMMITTED влияет на остаток", откат SOURCE_KEY-кэша после неудачного store().

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


# === Operation lock (PRIME_01_Runtime.PRIME_TryEnter/PRIME_Leave) - PRIME 2.0.1 regression ===
# 2.0.0 bug: PRIME_TryEnter always incremented a depth counter and always returned True - it
# never actually blocked re-entry. Model here as the fixed boolean-mutex semantics.
class OperationLock:
    def __init__(self):
        self._locked = False

    def try_enter(self) -> bool:
        if self._locked:
            return False
        self._locked = True
        return True

    def leave(self):
        self._locked = False


def test_operation_lock_rejects_reentry():
    lock = OperationLock()
    check(lock.try_enter() is True, "first entry must succeed")
    check(lock.try_enter() is False, "re-entry while locked must be rejected (double-click protection)")
    lock.leave()
    check(lock.try_enter() is True, "entry after leave() must succeed again")


def test_post_document_rejects_when_already_locked():
    # Models PRIME_PostDocument's new guard: it must check PRIME_TryEnter()'s result and bail
    # out WITHOUT writing anything if the lock is already held - not just acquire-and-ignore.
    lock = OperationLock()
    movements = []

    def post_document(qty):
        if not lock.try_enter():
            return None  # PRIME_LastPostError(), no movements written
        try:
            movements.append(qty)
            return f"DOC-{len(movements)}"
        finally:
            lock.leave()

    lock.try_enter()  # simulate an operation already in progress
    result = post_document(10)
    check(result is None, "PostDocument must refuse to run while the lock is held")
    check(movements == [], "no movement may be written when the operation is rejected as reentrant")


# === Event guard (PRIME_01_Runtime.PRIME_EventEnter/PRIME_EventLeave) - PRIME 2.0.1 regression ===
# 2.0.0 bug: on a blocked (nested) call, PRIME_EventEnter both incremented depth AND returned
# False. Every caller uses "If Not PRIME_EventEnter() Then Exit Sub" - so on False it NEVER
# calls PRIME_EventLeave, meaning the depth bump on the blocked path was never undone: the
# guard stayed permanently stuck after the first nested event (e.g. programmatic autofill
# inside a ContentChanged handler that itself raises ContentChanged).
class EventGuard:
    def __init__(self):
        self._locked = False

    def enter(self) -> bool:
        if self._locked:
            return False  # depth/state must NOT change on the blocked path
        self._locked = True
        return True

    def leave(self):
        self._locked = False


def test_event_guard_symmetric_after_blocked_nested_call():
    guard = EventGuard()

    # Outer user edit triggers ContentChanged.
    check(guard.enter() is True, "outer event must acquire the guard")
    # Handler autofills a cell, which programmatically raises a nested ContentChanged.
    nested_entered = guard.enter()
    check(nested_entered is False, "nested event while guard held must be rejected")
    # Per the real caller pattern, a rejected Enter is never paired with a Leave call.
    guard.leave()  # only the OUTER handler's cleanup runs
    # The guard must now be fully released - not stuck "in use" because of the nested attempt.
    check(guard.enter() is True, "guard must be free for the next real user edit after cleanup")


# === Committed-only stock (PRIME_04_Posting.PRIME_LotBalance et al.) - PRIME 2.0.1 regression ===
# 2.0.0 bug: stock/lot-balance summed EVERY row in DB_PRIME_MOVEMENTS regardless of the owning
# transaction's SYS_PRIME_TX.STATE. Movements are physically written to the sheet before the
# transaction is marked COMMITTED (see PRIME_PostDocument step order), so a crash or a
# store() failure between those steps left PREPARED/FAILED movements that still counted
# toward stock. Fix: filter by PRIME_IsOpIdCommitted(OP_ID).
def committed_only_balance(movements, committed_op_ids):
    return sum(qty for (op_id, qty) in movements if op_id in committed_op_ids)


def test_committed_only_stock_excludes_prepared_and_failed():
    committed_op_ids = {"OP-1"}
    movements_prepared_only = [("OP-2", 10.0)]
    movements_failed_only = [("OP-3", 10.0)]
    movements_committed_only = [("OP-1", 10.0)]

    check(committed_only_balance(movements_prepared_only, committed_op_ids) == 0,
          "a PREPARED (not yet committed) movement must not affect stock")
    check(committed_only_balance(movements_failed_only, committed_op_ids) == 0,
          "a FAILED movement must not affect stock")
    check(committed_only_balance(movements_committed_only, committed_op_ids) == 10.0,
          "a COMMITTED movement must affect stock")


# === Transaction/cache consistency (PRIME_02_Store SOURCE_KEY cache) - PRIME 2.0.1 regression ===
# 2.0.0 bug: PRIME_PostDocument marked SYS_PRIME_TX COMMITTED and registered the SOURCE_KEY as
# committed in memory BEFORE calling ThisComponent.store(). If store() then failed, the TX row
# was correctly downgraded to FAILED, but the in-memory SOURCE_KEY cache was never rolled back -
# so a retry of the same SOURCE_KEY was told "already posted" and silently did nothing, even
# though the operation was actually FAILED and never really went through.
def test_retry_after_failed_store_is_not_treated_as_already_posted():
    committed_cache = {}  # SOURCE_KEY -> DOC_ID, mirrors gCommittedKeyCache
    tx_state = {}         # OP_ID -> STATE

    def post_document(source_key, store_should_fail):
        if source_key in committed_cache:
            return committed_cache[source_key], False  # "already posted"
        op_id = f"OP-{len(tx_state) + 1}"
        tx_state[op_id] = "PREPARED"
        doc_id = f"DOC-{len(tx_state)}"
        tx_state[op_id] = "COMMITTED"
        committed_cache[source_key] = doc_id
        try:
            if store_should_fail:
                raise IOError("simulated store() failure")
        except IOError:
            tx_state[op_id] = "FAILED"
            del committed_cache[source_key]  # the 2.0.1 fix: PRIME_UnregisterCommittedKey
            return None, False
        return doc_id, True

    doc_id, was_new = post_document("KEY-1", store_should_fail=True)
    check(doc_id is None, "a failed store() must not report a successful DOC_ID")
    check("KEY-1" not in committed_cache, "cache must be rolled back after a failed store()")

    # Retry with the same SOURCE_KEY must be treated as a fresh attempt, not "already posted".
    doc_id2, was_new2 = post_document("KEY-1", store_should_fail=False)
    check(doc_id2 is not None, "retry after a rolled-back failure must be allowed to actually post")
    check(was_new2 is True, "retry after a rolled-back failure must not be reported as a no-op repeat")


# === Flexible date parsing (PRIME_00_Config.PRIME_ParseFlexibleDate) - PRIME 2.0.1 addition ===
def parse_flexible_date(text, today_year=2026):
    s = text.strip()
    if not s:
        return ""
    for sep in (".", "/", "-"):
        if sep in s:
            parts = s.split(sep)
            break
    else:
        return ""
    if len(parts) < 2 or not parts[0].isdigit() or not parts[1].isdigit():
        return ""
    day, month = int(parts[0]), int(parts[1])
    if len(parts) >= 3 and parts[2].isdigit():
        year = int(parts[2])
        if year < 100:
            year += 2000
    else:
        year = today_year
    if not (1 <= day <= 31) or not (1 <= month <= 12):
        return ""
    return f"{year:04d}-{month:02d}-{day:02d}"


def test_flexible_date_parsing_accepted_formats():
    check(parse_flexible_date("24.08") == "2026-08-24", "dot, no year")
    check(parse_flexible_date("24/08") == "2026-08-24", "slash, no year")
    check(parse_flexible_date("24-08") == "2026-08-24", "dash, no year")
    check(parse_flexible_date("24.08.2026") == "2026-08-24", "dot, with year")
    check(parse_flexible_date("24/08/2026") == "2026-08-24", "slash, with year")
    check(parse_flexible_date("24-08-2026") == "2026-08-24", "dash, with year")


def test_flexible_date_parsing_rejects_garbage():
    check(parse_flexible_date("") == "", "empty input must not produce a fabricated date")
    check(parse_flexible_date("not a date") == "", "unparseable text must return empty, not guess")
    check(parse_flexible_date("99.99") == "", "out-of-range day/month must be rejected")


# === Order status logic (PRIME_05_Orders.PRIME_Orders_RecomputeStatus) - PRIME 2.0.1 addition ===
def compute_order_status(ordered, received, expected_date, today, current_status=""):
    if current_status == "Отменено":
        return current_status  # manual override, never recomputed
    if ordered is None:
        return "Черновик" if current_status == "" else current_status
    if received and received >= ordered - 0.0000005:
        status = "Получено"
    elif received:
        status = "Частично получено"
    else:
        status = "Ожидается"
    if status != "Получено" and expected_date and expected_date < today:
        status = "Просрочено"
    return status


def test_order_status_progression():
    check(compute_order_status(None, None, "", "2026-01-01") == "Черновик", "nothing entered yet")
    check(compute_order_status(10, 0, "", "2026-01-01") == "Ожидается", "ordered, nothing received")
    check(compute_order_status(10, 6, "", "2026-01-01") == "Частично получено", "partial receipt")
    check(compute_order_status(10, 10, "", "2026-01-01") == "Получено", "fully received")


def test_order_status_overdue_rules():
    check(compute_order_status(10, 0, "2026-01-01", "2026-06-01") == "Просрочено",
          "expected date in the past with nothing received must be overdue")
    check(compute_order_status(10, 10, "2026-01-01", "2026-06-01") == "Получено",
          "a fully received order must never become overdue, even with a past expected date")
    check(compute_order_status(10, 0, "", "2026-06-01") == "Ожидается",
          "an empty expected date must never create an overdue status")


def test_order_status_respects_manual_cancellation():
    check(compute_order_status(10, 0, "2026-01-01", "2026-06-01", current_status="Отменено") == "Отменено",
          "a manually cancelled order must never be overwritten by the automatic recompute")


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
    test_operation_lock_rejects_reentry,
    test_post_document_rejects_when_already_locked,
    test_event_guard_symmetric_after_blocked_nested_call,
    test_committed_only_stock_excludes_prepared_and_failed,
    test_retry_after_failed_store_is_not_treated_as_already_posted,
    test_flexible_date_parsing_accepted_formats,
    test_flexible_date_parsing_rejects_garbage,
    test_order_status_progression,
    test_order_status_overdue_rules,
    test_order_status_respects_manual_cancellation,
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
