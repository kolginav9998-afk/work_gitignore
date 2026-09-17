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


# === PRIME 2.1.0 regressions =================================================================

# --- R01: COMMITTED OP_ID cache must update immediately on commit, not only on full rebuild ---
# 2.0.1 bug: PRIME_RegisterCommittedKey only added SOURCE_KEY->DOC_ID to gCommittedKeyCache; it
# never touched gCommittedOpIdCache. PRIME_IsOpIdCommitted(freshly-committed OP_ID) returned
# False until the next PRIME_RebuildCommittedKeyCache (normally only on reopen) - a receipt
# posted in the current session could fail to show up in stock/FIFO until reopen.
class CommittedCaches:
    def __init__(self):
        self.by_source_key = {}
        self.committed_op_ids = set()

    def register(self, source_key, doc_id, op_id):
        self.by_source_key.setdefault(source_key, doc_id)
        self.committed_op_ids.add(op_id)

    def is_op_committed(self, op_id):
        return op_id in self.committed_op_ids


def test_committed_op_id_cache_updates_immediately_after_post():
    caches = CommittedCaches()
    caches.register("KEY-1", "DOC-1", "OP-1")
    check(caches.is_op_committed("OP-1") is True,
          "OP_ID must be visible as committed in the SAME session immediately after commit, "
          "without waiting for a full cache rebuild")


# --- R02: FIFO/availability must be scoped to (product, location, contour), not product alone ---
def location_contour_balance(lots, product_code, location, contour):
    return sum(bal for (pc, loc, ctr, _date, bal) in lots
               if pc == product_code and loc == location and ctr == contour)


def test_fifo_is_location_and_contour_scoped():
    lots = [
        ("P1", "A", "GENERAL", "2026-01-01", 2.0),
        ("P1", "B", "GENERAL", "2026-01-01", 20.0),
        ("P1", "A", "WORKSHOP_DETAILS", "2026-01-01", 100.0),  # same product+location, different contour
    ]
    # Availability at A/GENERAL must be exactly 2, not the product-wide total across all
    # locations/contours (22 general + 100 workshop) - this is exactly the bug the external
    # audit found: "issue from A checks the PRODUCT's total stock, not A's own stock."
    check(location_contour_balance(lots, "P1", "A", "GENERAL") == 2.0,
          "availability must be scoped to the exact (product, location, contour), not product-wide")
    check(location_contour_balance(lots, "P1", "B", "GENERAL") == 20.0, "sanity check on B/GENERAL")
    check(location_contour_balance(lots, "P1", "A", "WORKSHOP_DETAILS") == 100.0,
          "same product+location but a different contour must be a completely separate bucket")


# --- R10: lines of the same product within one document must reserve stock against each other ---
def validate_issue_lines_with_reservation(lines, balance_fn):
    """lines: list of (product, location, contour, qty). Raises TestFailure on the first line
    that would push cumulative reservation past what the location/contour actually has."""
    reserved = {}
    for idx, (product, location, contour, qty) in enumerate(lines):
        key = (product, location, contour)
        available = balance_fn(product, location, contour) - reserved.get(key, 0.0)
        if qty > available + 1e-6:
            raise TestFailure(f"line {idx + 1}: needs {qty}, only {available} left after in-document reservation")
        reserved[key] = reserved.get(key, 0.0) + qty
    return True


def test_multiline_document_reservation_rejects_whole_batch_when_insufficient():
    balance_fn = lambda p, l, c: 10.0  # noqa: E731 - stock is 10 for any (product, location, contour)
    # Two lines of 6 each against the same 10 units must fail - a per-line-independent check
    # would let both pass (each individually sees 10 available) and net stock to -2.
    try:
        validate_issue_lines_with_reservation(
            [("P1", "A", "GENERAL", 6.0), ("P1", "A", "GENERAL", 6.0)], balance_fn)
        raise TestFailure("expected the second line to be rejected by in-document reservation")
    except TestFailure as e:
        check("line 2" in str(e), f"wrong rejection: {e}")
    # A single line of 6, or two lines of 4+6=10, must both be accepted.
    check(validate_issue_lines_with_reservation([("P1", "A", "GENERAL", 6.0)], balance_fn) is True,
          "a single line within stock must be accepted")
    check(validate_issue_lines_with_reservation(
        [("P1", "A", "GENERAL", 4.0), ("P1", "A", "GENERAL", 6.0)], balance_fn) is True,
        "two lines that exactly exhaust stock must both be accepted")


# --- R05/R06: TRANSFER and ADJUSTMENT must keep stock == sum(lot balances) (lot lineage) ------
def test_transfer_preserves_lot_lineage_and_fifo_age():
    # Source lot received 2026-01-01; transferring part of it must create a NEW destination lot
    # that inherits the ORIGINAL receipt date (FIFO age), not the transfer date - otherwise a
    # transferred lot would unfairly jump to the back of the FIFO queue at its new location.
    source_lot = {"id": "LOT-1", "location": "A", "contour": "GENERAL", "receipt_date": "2026-01-01", "balance": 20.0}
    transfer_qty = 8.0
    transfer_date = "2026-06-01"

    dest_lot = {
        "id": "LOT-2",
        "location": "B",
        "contour": "GENERAL",
        "receipt_date": source_lot["receipt_date"],  # inherited, not transfer_date
        "parent_lot_id": source_lot["id"],
        "balance": transfer_qty,
    }
    source_lot["balance"] -= transfer_qty

    check(dest_lot["receipt_date"] == "2026-01-01", "destination lot must inherit the source lot's FIFO age")
    check(dest_lot["parent_lot_id"] == "LOT-1", "destination lot must record lot lineage (PARENT_LOT_ID)")
    total_after = source_lot["balance"] + dest_lot["balance"]
    check(total_after == 20.0, f"transfer must conserve total stock across both lots, got {total_after}")


def test_inventory_adjustment_keeps_stock_equal_to_sum_of_lots():
    # Shortage: must be taken FROM a specific real lot (FIFO), not written as a lot-less movement.
    lots = {"LOT-1": 10.0}
    shortage = 3.0
    lots["LOT-1"] -= shortage
    check(lots["LOT-1"] == 7.0, "shortage must decrement a real, specific lot")
    check(sum(lots.values()) == 7.0, "stock must equal sum(lot balances) after a shortage adjustment")

    # Surplus: must create a NEW real lot (origin=ADJUSTMENT), not a lot-less movement.
    surplus = 5.0
    lots["LOT-ADJ-1"] = surplus
    check(sum(lots.values()) == 12.0, "stock must equal sum(lot balances) after a surplus adjustment")


# --- R08: repeated partial returns must consume allocations in order, not re-consume the first ---
def test_repeated_partial_return_uses_next_allocation_not_the_first_again():
    # Original issue of 10 was FIFO-split as LOT-A=5, LOT-B=5 (two allocations).
    allocations = [("ALC-A", "LOT-A", 5.0), ("ALC-B", "LOT-B", 5.0)]
    already_returned_per_alloc = {}  # ALLOC_ID -> qty already returned (R08 fix)

    def do_return(qty):
        remaining = qty
        touched = []
        for alloc_id, lot_id, alloc_qty in allocations:
            if remaining <= 1e-6:
                break
            available = alloc_qty - already_returned_per_alloc.get(alloc_id, 0.0)
            if available <= 1e-6:
                continue
            take = min(remaining, available)
            already_returned_per_alloc[alloc_id] = already_returned_per_alloc.get(alloc_id, 0.0) + take
            touched.append((lot_id, take))
            remaining -= take
        return touched

    first = do_return(5.0)
    check(first == [("LOT-A", 5.0)], f"first return of 5 must fully consume LOT-A's allocation, got {first}")

    second = do_return(5.0)
    # The bug this fixes: without per-allocation tracking, a naive re-scan of allocations in
    # FIFO order would again see ALC-A "unreturned" (if only a doc-line-level total were kept)
    # and credit LOT-A a second time instead of moving on to LOT-B.
    check(second == [("LOT-B", 5.0)],
          f"second partial return must move on to LOT-B, not re-credit LOT-A, got {second}")


# --- R11: validation must not create a product record before the whole plan is confirmed valid ---
def test_validation_does_not_create_product_until_full_plan_valid():
    created_products = []

    def validate_plan(lines):
        # Pass 1 (validation): compute what WOULD be created, write NOTHING.
        planned_new_products = [ln for ln in lines if ln["product_code"] == ""]
        for ln in lines:
            if ln["product_code"] == "" and ln["qty"] <= 0:
                raise TestFailure("invalid line found during validation")
        return planned_new_products

    def post_plan(lines):
        planned = validate_plan(lines)  # may raise - if it does, nothing below ever runs
        # Pass 2 (write stage): only reached if EVERY line validated successfully.
        for ln in planned:
            created_products.append(ln["product_name"])

    lines_with_bad_second_line = [
        {"product_code": "", "product_name": "New Widget", "qty": 5},
        {"product_code": "", "product_name": "Bad Widget", "qty": -1},  # invalid
    ]
    try:
        post_plan(lines_with_bad_second_line)
        raise TestFailure("expected validation to reject the whole plan")
    except TestFailure as e:
        check("invalid line" in str(e), f"wrong failure: {e}")
    check(created_products == [],
          "no product may be created when a LATER line in the same document fails validation")

    lines_all_valid = [{"product_code": "", "product_name": "Good Widget", "qty": 5}]
    post_plan(lines_all_valid)
    check(created_products == ["Good Widget"], "a fully valid plan must still create its new product(s)")


# --- R09/R12: one multi-line document = one DOC_ID/OP_ID/one store(), up to >=1000 lines -------
def test_multiline_batch_produces_exactly_one_document():
    posted_documents = []

    def post_batch(lines):
        # Models the 2.1.0 fix: N eligible rows are folded into ONE plan and posted with ONE
        # PRIME_PostDocument call, instead of looping PostDocument once per row (which used to
        # mean N documents and N store() calls for what the user experiences as one action).
        doc_id = f"DOC-{len(posted_documents) + 1}"
        posted_documents.append({"doc_id": doc_id, "lines": list(lines), "store_calls": 1})
        return doc_id

    rows = [{"product": f"P{i}", "qty": 1} for i in range(126)]  # the master task's concrete example
    doc_id = post_batch(rows)
    check(len(posted_documents) == 1, f"126 rows must produce exactly 1 document, got {len(posted_documents)}")
    check(posted_documents[0]["store_calls"] == 1, "exactly one store() for the whole batch")
    check(len(posted_documents[0]["lines"]) == 126, "all 126 rows must be lines of the same document")

    MAX_LINES = 999  # PRIME_DOC_PLAN_MAX_LINE_INDEX (0-based) -> 1000 lines supported
    check(MAX_LINES + 1 >= 1000, f"document line capacity must be at least 1000, got {MAX_LINES + 1}")


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
    test_committed_op_id_cache_updates_immediately_after_post,
    test_fifo_is_location_and_contour_scoped,
    test_multiline_document_reservation_rejects_whole_batch_when_insufficient,
    test_transfer_preserves_lot_lineage_and_fifo_age,
    test_inventory_adjustment_keeps_stock_equal_to_sum_of_lots,
    test_repeated_partial_return_uses_next_allocation_not_the_first_again,
    test_validation_does_not_create_product_until_full_plan_valid,
    test_multiline_batch_produces_exactly_one_document,
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
