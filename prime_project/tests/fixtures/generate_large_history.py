#!/usr/bin/env python3
"""
Генерирует синтетическую историю движений заданного размера (сценарии 1000/10000 движений)
без хранения такого объёма как статический JSON в репозитории. Все данные вымышленные:
небольшой пул синтетических товаров/партий/мест, циклически используемых с детерминированным
псевдослучайным количеством (фиксированный seed - воспроизводимо между запусками).

Использование:
    python3 generate_large_history.py 1000 > /tmp/history_1000.json
    python3 generate_large_history.py 10000 > /tmp/history_10000.json
"""
import json
import random
import sys

PRODUCTS = [f"ЕИ-{i:08d}" for i in range(1, 21)]  # 20 синтетических товаров
LOCATIONS = ["Sklad-1", "Sklad-2", "Sklad-3"]


def generate(n: int, seed: int = 42) -> dict:
    rnd = random.Random(seed)
    movements = []
    lots = []
    lot_counter = 1
    balances = {p: 0.0 for p in PRODUCTS}

    for i in range(n):
        product = rnd.choice(PRODUCTS)
        location = rnd.choice(LOCATIONS)
        # Каждое 3-е движение - расход (если есть остаток), иначе приход.
        is_issue = (i % 3 == 0) and balances[product] > 0
        if is_issue:
            qty = -min(balances[product], round(rnd.uniform(1, 20), 2))
        else:
            lot_id = f"LOT-{lot_counter:08d}"
            lot_counter += 1
            lots.append({"LOT_ID": lot_id, "PRODUCT_CODE": product, "LOCATION": location})
            qty = round(rnd.uniform(1, 50), 2)
        balances[product] = round(balances[product] + qty, 2)
        movements.append({
            "MOVE_ID": f"MOV-{i + 1:08d}",
            "PRODUCT_CODE": product,
            "LOCATION": location,
            "QTY_BASE": qty,
        })

    return {
        "scenario": f"movement_history_{n}",
        "description": f"Синтетическая история из {n} COMMITTED-движений на {len(PRODUCTS)} товарах "
                        f"и {len(LOCATIONS)} местах хранения, детерминированный seed={seed}.",
        "products": PRODUCTS,
        "locations": LOCATIONS,
        "lots_created": len(lots),
        "movements": movements,
        "expected": {"total_stock_by_product": balances},
    }


if __name__ == "__main__":
    count = int(sys.argv[1]) if len(sys.argv) > 1 else 1000
    print(json.dumps(generate(count), ensure_ascii=False, indent=2))
