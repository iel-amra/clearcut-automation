"""Generate a realistic fake sales CSV (no real data) for the demo."""
import csv, random
from datetime import date, timedelta
random.seed(7)
products = [("Desk lamp", 39.9), ("Office chair", 189.0), ("Monitor arm", 59.0), ("USB-C hub", 34.5),
            ("Standing desk", 420.0), ("Notebook pack", 12.0), ("Wireless mouse", 24.9), ("Webcam", 69.0)]
regions = ["North", "South", "East", "West"]
start = date(2026, 1, 1)
with open("sales.csv", "w", newline="") as fh:
    w = csv.writer(fh); w.writerow(["date", "order_id", "region", "product", "quantity", "unit_price"])
    for i in range(1, 601):
        d = start + timedelta(days=random.randint(0, 270))
        p, price = random.choice(products)
        w.writerow([d.isoformat(), f"ORD-{10000+i}", random.choice(regions), p, random.randint(1, 6), price])
print("sales.csv written (600 orders)")
