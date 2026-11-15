"""Create the sample data for the Data Quality Profiler.

Writes three files into data/ (fixed random seed, so every run gives the same files):

    orders.csv            about 2,000 customer orders, 15 columns, with problems planted on purpose
    customer_master.csv   150 customers: the list every order's customer_id should belong to
    rules.yaml            the business rules the orders are checked against

Everything is made up: names, customers, products and prices. Amounts are in rupees (Rs).

Problems planted in orders.csv (the exact counts are in PLANTED below, and the
end-to-end test checks that the profiler finds every one of them):

    order_id       8 written in the wrong shape, 6 reusing another order's number
    whole rows     10 exact copies of other rows, 3 completely blank rows
    order_date     15 written dd/mm/yyyy instead of yyyy-mm-dd, 4 unreadable, 5 in the future
    ship_date      30 blank, 9 earlier than the order date
    customer_id    12 that do not exist in the customer master
    customer_name  20 with stray spaces
    email          18 blank, 7 that are not email addresses
    city           25 in the wrong capital letters (MUMBAI, mumbai)
    category       6 values outside the allowed list
    quantity       7 extreme values, 5 zero or negative
    unit_price_rs  14 typed as text ("Rs 1,250.00"), 3 that are not numbers at all
    discount_pct   40 blank, 11 holding the placeholder "N/A"
    currency       the same value on every row (a constant column)
    remarks        completely empty

Usage:
    python make_sample_data.py

Author : Swaminathan Narayanan
Repo   : https://github.com/Swami529796/data-quality-profiler
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
DATA = Path(__file__).resolve().parent / "data"
AS_OF = date(2026, 3, 31)          # "today" for the sample: the last day of the financial year
N_ORDERS = 2000
N_CUSTOMERS = 150

# How many of each problem are planted. The tests assert the profiler finds exactly these.
PLANTED = {
    "order_id_wrong_shape": 8,
    "order_id_reused": 6,
    "exact_duplicate_rows": 10,
    "blank_rows": 3,
    "order_date_other_format": 15,
    "order_date_unreadable": 4,
    "order_date_future": 5,
    "ship_date_blank": 30,              # includes the 5 future orders, which cannot have shipped
    "ship_before_order": 9,
    "customer_not_in_master": 12,
    "name_stray_spaces": 20,
    "email_blank": 18,
    "email_invalid": 7,
    "city_wrong_case": 25,
    "category_not_allowed": 6,
    "quantity_extreme": 7,
    "quantity_not_positive": 5,
    "price_as_text": 14,
    "price_not_a_number": 3,
    "discount_blank": 40,
    "discount_placeholder": 11,
}

FIRST_NAMES = ["Aarav", "Aditi", "Akash", "Ananya", "Arjun", "Bhavna", "Chetan", "Deepa", "Dinesh", "Divya",
               "Farhan", "Gauri", "Harish", "Isha", "Jayesh", "Kavita", "Kiran", "Lakshmi", "Manoj", "Meera",
               "Nikhil", "Pooja", "Prakash", "Priya", "Rahul", "Rekha", "Sanjay", "Sneha", "Suresh", "Vidya"]
LAST_NAMES = ["Agarwal", "Bhat", "Chopra", "Desai", "Fernandes", "Gupta", "Iyer", "Joshi", "Kapoor", "Kulkarni",
              "Menon", "Mishra", "Nair", "Patel", "Pillai", "Rao", "Reddy", "Shah", "Sharma", "Singh",
              "Sinha", "Thomas", "Trivedi", "Varma", "Verma"]
CITIES = ["Mumbai", "Pune", "Bengaluru", "Chennai", "Hyderabad", "Delhi", "Kolkata", "Ahmedabad", "Jaipur", "Kochi"]
SEGMENTS = ["Retail", "Wholesale", "Online"]
CATEGORIES = ["Electronics", "Furniture", "Stationery", "Kitchen", "Apparel"]

# product code, category, unit price in Rs
PRODUCTS = [
    ("ELC-1042", "Electronics", 1499.00), ("ELC-1077", "Electronics", 2999.00), ("ELC-1103", "Electronics", 899.00),
    ("ELC-1150", "Electronics", 12499.00), ("ELC-1188", "Electronics", 54999.00),
    ("FUR-2011", "Furniture", 4250.00), ("FUR-2034", "Furniture", 7999.00), ("FUR-2056", "Furniture", 1850.00),
    ("FUR-2090", "Furniture", 15500.00), ("FUR-2121", "Furniture", 3299.00),
    ("STN-3005", "Stationery", 149.00), ("STN-3018", "Stationery", 249.00), ("STN-3027", "Stationery", 399.00),
    ("STN-3044", "Stationery", 99.00), ("STN-3060", "Stationery", 549.00),
    ("KIT-4002", "Kitchen", 1250.00), ("KIT-4019", "Kitchen", 2199.00), ("KIT-4033", "Kitchen", 649.00),
    ("KIT-4051", "Kitchen", 3499.00), ("KIT-4078", "Kitchen", 899.50),
    ("APP-5010", "Apparel", 799.00), ("APP-5023", "Apparel", 1299.00), ("APP-5041", "Apparel", 1999.00),
    ("APP-5066", "Apparel", 599.00), ("APP-5084", "Apparel", 2499.00),
]

RULES_YAML = """\
# Business rules for data/orders.csv. Every rule names a column and a check.
# Paths are relative to this file. See the README for the full list of checks.

key: order_id            # the column that should identify a row
as_of: 2026-03-31        # "today" for the not_in_future check (leave out to use the real date)

rules:
  - column: order_id
    check: not_null

  - column: order_id
    check: unique

  - column: category
    check: allowed_values
    values: [Electronics, Furniture, Stationery, Kitchen, Apparel]

  - column: quantity
    check: range
    min: 1

  - column: discount_pct
    check: range
    min: 0
    max: 50
    severity: Medium

  - column: email
    check: regex
    pattern: '[\\w.+-]+@[\\w-]+(\\.[\\w-]+)+'
    severity: Medium

  - column: order_date
    check: not_in_future

  - column: order_date
    check: less_or_equal
    other: ship_date
    name: order_date is on or before ship_date

  - column: customer_id
    check: exists_in
    file: customer_master.csv
    other: customer_id
"""


def make_master(rng: np.random.Generator) -> pd.DataFrame:
    """150 customers, each with one name, email, city and segment."""
    rows = []
    for number in range(1, N_CUSTOMERS + 1):
        first, last = rng.choice(FIRST_NAMES), rng.choice(LAST_NAMES)
        rows.append({
            "customer_id": f"C{number:04d}",
            "customer_name": f"{first} {last}",
            "email": f"{first}.{last}{number}@example.com".lower(),
            "city": rng.choice(CITIES),
            "segment": rng.choice(SEGMENTS, p=[0.5, 0.2, 0.3]),
        })
    return pd.DataFrame(rows)


def make_orders(rng: np.random.Generator, master: pd.DataFrame, rows: int = N_ORDERS,
                plant_problems: bool = True) -> pd.DataFrame:
    """Build the orders table as text, exactly as it will sit in the CSV."""
    start = date(2025, 4, 1)
    days = (AS_OF - start).days
    order_dates = [start + timedelta(days=int(d)) for d in rng.integers(0, days + 1, rows)]
    ship_dates = [d + timedelta(days=int(g)) for d, g in zip(order_dates, rng.integers(1, 8, rows))]
    customers = master.iloc[rng.integers(0, len(master), rows)].reset_index(drop=True)
    products = [PRODUCTS[i] for i in rng.integers(0, len(PRODUCTS), rows)]

    orders = pd.DataFrame({
        "order_id": [f"ORD-{100001 + i}" for i in range(rows)],
        "order_date": [d.isoformat() for d in order_dates],
        "ship_date": [d.isoformat() for d in ship_dates],
        "customer_id": customers["customer_id"],
        "customer_name": customers["customer_name"],
        "email": customers["email"],
        "city": customers["city"],
        "product_code": [p[0] for p in products],
        "category": [p[1] for p in products],
        "quantity": [str(q) for q in rng.integers(1, 13, rows)],
        "unit_price_rs": [f"{p[2]:.2f}" for p in products],
        "discount_pct": [str(d) for d in rng.choice([0, 5, 10, 15], rows, p=[0.4, 0.3, 0.2, 0.1])],
        "prepaid": rng.choice(["Yes", "No"], rows, p=[0.65, 0.35]),
        "currency": "INR",
    })
    if not plant_problems:
        return orders
    orders["remarks"] = ""                                   # a completely empty column

    # Every problem goes on its own rows, so the planted counts never overlap.
    spare = list(rng.permutation(rows))

    def take(count: int) -> list[int]:
        chosen, spare[:] = spare[:count], spare[count:]
        return [int(i) for i in chosen]

    def put(column: str, row: int, value: str) -> None:
        orders.loc[row, column] = value

    for n, row in enumerate(take(PLANTED["order_id_wrong_shape"])):
        good = orders.loc[row, "order_id"]
        put("order_id", row, good.replace("-", "") if n % 2 == 0 else good.lower())
    for row, source in zip(take(PLANTED["order_id_reused"]), take(PLANTED["order_id_reused"])):
        put("order_id", row, orders.loc[source, "order_id"])

    for row in take(PLANTED["order_date_other_format"]):
        put("order_date", row, date.fromisoformat(orders.loc[row, "order_date"]).strftime("%d/%m/%Y"))
    for row, junk in zip(take(PLANTED["order_date_unreadable"]), ["31/02/2026", "TBC", "2026-13-05", "last week"]):
        put("order_date", row, junk)
    future = take(PLANTED["order_date_future"])
    for row in future:
        put("order_date", row, (AS_OF + timedelta(days=int(rng.integers(200, 2000)))).isoformat())
    for row in future + take(PLANTED["ship_date_blank"] - len(future)):
        put("ship_date", row, "")
    for row in take(PLANTED["ship_before_order"]):
        ordered = date.fromisoformat(orders.loc[row, "order_date"])
        put("ship_date", row, (ordered - timedelta(days=int(rng.integers(1, 30)))).isoformat())

    for n, row in enumerate(take(PLANTED["customer_not_in_master"])):
        put("customer_id", row, f"C{9001 + n}")
    for n, row in enumerate(take(PLANTED["name_stray_spaces"])):
        name = orders.loc[row, "customer_name"]
        put("customer_name", row, [" " + name, name + "  ", name.replace(" ", "  ")][n % 3])
    for row in take(PLANTED["email_blank"]):
        put("email", row, "")
    for n, row in enumerate(take(PLANTED["email_invalid"])):
        email = orders.loc[row, "email"]
        put("email", row, [email.replace("@", "."), email.split("@")[0] + "@", "not provided"][n % 3])
    for n, row in enumerate(take(PLANTED["city_wrong_case"])):
        city = orders.loc[row, "city"]
        put("city", row, city.upper() if n % 2 == 0 else city.lower())
    for n, row in enumerate(take(PLANTED["category_not_allowed"])):
        put("category", row, ["Misc", "Electronic", "Home"][n % 3])

    for row in take(PLANTED["quantity_extreme"]):
        put("quantity", row, str(int(rng.integers(250, 900))))
    for n, row in enumerate(take(PLANTED["quantity_not_positive"])):
        put("quantity", row, ["0", "-1", "-2"][n % 3])
    for row in take(PLANTED["price_as_text"]):
        put("unit_price_rs", row, f"Rs {float(orders.loc[row, 'unit_price_rs']):,.2f}")
    for row, junk in zip(take(PLANTED["price_not_a_number"]), ["TBC", "on request", "see quote"]):
        put("unit_price_rs", row, junk)
    for row in take(PLANTED["discount_blank"]):
        put("discount_pct", row, "")
    for row in take(PLANTED["discount_placeholder"]):
        put("discount_pct", row, "N/A")

    copies = orders.loc[take(PLANTED["exact_duplicate_rows"])]            # copies of untouched rows
    blanks = pd.DataFrame("", index=range(PLANTED["blank_rows"]), columns=orders.columns)
    orders = pd.concat([orders, copies, blanks], ignore_index=True)
    return orders.iloc[rng.permutation(len(orders))].reset_index(drop=True)


def main() -> int:
    rng = np.random.default_rng(SEED)
    DATA.mkdir(exist_ok=True)
    master = make_master(rng)
    orders = make_orders(rng, master)
    master.to_csv(DATA / "customer_master.csv", index=False, lineterminator="\n")
    orders.to_csv(DATA / "orders.csv", index=False, lineterminator="\n")
    (DATA / "rules.yaml").write_text(RULES_YAML, encoding="utf-8")
    print(f"Wrote data/orders.csv           {len(orders):,} rows, {len(orders.columns)} columns")
    print(f"Wrote data/customer_master.csv  {len(master):,} rows")
    print("Wrote data/rules.yaml")
    print(f"Planted problems: {sum(PLANTED.values())} in {len(PLANTED)} kinds (see PLANTED in this file)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
