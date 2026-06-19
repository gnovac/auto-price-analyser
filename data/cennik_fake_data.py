import random
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

# 1. Setup variables
num_records = 5000
parts_names = [
    "Tarcza hamulcowa",
    "Klocki hamulcowe",
    "Filtr oleju",
    "Filtr powietrza",
    "Amortyzator",
    "Wahacz",
    "Rozrzad",
    "Pompa wody",
    "Swieca zaplonowa",
    "Cewka",
    "Sprzeglo",
    "Akumulator",
    "Zarowka",
    "Wycieraczki",
]
branches = ["Warszawa", "Krakow", "Poznan", "Gdansk", "Katowice", "Wroclaw"]


# Helper to generate random dates
def random_date(days_back=365):
    start = datetime.now() - timedelta(days=days_back)
    return start + timedelta(days=random.randint(0, days_back))


# 2. Generate Base Data (Page1)
np.random.seed(42)  # For reproducibility
part_codes = random.sample(range(100000, 9999999999), num_records)
part_codes = [str(code) for code in part_codes]

data_page1 = []
for code in part_codes:
    name = f"{random.choice(parts_names)} {random.randint(100, 999)}"
    has_sub = random.random() < 0.05
    sub_part = str(random.randint(100000, 9999999999)) if has_sub else ""

    # Base price generation
    stock_price = round(random.uniform(20.0, 1500.0), 2)
    urgent_price = round(stock_price * random.uniform(1.1, 1.3), 2)
    retail_price = round(urgent_price * random.uniform(1.2, 1.5), 2)
    pack_qty = random.choice([1, 1, 1, 2, 4, 10, 50])  # Weighted towards 1

    data_page1.append(
        [code, name, sub_part, stock_price, urgent_price, retail_price, pack_qty]
    )

df_page1 = pd.DataFrame(
    data_page1,
    columns=[
        "Part Code",
        "Part_Description_PL",
        "Sub_To_Part_Number",
        "cena stock",
        "cena pilne",
        "LIST PRICE  Detal",
        "Pack Quantity",
    ],
)

# 3. Generate subsets for other sheets
# 30% of records = 1500, 15% = 750
sample_30 = 1500
sample_15 = 750

# AFTER (Lowest price: 60-80% of stock)
after_codes = random.sample(part_codes, sample_30)
df_after = pd.DataFrame(
    {
        "PN": after_codes,
        "cena ": [
            round(
                df_page1[df_page1["Part Code"] == c]["cena stock"].values[0]
                * random.uniform(0.6, 0.8),
                2,
            )
            for c in after_codes
        ],
        "data": [random_date().strftime("%Y-%m-%d") for _ in range(sample_30)],
    }
)

# UZYSKANE CENY (80-95% of stock)
uzyskane_codes = random.sample(part_codes, sample_30)
df_uzyskane = pd.DataFrame(
    {
        "Numer części": uzyskane_codes,
        "Cena pln Stock": [
            round(
                df_page1[df_page1["Part Code"] == c]["cena stock"].values[0]
                * random.uniform(0.8, 0.95),
                2,
            )
            for c in uzyskane_codes
        ],
        "Data uzyskanej ceny": [
            random_date().strftime("%Y-%m-%d") for _ in range(sample_30)
        ],
    }
)

# KOSZYK (85-95% of stock)
koszyk_codes = random.sample(part_codes, sample_30)
df_koszyk = pd.DataFrame(
    {
        "Numer części": koszyk_codes,
        "Opis produktu": [
            df_page1[df_page1["Part Code"] == c]["Part_Description_PL"].values[0]
            for c in koszyk_codes
        ],
        "Promocja - cena": [
            round(
                df_page1[df_page1["Part Code"] == c]["cena stock"].values[0]
                * random.uniform(0.85, 0.95),
                2,
            )
            for c in koszyk_codes
        ],
    }
)

# PH (90-98% of stock)
ph_codes = random.sample(part_codes, sample_30)
df_ph = pd.DataFrame(
    {
        "PN": ph_codes,
        "PH aktualne": [
            round(
                df_page1[df_page1["Part Code"] == c]["cena stock"].values[0]
                * random.uniform(0.9, 0.98),
                2,
            )
            for c in ph_codes
        ],
    }
)

# DEADSTOCK (15% coverage)
deadstock_codes = random.sample(part_codes, sample_15)
df_deadstock = pd.DataFrame(
    {
        "PN": deadstock_codes,
        "Oddział": [random.choice(branches) for _ in range(sample_15)],
        "Stan magazynu dostępny": [random.randint(1, 150) for _ in range(sample_15)],
        "średnia cena zak": [
            round(
                df_page1[df_page1["Part Code"] == c]["cena stock"].values[0]
                * random.uniform(0.9, 1.1),
                2,
            )
            for c in deadstock_codes
        ],
    }
)

# ZAPAS 3 MCE< (30% coverage)
zapas_codes = random.sample(part_codes, sample_30)
df_zapas = pd.DataFrame(
    {
        "PN": zapas_codes,
        "nazwa": [
            df_page1[df_page1["Part Code"] == c]["Part_Description_PL"].values[0]
            for c in zapas_codes
        ],
        "stan mag.": [random.randint(10, 500) for _ in range(sample_30)],
        "Zapas powyżej 3 mc-y": [
            round(random.uniform(3.1, 24.0), 2) for _ in range(sample_30)
        ],
    }
)

# 4. Save to Excel
output_filename = "cennik_fake_data.xlsx"
with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
    df_page1.to_excel(writer, sheet_name="Page1", index=False)
    df_uzyskane.to_excel(writer, sheet_name="Uzyskane ceny...", index=False)
    df_koszyk.to_excel(writer, sheet_name="KOSZYK ....", index=False)
    df_ph.to_excel(writer, sheet_name="PH", index=False)
    df_after.to_excel(writer, sheet_name="After", index=False)
    df_deadstock.to_excel(writer, sheet_name="DEADSTOCK", index=False)
    df_zapas.to_excel(writer, sheet_name="Zapas 3mce<", index=False)

print(f"Plik {output_filename} został wygenerowany pomyślnie!")
