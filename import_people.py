import json
import pandas as pd
import os

PEOPLE_FILE = "people.xlsx"
STORAGE_FILE = "people_storage.json"

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def main():
    if not os.path.exists(PEOPLE_FILE):
        print("❌ people.xlsx не знайдено")
        return

    df = pd.read_excel(PEOPLE_FILE)

    required_cols = {"surname", "address", "station"}
    if not required_cols.issubset(df.columns):
        print("❌ В Excel мають бути колонки: surname, address, station")
        return

    storage = {}

    for _, row in df.iterrows():
        surname_key = str(row["surname"]).strip().lower()
        storage[surname_key] = {
            "surname": str(row["surname"]).strip(),
            "address": str(row["address"]).strip(),
            "station": str(row["station"]).strip()
        }

    save_json(STORAGE_FILE, storage)
    print(f"✅ Імпортовано {len(storage)} людей у {STORAGE_FILE}")

if __name__ == "__main__":
    main()
