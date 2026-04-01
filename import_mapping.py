import csv
import sqlite3
import os
from werkzeug.security import generate_password_hash

DB_PATH = os.path.join('instance', 'survey.db')
CSV_FILE = 'Mapping2.csv'

def import_csv():
    if not os.path.exists(DB_PATH):
        print("Database not found! Please run app.py once to create it.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    with open(CSV_FILE, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        added_count = 0
        skipped = 0

        for row in reader:
            se_id        = row.get('SE ID', '').strip()               # e.g. 118539
            name         = row.get('SALES EXEC', '').strip()          # e.g. AJITH K E .
            region       = row.get('REGION', '').strip()              # e.g. KA, TG-2
            sales_office = row.get('SALES OFFICE', '').strip().replace(' ', '')  # BSO 5 → BSO5, HSO 11 → HSO11

            if not se_id or not name:
                skipped += 1
                continue

            cursor.execute("SELECT id FROM users WHERE username=?", (se_id,))
            if cursor.fetchone():
                print(f"  ⚠️  Skipping duplicate: {se_id} ({name})")
                skipped += 1
                continue

            hashed_pw = generate_password_hash(se_id)

            cursor.execute(
                """INSERT INTO users 
                   (username, password, full_name, employee_id, role, region, sales_office)
                   VALUES (?, ?, ?, ?, 'se', ?, ?)""",
                (se_id, hashed_pw, name, se_id, region, sales_office)
            )
            added_count += 1
            print(f"  ✅ Added: {se_id} | {name} | {region} | {sales_office}")

    conn.commit()
    conn.close()
    print(f"\n✅ Import complete — Added: {added_count}, Skipped: {skipped}")

if __name__ == '__main__':
    import_csv()