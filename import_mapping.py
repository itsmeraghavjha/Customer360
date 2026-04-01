import csv
import sqlite3
import os
from werkzeug.security import generate_password_hash

DB_PATH = os.path.join('instance', 'survey.db')
CSV_FILE = 'Mapping2.csv' # Make sure your CSV file is in the same folder

def import_csv():
    if not os.path.exists(DB_PATH):
        print("Database not found! Please run app.py once to create it.")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    with open(CSV_FILE, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        added_count = 0
        
        for row in reader:
            se_id = row.get('SE ID', '').strip()
            name = row.get('SALES EXEC', '').strip()
            
            # Combine Region and Sales Office for better context
            region_str = f"{row.get('REGION','').strip()} - {row.get('SALES OFFICE','').strip()}"

            # Skip empty rows
            if not se_id or not name:
                continue

            # Check if user already exists to prevent duplicates
            cursor.execute("SELECT id FROM users WHERE username=?", (se_id,))
            if cursor.fetchone():
                continue

            # SE ID is used as BOTH username and password
            hashed_pw = generate_password_hash(se_id)
            
            cursor.execute(
                "INSERT INTO users (username, password, full_name, employee_id, role, region) VALUES (?, ?, ?, ?, ?, ?)",
                (se_id, hashed_pw, name, se_id, 'se', region_str)
            )
            added_count += 1

    conn.commit()
    conn.close()
    print(f"✅ Successfully imported {added_count} Sales Executives into the database!")

if __name__ == '__main__':
    import_csv()