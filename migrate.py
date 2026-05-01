import sqlite3

db = sqlite3.connect('database.db')
try:
    db.execute("ALTER TABLE fees ADD COLUMN late_fee DECIMAL(10, 2) DEFAULT 0")
    db.execute("ALTER TABLE students ADD COLUMN profile_pic TEXT")
    db.commit()
    print("Migration successful: Added late_fee and profile_pic columns.")
except sqlite3.OperationalError as e:
    print(f"Migration skipped (might already exist): {e}")
finally:
    db.close()
