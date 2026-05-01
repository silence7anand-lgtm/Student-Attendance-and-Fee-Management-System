import os
import sqlite3
from werkzeug.security import generate_password_hash
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, 'database.db')

def migrate():
    if not os.path.exists(DATABASE):
        print(f"Database {DATABASE} does not exist. Please run app.py first to initialize it.")
        return

    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    cursor = db.cursor()
    
    # Check if 'users' table exists
    table_exists = cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='users'").fetchone()
    if not table_exists:
        print("'users' table does not exist. Skipping migration.")
        db.close()
        return

    # 1. Add 'role' to users table
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'admin'")
        print("Added 'role' column to users table.")
    except sqlite3.OperationalError:
        print("'role' column likely already exists in users table.")
        
    # 2. Add 'password_hash' to students table
    try:
        cursor.execute("ALTER TABLE students ADD COLUMN password_hash TEXT")
        print("Added 'password_hash' column to students table.")
    except sqlite3.OperationalError:
        print("'password_hash' column likely already exists in students table.")

    # 3. Update existing admin user to have role 'admin'
    cursor.execute("UPDATE users SET role = 'admin' WHERE username = 'admin'")
    
    # 4. Create a default 'management' user
    try:
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                       ('manager', generate_password_hash('manager'), 'management'))
        print("Created default 'manager' user.")
    except sqlite3.IntegrityError:
        print("'manager' user likely already exists.")
    except sqlite3.OperationalError as e:
        print(f"Error inserting manager: {e}")

    # 5. Update existing students to have default password (same as roll_no)
    try:
        students = cursor.execute("SELECT id, roll_no FROM students WHERE password_hash IS NULL").fetchall()
        count = 0
        for s in students:
            p_hash = generate_password_hash(s['roll_no'])
            cursor.execute("UPDATE students SET password_hash = ? WHERE id = ?", (p_hash, s['id']))
            count += 1
        print(f"Updated {count} students with default passwords.")
    except sqlite3.OperationalError:
        print("'students' table might missing password_hash or not exist.")
    
    db.commit()
    db.close()

if __name__ == '__main__':
    migrate()
