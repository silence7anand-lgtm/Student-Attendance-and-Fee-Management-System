from app import app, export_attendance, get_db
import sqlite3
import pandas as pd

with app.app_context():
    db = get_db()
    students = db.execute('SELECT count(*) as c FROM students').fetchone()['c']
    att = db.execute('SELECT count(*) as c FROM attendance').fetchone()['c']
    print(f"Students count: {students}, Attendance count: {att}")

    with app.test_request_context('/export_attendance'):
        app.preprocess_request()
        import flask
        flask.session['role'] = 'admin'
        flask.session['user_id'] = 1
        
        try:
            res = export_attendance()
            print("Response:", res.status_code, len(res.data))
            
            # actually read the excel to see if it's readable
            import io
            df = pd.read_excel(io.BytesIO(res.data))
            print("Excel looks like:")
            print(df.head())
        except Exception as e:
            import traceback
            traceback.print_exc()
