from app import app, export_attendance
from flask import session
import traceback

with app.app_context():
    with app.test_request_context('/export_attendance'):
        app.preprocess_request()
        session['role'] = 'admin'
        session['user_id'] = 1
        try:
            res = export_attendance()
            print("SUCCESS. Return type:", type(res))
        except Exception as e:
            print("ERROR IN EXPORT:")
            traceback.print_exc()
