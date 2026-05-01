import urllib.request
import urllib.parse
from http.cookiejar import CookieJar

URL = 'http://127.0.0.1:5000'

def debug():
    cj = CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    
    # 1. Login
    login_data = urllib.parse.urlencode({
        'role': 'student',
        'username': 'a03',
        'password': 'a03'
    }).encode('utf-8')
    
    print("Logging in...")
    try:
        r = opener.open(f"{URL}/login", data=login_data)
        print(f"Login Response: {r.getcode()}")
    except Exception as e:
        print(f"Login Failed: {e}")
        return

    # 2. Access Dashboard
    print("Accessing Dashboard...")
    try:
        r = opener.open(f"{URL}/")
        print(f"Dashboard Response: {r.getcode()}")
        print("Dashboard loaded successfully.")
    except urllib.error.HTTPError as e:
        print(f"Dashboard Error ({e.code}):")
        print(e.read().decode('utf-8'))
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == '__main__':
    debug()
