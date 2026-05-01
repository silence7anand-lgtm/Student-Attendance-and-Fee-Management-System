import urllib.request
import urllib.parse
from http.cookiejar import CookieJar

cj = CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

# 1. Login
data = urllib.parse.urlencode({'username': 'admin', 'password': 'admin'}).encode('ascii')
req_login = urllib.request.Request('http://127.0.0.1:5000/login', data=data)
res_login = opener.open(req_login)
print("Login Status:", res_login.getcode())

# 2. Export
try:
    req_export = urllib.request.Request('http://127.0.0.1:5000/export_attendance')
    res_export = opener.open(req_export)
    print("Export Status:", res_export.getcode())
    print("Headers:", res_export.headers)
except urllib.error.HTTPError as e:
    print("HTTP ERROR:", e.code)
    print("HTML Content:")
    print(e.read().decode('utf-8')[:2500])
