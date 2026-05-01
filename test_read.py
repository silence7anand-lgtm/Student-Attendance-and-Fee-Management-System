import urllib.request
import urllib.parse
from http.cookiejar import CookieJar
import pandas as pd
import io

cj = CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

# Login
data = urllib.parse.urlencode({'username': 'admin', 'password': 'admin'}).encode('ascii')
opener.open(urllib.request.Request('http://127.0.0.1:5000/login', data=data))

# Download
res = opener.open(urllib.request.Request('http://127.0.0.1:5000/export_attendance'))
file_bytes = res.read()
print('Size downloaded:', len(file_bytes))

# Try parsing
try:
    df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    print('DF SHAPE:', df.shape)
    print("Columns:", df.columns.tolist())
    print("Success reading Excel!")
except Exception as e:
    import traceback
    traceback.print_exc()
