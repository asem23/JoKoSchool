# -*- coding: utf-8 -*-
"""سيرفر الامتحان: يقدّم الصفحة والأسئلة (بدون إجابات)، ويصحّح ويحفظ النتائج ويصدّرها Excel.
التشغيل:  python server.py
المتغيرات (اختيارية): PORT, DATA_DIR, START_PASSWORD, END_PASSWORD"""
import random, http.server, json, hashlib, hmac, zipfile, io, os, threading, time, socket
from urllib.parse import quote
from xml.sax.saxutils import escape

DIR = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("PORT", 8000))
DATA_DIR = os.environ.get("DATA_DIR", DIR)
os.makedirs(DATA_DIR, exist_ok=True)
DB = os.path.join(DATA_DIR, "results.json")
START_PW = os.environ.get("START_PASSWORD", "Start@2026")
END_PW = os.environ.get("END_PASSWORD", "Moraqeb@2026")
EXAM_SIZE = min(int(os.environ.get("EXAM_SIZE", 30)), 10**6)   # عدد الأسئلة لكل طالب
DEFAULT_PW = "START_PASSWORD" not in os.environ or "END_PASSWORD" not in os.environ

with open(os.path.join(DIR, "questions.json"), encoding="utf-8") as f:
    QS = json.load(f)
KEY = {q["id"]: q["a"] for q in QS}
PUBLIC = [{"id": q["id"], "q": q["q"], "o": q["o"]} for q in QS]
lock = threading.Lock()
fails = {}

def same(a, b):
    return hmac.compare_digest(hashlib.sha256(str(a).encode()).digest(), hashlib.sha256(b.encode()).digest())

def load():
    try:
        with open(DB, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def make_xlsx(rows):
    data = [["#", "الاسم", "التخصص", "الشعبة", "العلامة", "من", "النسبة %", "وقت الإنهاء"]]
    for n, r in enumerate(rows, 1):
        data.append([n, r["name"], r["major"], r["sec"], r["score"], r["total"], r["pct"], r["time"]])
    out = []
    for ri, row in enumerate(data, 1):
        cs = []
        for ci, v in enumerate(row):
            ref = "ABCDEFGH"[ci] + str(ri)
            cs.append('<c r="%s"><v>%s</v></c>' % (ref, v) if isinstance(v, (int, float))
                      else '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, escape(str(v))))
        out.append('<row r="%d">%s</row>' % (ri, "".join(cs)))
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>'
             '<cols><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="4" width="24" customWidth="1"/><col min="8" max="8" width="22" customWidth="1"/></cols>'
             '<sheetData>%s</sheetData></worksheet>') % "".join(out)
    R = "http://schemas.openxmlformats.org/"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="%spackage/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % R)
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="%spackage/2006/relationships"><Relationship Id="rId1" Type="%sofficeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' % (R, R))
        z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="%sspreadsheetml/2006/main" xmlns:r="%sofficeDocument/2006/relationships"><sheets><sheet name="Results" sheetId="1" r:id="rId1"/></sheets></workbook>' % (R, R))
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="%spackage/2006/relationships"><Relationship Id="rId1" Type="%sofficeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>' % (R, R))
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()

class H(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def ip(self):
        return (self.headers.get("X-Forwarded-For") or self.client_address[0]).split(",")[0].strip()

    def send(self, code, body=b"", ctype="application/json", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items(): self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def js(self, code, obj):
        self.send(code, json.dumps(obj, ensure_ascii=False).encode())

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(os.path.join(DIR, "index.html"), "rb") as f:
                self.send(200, f.read(), "text/html; charset=utf-8")
        elif self.path == "/healthz":
            self.send(200, b"ok", "text/plain")
        else:
            self.send(404, b"not found", "text/plain")

    def auth(self, b, real):
        now, ip = time.time(), self.ip()
        fl = [t for t in fails.get(ip, []) if now - t < 60]
        fails[ip] = fl
        if len(fl) >= 8:
            self.js(429, {"error": "محاولات كثيرة، انتظر دقيقة"}); return False
        if not same(b.get("pw", ""), real):
            fl.append(now); self.js(403, {"error": "كلمة المرور غير صحيحة"}); return False
        return True

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 200000:
            return self.js(413, {"error": "too large"})
        try:
            b = json.loads(self.rfile.read(n) or b"{}")
        except Exception:
            return self.js(400, {"error": "bad request"})
        if self.path == "/api/start":
            if not self.auth(b, START_PW): return
            self.js(200, {"questions": random.sample(PUBLIC, min(EXAM_SIZE, len(PUBLIC)))})
        elif self.path == "/api/finish":
            if not self.auth(b, END_PW): return
            ans = b.get("ans") if isinstance(b.get("ans"), dict) else {}
            ids = b.get("ids")
            if (not isinstance(ids, list) or len(set(ids)) != len(ids) or len(ids) != min(EXAM_SIZE, len(KEY))
                    or any(i not in KEY for i in ids)):
                return self.js(400, {"error": "قائمة الأسئلة غير صالحة"})
            score = sum(1 for k in ids if ans.get(str(k)) == KEY[k])
            total = len(ids)
            row = {"name": str(b.get("name", ""))[:80], "major": str(b.get("major", ""))[:60], "sec": str(b.get("sec", ""))[:20],
                   "score": score, "total": total, "pct": round(score * 100 / total),
                   "time": time.strftime("%Y-%m-%d %H:%M:%S")}
            with lock:
                rows = load(); rows.append(row)
                with open(DB, "w", encoding="utf-8") as f: json.dump(rows, f, ensure_ascii=False)
            self.js(200, {"score": score, "total": total, "pct": row["pct"], "correct": {str(k): KEY[k] for k in ids}})
        elif self.path == "/api/export":
            if not self.auth(b, END_PW): return
            with lock: data = make_xlsx(load())
            self.send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      {"Content-Disposition": "attachment; filename*=UTF-8''" + quote("نتائج_الامتحان.xlsx")})
        else:
            self.send(404, b"not found", "text/plain")

if __name__ == "__main__":
    if DEFAULT_PW:
        print("تنبيه: تستخدم كلمات المرور الافتراضية. عيّن START_PASSWORD و END_PASSWORD قبل الاستخدام الفعلي.")
    try: ip = socket.gethostbyname(socket.gethostname())
    except Exception: ip = "IP"
    print("localhost: http://localhost:%d   |   الشبكة: http://%s:%d" % (PORT, ip, PORT))
    http.server.ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()
