"""«Рассылка селлерам» — простое окно с кнопками поверх mailer.py.

Запускает маленький сервер только на этом компьютере (127.0.0.1) и открывает
страницу приложения в браузере. Сама отправка писем выполняется mailer.py,
поэтому всё, что проверено там (лимиты, журнал, остановка), работает и здесь.
"""

import json
import os
import re
import secrets
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
APP_ID = "seller-mailer-app"
PORTS = range(8765, 8776)
TOKEN = secrets.token_hex(16)
IS_WIN = os.name == "nt"

try:
    import mailer as M  # noqa: E402
except SystemExit:
    raise
except Exception as e:  # openpyxl missing etc.
    print("Не удалось загрузить mailer.py:", e)
    sys.exit(1)


# ---------------------------------------------------------------- config

CONFIG = HERE / "config.ini"
FIELDS = ["provider", "email", "app_password", "from_name", "your_name", "telegram", "max_per_day", "leads_file"]
PLACEHOLDERS = ("ВАШ", "@ваш_ник")


def read_config_text():
    return CONFIG.read_text(encoding="utf-8")


def set_key(text, key, value):
    value = str(value).replace("\n", " ").strip()
    pat = re.compile(rf"^({re.escape(key)}\s*=).*$", re.M)
    if pat.search(text):
        return pat.sub(lambda m: f"{m.group(1)} {value}", text, count=1)
    return text.rstrip("\n") + f"\n{key} = {value}\n"


def config_view():
    c = M.load_config(strict=False)
    missing = []
    for k, label in [("email", "почта"), ("app_password", "пароль приложения"), ("your_name", "имя"),
                     ("from_name", "имя отправителя"), ("telegram", "Telegram")]:
        v = c.get(k, "")
        if not v or any(p in v for p in PLACEHOLDERS):
            missing.append(label)
    clean = lambda v: "" if any(p in (v or "") for p in PLACEHOLDERS) else (v or "")
    return c, {
        "provider": c.get("provider", "yandex"),
        "email": clean(c.get("email")),
        "has_password": bool(clean(c.get("app_password"))),
        "from_name": clean(c.get("from_name")),
        "your_name": clean(c.get("your_name")),
        "telegram": clean(c.get("telegram")),
        "max_per_day": c.get("max_per_day", "50"),
        "work_hours": c.get("work_hours", "9-19"),
        "leads_file": c.get("leads_file", ""),
    }, missing


def save_config(data):
    text = read_config_text()
    for k in FIELDS:
        if k not in data:
            continue
        v = str(data[k]).strip()
        if k == "app_password" and not v:
            continue  # empty field = keep the saved password
        if k == "max_per_day":
            v = str(max(5, min(100, int(v or 50))))
        if k == "provider" and v not in ("yandex", "mailru"):
            v = "yandex"
        text = set_key(text, k, v)
    if "from_name" not in data and data.get("your_name"):
        text = set_key(text, "from_name", f"{data['your_name']} — ИИ-фото для маркетплейсов")
    CONFIG.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------- leads & letter

def find_leads(c):
    try:
        return M.find_leads_file(c), ""
    except SystemExit:
        if c.get("leads_file"):
            return None, f"Файл «{c['leads_file']}» не найден."
        return None, "Не нашёл таблицу wb-leads-….xlsx в «Загрузках». Скачайте её из расширения WB Lead Helper."


def leads_info(c):
    path, err = find_leads(c)
    if not path:
        return None, {"file": "", "count": 0, "error": err}
    try:
        leads = M.read_leads(path)
    except Exception as e:
        return None, {"file": path.name, "count": 0, "error": f"Не получилось прочитать таблицу: {e}"}
    good = [l for l in leads if M.EMAIL_RE.match(l["email"])]
    return leads, {"file": path.name, "path": str(path), "count": len(good),
                   "modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y %H:%M"), "error": ""}


def letter_get():
    subject, body = M.read_template("letter.txt")
    return {"subject": subject, "body": body}


def letter_save(subject, body):
    text = f"Тема: {subject.strip()}\n---\n{body.strip()}\n"
    (HERE / "letter.txt").write_text(text, encoding="utf-8")


def letter_preview(c, leads, subject=None, body=None):
    lead = None
    for l in leads or []:
        if M.EMAIL_RE.match(l["email"]):
            lead = l
            break
    lead = lead or {"email": "example@mail.ru", "company": "ИП Иванов Иван Иванович", "shop": "Пример", "inn": "", "lead_status": ""}
    if subject is None:
        subject, body = M.read_template("letter.txt")
    filled_subject = M.fill(subject, lead, c)
    filled_body = M.fill(body, lead, c).replace("{ФОТО}", "[здесь будут фото из папки photos]")
    return {"to": lead["email"], "company": lead["company"] or lead["shop"], "subject": filled_subject, "body": filled_body}


def plan(c, leads, j):
    stop = M.load_stoplist()
    queue = M.build_queue(c, j, leads or [], stop)
    limit, day_no = M.daily_limit(c)
    left = max(0, limit - M.sent_today(j))
    first = sum(1 for k, _ in queue if k == "first")
    follow = sum(1 for k, _ in queue if k == "followup")
    return {"new": first, "followups": follow, "left": left, "will_send": min(left, len(queue)),
            "limit": limit, "day": day_no + 1, "within_hours": M.within_hours(c)}


# ---------------------------------------------------------------- runner

class Runner:
    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None
        self.lines = []
        self.kind = ""

    def python(self):
        exe = Path(sys.executable)
        if exe.name.lower() == "pythonw.exe" and (exe.parent / "python.exe").exists():
            exe = exe.parent / "python.exe"
        return str(exe)

    def _popen(self, args):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        flags = subprocess.CREATE_NO_WINDOW if IS_WIN else 0
        return subprocess.Popen([self.python(), "-u", str(HERE / "mailer.py"), *args], cwd=str(HERE),
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                encoding="utf-8", errors="replace", env=env, creationflags=flags)

    @property
    def running(self):
        return self.proc is not None and self.proc.poll() is None

    def start(self, kind, args):
        with self.lock:
            if self.running:
                return False
            self.lines = []
            self.kind = kind
            self.proc = self._popen(args)
            threading.Thread(target=self._pump, args=(self.proc,), daemon=True).start()
            return True

    def _pump(self, proc):
        for line in proc.stdout:
            with self.lock:
                self.lines.append(line.rstrip("\n"))
        proc.wait()

    def run_sync(self, args, timeout=180):
        if self.running:
            return "Сейчас идёт отправка — дождитесь окончания или нажмите «Остановить»."
        p = self._popen(args)
        try:
            out, _ = p.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            p.kill()
            out = "Почта не ответила вовремя. Проверьте интернет и попробуйте ещё раз."
        return out.strip()

    def stop(self):
        if self.running:
            M.STOP_FILE.write_text("stop", encoding="utf-8")
            return True
        return False

    def log(self, since):
        with self.lock:
            return self.lines[since:], len(self.lines), self.running, self.kind


RUN = Runner()
LAST_SEEN = [time.time()]


def open_path(what):
    targets = {"journal": HERE / "journal.xlsx", "photos": HERE / "photos", "folder": HERE,
               "stoplist": M.STOPLIST, "reply": HERE / "ОТВЕТ_с_примерами.txt"}
    p = targets.get(what)
    if p is None:
        return "Неизвестный объект."
    if what == "stoplist":
        M.load_stoplist()  # creates the file if missing
    if not p.exists():
        return "Журнал появится после первой отправки." if what == "journal" else f"Не найдено: {p.name}"
    if IS_WIN:
        os.startfile(str(p))  # noqa
    else:
        subprocess.Popen(["xdg-open", str(p)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return ""


def make_shortcut():
    if not IS_WIN:
        return "Ярлык можно создать только в Windows."
    exe = Path(sys.executable)
    pyw = exe.parent / "pythonw.exe"
    target = str(pyw if pyw.exists() else exe)
    ps = (
        "$d=[Environment]::GetFolderPath('Desktop');"
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'Рассылка селлерам.lnk'));"
        f"$s.TargetPath='{target}';$s.Arguments='\"{HERE / 'app.py'}\"';"
        f"$s.WorkingDirectory='{HERE}';$s.Save()"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
    return "" if r.returncode == 0 else (r.stderr or "Не удалось создать ярлык.")


def reply_text():
    p = HERE / "ОТВЕТ_с_примерами.txt"
    if not p.exists():
        return ""
    text = p.read_text(encoding="utf-8")
    return text.split("\n---\n", 1)[-1].strip()


# ---------------------------------------------------------------- http

class Handler(BaseHTTPRequestHandler):
    server_version = "SellerMailer"

    def log_message(self, *a):
        pass

    def _host_ok(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        return host in ("127.0.0.1", "localhost")

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, {"error": "forbidden"})
        LAST_SEEN[0] = time.time()
        path = self.path.split("?")[0]
        q = dict(p.split("=", 1) for p in self.path.split("?", 1)[1].split("&") if "=" in p) if "?" in self.path else {}
        try:
            if path == "/":
                html = (HERE / "ui.html").read_text(encoding="utf-8").replace("__TOKEN__", TOKEN)
                return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            if path == "/api/hello":
                return self._send(200, {"app": APP_ID})
            if path == "/api/status":
                c, cfg, missing = config_view()
                leads, li = leads_info(c)
                j = M.load_journal()
                return self._send(200, {"config": cfg, "missing": missing, "leads": li,
                                        "stats": M.stats(c), "plan": plan(c, leads, j),
                                        "running": RUN.running, "is_windows": IS_WIN})
            if path == "/api/letter":
                c, _, _ = config_view()
                leads, _ = leads_info(c)
                return self._send(200, {**letter_get(), "preview": letter_preview(c, leads)})
            if path == "/api/log":
                lines, total, running, kind = RUN.log(int(q.get("since", "0") or 0))
                return self._send(200, {"lines": lines, "total": total, "running": running, "kind": kind})
            if path == "/api/reply":
                return self._send(200, {"text": reply_text()})
            if path == "/api/ping":
                return self._send(200, {"ok": True})
        except Exception as e:
            return self._send(500, {"error": str(e)})
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        if not self._host_ok() or self.headers.get("X-Token") != TOKEN:
            return self._send(403, {"error": "forbidden"})
        LAST_SEEN[0] = time.time()
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n) if n else b""
        path = self.path.split("?")[0]
        try:
            if path == "/api/upload":
                if not raw.startswith(b"PK"):
                    return self._send(400, {"error": "Это не файл Excel (.xlsx)."})
                dest = HERE / "wb-leads-загружено.xlsx"
                dest.write_bytes(raw)
                save_config({"leads_file": dest.name})
                return self._send(200, {"ok": True})
            data = json.loads(raw.decode("utf-8") or "{}")
            if path == "/api/config":
                save_config(data)
                return self._send(200, {"ok": True})
            if path == "/api/letter":
                letter_save(data.get("subject", ""), data.get("body", ""))
                return self._send(200, {"ok": True})
            if path == "/api/letter/preview":
                c, _, _ = config_view()
                leads, _ = leads_info(c)
                return self._send(200, letter_preview(c, leads, data.get("subject", ""), data.get("body", "")))
            if path == "/api/leads/auto":
                save_config({"leads_file": ""})
                return self._send(200, {"ok": True})
            if path == "/api/test":
                to = (data.get("to") or "").strip()
                out = RUN.run_sync(["--test"] + ([to] if to else []))
                return self._send(200, {"output": out, "ok": "Тестовое письмо отправлено" in out})
            if path == "/api/preview-files":
                out = RUN.run_sync(["--preview"])
                open_path("folder")
                return self._send(200, {"output": out})
            if path == "/api/send":
                ok = RUN.start("send", [])
                return self._send(200, {"ok": ok})
            if path == "/api/stop":
                return self._send(200, {"ok": RUN.stop()})
            if path == "/api/open":
                return self._send(200, {"error": open_path(data.get("what", ""))})
            if path == "/api/shortcut":
                return self._send(200, {"error": make_shortcut()})
            if path == "/api/quit":
                self._send(200, {"ok": True})
                threading.Thread(target=lambda: (time.sleep(0.5), os._exit(0)), daemon=True).start()
                return
        except Exception as e:
            return self._send(500, {"error": str(e)})
        return self._send(404, {"error": "not found"})


def already_running():
    for port in PORTS:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/hello", timeout=0.5) as r:
                if json.loads(r.read()).get("app") == APP_ID:
                    return port
        except Exception:
            continue
    return None


def watchdog():
    # Close by itself when the page has been closed for a while and nothing is being sent.
    while True:
        time.sleep(15)
        if time.time() - LAST_SEEN[0] > 120 and not RUN.running:
            os._exit(0)


def main():
    port = already_running()
    if port:
        webbrowser.open(f"http://127.0.0.1:{port}/")
        return
    for port in PORTS:
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    else:
        print("Не нашёл свободный порт для приложения.")
        sys.exit(1)
    url = f"http://127.0.0.1:{port}/"
    print("Приложение открыто:", url)
    if "--no-browser" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
        threading.Thread(target=watchdog, daemon=True).start()
    srv.serve_forever()


if __name__ == "__main__":
    main()
