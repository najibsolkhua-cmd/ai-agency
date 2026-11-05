"""Personal email outreach from ONE mailbox, with warm-up limits.

Reads the xlsx exported by WB Lead Helper, writes each seller a personal
letter (company name, shop name, example photos), keeps a journal so nobody
gets the first letter twice, sends one follow-up after N days to those who
did not answer, and stops for the day if the mail server starts refusing.

Usage (from the .bat files):
    mailer.py            send today's portion
    mailer.py --preview  build .eml previews in the "preview" folder, send nothing
    mailer.py --test     send one letter to your own address
"""

import argparse
import configparser
import email.utils
import glob
import html
import imaplib
import os
import random
import re
import smtplib
import ssl
import sys
import time
from datetime import date, datetime, timedelta
from email.header import decode_header, make_header
from email.message import EmailMessage
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    print("Не установлена библиотека openpyxl. Запустите install.bat и попробуйте снова.")
    sys.exit(1)

HERE = Path(__file__).resolve().parent
JOURNAL = HERE / "journal.xlsx"
STOPLIST = HERE / "stoplist.txt"
STATE = HERE / "state.ini"
STOP_FILE = HERE / "STOP"

PRESETS = {
    "yandex": ("smtp.yandex.ru", 465, "ssl", "imap.yandex.ru"),
    "mailru": ("smtp.mail.ru", 465, "ssl", "imap.mail.ru"),
}
WARMUP = [(7, 10), (14, 20), (21, 35)]  # (days since start, letters per day)
HARD_CAP = 100
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
J_COLS = ["Email", "Компания", "Магазин", "ИНН", "Письмо 1", "Message-ID", "Напоминание", "Статус", "Ошибка", "Комментарий"]
DONE_STATUSES = {"ответил", "отказ", "не доставлено"}


# ---------------------------------------------------------------- helpers

def say(*a):
    print(*a, flush=True)


def load_config(strict=True):
    path = HERE / "config.ini"
    if not path.exists():
        say("Нет файла config.ini рядом с программой.")
        sys.exit(1)
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(path, encoding="utf-8")
    c = {k: v.strip() for k, v in cfg["mail"].items()}
    c.update({k: v.strip() for k, v in cfg["letter"].items()})
    c.update({k: v.strip() for k, v in cfg["schedule"].items()})
    prov = c.get("provider", "yandex").lower()
    if prov in PRESETS:
        host, port, sec, imap = PRESETS[prov]
        c["smtp_host"] = c.get("smtp_host") or host
        c["smtp_port"] = c.get("smtp_port") or str(port)
        c["security"] = c.get("security") or sec
        c["imap_host"] = c.get("imap_host") or imap
    for key in ("email", "app_password", "smtp_host", "your_name", "from_name", "telegram") if strict else ():
        if not c.get(key) or "ВАШ" in c.get(key, "") or c.get(key) == "@ваш_ник":
            say(f"В config.ini не заполнено поле «{key}».")
            sys.exit(1)
    return c


def company_name(legal):
    s = re.sub(r"\s+", " ", (legal or "").strip())
    if not s:
        return ""
    reps = [
        (r"^ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ПАО"),
        (r"^АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО"),
        (r"^ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО"),
        (r"^ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП"),
    ]
    for pat, short in reps:
        s = re.sub(pat, short, s, flags=re.I)
    s = re.sub(r"[\"“”„]([^\"“”„]+)[\"“”„]", r"«\1»", s)
    if s.startswith("ИП "):
        rest = s[3:]
        if rest.isupper():
            rest = " ".join("-".join(p.capitalize() for p in w.split("-")) for w in rest.split())
        s = "ИП " + rest
    return s


def read_leads(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Продавцы"] if "Продавцы" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [str(h or "").strip() for h in rows[0]]
    idx = {h: i for i, h in enumerate(head)}

    def g(r, name):
        i = idx.get(name)
        return str(r[i]).strip() if i is not None and i < len(r) and r[i] is not None else ""

    leads = []
    for r in rows[1:]:
        mail = g(r, "Email").lower()
        if not mail:
            continue
        leads.append({
            "email": mail,
            "company": company_name(g(r, "Юрлицо / ИП")),
            "shop": g(r, "Магазин на WB"),
            "inn": g(r, "ИНН"),
            "lead_status": g(r, "Статус"),
        })
    return leads


def find_leads_file(c):
    if c.get("leads_file"):
        p = Path(c["leads_file"])
        p = p if p.is_absolute() else HERE / p
        if p.exists():
            return p
        say(f"Не найден файл с базой: {p}")
        sys.exit(1)
    cands = glob.glob(str(HERE / "wb-leads-*.xlsx"))
    cands += glob.glob(str(Path.home() / "Downloads" / "wb-leads-*.xlsx"))
    if not cands:
        say("Не нашёл файл wb-leads-*.xlsx ни рядом с программой, ни в «Загрузках».")
        say("Выгрузите таблицу из расширения WB Lead Helper («Скачать xlsx») и запустите снова.")
        sys.exit(1)
    return Path(max(cands, key=os.path.getmtime))


def load_journal():
    if not JOURNAL.exists():
        return {}
    wb = load_workbook(JOURNAL)
    ws = wb.active
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        rec = {k: ("" if v is None else str(v)) for k, v in zip(J_COLS, r)}
        out[rec["Email"].lower()] = rec
    return out


def save_journal(j):
    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"
    ws.append(J_COLS)
    for rec in j.values():
        ws.append([rec.get(k, "") for k in J_COLS])
    widths = [30, 40, 22, 14, 17, 40, 17, 15, 40, 30]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "B2"
    try:
        wb.save(JOURNAL)
    except PermissionError:
        say("!!! Не могу записать journal.xlsx — закройте его в Excel. Повторю через 10 секунд.")
        time.sleep(10)
        wb.save(JOURNAL)


def load_stoplist():
    if not STOPLIST.exists():
        STOPLIST.write_text("# Адреса, которым писать нельзя (по одному в строке)\n", encoding="utf-8")
    return {l.strip().lower() for l in STOPLIST.read_text(encoding="utf-8").splitlines() if l.strip() and not l.startswith("#")}


def start_date():
    cfg = configparser.ConfigParser()
    if STATE.exists():
        cfg.read(STATE, encoding="utf-8")
    if not cfg.has_section("state"):
        cfg.add_section("state")
    if not cfg.get("state", "start", fallback=""):
        cfg.set("state", "start", date.today().isoformat())
        with open(STATE, "w", encoding="utf-8") as f:
            cfg.write(f)
    return date.fromisoformat(cfg.get("state", "start"))


def remember_refusal(text):
    cfg = configparser.ConfigParser(interpolation=None)
    if STATE.exists():
        cfg.read(STATE, encoding="utf-8")
    if not cfg.has_section("state"):
        cfg.add_section("state")
    cfg.set("state", "last_refusal", datetime.now().isoformat(timespec="minutes"))
    cfg.set("state", "last_refusal_text", text[:200].replace("%", "%%"))
    with open(STATE, "w", encoding="utf-8") as f:
        cfg.write(f)


def stats(c):
    j = load_journal()
    first = [r for r in j.values() if r["Письмо 1"]]
    n = len(first)
    st = lambda s: sum(r["Статус"] == s for r in first)
    replied, refused, bounced = st("ответил"), st("отказ"), st("не доставлено")
    limit, day_no = daily_limit(c)
    days = []
    for k in range(6, -1, -1):
        d = (date.today() - timedelta(days=k)).isoformat()
        days.append({"date": d,
                     "first": sum(r["Письмо 1"].startswith(d) for r in first),
                     "followup": sum(r["Напоминание"].startswith(d) for r in first)})
    cfg = configparser.ConfigParser(interpolation=None)
    if STATE.exists():
        cfg.read(STATE, encoding="utf-8")
    last = cfg.get("state", "last_refusal", fallback="") if cfg.has_section("state") else ""
    risk, why = 0, []
    if last and datetime.fromisoformat(last) > datetime.now() - timedelta(days=7):
        risk += 3
        why.append(f"почта отказывала в отправке {last[:10]}")
    if n >= 20 and bounced / n > 0.05:
        risk += 2
        why.append("больше 5% адресов не существуют — чистите базу")
    elif n >= 20 and bounced / n > 0.02:
        risk += 1
        why.append("2–5% адресов не существуют")
    if n >= 50 and replied / n < 0.01:
        risk += 1
        why.append("почти никто не отвечает — письма могут уходить в спам, сделайте тест на себя")
    if limit >= 50 and day_no < 28:
        risk += 1
        why.append("большой лимит для молодого ящика")
    level = "низкий" if risk == 0 else "средний" if risk <= 2 else "высокий"
    return {
        "sent": n, "followups": sum(bool(r["Напоминание"]) for r in first),
        "replied": replied, "refused": refused, "bounced": bounced, "waiting": st("отправлено"),
        "limit": limit, "day": day_no + 1, "today": sent_today(j), "days": days,
        "risk": level, "why": why,
    }


def report(c):
    s = stats(c)
    n = s["sent"]
    if not n:
        say("Пока ничего не отправлено.")
        return
    pct = lambda x: f"{x / n * 100:.1f}%"
    say("=" * 52)
    say(f"Отправлено первых писем: {n}   напоминаний: {s['followups']}")
    say(f"Ответили: {s['replied']} ({pct(s['replied'])})   из них «нет»: {s['refused']}")
    say(f"Не доставлено (адрес не существует): {s['bounced']} ({pct(s['bounced'])})")
    say(f"Ждут ответа: {s['waiting']}")
    say(f"День прогрева: {s['day']}, лимит сегодня: {s['limit']}, отправлено сегодня: {s['today']}")
    say("-" * 52)
    say("По дням (последние 7):")
    for d in s["days"]:
        if d["first"] or d["followup"]:
            say(f"  {d['date']}: писем {d['first']}, напоминаний {d['followup']}")
    say("-" * 52)
    say(f"Риск блокировки (оценка программы): {s['risk']}")
    for w in s["why"]:
        say("  • " + w)
    if s["risk"] == "высокий":
        say("  Совет: 2–3 дня не отправлять, затем снизить max_per_day вдвое.")
    say("=" * 52)


def daily_limit(c):
    days = (date.today() - start_date()).days
    want = min(int(c.get("max_per_day", "50") or 50), HARD_CAP)
    for until, limit in WARMUP:
        if days < until:
            return min(limit, want), days
    return want, days


def sent_today(j):
    today = date.today().isoformat()
    return sum(rec["Письмо 1"].startswith(today) for rec in j.values()) + sum(rec["Напоминание"].startswith(today) for rec in j.values())


# ---------------------------------------------------------------- letters

def read_template(name):
    text = (HERE / name).read_text(encoding="utf-8").replace("\r\n", "\n")
    head, _, body = text.partition("\n---\n")
    subject = head.replace("Тема:", "", 1).strip()
    return subject, body.strip("\n")


def fill(text, lead, c):
    if not lead["company"]:
        text = text.replace(", {Компания}", "")
    values = {
        "{Компания}": lead["company"],
        "{Магазин}": lead["shop"] or "на Wildberries",
        "{Имя}": c.get("your_name", ""),
        "{Telegram}": c.get("telegram", ""),
        "{Телефон}": c.get("phone", ""),
    }
    for k, v in values.items():
        text = text.replace(k, v)
    return text


def photos():
    files = sorted(p for p in (HERE / "photos").glob("*") if p.suffix.lower() in (".jpg", ".jpeg", ".png"))
    total = sum(p.stat().st_size for p in files)
    if total > 3 * 1024 * 1024:
        say(f"Внимание: фото весят {total // 1024} КБ. Для холодных писем лучше до 1,5 МБ всего.")
    return files


def build_message(c, lead, template, reply_to_id=None, reply_subject=None):
    subject_t, body_t = read_template(template)
    subject = reply_subject or fill(subject_t, lead, c)
    body = fill(body_t, lead, c)
    pics = photos() if "{ФОТО}" in body else []

    msg = EmailMessage()
    msg["From"] = email.utils.formataddr((c.get("from_name") or c.get("your_name") or "", c["email"]))
    msg["To"] = lead["email"]
    msg["Subject"] = subject
    msg["Date"] = email.utils.formatdate(localtime=True)
    domain = c["email"].split("@")[-1]
    msg["Message-ID"] = email.utils.make_msgid(domain=domain)
    if reply_to_id:
        msg["In-Reply-To"] = reply_to_id
        msg["References"] = reply_to_id

    plain = body.replace("{ФОТО}", "(примеры работ — во вложении к письму)")
    msg.set_content(plain)

    cids = [email.utils.make_msgid(domain=domain)[1:-1] for _ in pics]
    parts = []
    for block in body.split("\n\n"):
        if block.strip() == "{ФОТО}":
            for cid in cids:
                parts.append(f'<p><img src="cid:{cid}" alt="Пример работы" width="560" style="max-width:100%;height:auto;border-radius:6px"></p>')
        else:
            parts.append("<p>" + html.escape(block).replace("\n", "<br>") + "</p>")
    html_body = '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5;color:#222">' + "".join(parts) + "</div>"
    msg.add_alternative(html_body, subtype="html")
    html_part = msg.get_payload()[1]
    for p, cid in zip(pics, cids):
        sub = "png" if p.suffix.lower() == ".png" else "jpeg"
        html_part.add_related(p.read_bytes(), maintype="image", subtype=sub, cid=f"<{cid}>", filename=p.name)
    return msg


# ---------------------------------------------------------------- transport

class StopForToday(Exception):
    pass


def smtp_connect(c):
    host, port, sec = c["smtp_host"], int(c.get("smtp_port") or 465), c.get("security", "ssl").lower()
    ctx = ssl.create_default_context()
    if sec == "ssl":
        s = smtplib.SMTP_SSL(host, port, context=ctx, timeout=60)
    else:
        s = smtplib.SMTP(host, port, timeout=60)
        if sec == "starttls":
            s.starttls(context=ctx)
    if c["app_password"] != "-":
        s.login(c["email"], c["app_password"])
    return s


def send(c, msg):
    """Returns (ok, error_text). Raises StopForToday on server-side refusals."""
    try:
        s = smtp_connect(c)
    except smtplib.SMTPAuthenticationError:
        say("\n!!! Почта не приняла логин или пароль.")
        say("    Нужен ПАРОЛЬ ПРИЛОЖЕНИЯ, а не обычный пароль (см. инструкцию, шаг 3).")
        sys.exit(1)
    except OSError as e:  # includes other SMTP errors and network failures
        raise StopForToday(f"нет связи с почтовым сервером: {e}")
    try:
        s.send_message(msg)
        return True, ""
    except smtplib.SMTPRecipientsRefused as e:
        return False, "адрес не принят: " + str(list(e.recipients.values())[0])[:200]
    except smtplib.SMTPDataError as e:
        text = f"{e.smtp_code} {e.smtp_error!r}"
        raise StopForToday(text)
    except smtplib.SMTPSenderRefused as e:
        raise StopForToday(f"{e.smtp_code} {e.smtp_error!r}")
    finally:
        try:
            s.quit()
        except Exception:
            pass


def imap_check(c, j):
    """Mark replies and bounces using the INBOX. Best effort."""
    pending = {m: r for m, r in j.items() if r["Статус"] == "отправлено"}
    if not pending or not c.get("imap_host") or c.get("check_replies", "yes").lower() != "yes":
        return
    oldest = min(r["Письмо 1"][:10] for r in pending.values())
    since = datetime.fromisoformat(oldest).strftime("%d-%b-%Y")
    try:
        im = imaplib.IMAP4_SSL(c["imap_host"], 993, ssl_context=ssl.create_default_context())
        im.login(c["email"], c["app_password"])
        im.select("INBOX", readonly=True)
        typ, data = im.search(None, "SINCE", since)
        ids = data[0].split()[-800:]
        replied, bounced = 0, 0
        for mid in ids:
            typ, d = im.fetch(mid, "(BODY.PEEK[HEADER.FIELDS (FROM)] BODY.PEEK[TEXT]<0.6000>)")
            raw_from, raw_text = b"", b""
            for item in d:
                if isinstance(item, tuple):
                    if b"HEADER.FIELDS" in item[0]:
                        raw_from = item[1]
                    else:
                        raw_text = item[1]
            frm = str(make_header(decode_header(raw_from.decode("utf-8", "ignore")))).lower()
            addr = email.utils.parseaddr(frm.replace("from:", "").strip())[1]
            if addr in pending:
                pending[addr]["Статус"] = "ответил"
                replied += 1
            elif re.search(r"mailer-daemon|postmaster|devnull", frm):
                text = raw_text.decode("utf-8", "ignore").lower()
                for m in list(pending):
                    if m in text and pending[m]["Статус"] == "отправлено":
                        pending[m]["Статус"] = "не доставлено"
                        bounced += 1
        im.logout()
        if replied or bounced:
            say(f"Проверка почты: ответили {replied}, не доставлено {bounced}.")
    except Exception as e:
        say(f"(Не удалось проверить входящие: {e}. Напоминания отправлю только тем, у кого статус «отправлено».)")


# ---------------------------------------------------------------- main

def build_queue(c, j, leads, stop):
    """Follow-ups that are due first, then new addresses."""
    follow_days = int(c.get("followup_after_days", "3") or 3)
    queue = []
    for rec in j.values():
        if rec["Статус"] == "отправлено" and not rec["Напоминание"] and rec["Письмо 1"]:
            if datetime.fromisoformat(rec["Письмо 1"]) <= datetime.now() - timedelta(days=follow_days):
                if rec["Email"].lower() not in stop:
                    queue.append(("followup", rec))
    for lead in leads:
        m = lead["email"]
        if m in j or m in stop or not EMAIL_RE.match(m) or lead["lead_status"] in ("Отказ", "Не подходит"):
            continue
        queue.append(("first", lead))
    return queue


def within_hours(c):
    try:
        a, b = [int(x) for x in c.get("work_hours", "9-19").split("-")]
    except ValueError:
        a, b = 9, 19
    return a <= datetime.now().hour < b


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--preview", action="store_true")
    ap.add_argument("--test", nargs="?", const="", default=None, metavar="ADDRESS")
    ap.add_argument("--report", action="store_true")
    args = ap.parse_args()

    c = load_config(strict=not args.report)
    if args.report:
        report(c)
        return
    leads_path = find_leads_file(c)
    leads = read_leads(leads_path)
    say(f"База: {leads_path.name} — продавцов с почтой: {len(leads)}")

    if args.preview:
        out = HERE / "preview"
        out.mkdir(exist_ok=True)
        good = [l for l in leads if EMAIL_RE.match(l["email"]) and l["lead_status"] not in ("Отказ", "Не подходит")]
        for lead in good[:5]:
            msg = build_message(c, lead, "letter.txt")
            (out / f"{lead['email']}.eml").write_bytes(bytes(msg))
        msg = build_message(c, good[0], "followup.txt", reply_subject="Re: preview") if good else None
        if msg:
            (out / "_напоминание.eml").write_bytes(bytes(msg))
        say(f"Готово: примеры писем в папке preview ({min(5, len(good))} шт.). Откройте .eml двойным щелчком.")
        return

    if args.test is not None:
        if not leads:
            say("В базе нет ни одного адреса.")
            return
        to = (args.test or c["email"]).strip()
        if not EMAIL_RE.match(to):
            say(f"Адрес «{to}» записан с ошибкой. Запустите снова и проверьте адрес (или просто нажмите Enter).")
            return
        lead = dict(leads[0], email=to)
        try:
            ok, err = send(c, build_message(c, lead, "letter.txt"))
        except StopForToday as e:
            remember_refusal(str(e))
            say("\n!!! Почта не выпустила письмо:", e)
            if "SPAM" in str(e).upper():
                say("    Фильтр Яндекса посчитал письмо похожим на рекламу.")
                say("    Что сделать: оставьте в letter.txt короткий текст без фото и цен, привяжите телефон")
                say("    к аккаунту, отправьте вручную из браузера 3–5 обычных писем знакомым и повторите тест завтра.")
            return
        say("Тестовое письмо отправлено на " + to if ok else "Ошибка: " + err)
        return

    if not within_hours(c):
        say(f"Сейчас нерабочее время ({c.get('work_hours', '9-19')}). Письма лучше отправлять днём — запустите позже.")
        return

    j = load_journal()
    stop = load_stoplist()
    imap_check(c, j)
    save_journal(j)

    limit, day_no = daily_limit(c)
    left = limit - sent_today(j)
    say(f"День прогрева: {day_no + 1}. Лимит на сегодня: {limit}. Осталось: {max(left, 0)}.")
    if left <= 0:
        say("На сегодня всё. Запустите завтра.")
        return

    queue = build_queue(c, j, leads, stop)[:left]
    if not queue:
        say("Новых адресов нет. Выгрузите свежую таблицу из расширения.")
        return

    STOP_FILE.unlink(missing_ok=True)
    dmin, dmax = int(c.get("delay_min_sec", "60") or 60), int(c.get("delay_max_sec", "240") or 240)
    say(f"Отправлю: {len(queue)}. Пауза между письмами {dmin}–{dmax} сек. Остановить: кнопка «Остановить» в приложении (или Ctrl+C).\n")
    try:
        for n, (kind, item) in enumerate(queue, 1):
            now = datetime.now().isoformat(timespec="minutes")
            if kind == "first":
                lead = item
                msg = build_message(c, lead, "letter.txt")
                ok, err = send(c, msg)
                j[lead["email"]] = {
                    "Email": lead["email"], "Компания": lead["company"], "Магазин": lead["shop"], "ИНН": lead["inn"],
                    "Письмо 1": now, "Message-ID": msg["Message-ID"], "Напоминание": "",
                    "Статус": "отправлено" if ok else "не доставлено", "Ошибка": err, "Комментарий": "",
                }
                label = f"письмо → {lead['email']} ({lead['company'] or lead['shop']})"
            else:
                rec = item
                lead = {"email": rec["Email"], "company": rec["Компания"], "shop": rec["Магазин"], "inn": rec["ИНН"]}
                subj, _ = read_template("letter.txt")
                msg = build_message(c, lead, "followup.txt", reply_to_id=rec["Message-ID"] or None,
                                    reply_subject="Re: " + fill(subj, lead, c))
                ok, err = send(c, msg)
                rec["Напоминание"] = now
                if not ok:
                    rec["Статус"], rec["Ошибка"] = "не доставлено", err
                label = f"напоминание → {rec['Email']}"
            save_journal(j)
            say(f"[{n}/{len(queue)}] {'✓' if ok else '✗'} {label}{'' if ok else ' — ' + err}")
            if n < len(queue):
                pause = random.randint(dmin, dmax)
                say(f"    пауза {pause} сек…")
                for _ in range(pause):
                    if STOP_FILE.exists():
                        raise KeyboardInterrupt
                    time.sleep(1)
            if STOP_FILE.exists():
                raise KeyboardInterrupt
    except StopForToday as e:
        save_journal(j)
        remember_refusal(str(e))
        say("\n!!! Почтовый сервер отказался принимать письма:", e)
        say("    Отправка остановлена до завтра. Не создавайте новые ящики — проверьте текст письма,")
        say("    уменьшите max_per_day в config.ini и посмотрите статистику в Яндекс Постмастере.")
        return
    except KeyboardInterrupt:
        save_journal(j)
        STOP_FILE.unlink(missing_ok=True)
        say("\nОстановлено. Журнал сохранён.")
        return
    say("\nГотово на сегодня. Журнал: journal.xlsx")


if __name__ == "__main__":
    main()
