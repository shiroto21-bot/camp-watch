#!/usr/bin/env python3
"""城南島海浜公園キャンプ場(なっぷ)の空きカレンダーを監視 → 空きがあればメール通知

使い方:
  python check_camp.py --dates 2026-11-28            # 11/28の1泊
  python check_camp.py --dates 2026-11-28:2026-11-29 # チェックイン:チェックアウト形式でもOK
  python check_camp.py --dates 2026-11-28,2026-12-05 # 複数日はカンマ区切り
  python check_camp.py --dump --dates 2026-11-28     # 読み取った表を表示(動作確認用)
  環境変数 TARGETS でも日付指定可
"""
import argparse, json, os, re, smtplib, subprocess, sys, time
from datetime import date, datetime, timedelta
from email.header import Header
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ---- 設定(環境変数で上書き可) -------------------------------------------
FACILITY_ID = "12464"
PAGE_URL = f"https://www.nap-camp.com/tokyo/{FACILITY_ID}"
CALENDAR_URL = f"https://www.nap-camp.com/widget/calendar/{FACILITY_ID}?m_shift={{shift}}"
# 監視するプラン名(部分一致)。カンマ区切りで複数可。
# 既定は「【宿泊】」を含む全プラン(オートサイト・第一キャンプ場)
PLAN_FILTER = [s.strip() for s in (os.getenv("PLAN_FILTER") or "【宿泊】").split(",") if s.strip()]
# 通知の送信先(環境変数 MAIL_TO で変更。複数はカンマ区切り)
MAIL_TO = os.getenv("MAIL_TO") or "sabre.mic.o@gmail.com"
# 送信方式: smtp(Gmail) / outlook(Windowsのデスクトップ版Outlook経由)
MAIL_BACKEND = (os.getenv("MAIL_BACKEND") or "smtp").lower()
MAIL_SUBJECT = "城南島海浜公園キャンプ場 予約できるよ！"
MAIL_BODY = f"予約ページ\n{PAGE_URL}"
# 曜日指定モードで先読みする月数(1ページで3か月分見えるので 3 刻み)
MONTHS_AHEAD = int(os.getenv("MONTHS_AHEAD") or 6)
COOLDOWN_MIN = int(os.getenv("COOLDOWN_MIN") or 60)               # 空き継続中の再通知間隔
UNKNOWN_ALERT_AFTER = int(os.getenv("UNKNOWN_ALERT_AFTER") or 6)  # 判定不能が連続したら異常通知
STATE_FILE = Path(os.getenv("STATE_FILE", "state.json"))
# ---------------------------------------------------------------------------

AVAILABLE = set("○◯〇△▲")           # 空きあり(◯=余裕あり, △/▲=残りわずか)
UNAVAILABLE = set("×✕✖xX-−ー―–—")   # 満室・受付なし
UA = {"User-Agent": "Mozilla/5.0 (camp-watch personal use)"}


PS_SCRIPT = r"""
$ol = New-Object -ComObject Outlook.Application
$m = $ol.CreateItem(0)
$to = $env:CW_TO
if (-not $to) { $to = $ol.Session.CurrentUser.Name }   # 宛先未指定なら自分宛て
$m.To = $to
$m.Subject = $env:CW_SUBJECT
$m.Body = $env:CW_BODY
$m.Send()
"""


def send_mail_outlook(subject, body):
    """デスクトップ版Outlook(COM)経由で送信。会社のメールサーバーから送られる"""
    to = (os.getenv("MAIL_TO") or "").replace(",", ";")   # 未設定なら自分宛て
    try:  # pywin32 があればそれを使う
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application")
        m = ol.CreateItem(0)
        m.To = to or ol.Session.CurrentUser.Name
        m.Subject, m.Body = subject, body
        m.Send()
        return
    except ImportError:
        pass
    env = dict(os.environ, CW_TO=to, CW_SUBJECT=subject, CW_BODY=body)
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", PS_SCRIPT],
        env=env, capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError("Outlook経由の送信に失敗: " + (r.stderr or r.stdout).strip())


def send_mail(subject, body):
    if MAIL_BACKEND == "outlook":
        return send_mail_outlook(subject, body)
    user = os.environ["GMAIL_USER"]
    pw = os.environ["GMAIL_APP_PASSWORD"]
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = user
    recipients = [x.strip() for x in MAIL_TO.split(",") if x.strip()]
    msg["To"] = ", ".join(recipients)
    errors = []
    # まず465(SSL)、だめなら587(STARTTLS)で送信を試みる
    for mode, port in (("ssl", 465), ("starttls", 587)):
        try:
            if mode == "ssl":
                s = smtplib.SMTP_SSL("smtp.gmail.com", port, timeout=30)
            else:
                s = smtplib.SMTP("smtp.gmail.com", port, timeout=30)
                s.starttls()
            with s:
                s.login(user, pw)
                s.send_message(msg, to_addrs=recipients)
            return
        except Exception as e:
            errors.append(f"{mode}:{port} -> {e!r}")
    raise RuntimeError("メール送信に失敗しました: " + " / ".join(errors))


def month_index(d):
    return d.year * 12 + d.month


def fetch_html(shift):
    url = CALENDAR_URL.format(shift=shift)
    last = None
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=30)
            r.raise_for_status()
            return r.content
        except Exception as e:
            last = e
            time.sleep(3)
    raise last


def expand_table(table):
    """rowspan/colspan を展開して、行×列の文字列グリッドにする"""
    grid, pending = [], {}   # pending[(r,c)] = text
    for r, tr in enumerate(table.find_all("tr")):
        row, c = [], 0
        cells = tr.find_all(["td", "th"], recursive=False) or tr.find_all(["td", "th"])
        def fill_pending():
            nonlocal c
            while (r, c) in pending:
                row.append(pending.pop((r, c))); c += 1
        for cell in cells:
            fill_pending()
            text = re.sub(r"\s+", "", cell.get_text())
            rs = int(cell.get("rowspan", 1) or 1)
            cs = int(cell.get("colspan", 1) or 1)
            for dc in range(cs):
                row.append(text)
                for dr in range(1, rs):
                    pending[(r + dr, c)] = text
                c += 1
        fill_pending()
        grid.append(row)
    return grid


def parse_calendar(html, base_month_first):
    """return {プラン名: {date: 記号}}"""
    soup = BeautifulSoup(html, "html.parser")
    for table in soup.find_all("table"):
        grid = expand_table(table)
        hdr = next((i for i, row in enumerate(grid) if row and row[0] == "日付"), None)
        if hdr is None:
            continue
        # --- 日付ヘッダ → 列ごとの date ---
        cur_y, cur_m, last_day = base_month_first.year, base_month_first.month, 0
        col_dates = {}
        for ci, txt in enumerate(grid[hdr]):
            if ci == 0:
                continue
            m = re.fullmatch(r"(\d{1,2})/(\d{1,2})", txt)       # 例: 12/01
            sep = re.fullmatch(r"(\d{1,2})月", txt)              # 例: 12月(月の区切り列)
            if m or sep:
                nm = int((m or sep).group(1))
                if nm < cur_m:
                    cur_y += 1
                cur_m = nm
                last_day = 0
                if sep:
                    continue
                day = int(m.group(2))
            elif re.fullmatch(r"\d{1,2}", txt):
                day = int(txt)
                if day < last_day:                                # 区切り列なしで月が替わった場合
                    cur_m += 1
                    if cur_m > 12:
                        cur_m, cur_y = 1, cur_y + 1
            else:
                continue
            last_day = day
            try:
                col_dates[ci] = date(cur_y, cur_m, day)
            except ValueError:
                pass
        # --- プラン行 ---
        result = {}
        for row in grid[hdr + 1:]:
            if not row or not row[0] or row[0] in ("曜日", "日付"):
                continue
            result[row[0]] = {d: row[ci] for ci, d in col_dates.items() if ci < len(row)}
        if result:
            return result
    return {}


def nights_of(spec):
    """'2026-11-28' or '2026-11-28:2026-11-29' → 泊まる日のリスト"""
    parts = spec.split(":")
    ci = datetime.strptime(parts[0], "%Y-%m-%d").date()
    co = datetime.strptime(parts[1], "%Y-%m-%d").date() if len(parts) > 1 else ci + timedelta(days=1)
    return [ci + timedelta(days=i) for i in range((co - ci).days)] or [ci]


def check(spec, dump=False):
    """return ('available'|'unavailable'|'unknown', [空きプラン名])"""
    nights = nights_of(spec)
    today = date.today()
    base_shift = month_index(nights[0]) - month_index(today)
    tables = None
    for shift in (base_shift, base_shift - 1):     # 見つからなければ1つ前の月からも探す
        if shift < 0:
            continue
        first = date(today.year + (today.month - 1 + shift) // 12, (today.month - 1 + shift) % 12 + 1, 1)
        tables = parse_calendar(fetch_html(shift), first)
        if dump:
            for plan, days in tables.items():
                print(plan, " ".join(f"{d.month}/{d.day}{s}" for d, s in sorted(days.items())))
        if tables and any(n in next(iter(tables.values())) for n in nights):
            break
    if not tables:
        return "unknown", []
    plans = [p for p in tables if any(f in p for f in PLAN_FILTER)]
    if not plans:
        return "unknown", []
    ok_plans, unknown = [], False
    for p in plans:
        syms = [tables[p].get(n) for n in nights]
        if all(s and s[0] in AVAILABLE for s in syms):
            ok_plans.append(p)
        elif any(s is None or (s and s[0] not in AVAILABLE | UNAVAILABLE) for s in syms):
            unknown = True
    if ok_plans:
        return "available", ok_plans
    return ("unknown" if unknown else "unavailable"), []


WEEKDAY_NAMES = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6,
                 "月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}
WD_LABEL = "月火水木金土日"


def parse_weekdays(text):
    out = set()
    for t in re.split(r"[,\s、]+", (text or "").strip().lower()):
        if t in WEEKDAY_NAMES:
            out.add(WEEKDAY_NAMES[t])
        elif t[:3] in WEEKDAY_NAMES:
            out.add(WEEKDAY_NAMES[t[:3]])
    return out


def find_open_weekdays(weekdays, dump=False):
    """今日以降で、指定曜日に空きのある日を探す。return ('ok'|'unknown', {date: [プラン名]})"""
    today = date.today()
    merged = {}
    for shift in range(0, MONTHS_AHEAD, 3):
        first = date(today.year + (today.month - 1 + shift) // 12, (today.month - 1 + shift) % 12 + 1, 1)
        try:
            tables = parse_calendar(fetch_html(shift), first)
        except Exception:
            if shift == 0:
                raise
            continue   # 先の月がまだ表示できない場合は無視
        for plan, days in tables.items():
            merged.setdefault(plan, {}).update(days)
    plans = [p for p in merged if any(f in p for f in PLAN_FILTER)]
    if not plans:
        return "unknown", {}
    found = {}
    for p in plans:
        for d, sym in sorted(merged[p].items()):
            if d >= today and d.weekday() in weekdays and sym and sym[0] in AVAILABLE:
                found.setdefault(d, []).append(p)
    if dump:
        for d in sorted(found):
            print(f"  {d} ({WD_LABEL[d.weekday()]}): {found[d]}")
    return "ok", dict(sorted(found.items()))


def load_state():
    try:
        return json.loads(STATE_FILE.read_text())
    except Exception:
        return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dates", default=os.getenv("TARGETS", ""))
    ap.add_argument("--weekday", default=os.getenv("WEEKDAY", ""),
                    help="曜日指定。例: sat / sat,sun / 土  (今日以降で空いている該当曜日を通知)")
    ap.add_argument("--no-mail", action="store_true", help="通知せず判定のみ")
    ap.add_argument("--dump", action="store_true", help="読み取った表や空き日を表示")
    ap.add_argument("--test-mail", action="store_true", help="テストメールだけ送って終了")
    a = ap.parse_args()
    if a.test_mail:
        send_mail("【テスト】" + MAIL_SUBJECT, MAIL_BODY)
        print("テストメールを送信しました")
        return
    specs = [x.strip() for x in a.dates.split(",") if x.strip()]
    weekdays = parse_weekdays(a.weekday)
    if not specs and not weekdays:
        sys.exit("日付か曜日が未指定です (--dates 2026-11-28 / --weekday sat)")

    state, now, unknown_count = load_state(), time.time(), 0

    # ---- 日付指定モード ----
    for spec in specs:
        try:
            result, plans = check(spec, a.dump)
        except Exception as e:
            print(f"error: {e}")
            result, plans = "unknown", []
        print(f"{datetime.now():%F %T} {spec}: {result} {plans if plans else ''}")
        if result == "available":
            if now - state.get(spec, 0) >= COOLDOWN_MIN * 60:
                if a.no_mail:
                    print("  (--no-mail) 通知スキップ")
                else:
                    send_mail(MAIL_SUBJECT, MAIL_BODY)
                    print("  メール送信しました")
                state[spec] = now
        else:
            state.pop(spec, None)
            if result == "unknown":
                unknown_count += 1

    # ---- 曜日指定モード ----
    if weekdays:
        names = "".join(WD_LABEL[i] for i in sorted(weekdays))
        try:
            result, found = find_open_weekdays(weekdays, a.dump)
        except Exception as e:
            print(f"error: {e}")
            result, found = "unknown", {}
        if result == "unknown":
            unknown_count += 1
            print(f"{datetime.now():%F %T} 曜日[{names}]: unknown")
        else:
            lines = [f"{d:%Y-%m-%d}({WD_LABEL[d.weekday()]}) {' / '.join(ps)}" for d, ps in found.items()]
            print(f"{datetime.now():%F %T} 曜日[{names}]: 空き{len(found)}日 {', '.join(f'{d:%m/%d}' for d in found)}")
            keys = {f"wd:{d}" for d in found}
            for k in [k for k in state if k.startswith("wd:") and k not in keys]:
                state.pop(k)                      # 埋まった日はリセット → また空いたら即通知
            due = [d for d in found if now - state.get(f"wd:{d}", 0) >= COOLDOWN_MIN * 60]
            if due:
                body = MAIL_BODY + "\n\n空きがある日:\n" + "\n".join(lines)
                if a.no_mail:
                    print("  (--no-mail) 通知スキップ")
                else:
                    send_mail(MAIL_SUBJECT, body)
                    print("  メール送信しました")
                for d in due:
                    state[f"wd:{d}"] = now

    if unknown_count:
        state["unknown_streak"] = state.get("unknown_streak", 0) + 1
    else:
        state["unknown_streak"] = 0
        state.pop("unknown_alerted", None)
    if (state["unknown_streak"] >= UNKNOWN_ALERT_AFTER
            and not state.get("unknown_alerted") and not a.no_mail):
        send_mail("【監視異常】城南島キャンプ場チェッカー",
                  "空き状況を判定できない状態が続いています。\n"
                  "なっぷのカレンダー仕様変更の可能性があります。")
        state["unknown_alerted"] = True
    STATE_FILE.write_text(json.dumps(state))


if __name__ == "__main__":
    main()
