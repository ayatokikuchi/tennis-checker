import datetime, time, requests, json, re, smtplib, os, jpholiday
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait

CONFIG = {
    "TO_EMAIL"    : "ayato.kikuchi@accenture.com",
    "FROM_EMAIL"  : os.environ["FROM_EMAIL"],
    "APP_PASSWORD": os.environ["APP_PASSWORD"],
}

TOKYO_PARKS = {
    "1280": ("東大和南公園", "12800030", "人工芝"),  # ※稼働確認用
}

MINATO_PARKS = {}  # ※稼働確認中は港区除外

TOKYO_AJAX_URL  = "https://kouen.sports.metro.tokyo.lg.jp/web/rsvWOpeInstSrchMonthVacantAjaxAction.do"
MINATO_AJAX_URL = "https://web101.rsv.ws-scs.jp/web/rsvWOpeInstSrchMonthVacantAjaxAction.do"
TOKYO_WEEK_URL  = "https://kouen.sports.metro.tokyo.lg.jp/web/rsvWOpeInstSrchVacantAjaxAction.do"
MINATO_WEEK_URL = "https://web101.rsv.ws-scs.jp/web/rsvWOpeInstSrchVacantAjaxAction.do"


def get_holiday_and_weekend_dates(base_date):
    dates = []
    for month_offset in [0, 1]:
        year  = (base_date.replace(day=1) + datetime.timedelta(days=31 * month_offset)).year
        month = (base_date.replace(day=1) + datetime.timedelta(days=31 * month_offset)).month
        last_day = (datetime.date(year, month % 12 + 1, 1) - datetime.timedelta(days=1)).day if month < 12 else 31
        for day in range(1, last_day + 1):
            d = datetime.date(year, month, day)
            if d < base_date:
                continue
            if d.weekday() >= 5 or jpholiday.is_holiday(d):
                dates.append(d)
    return dates


def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(options=options)


def wait_page(driver, t=3):
    WebDriverWait(driver, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    time.sleep(t)


def parse_ajax(resp_text):
    raw = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', resp_text)
    return json.loads(raw)


def status_to_disp(statuses):
    if not statuses:
        return "⚠️ 情報なし", "unknown"
    if any(s == 100 for s in statuses):
        return "▲ 一部空きあり", "partial"
    if all(s == 700 for s in statuses):
        return "－ 営業時間外", "closed"
    return "❌ 満杯", "full"


def get_session():
    print("東京都セッション確立中...")
    driver = get_driver()
    try:
        driver.get("https://kouen.sports.metro.tokyo.lg.jp/web/")
        time.sleep(4)
        driver.execute_script("doAction(document.form1, gRsvWOpeInstSrchVacantAction)")
        time.sleep(5)
        driver.execute_script("""
            var sel = document.getElementById('purpose');
            for (var i = 0; i < sel.options.length; i++) {
                if (sel.options[i].text === 'テニス（人工芝）') {
                    sel.selectedIndex = i;
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                    break;
                }
            }
        """)
        time.sleep(3)
        driver.execute_script("""
            var sel = document.getElementById('bname');
            for (var i = 0; i < sel.options.length; i++) {
                if (sel.options[i].value === '1140') {
                    sel.selectedIndex = i;
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                    break;
                }
            }
        """)
        time.sleep(3)
        driver.execute_script("doSearch(document.form1, gRsvWOpeInstSrchVacantAction)")
        time.sleep(6)
        tokyo_cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        tokyo_referer = driver.current_url
        print("✅ 東京都セッション確立完了")
    finally:
        driver.quit()

    print("港区セッション確立中...")
    driver = get_driver()
    try:
        driver.get("https://web101.rsv.ws-scs.jp/web/")
        time.sleep(4)
        driver.execute_script("doAction(document.form1, gRsvWOpeInstSrchVacantAction)")
        time.sleep(5)
        driver.execute_script("""
            var sel = document.getElementById('purpose') || document.querySelector('[name=purpose]');
            for (var i = 0; i < sel.options.length; i++) {
                if (sel.options[i].text === 'テニス') {
                    sel.selectedIndex = i;
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                    break;
                }
            }
        """)
        time.sleep(3)
        driver.execute_script("""
            var sel = document.getElementById('bname') || document.querySelector('[name=bname]');
            for (var i = 0; i < sel.options.length; i++) {
                if (sel.options[i].value === '1000_70100') {
                    sel.selectedIndex = i;
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                    break;
                }
            }
        """)
        time.sleep(3)
        driver.execute_script("doSearch(document.form1, gRsvWOpeInstSrchVacantAction)")
        time.sleep(6)
        minato_cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        minato_referer = driver.current_url
        print("✅ 港区セッション確立完了")
    finally:
        driver.quit()

    return tokyo_cookies, tokyo_referer, minato_cookies, minato_referer


def fetch_vacancy(ajax_url, bld_cd, inst_cd, use_day, cookies, referer, retry=2):
    payload = {
        "displayNo": "prwrc2000", "useDay": use_day,
        "bldCd": bld_cd, "instCd": inst_cd, "transVacantMode": "0",
    }
    headers = {"Referer": referer, "X-Requested-With": "XMLHttpRequest"}
    for attempt in range(retry):
        try:
            resp = requests.post(ajax_url, data=payload, cookies=cookies,
                                 headers=headers, timeout=30)
            return parse_ajax(resp.text).get("result", [])
        except Exception as e:
            print(f"  Ajax失敗 bldCd={bld_cd} (試行{attempt+1}/{retry}): {e}")
            time.sleep(3)
    return []


def fetch_timeslots(week_url, bld_cd, inst_cd, use_day, cookies, referer):
    payload = {
        "displayNo": "prwrc2000", "useDay": use_day,
        "bldCd": bld_cd, "instCd": inst_cd, "transVacantMode": "0",
    }
    headers = {"Referer": referer, "X-Requested-With": "XMLHttpRequest"}
    try:
        resp = requests.post(week_url, data=payload, cookies=cookies,
                             headers=headers, timeout=15)
        data = parse_ajax(resp.text)
        ymd  = int(use_day)
        slots_stable = set()  # rsvNum >= 2
        slots_tight  = set()  # rsvNum == 1
        for zone in data.get("result", []):
            for tr in zone.get("timeResult", []):
                if tr.get("useDay") == ymd and tr.get("status") == 0:
                    st      = tr.get("startTime", 0)
                    et      = tr.get("endTime", 0)
                    rsv_num = tr.get("rsvNum", 0)
                    slot    = f"{st//100}:{st%100:02d}〜{et//100}:{et%100:02d}"
                    if rsv_num >= 2:
                        slots_stable.add(slot)
                    elif rsv_num == 1:
                        slots_tight.add(slot)
        return sorted(slots_stable), sorted(slots_tight)
    except Exception as e:
        print(f"  時間帯取得失敗 bldCd={bld_cd} {use_day}: {e}")
        return [], []


def check_tokyo(target_dates, cookies, referer):
    results = []
    use_day = target_dates[0].strftime("%Y%m%d")
    target_ymds = {int(d.strftime("%Y%m%d")) for d in target_dates}
    for bld_cd, (park_name, inst_cd, court_type) in TOKYO_PARKS.items():
        print(f"  チェック中: {park_name}（{court_type}）")
        day_results = fetch_vacancy(TOKYO_AJAX_URL, bld_cd, inst_cd, use_day, cookies, referer)
        status_map = {r["dayYMD"]: r["status"] for r in day_results}
        for date in target_dates:
            ymd = int(date.strftime("%Y%m%d"))
            weekday = ["月","火","水","木","金","土","日"][date.weekday()]
            holiday = jpholiday.is_holiday_name(date)
            label = f"{date.strftime('%-m月%-d日')}（{weekday}）{' ' + holiday if holiday else ''}"
            status = status_map.get(ymd)
            if status == 100:
                slots_stable, slots_tight = fetch_timeslots(
                    TOKYO_WEEK_URL, bld_cd, inst_cd,
                    date.strftime("%Y%m%d"), cookies, referer)
                if slots_stable:
                    slot_str = "　".join(slots_stable)
                    disp = f"✅ 空きあり（{slot_str}）"
                    ck   = "partial"
                else:
                    # rsvNum=1は信頼性が低いため除外
                    disp = "❌ 満杯"
                    ck   = "full"
            elif status == 200:
                disp, ck = "❌ 満杯", "full"
            elif status == 700:
                disp, ck = "－ 営業時間外", "closed"
            else:
                disp, ck = "⚠️ 情報なし", "unknown"
            results.append({
                "site": f"東京都 {park_name}（{court_type}）",
                "date": label, "status": disp, "color_key": ck,
                "sort_key": date,
                "url": "https://kouen.sports.metro.tokyo.lg.jp/web/rsvWOpeInstSrchVacantAction.do"
            })
        time.sleep(0.5)
    return results


def check_minato(target_dates, cookies, referer):
    results = []
    use_day = target_dates[0].strftime("%Y%m%d")
    target_ymds = {int(d.strftime("%Y%m%d")) for d in target_dates}
    for bld_cd, (park_name, inst_cds) in MINATO_PARKS.items():
        print(f"  チェック中: 港区 {park_name}")
        date_status = {}
        for inst_cd in inst_cds:
            day_results = fetch_vacancy(MINATO_AJAX_URL, bld_cd, inst_cd, use_day, cookies, referer)
            for r in day_results:
                ymd = r.get("dayYMD")
                if ymd in target_ymds:
                    date_status.setdefault(ymd, []).append(r.get("status"))
            time.sleep(0.3)
        for date in target_dates:
            ymd = int(date.strftime("%Y%m%d"))
            weekday = ["月","火","水","木","金","土","日"][date.weekday()]
            holiday = jpholiday.is_holiday_name(date)
            label = f"{date.strftime('%-m月%-d日')}（{weekday}）{' ' + holiday if holiday else ''}"
            disp, ck = status_to_disp(date_status.get(ymd, []))
            if ck == "partial":
                slots = fetch_timeslots(MINATO_WEEK_URL, bld_cd, inst_cds[0],
                                        date.strftime("%Y%m%d"), cookies, referer)
                if slots:
                    disp = "✅ 空きあり（" + "　".join(slots) + "）"
                else:
                    disp, ck = "❌ 満杯", "full"
            results.append({
                "site": f"港区 {park_name}",
                "date": label, "status": disp, "color_key": ck,
                "sort_key": date,
                "url": "https://web101.rsv.ws-scs.jp/web/rsvWOpeInstSrchVacantAction.do"
            })
    return results


def send_email(results, target_dates):
    today_str  = datetime.date.today().strftime("%Y年%m月%d日")
    date_range = f"{target_dates[0].strftime('%-m/%-d')}〜{target_dates[-1].strftime('%-m/%-d')}（土日祝）"
    color_map  = {"partial":"#d4edda","full":"#f0f0f0","closed":"#e9ecef","unknown":"#f8d7da"}

    partial = [r for r in results if r["color_key"] == "partial"]
    summary = f"▲ 一部空きあり：{len(partial)}件"

    tokyo_results  = sorted([r for r in results if r["site"].startswith("東京都")],
                             key=lambda r: r["sort_key"])
    minato_results = sorted([r for r in results if r["site"].startswith("港区")],
                             key=lambda r: r["sort_key"])

    def build_table(data, title, color):
        rows = ""
        for r in data:
            bg = color_map.get(r["color_key"], "#fff")
            rows += (
                "<tr style=\"background:" + bg + "\">"
                "<td style=\"padding:6px 8px;border:1px solid #ddd;font-size:12px;\">" + r["date"] + "</td>"
                "<td style=\"padding:6px 8px;border:1px solid #ddd;font-size:12px;\">" + r["site"] + "</td>"
                "<td style=\"padding:6px 8px;border:1px solid #ddd;font-size:12px;\">" + r["status"] + "</td>"
                "<td style=\"padding:6px 8px;border:1px solid #ddd;font-size:12px;\"><a href=\"" + r["url"] + "\">確認</a></td>"
                "</tr>"
            )
        return (
            "<h2 style=\"color:" + color + ";margin:20px 0 8px;font-size:16px;\">" + title + "</h2>"
            "<table style=\"width:100%;border-collapse:collapse;margin-bottom:24px;\">"
            "<tr style=\"background:" + color + ";color:#fff;\">"
            "<th style=\"padding:8px;border:1px solid #ddd;\">日付</th>"
            "<th style=\"padding:8px;border:1px solid #ddd;\">施設</th>"
            "<th style=\"padding:8px;border:1px solid #ddd;\">状況・空き時間帯</th>"
            "<th style=\"padding:8px;border:1px solid #ddd;\">リンク</th>"
            "</tr>" + rows + "</table>"
        )

    html = (
        "<div style=\"font-family:Arial,sans-serif;max-width:800px;margin:auto;background:#f5f5f5;padding:20px;\">"
        "<div style=\"background:#2e7d32;padding:20px;border-radius:8px 8px 0 0;\">"
        "<h1 style=\"color:#fff;margin:0;font-size:20px;\">🎾 テニスコート空き状況（土日祝・当月〜翌月）</h1>"
        "<p style=\"color:#c8e6c9;margin:4px 0 0;\">" + today_str + " 配信　" + date_range + "　" + summary + "</p>"
        "</div>"
        "<div style=\"background:#fff;padding:20px;border-radius:0 0 8px 8px;overflow-x:auto;\">"
        + build_table(tokyo_results,  "🌳 東京都のコート", "#2e7d32")
        + build_table(minato_results, "🏙️ 港区のコート",   "#1565c0")
        +         "<p style=\"margin-top:8px;font-size:12px;color:#e65100;background:#fff3e0;padding:10px;border-radius:4px;\">"
        "⚠️ <b>注意</b>：空き情報は取得時点のものです。実際に予約する前に必ずサイトで最新の空き状況をご確認ください。"
        "</p>"
        "<p style=\"margin-top:8px;font-size:11px;color:#aaa;\">✅空きあり（2枠以上）＝緑　❌満杯＝グレー　－営業時間外＝薄グレー</p>"
        "</div></div>"
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"【🎾テニスコート空き情報・土日祝】{today_str}　{summary}"
    msg["From"]    = CONFIG["FROM_EMAIL"]
    msg["To"]      = CONFIG["TO_EMAIL"]
    msg.attach(MIMEText(html, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(CONFIG["FROM_EMAIL"], CONFIG["APP_PASSWORD"])
        server.sendmail(CONFIG["FROM_EMAIL"], CONFIG["TO_EMAIL"], msg.as_string())
    print("✅ メール送信完了！")


if __name__ == "__main__":
    today = datetime.date.today()
    target_dates = get_holiday_and_weekend_dates(today)
    print(f"チェック対象: {len(target_dates)}日")

    tokyo_cookies, tokyo_referer, minato_cookies, minato_referer = get_session()

    print("\n🔍 東京都チェック中...")
    tokyo_results  = check_tokyo(target_dates, tokyo_cookies, tokyo_referer)

    print("\n🔍 港区チェック中...")
    minato_results = check_minato(target_dates, minato_cookies, minato_referer)

    all_results = tokyo_results + minato_results
    print("\n--- 空きありのみ ---")
    for r in all_results:
        if r["color_key"] == "partial":
            print(f"  {r['site']} {r['date']}: {r['status']}")

    # JSON保存（GitHub Actionsのワークスペースに保存）
    json_data = {
        "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "results": [
            {k: v for k, v in r.items() if k != "sort_key"}
            for r in all_results
        ]
    }
    with open("tennis_results.json", "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    print("✅ JSON保存完了: tennis_results.json")

    print("\n📧 メール送信中...")
    send_email(all_results, target_dates)
