"""
Zocial Eye Daily Crisis Monitor
1. Login Zocial Eye → Export ALL messages to Excel (no sentiment filter)
2. Wait for Excel email → download attachment via IMAP
3. Analyze every row for crisis keywords (don't trust ZE sentiment alone)
4. Send daily summary to team
"""

import asyncio, os, imaplib, email, smtplib, time, tempfile
from email.mime.text import MIMEText
from datetime import datetime
from playwright.async_api import async_playwright
import pandas as pd

# ─── Config ───────────────────────────────────────────────────────────────────
ZOCIAL_ID       = os.environ.get("ZOCIAL_ID",          "Nativejump01")
ZOCIAL_PASS     = os.environ.get("ZOCIAL_PASS",         "Nativejump123")
CAMPAIGN_ID     = os.environ.get("CAMPAIGN_ID",         "93082")
EXPORT_EMAIL    = os.environ.get("EXPORT_EMAIL",        "kanthorn@nativejump.co")
NOTIFY_EMAIL    = os.environ.get("NOTIFY_EMAIL",        "kanthornb@gmail.com")
GMAIL_USER      = os.environ.get("GMAIL_USER",          "")
GMAIL_APP_PASS  = os.environ.get("GMAIL_APP_PASSWORD",  "")

IMAP_HOST       = "imap.gmail.com"
IMAP_MAX_WAIT   = 10   # นาที รอ Excel email
IMAP_POLL_SEC   = 30   # วินาที poll แต่ละครั้ง

CRISIS_KEYWORDS = [
    # ภาษาไทย
    "อันตราย", "เสียชีวิต", "ตาย", "เสียหาย", "บาดเจ็บ",
    "ฟ้องร้อง", "แจ้งความ", "ร้องเรียน",
    "ปลอม", "หลอกลวง", "โกง", "ทุจริต",
    "ผลข้างเคียง", "แพ้", "อาการแพ้", "ติดเชื้อ", "ฟกช้ำ",
    "ไม่ปลอดภัย", "ระวัง", "เตือน", "แย่มาก", "ห่วย",
    "ไม่ได้เรื่อง", "ผิดพลาด", "ล้มเหลว", "boycott",
    # English
    "danger", "dangerous", "death", "died", "injury", "injured",
    "lawsuit", "sue", "complaint", "fraud", "fake", "scam",
    "side effect", "allergic", "allergy", "infection", "bruise",
    "unsafe", "warning", "terrible", "horrible", "worst",
]

# ─── Helpers ──────────────────────────────────────────────────────────────────
def today_str():
    return datetime.now().strftime("%-d %b %Y")

def all_messages_url():
    d = today_str()
    return (
        f"https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message"
        f"?start={d.replace(' ', '+')}&end={d.replace(' ', '+')}&action=filter"
    )

# ─── Step 1: Playwright — login + trigger export ───────────────────────────────
async def trigger_export():
    print("  → Logging in to Zocial Eye...")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page    = await browser.new_page()

        await page.goto("https://zocialeye.wisesight.com/login")
        await page.wait_for_selector("input[name='username']")
        await page.fill("input[name='username']", ZOCIAL_ID)
        await page.fill("input[name='passwd']",   ZOCIAL_PASS)
        await page.click("#btn-login")
        await page.wait_for_url("**/home", timeout=15000)

        # Navigate to ALL messages (no sentiment filter)
        print("  → Loading all messages...")
        await page.goto("https://zocialeye.wisesight.com/campaigns")
        await page.wait_for_load_state("domcontentloaded")
        await page.wait_for_timeout(2000)
        await page.goto(all_messages_url())
        await page.wait_for_load_state("domcontentloaded")
        await page.wait_for_timeout(5000)

        # Get total message count
        total = await page.evaluate("""() => {
            const el = document.querySelector('.nav-tabs .active .badge, [class*="tab-active"] .badge');
            return el ? el.innerText.trim() : '?';
        }""")
        print(f"  → Total messages today: {total}")

        # Export → All channel (Excel)
        print("  → Triggering Excel export...")
        await page.locator("a.dropdown-toggle:has-text('Export')").click()
        await page.wait_for_timeout(1000)
        await page.locator("a[data-target='#modal-input-export-emails']").click()
        await page.wait_for_timeout(2000)
        await page.locator("#input-export-emails").wait_for(state="visible", timeout=10000)
        await page.fill("#input-export-emails", EXPORT_EMAIL)
        await page.wait_for_timeout(500)
        await page.click("#btn_submit_emails")
        await page.wait_for_timeout(2000)

        await browser.close()
        print(f"  → Export requested → {EXPORT_EMAIL}")
        return total

# ─── Step 2: IMAP — รอและดาวน์โหลด Excel attachment ──────────────────────────
def fetch_excel_from_email(triggered_at: datetime) -> str | None:
    print(f"  → Waiting for Excel email (up to {IMAP_MAX_WAIT} min)...")
    deadline = time.time() + IMAP_MAX_WAIT * 60

    with imaplib.IMAP4_SSL(IMAP_HOST) as mail:
        mail.login(EXPORT_EMAIL, GMAIL_APP_PASS)
        mail.select("inbox")

        while time.time() < deadline:
            mail.noop()
            _, data = mail.search(None, 'FROM "noreply@wisesight.com" SUBJECT "Zocial eye export data"')
            ids = data[0].split()

            for uid in reversed(ids):
                _, msg_data = mail.fetch(uid, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])

                # ตรวจว่าอีเมลมาหลังจาก trigger
                date_str = msg.get("Date", "")
                try:
                    from email.utils import parsedate_to_datetime
                    mail_dt = parsedate_to_datetime(date_str)
                    # normalize timezone
                    if mail_dt.tzinfo:
                        import datetime as dt_mod
                        mail_dt = mail_dt.astimezone(dt_mod.timezone.utc).replace(tzinfo=None)
                    if mail_dt < triggered_at:
                        continue
                except Exception:
                    pass

                # ดึง Excel attachment
                for part in msg.walk():
                    ct = part.get_content_type()
                    fn = part.get_filename() or ""
                    if "spreadsheet" in ct or fn.endswith(".xlsx"):
                        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                        tmp.write(part.get_payload(decode=True))
                        tmp.close()
                        print(f"  → Excel downloaded: {fn}")
                        return tmp.name

            print(f"  → Not yet, retrying in {IMAP_POLL_SEC}s...")
            time.sleep(IMAP_POLL_SEC)

    print("  ✗ Timed out waiting for Excel email")
    return None

# ─── Step 3: วิเคราะห์ Excel ทุก row ─────────────────────────────────────────
def analyze_excel(xlsx_path: str) -> dict:
    df = pd.read_excel(xlsx_path)
    total = len(df)

    neg_by_ze    = df[df["Sentiment"].str.lower() == "negative"] if "Sentiment" in df.columns else pd.DataFrame()
    pos_by_ze    = df[df["Sentiment"].str.lower() == "positive"] if "Sentiment" in df.columns else pd.DataFrame()
    neutral_by_ze = df[df["Sentiment"].str.lower() == "neutral"]  if "Sentiment" in df.columns else pd.DataFrame()

    # วิเคราะห์ crisis keyword ทุก row ไม่เชื่อ sentiment ของ ZE
    crisis_rows = []
    for _, row in df.iterrows():
        msg = str(row.get("Message", "")).lower()
        for kw in CRISIS_KEYWORDS:
            if kw.lower() in msg:
                crisis_rows.append({
                    "keyword":   kw,
                    "account":   row.get("Account", "-"),
                    "source":    row.get("Source", "-"),
                    "sentiment": row.get("Sentiment", "-"),
                    "message":   str(row.get("Message", ""))[:200],
                    "post_time": str(row.get("Post time", "")),
                    "main_kw":   row.get("Main keyword", "-"),
                })
                break

    # Top brands mentioned in crisis rows
    brand_counts = {}
    for r in crisis_rows:
        b = str(r["main_kw"])
        brand_counts[b] = brand_counts.get(b, 0) + 1

    return {
        "date":           datetime.now().strftime("%d %b %Y"),
        "total":          total,
        "neg_ze":         len(neg_by_ze),
        "pos_ze":         len(pos_by_ze),
        "neutral_ze":     len(neutral_by_ze),
        "crisis_count":   len(crisis_rows),
        "crisis_rows":    crisis_rows[:5],
        "brand_counts":   brand_counts,
    }

# ─── Step 4: ส่งอีเมลสรุป ────────────────────────────────────────────────────
def send_summary(result: dict):
    date          = result["date"]
    total         = result["total"]
    neg_ze        = result["neg_ze"]
    crisis_count  = result["crisis_count"]
    crisis_rows   = result["crisis_rows"]
    brand_counts  = result["brand_counts"]
    crisis        = crisis_count > 0

    if crisis:
        subject = f"[CRISIS ALERT] พบข้อความน่าเป็นห่วง {crisis_count} รายการ — {date}"
        hits_txt = "\n".join([
            f"  [{i+1}] @{r['account']} ({r['source']}) | keyword: {r['keyword']}\n"
            f"       ZE sentiment: {r['sentiment']} | brand: {r['main_kw']}\n"
            f"       \"{r['message'][:120]}...\"\n"
            for i, r in enumerate(crisis_rows)
        ])
        brands_txt = "\n".join([f"  - {b}: {c} ข้อความ" for b, c in sorted(brand_counts.items(), key=lambda x: -x[1])])
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: CRISIS DETECTED

ภาพรวมวันนี้
  - ข้อความทั้งหมด:       {total} รายการ
  - ZE ระบุ Negative:    {neg_ze} รายการ
  - Crisis keyword hit:  {crisis_count} รายการ

แบรนด์ที่ถูกพูดถึงใน crisis:
{brands_txt}

ตัวอย่างข้อความน่าเป็นห่วง (top 5):
{hits_txt}
================================================
ดูรายละเอียดทั้งหมด:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""
    else:
        subject = f"[No Crisis] Daily Brand Monitor — {date}"
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: ไม่พบ crisis

ภาพรวมวันนี้
  - ข้อความทั้งหมด:       {total} รายการ
  - ZE ระบุ Negative:    {neg_ze} รายการ
  - Crisis keyword hit:  0 รายการ

ไม่พบข้อความที่เข้าข่าย crisis ในวันนี้

================================================
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""

    if not GMAIL_USER or not GMAIL_APP_PASS:
        path = f"/tmp/crisis_report_{datetime.now():%Y%m%d}.txt"
        open(path, "w").write(f"Subject: {subject}\n\n{body}")
        print(f"  ⚠ No Gmail config — saved to {path}")
        return

    msg             = MIMEText(body, "plain", "utf-8")
    msg["Subject"]  = subject
    msg["From"]     = GMAIL_USER
    msg["To"]       = NOTIFY_EMAIL

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PASS)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAIL, msg.as_string())
        print(f"  ✓ Summary sent to {NOTIFY_EMAIL}")
    except Exception as e:
        print(f"  ✗ Email error: {e}")

# ─── Main ─────────────────────────────────────────────────────────────────────
async def main():
    print(f"[{datetime.now():%H:%M}] Zocial Eye Crisis Monitor starting...")
    triggered_at = datetime.utcnow()

    # 1. Trigger export
    await trigger_export()

    # 2. Fetch Excel from email
    xlsx_path = fetch_excel_from_email(triggered_at)

    if not xlsx_path:
        print("  ✗ Could not get Excel — sending fallback email")
        send_summary({"date": datetime.now().strftime("%d %b %Y"),
                      "total": "?", "neg_ze": "?", "pos_ze": "?", "neutral_ze": "?",
                      "crisis_count": 0, "crisis_rows": [], "brand_counts": {}})
        return

    # 3. Analyze
    print("  → Analyzing Excel...")
    result = analyze_excel(xlsx_path)
    print(f"  → Total: {result['total']} | ZE Negative: {result['neg_ze']} | Crisis hits: {result['crisis_count']}")

    # 4. Send summary
    send_summary(result)
    print(f"[{datetime.now():%H:%M}] Done.")

if __name__ == "__main__":
    asyncio.run(main())
