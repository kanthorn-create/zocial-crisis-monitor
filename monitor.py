"""
Zocial Eye Daily Crisis Monitor
- Login to Zocial Eye
- Filter negative sentiment messages for today
- Export to Excel via email
- Analyze for crisis and send summary to team
"""

import asyncio
import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
from playwright.async_api import async_playwright

# ─── Config (อ่านจาก env vars เมื่อรันบน GitHub Actions) ─────────────────────
ZOCIAL_ID        = os.environ.get("ZOCIAL_ID",        "Nativejump01")
ZOCIAL_PASS      = os.environ.get("ZOCIAL_PASS",       "Nativejump123")
CAMPAIGN_ID      = os.environ.get("CAMPAIGN_ID",       "93082")
EXPORT_EMAIL     = os.environ.get("EXPORT_EMAIL",      "kanthorn@nativejump.co")
NOTIFY_EMAIL     = os.environ.get("NOTIFY_EMAIL",      "kanthornb@gmail.com")
GMAIL_USER       = os.environ.get("GMAIL_USER",        "")   # Gmail address ที่ใช้ส่ง
GMAIL_APP_PASS   = os.environ.get("GMAIL_APP_PASSWORD", "")  # App Password (ไม่ใช่ password จริง)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def today_str():
    return datetime.now().strftime("%-d %b %Y")   # e.g. "29 May 2026"

def messages_url(sentiment="neg"):
    d = today_str()
    return (
        f"https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message"
        f"?start={d.replace(' ', '+')}&end={d.replace(' ', '+')}"
        f"&action=filter&sentiment={sentiment}"
    )

# ─── Main ─────────────────────────────────────────────────────────────────────
async def run():
    print(f"[{datetime.now():%H:%M}] Starting Zocial Eye crisis monitor...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 1. Login
        print("  → Logging in...")
        await page.goto("https://zocialeye.wisesight.com/login")
        await page.wait_for_selector("input[name='username']")
        await page.fill("input[name='username']", ZOCIAL_ID)
        await page.fill("input[name='passwd']", ZOCIAL_PASS)
        await page.click("#btn-login")
        await page.wait_for_url("**/home", timeout=15000)
        await page.wait_for_timeout(1000)

        # 2. Go to negative sentiment messages (via campaigns first to set session)
        print("  → Loading negative messages...")
        await page.goto("https://zocialeye.wisesight.com/campaigns")
        await page.wait_for_load_state("domcontentloaded")
        await page.wait_for_timeout(2000)
        await page.goto(messages_url("neg"))
        await page.wait_for_load_state("domcontentloaded")
        await page.wait_for_timeout(5000)

        # Count messages
        msg_count_el = page.locator(".tab-active .badge, span:has-text('messages')")
        msg_count_text = ""
        try:
            msg_count_text = await msg_count_el.first.inner_text()
        except:
            pass

        # 3. Click Export dropdown button, then click "All channel" (To Excel)
        print("  → Opening Export dropdown...")
        # Click the Export dropdown toggle
        await page.locator("a.dropdown-toggle:has-text('Export')").click()
        await page.wait_for_timeout(1000)

        # Click "All channel" in the dropdown (To Excel section)
        await page.locator("a[data-target='#modal-input-export-emails']").click()
        await page.wait_for_timeout(2000)

        # 4. Fill email in export modal (clear old value first)
        print("  → Filling export email...")
        await page.locator("#input-export-emails").wait_for(state="visible", timeout=10000)
        await page.fill("#input-export-emails", EXPORT_EMAIL)
        await page.wait_for_timeout(500)

        # Submit export
        await page.click("#btn_submit_emails")
        await page.wait_for_timeout(2000)

        print(f"  → Export sent to {EXPORT_EMAIL}")

        # 5. Scrape message previews for quick crisis analysis
        print("  → Scraping message text for crisis analysis...")
        messages = await page.eval_on_selector_all(
            ".message-content, td.message-text, [class*='message'] p, .msg-text",
            "els => els.map(e => e.innerText.trim()).filter(t => t.length > 10)"
        )

        await browser.close()

    # 6. Analyze for crisis
    crisis_detected, summary = analyze_crisis(messages, msg_count_text)

    # 7. Send email summary
    print("  → Sending crisis summary email...")
    send_email_summary(crisis_detected, summary, messages)
    print("  ✓ Done.")


def analyze_crisis(messages, count_text):
    """Simple keyword-based crisis detection on negative messages."""
    CRISIS_KEYWORDS = [
        # Thai
        "อันตราย", "เสียชีวิต", "ตาย", "ฟ้องร้อง", "แจ้งความ",
        "ปลอม", "หลอกลวง", "โกง", "ผลข้างเคียง", "บาดเจ็บ",
        "เสียหาย", "ผิดพลาด", "ระวัง", "เตือน", "boycott",
        "ไม่ปลอดภัย", "อาการแพ้", "รอยฟกช้ำ", "แพ้", "ติดเชื้อ",
        # English
        "danger", "death", "lawsuit", "fake", "fraud", "side effect",
        "injury", "warning", "unsafe", "allergy", "infection",
    ]

    hits = []
    for msg in messages:
        msg_lower = msg.lower()
        for kw in CRISIS_KEYWORDS:
            if kw.lower() in msg_lower:
                hits.append({"message": msg[:200], "keyword": kw})
                break

    crisis = len(hits) > 0
    neg_count = count_text if count_text else f"{len(messages)} (scraped)"

    summary = {
        "date": datetime.now().strftime("%d %b %Y"),
        "negative_count": neg_count,
        "crisis_hits": len(hits),
        "hits": hits[:5],  # top 5
    }
    return crisis, summary


def send_email_summary(crisis_detected, summary, all_messages):
    """Send email via Gmail SMTP (ทำงานได้ทั้ง macOS และ Linux/GitHub Actions)"""
    date         = summary["date"]
    neg_count    = summary["negative_count"]
    crisis_hits  = summary["crisis_hits"]

    if crisis_detected:
        subject = f"[CRISIS ALERT] Negative Brand Mentions Detected — {date}"
        hits_section = "\n".join(
            [f"  - [{h['keyword']}] {h['message'][:150]}..." for h in summary["hits"]]
        )
        body = f"""สรุปประจำวัน: {date}
================================================
สถานะ: CRISIS DETECTED
Negative mentions วันนี้: {neg_count}
พบ {crisis_hits} ข้อความที่อาจเป็น crisis

ข้อความที่น่าเป็นห่วง:
{hits_section}

================================================
ตรวจสอบรายละเอียดที่:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""
    else:
        subject = f"[No Crisis] Daily Brand Monitor — {date}"
        body = f"""สรุปประจำวัน: {date}
================================================
สถานะ: ไม่พบ crisis
Negative mentions วันนี้: {neg_count}

ไม่พบข้อความที่เข้าข่าย crisis ในวันนี้

================================================
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""

    if not GMAIL_USER or not GMAIL_APP_PASS:
        # ไม่มี Gmail credentials — save to file แทน
        path = f"/tmp/crisis_report_{datetime.now():%Y%m%d}.txt"
        with open(path, "w") as f:
            f.write(f"Subject: {subject}\n\n{body}")
        print(f"  ⚠ No Gmail config — saved to {path}")
        return

    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = NOTIFY_EMAIL

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PASS)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAIL, msg.as_string())
        print(f"  ✓ Email sent to {NOTIFY_EMAIL}")
    except Exception as e:
        print(f"  ✗ Email error: {e}")


if __name__ == "__main__":
    asyncio.run(run())
