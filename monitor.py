"""
Zocial Eye Daily Crisis Monitor
1. Login Zocial Eye → Export ALL messages to Excel (no sentiment filter)
2. Wait for Excel email → download attachment via IMAP
3. Analyze every row for crisis keywords (don't trust ZE sentiment alone)
4. Send daily summary to team
"""

import asyncio, os, imaplib, email, smtplib, time, tempfile, json, re, urllib.request
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from playwright.async_api import async_playwright
import pandas as pd
import anthropic

# ─── Config ───────────────────────────────────────────────────────────────────
ZOCIAL_ID       = os.environ.get("ZOCIAL_ID",          "Nativejump01")
ZOCIAL_PASS     = os.environ.get("ZOCIAL_PASS",         "Nativejump123")
CAMPAIGN_ID     = os.environ.get("CAMPAIGN_ID",         "93082")
EXPORT_EMAIL    = os.environ.get("EXPORT_EMAIL",        "kanthorn@nativejump.co")
NOTIFY_EMAIL    = os.environ.get("NOTIFY_EMAIL",        "kanthornb@gmail.com")
GMAIL_USER      = os.environ.get("GMAIL_USER",          "")
GMAIL_APP_PASS  = os.environ.get("GMAIL_APP_PASSWORD",  "")

ANTHROPIC_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
IMAP_HOST       = "imap.gmail.com"
IMAP_MAX_WAIT   = 10   # นาที รอ Excel email
IMAP_POLL_SEC   = 30   # วินาที poll แต่ละครั้ง

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
            _, data = mail.search(None, 'FROM "no-reply@zocialeye.com" SUBJECT "Zocial eye export data"')
            ids = data[0].split()

            for uid in reversed(ids):
                _, msg_data = mail.fetch(uid, "(RFC822)")
                raw = msg_data[0][1] if isinstance(msg_data[0], tuple) else None
                if not raw:
                    continue
                msg = email.message_from_bytes(raw)

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

                # ดึง download URL จาก email body
                body = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")

                url_match = re.search(r'https://downloads\.zocialeye\.com/[^\s\'"<>]+\.xlsx', body)
                if url_match:
                    url = url_match.group(0)
                    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    urllib.request.urlretrieve(url, tmp.name)
                    print(f"  → Excel downloaded from: {url}")
                    return tmp.name

            print(f"  → Not yet, retrying in {IMAP_POLL_SEC}s...")
            time.sleep(IMAP_POLL_SEC)

    print("  ✗ Timed out waiting for Excel email")
    return None

# ─── Step 3: Claude วิเคราะห์ทุก row ─────────────────────────────────────────
def analyze_excel(xlsx_path: str) -> dict:
    df = pd.read_excel(xlsx_path)
    total = len(df)

    neg_ze = len(df[df["Sentiment"].str.lower() == "negative"]) if "Sentiment" in df.columns else 0
    pos_ze = len(df[df["Sentiment"].str.lower() == "positive"]) if "Sentiment" in df.columns else 0

    # สร้าง message list ส่งให้ Claude
    messages_for_claude = []
    for i, row in df.iterrows():
        messages_for_claude.append({
            "id":        i,
            "account":   str(row.get("Account", "-")),
            "source":    str(row.get("Source", "-")),
            "brand":     str(row.get("Main keyword", "-")),
            "sentiment": str(row.get("Sentiment", "-")),
            "message":   str(row.get("Message", ""))[:500],
            "post_time": str(row.get("Post time", "")),
        })

    print(f"  → Sending {len(messages_for_claude)} messages to Claude for analysis...")
    claude_result = claude_analyze(messages_for_claude)

    return {
        "date":         datetime.now().strftime("%d %b %Y"),
        "total":        total,
        "neg_ze":       neg_ze,
        "pos_ze":       pos_ze,
        "crisis_count": claude_result["crisis_count"],
        "crisis_rows":  claude_result["crisis_items"][:5],
        "summary":      claude_result["summary"],
        "brand_counts": claude_result["brand_counts"],
    }


def claude_analyze(messages: list) -> dict:
    """ส่งทุก message ให้ Claude วิเคราะห์ crisis ในครั้งเดียว"""
    if not ANTHROPIC_KEY:
        return {"crisis_count": 0, "crisis_items": [], "summary": "No API key", "brand_counts": {}}

    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    msgs_text = "\n".join([
        f"[{m['id']}] @{m['account']} ({m['source']}) brand={m['brand']} ZE_sentiment={m['sentiment']}\n{m['message']}"
        for m in messages
    ])

    prompt = f"""คุณเป็น Social Media Crisis Analyst สำหรับแบรนด์ความงามและสุขภาพในไทย

วิเคราะห์ข้อความ social media ด้านล่างนี้ทุกข้อความ แล้วระบุว่าข้อความไหนเป็น "crisis" หรือ "น่ากังวล" สำหรับแบรนด์

นิยาม crisis:
- มีการพูดถึงแบรนด์/ผลิตภัณฑ์ในทางลบอย่างชัดเจน
- มีการร้องเรียน แจ้งความ ฟ้องร้อง
- มีผลข้างเคียงที่อันตราย หรืออาการบาดเจ็บหลังใช้ผลิตภัณฑ์
- มีการกล่าวหาว่าปลอม หลอกลวง หรือโกง
- มีการแพร่กระจายข่าวเชิงลบเกี่ยวกับแบรนด์

ไม่ใช่ crisis:
- ข่าวสุขภาพทั่วไปที่ไม่เกี่ยวกับแบรนด์โดยตรง
- โพสต์โปรโมชั่น/ขาย
- รีวิวบวก
- ข่าวธุรกิจทั่วไป

ข้อความทั้งหมด:
{msgs_text}

ตอบเป็น JSON ในรูปแบบนี้เท่านั้น:
{{
  "crisis_items": [
    {{
      "id": <id ของข้อความ>,
      "account": "<account>",
      "source": "<source>",
      "brand": "<brand>",
      "reason": "<เหตุผลสั้นๆ ว่าทำไมถึงเป็น crisis>",
      "severity": "low|medium|high",
      "message_preview": "<ข้อความ 100 ตัวอักษรแรก>"
    }}
  ],
  "summary": "<สรุปภาพรวมสั้นๆ 2-3 ประโยค>",
  "brand_counts": {{"<brand>": <จำนวน crisis>}}
}}

ถ้าไม่มี crisis เลย ให้ crisis_items เป็น [] และ summary บอกว่าไม่พบ crisis"""

    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    # ตัด markdown code block ถ้ามี
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    try:
        result = json.loads(raw)
        result["crisis_count"] = len(result.get("crisis_items", []))
        return result
    except Exception as e:
        print(f"  ✗ Claude JSON parse error: {e}\n  Raw: {raw[:200]}")
        return {"crisis_count": 0, "crisis_items": [], "summary": raw[:300], "brand_counts": {}}

# ─── Step 3b: สร้าง PDF รายงาน ────────────────────────────────────────────────
def create_pdf_report(result: dict) -> str:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import os

    # ลองใช้ฟอนต์ที่รองรับภาษาไทย (Linux/GitHub Actions)
    thai_font = "Helvetica"
    for font_path in [
        "/usr/share/fonts/truetype/tlwg/Sarabun.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
    ]:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("Thai", font_path))
                thai_font = "Thai"
            except Exception:
                pass
            break

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.close()

    doc = SimpleDocTemplate(tmp.name, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    styles = getSampleStyleSheet()
    title_style  = ParagraphStyle("title",  fontName=thai_font, fontSize=18, spaceAfter=6,  textColor=colors.HexColor("#1a1a2e"), leading=22)
    head_style   = ParagraphStyle("head",   fontName=thai_font, fontSize=13, spaceAfter=4,  textColor=colors.HexColor("#16213e"), leading=16)
    normal_style = ParagraphStyle("normal", fontName=thai_font, fontSize=10, spaceAfter=4,  textColor=colors.HexColor("#333333"), leading=14)
    small_style  = ParagraphStyle("small",  fontName=thai_font, fontSize=9,  spaceAfter=2,  textColor=colors.HexColor("#555555"), leading=12)

    crisis       = result["crisis_count"] > 0
    status_color = colors.HexColor("#c0392b") if crisis else colors.HexColor("#27ae60")
    status_text  = "CRISIS DETECTED" if crisis else "ไม่พบ crisis"

    story = []

    # Header
    story.append(Paragraph("Zocial Eye Crisis Monitor", title_style))
    story.append(Paragraph(f"รายงานประจำวัน: {result['date']}", normal_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cccccc")))
    story.append(Spacer(1, 0.3*cm))

    # Status badge
    status_style = ParagraphStyle("status", fontName=thai_font, fontSize=14,
                                  textColor=colors.white, backColor=status_color,
                                  borderPadding=(4, 8, 4, 8), leading=18)
    story.append(Paragraph(f"สถานะ: {status_text}", status_style))
    story.append(Spacer(1, 0.4*cm))

    # Overview table
    story.append(Paragraph("ภาพรวมวันนี้", head_style))
    overview_data = [
        ["รายการ", "จำนวน"],
        ["ข้อความทั้งหมด", str(result["total"])],
        ["ZE ระบุ Negative", str(result["neg_ze"])],
        ["Claude พบ Crisis", str(result["crisis_count"])],
    ]
    t = Table(overview_data, colWidths=[10*cm, 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
        ("FONTNAME",     (0,0), (-1,-1), thai_font),
        ("FONTSIZE",     (0,0), (-1,-1), 10),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.HexColor("#f9f9f9"), colors.white]),
        ("GRID",         (0,0), (-1,-1), 0.5, colors.HexColor("#dddddd")),
        ("ALIGN",        (1,0), (1,-1), "CENTER"),
        ("PADDING",      (0,0), (-1,-1), 6),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.4*cm))

    # Claude summary
    story.append(Paragraph("สรุปจาก Claude AI", head_style))
    story.append(Paragraph(result.get("summary", "-"), normal_style))
    story.append(Spacer(1, 0.4*cm))

    # Crisis details
    if crisis and result.get("crisis_rows"):
        story.append(Paragraph("รายละเอียด Crisis (Top 5)", head_style))
        severity_map = {"high": "สูง", "medium": "กลาง", "low": "ต่ำ"}
        severity_color = {"high": "#c0392b", "medium": "#e67e22", "low": "#f1c40f"}

        for i, r in enumerate(result["crisis_rows"], 1):
            sev = r.get("severity", "low")
            sev_th = severity_map.get(sev, sev)
            sev_c  = colors.HexColor(severity_color.get(sev, "#f1c40f"))

            row_data = [[
                Paragraph(f"<b>[{i}] @{r.get('account','-')}</b> ({r.get('source','-')}) — แบรนด์: {r.get('brand','-')}", normal_style),
                Paragraph(f"ระดับ: {sev_th}", ParagraphStyle("sev", fontName=thai_font, fontSize=9, textColor=sev_c, leading=12)),
            ]]
            rt = Table(row_data, colWidths=[11*cm, 3*cm])
            rt.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#fff8f8")),
                ("GRID",       (0,0), (-1,-1), 0.5, colors.HexColor("#ffcccc")),
                ("PADDING",    (0,0), (-1,-1), 6),
                ("VALIGN",     (0,0), (-1,-1), "TOP"),
            ]))
            story.append(rt)
            story.append(Paragraph(f"เหตุผล: {r.get('reason','')}", small_style))
            story.append(Paragraph(f"\"{r.get('message_preview','')[:120]}\"", small_style))
            story.append(Spacer(1, 0.2*cm))

    # Footer
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")))
    story.append(Paragraph("ส่งโดย: Zocial Eye Crisis Monitor + Claude Sonnet AI (อัตโนมัติ)", small_style))

    doc.build(story)
    return tmp.name


# ─── Step 4: ส่งอีเมลสรุป ────────────────────────────────────────────────────
def send_summary(result: dict, xlsx_path: str = None, pdf_path: str = None):
    date          = result["date"]
    total         = result["total"]
    neg_ze        = result["neg_ze"]
    crisis_count  = result["crisis_count"]
    crisis_rows   = result["crisis_rows"]
    brand_counts  = result["brand_counts"]
    crisis        = crisis_count > 0

    claude_summary = result.get("summary", "")
    severity_map   = {"high": "สูง", "medium": "กลาง", "low": "ต่ำ"}

    if crisis:
        subject = f"[CRISIS ALERT] พบข้อความน่าเป็นห่วง {crisis_count} รายการ — {date}"
        hits_txt = "\n".join([
            f"  [{i+1}] @{r.get('account','-')} ({r.get('source','-')}) "
            f"| แบรนด์: {r.get('brand','-')} | ระดับ: {severity_map.get(r.get('severity',''),'?')}\n"
            f"       เหตุผล: {r.get('reason','')}\n"
            f"       \"{r.get('message_preview', r.get('message',''))[:120]}\"\n"
            for i, r in enumerate(crisis_rows)
        ])
        brands_txt = "\n".join([f"  - {b}: {c} ข้อความ" for b, c in sorted(brand_counts.items(), key=lambda x: -x[1])])
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: CRISIS DETECTED

สรุปจาก Claude:
{claude_summary}

ภาพรวมวันนี้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
  - Claude พบ crisis:   {crisis_count} รายการ

แบรนด์ที่ถูกพูดถึงใน crisis:
{brands_txt}

รายละเอียด (top 5):
{hits_txt}
================================================
ดูทั้งหมดที่:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ส่งโดย: Zocial Eye Crisis Monitor + Claude AI (อัตโนมัติ)"""
    else:
        subject = f"[No Crisis] Daily Brand Monitor — {date}"
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: ไม่พบ crisis

สรุปจาก Claude:
{claude_summary}

ภาพรวมวันนี้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
  - Claude พบ crisis:   0 รายการ

================================================
ส่งโดย: Zocial Eye Crisis Monitor + Claude AI (อัตโนมัติ)"""

    if not GMAIL_USER or not GMAIL_APP_PASS:
        path = f"/tmp/crisis_report_{datetime.now():%Y%m%d}.txt"
        open(path, "w").write(f"Subject: {subject}\n\n{body}")
        print(f"  ⚠ No Gmail config — saved to {path}")
        return

    msg = MIMEMultipart()
    msg["Subject"]  = subject
    msg["From"]     = GMAIL_USER
    msg["To"]       = NOTIFY_EMAIL
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for fpath, fname in [
        (pdf_path,  f"ZE_report_{result['date'].replace(' ', '_')}.pdf"),
        (xlsx_path, f"ZE_export_{result['date'].replace(' ', '_')}.xlsx"),
    ]:
        if fpath and os.path.exists(fpath):
            with open(fpath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={fname}")
            msg.attach(part)

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

    # 4. Generate PDF report
    print("  → Generating PDF report...")
    pdf_path = create_pdf_report(result)

    # 5. Send summary
    send_summary(result, xlsx_path, pdf_path)
    print(f"[{datetime.now():%H:%M}] Done.")

if __name__ == "__main__":
    asyncio.run(main())
