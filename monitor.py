"""
Zocial Eye Daily Crisis Monitor
1. Login Zocial Eye → Export ALL messages to Excel (no sentiment filter)
2. Wait for Excel email → download attachment via IMAP
3. Analyze every row for crisis keywords (don't trust ZE sentiment alone)
4. Send daily summary to team
"""

import asyncio, os, imaplib, email, smtplib, time, tempfile, json, re, urllib.request, sys, traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# บังคับเวลาไทยทั้งระบบ (runner เป็น UTC) — กัน crisis ช่วงเย็น-ดึกหลุด
TH = ZoneInfo("Asia/Bangkok")
def now_th():
    return datetime.now(TH)
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
def yesterday_str():
    return (now_th() - timedelta(days=1)).strftime("%-d %b %Y")

def all_messages_url():
    d = yesterday_str()
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
        print(f"  → Total messages yesterday: {total}")

        # ถ้าไม่มีข้อมูล ไม่ต้อง export
        if total == '0':
            print("  → No messages yesterday, skipping export")
            await browser.close()
            return '0'

        # รอให้ Export button ไม่ disabled
        print("  → Triggering Excel export...")
        export_btn = page.locator("a.dropdown-toggle:has-text('Export')")
        await export_btn.wait_for(state="visible", timeout=15000)
        # รอให้ disabled หาย
        for _ in range(10):
            is_disabled = await export_btn.get_attribute("disabled")
            if not is_disabled:
                break
            await page.wait_for_timeout(2000)
        await export_btn.click()
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
            "message":   str(row.get("Message", ""))[:2000],
            "post_time": str(row.get("Post time", "")),
        })

    print(f"  → Sending {len(messages_for_claude)} messages to Claude for analysis...")
    claude_result = claude_analyze(messages_for_claude)

    # เรียง crisis ตามความรุนแรง high→medium→low (กัน high หลุดท้ายแถว)
    sev_rank = {"high": 0, "medium": 1, "low": 2}
    items = sorted(claude_result.get("crisis_items", []),
                   key=lambda r: sev_rank.get(str(r.get("severity", "low")).lower(), 3))

    return {
        "date":         (now_th() - timedelta(days=1)).strftime("%d %b %Y"),
        "total":        total,
        "neg_ze":       neg_ze,
        "pos_ze":       pos_ze,
        "crisis_count": claude_result["crisis_count"],
        "crisis_rows":  items[:15],          # โชว์ได้ถึง 15 (เรียงตามรุนแรง)
        "all_crisis":   items,               # เก็บครบไว้แนบไฟล์
        "summary":      claude_result["summary"],
        "brand_counts": claude_result["brand_counts"],
        "error":        claude_result.get("error"),
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

    prompt = f"""คุณเป็น Social Media Crisis Analyst สำหรับแบรนด์ความงามและสุขภาพในไทย งานนี้เป็น safety-critical: พลาด crisis ไม่ได้เด็ดขาด เมื่อสงสัยให้ flag ไว้ก่อน (เลือก false positive ดีกว่า false negative)

แบรนด์ที่ต้องติดตาม (alert เฉพาะที่เกี่ยวกับแบรนด์เหล่านี้):
- Xeomin = โบทูลินั่มท็อกซิน (เรียกว่า "โบ", ซีโอมิน, ซีโอมิน, Xeomin)
- Belotero / Belotero Revive = ฟิลเลอร์ HA (เบลโลเทโร่, เบโลเทโร, เบลโลเทโร่ รีไวฟ์)
- Ulthera / Ulthera Prime = HIFU ยกกระชับ (อัลเทอร่า, อัลเธอร่า, อัลทีร่า, ไฮฟู่/HIFU ยกกระชับ, "เครื่องคลื่นเสียงยกกระชับ")
- Radiesse = ฟิลเลอร์/คอลลาเจนสติมูเลเตอร์ CaHA (เรดิเอส, เรดิแอส, เรดิเอซ)

สำคัญมาก — ต้องจับให้ได้แม้เขียนเลี่ยง:
1. อ้างถึงแบรนด์เป็นภาษาไทย/ทับศัพท์/สะกดผิด/ชื่อเล่น หรือเรียกแบบ generic ("โบ", "ฟิลเลอร์ใต้ตา", "ทำไฮฟู่", "ฉีดสารเติมเต็มร่องแก้ม", "เครื่องที่ดารารีวิว") — ถ้ามีโอกาสหมายถึงสินค้าข้างต้น ให้ flag (severity low/medium ระบุว่า "ต้องตรวจสอบ")
2. โทนประชดประชัน/แดกดัน ("ขอบคุณคลินิกที่ทำให้ได้นอนรพ.", "บริการดีมากค่ะ" ทั้งที่บรรยายอาการแย่)
3. romanized/karaoke Thai ("na pang", "chit laew jeb free"), หรือศัพท์การแพทย์อังกฤษ (necrosis, vascular occlusion, blindness, lumps, botched)
4. เสียงคนอื่นเล่าแทน ("แม่ไปฉีดมา", "เพื่อนทำแล้วหน้า...", "พาญาติไปแก้")
5. ชี้ไปรูป/คลิป/คอมเมนต์ ("ดูรูป", "ดูคลิป", "อ่านในคอมเมนต์")

นิยาม crisis (จัด severity = high ถ้าอันตรายถึงชีวิต/ตา/ถาวร หรือมีแนวโน้มไวรัล):
- อาการไม่พึงประสงค์ร้ายแรง: เนื้อตาย/หน้าเน่า/ผิวซีดเป็นลายๆ/ปวดมาก (vascular occlusion), ตาบอด/ตามัว/มองไม่เห็น, โบกระจาย→หายใจ/กลืน/พูด/เคี้ยวลำบาก กล้ามเนื้ออ่อนแรง (iatrogenic botulism), แพ้รุนแรง/หายใจไม่ออก/ช็อก (anaphylaxis), แผลไหม้/เนื้อยุบถาวร (HIFU), ติดเชื้อ/เป็นหนอง, หน้าเบี้ยว/ปากเบี้ยว, เป็นก้อน/ฟิลเลอร์ไหล
- ของปลอม/หิ้ว/ไม่มี อย., หมอเถื่อน/พยาบาลฉีดเอง/คลินิกในคอนโด, ฉีดเด็ก/คนท้อง
- ร้องเรียน/แจ้งความ/ฟ้อง/อย./สคบ./สั่งปิด/เรียกคืนสินค้า
- คนถามหา "ฉีดสลาย/hyalase/เลาะฟิลเลอร์/หาหมอแก้เคสพัง" (= มีเคสเสียหายมาก่อน)
- สัญญาณไวรัล/รวมพลัง: โหนกระแส, สายไหมต้องรอด, รวมตัวผู้เสียหาย, ผู้เสียหายหลายราย, ล่ารายชื่อ, ทัวร์ลง, แฉ
- ภาวะซึมเศร้า/ทำร้ายตัวเองหลังหน้าพัง

ไม่ใช่ crisis: ข้อความที่ไม่เกี่ยวกับสินค้าข้างต้นเลย, ข่าวสุขภาพทั่วไป, โปรโมชั่น/ขายปกติ, รีวิวบวกแท้จริง

ข้อความทั้งหมด:
{msgs_text}

ตอบเป็น JSON ในรูปแบบนี้เท่านั้น (ห้ามมีข้อความอื่นนอก JSON):
{{
  "crisis_items": [
    {{
      "id": <id>,
      "account": "<account>",
      "source": "<source>",
      "brand": "<แบรนด์ที่เกี่ยว หรือ 'ต้องตรวจสอบ'>",
      "reason": "<เหตุผลสั้นๆ>",
      "severity": "low|medium|high",
      "message_preview": "<ข้อความ 120 ตัวอักษรแรก>"
    }}
  ],
  "summary": "<สรุปภาพรวม 2-3 ประโยค>",
  "brand_counts": {{"<brand>": <จำนวน>}}
}}

ถ้าไม่มี crisis เลย ให้ crisis_items เป็น [] และ summary บอกว่าไม่พบ crisis"""

    last_err = None
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=8000,
                messages=[{"role": "user", "content": prompt}]
            )
            # ถ้าคำตอบโดนตัดกลางคัน = วันที่ข้อความเยอะ → อย่าไว้ใจ
            if response.stop_reason == "max_tokens":
                raise ValueError("response truncated (max_tokens) — too many messages for one call")

            raw = response.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]

            result = json.loads(raw)
            result["crisis_count"] = len(result.get("crisis_items", []))
            return result
        except Exception as e:
            last_err = e
            print(f"  ✗ analyze attempt {attempt+1}/3 failed: {e}")
            time.sleep(5 * (attempt + 1))

    # พังทุก attempt → ห้ามรายงาน 0 crisis เด็ดขาด ส่งสัญญาณ error ให้ส่งอีเมลแจ้งเตือนแทน
    print(f"  ✗✗ Claude analysis FAILED after 3 attempts — flagging for manual review")
    return {"crisis_count": 0, "crisis_items": [], "brand_counts": {},
            "summary": f"วิเคราะห์ไม่สำเร็จ: {last_err}",
            "error": f"การวิเคราะห์ล้มเหลว ({last_err}) — ต้องตรวจสอบด้วยตนเอง"}

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

    # ใช้ฟอนต์ภาษาไทยที่อยู่ใน repo
    thai_font = "Helvetica"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    font_candidates = [
        os.path.join(script_dir, "fonts", "NotoSansThai-Regular.ttf"),
        "/usr/share/fonts/truetype/tlwg/Sarabun.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
    ]
    for font_path in font_candidates:
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
        ["พบ Crisis", str(result["crisis_count"])],
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
    story.append(Paragraph("สรุปผลการวิเคราะห์", head_style))
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
    story.append(Paragraph("ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)", small_style))

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
    err            = result.get("error")

    if err:
        # วิเคราะห์/ดึงข้อมูลพัง — ห้ามทำเป็น "ไม่พบ crisis" ต้องให้คนเช็คเอง
        subject = f"[⚠️ ต้องตรวจสอบเอง] Monitor มีปัญหา — {date}"
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: ⚠️ ระบบมีปัญหา — กรุณาตรวจสอบด้วยตนเอง

ปัญหา:
{err}

หมายเหตุ: วันนี้ระบบ "ยังไม่ได้" ยืนยันว่าไม่มี crisis
กรุณาเข้าไปดูข้อมูลด้วยตนเองที่:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ภาพรวมเท่าที่ได้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
================================================
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""
    elif crisis:
        shown = len(crisis_rows)
        more_note = f"\n(แสดง {shown} จากทั้งหมด {crisis_count} รายการ — ดูครบในไฟล์แนบ)" if crisis_count > shown else ""
        subject = f"[CRISIS ALERT] พบข้อความน่าเป็นห่วง {crisis_count} รายการ — {date}"
        hits_txt = "\n".join([
            f"  [{i+1}] @{r.get('account','-')} ({r.get('source','-')}) "
            f"| แบรนด์: {r.get('brand','-')} | ระดับ: {severity_map.get(str(r.get('severity','')).lower(),'?')}\n"
            f"       เหตุผล: {r.get('reason','')}\n"
            f"       \"{r.get('message_preview', r.get('message',''))[:120]}\"\n"
            for i, r in enumerate(crisis_rows)
        ])
        brands_txt = "\n".join([f"  - {b}: {c} ข้อความ" for b, c in sorted(brand_counts.items(), key=lambda x: -x[1])])
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: CRISIS DETECTED

สรุปผลการวิเคราะห์:
{claude_summary}

ภาพรวมวันนี้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
  - พบ crisis:          {crisis_count} รายการ

แบรนด์ที่ถูกพูดถึงใน crisis:
{brands_txt}

รายละเอียด (เรียงตามความรุนแรง):{more_note}
{hits_txt}
================================================
ดูทั้งหมดที่:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""
    else:
        subject = f"[No Crisis] Daily Brand Monitor — {date}"
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: ไม่พบ crisis

สรุปผลการวิเคราะห์:
{claude_summary}

ภาพรวมวันนี้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
  - พบ crisis:          0 รายการ

================================================
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""

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
async def run_pipeline(report_date: str):
    triggered_at = datetime.utcnow()

    # 1. Trigger export
    total = await trigger_export()

    if total == '0':
        # อ่านได้ 0 — อาจเป็นวันเงียบจริง หรือ badge scrape พลาด → แจ้งแบบให้ยืนยัน ไม่ใช่ all-clear มั่นใจ
        print("  → Read 0 messages — sending cautionary email")
        send_summary({"date": report_date, "total": 0, "neg_ze": 0,
                      "crisis_count": 0, "crisis_rows": [], "brand_counts": {},
                      "error": "ระบบอ่านว่าไม่มีข้อความเมื่อวาน (0 รายการ) — หากผิดปกติโปรดเข้าไปตรวจสอบใน Zocial Eye ด้วยตนเอง"})
        return

    # 2. Fetch Excel from email
    xlsx_path = fetch_excel_from_email(triggered_at)

    if not xlsx_path:
        print("  ✗ Could not get Excel — sending error email")
        send_summary({"date": report_date, "total": "?", "neg_ze": "?",
                      "crisis_count": 0, "crisis_rows": [], "brand_counts": {},
                      "error": "ดึงไฟล์ Excel จากอีเมลไม่สำเร็จภายในเวลาที่กำหนด — ยังไม่ได้วิเคราะห์ข้อมูลวันนี้"})
        return

    # 3. Analyze
    print("  → Analyzing Excel...")
    result = analyze_excel(xlsx_path)
    print(f"  → Total: {result['total']} | ZE Negative: {result['neg_ze']} | Crisis hits: {result['crisis_count']}")

    # 4. Generate PDF report (ข้ามถ้าวิเคราะห์พัง — จะส่งอีเมลแจ้งเตือนแทน)
    pdf_path = None
    if not result.get("error"):
        print("  → Generating PDF report...")
        pdf_path = create_pdf_report(result)

    # 5. Send summary
    send_summary(result, xlsx_path, pdf_path)


async def main():
    print(f"[{now_th():%H:%M}] Zocial Eye Crisis Monitor starting...")
    report_date = (now_th() - timedelta(days=1)).strftime("%d %b %Y")
    try:
        await run_pipeline(report_date)
        print(f"[{now_th():%H:%M}] Done.")
    except Exception as e:
        # พังที่ไหนก็ตาม → ต้องส่งอีเมลแจ้ง ห้ามเงียบ (เงียบ = เข้าใจผิดว่าไม่มี crisis)
        tb = traceback.format_exc()
        print(f"  ✗✗ PIPELINE CRASHED:\n{tb}")
        try:
            send_summary({"date": report_date, "total": "?", "neg_ze": "?",
                          "crisis_count": 0, "crisis_rows": [], "brand_counts": {},
                          "error": f"ระบบทำงานล้มเหลว (crash): {e}\nกรุณาตรวจสอบ Zocial Eye ด้วยตนเองวันนี้"})
        except Exception as e2:
            print(f"  ✗✗ Even the failure email failed: {e2}")
        sys.exit(1)   # ทำให้ GitHub Action ขึ้นแดง จะได้รู้ว่าพัง

if __name__ == "__main__":
    asyncio.run(main())
