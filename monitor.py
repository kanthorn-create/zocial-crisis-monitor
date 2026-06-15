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
# ผู้รับรายงานประจำวันที่สำเร็จ — ลูกค้า Merz + ทีม NativeJump (เป็น config ไม่ใช่ความลับ)
# *** TEMP: เช็ก 14 มิ.ย. ส่งเฉพาะทีม NativeJump — จะ revert กลับเป็น 6 คน ***
REPORT_RECIPIENTS = [
    # "kamolrat.p@merz.com",       # Merz — ปิดชั่วคราว
    # "sarun.chompaisal@merz.com", # Merz — ปิดชั่วคราว
    # "maytita.t@merz.com",        # Merz — ปิดชั่วคราว
    "kanthorn@nativejump.co",
    "varithorn@nativejump.co",
    "nawarat@nativejump.co",
]
# ผู้รับแจ้งเตือน error เท่านั้น — เฉพาะทีม NativeJump ที่เข้า ZE ได้ (ไม่ส่งหาลูกค้า Merz)
ADMIN_RECIPIENTS = [
    "kanthorn@nativejump.co",
    "varithorn@nativejump.co",
    "nawarat@nativejump.co",
]
GMAIL_USER      = os.environ.get("GMAIL_USER",          "")
GMAIL_APP_PASS  = os.environ.get("GMAIL_APP_PASSWORD",  "")

ANTHROPIC_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
IMAP_HOST       = "imap.gmail.com"
IMAP_MAX_WAIT   = 8    # นาที รอ Excel email (เผื่อ retry 2 รอบให้พอใน 25 นาที)
IMAP_POLL_SEC   = 30   # วินาที poll แต่ละครั้ง
PIPELINE_RETRIES = 2   # ลองรันทั้ง pipeline กี่รอบก่อนแจ้ง error
MIN_MESSAGES     = 5    # ถ้าข้อความที่ใช้ได้น้อยกว่านี้ = ข้อมูลผิดปกติ แจ้งทีม verify (ปกติ 80-120/วัน)

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

        # Login (retry สูงสุด 3 ครั้ง — ZE บางทีช้า/ไม่ redirect ไป /home)
        logged_in = False
        for attempt in range(3):
            try:
                await page.goto("https://zocialeye.wisesight.com/login", wait_until="domcontentloaded")
                await page.wait_for_selector("input[name='username']", timeout=20000)
                await page.fill("input[name='username']", ZOCIAL_ID)
                await page.fill("input[name='passwd']",   ZOCIAL_PASS)
                await page.click("#btn-login")
                # รอแค่ "ออกจากหน้า login" (อาจไป /home หรือ /campaigns ก็ได้)
                await page.wait_for_function(
                    "() => !location.pathname.includes('/login')", timeout=30000)
                logged_in = True
                break
            except Exception as e:
                print(f"  ✗ login attempt {attempt+1}/3 failed: {e}")
                await page.wait_for_timeout(3000)
        if not logged_in:
            raise RuntimeError("เข้าสู่ระบบ Zocial Eye ไม่สำเร็จหลังพยายาม 3 ครั้ง")

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

    # ดึงลิงก์โพสต์จริงจากแต่ละ row (ไล่ตามลำดับความตรง)
    def best_link(row):
        for col in ("Direct URL", "Post URL", "Comment URL", "Reply comment URL"):
            v = row.get(col)
            if pd.notna(v) and str(v).strip().lower().startswith("http"):
                return str(v).strip()
        return ""

    # สร้าง message list ส่งให้ Claude + เก็บลิงก์ไว้ map กลับด้วย id
    # ข้ามแถวที่ไม่มีเนื้อหา (NaN/ว่าง) — ไม่มีอะไรให้วิเคราะห์
    messages_for_claude = []
    id_to_link = {}
    for i, row in df.iterrows():
        msg_text = row.get("Message", "")
        if pd.isna(msg_text) or not str(msg_text).strip():
            continue
        id_to_link[i] = best_link(row)
        messages_for_claude.append({
            "id":        i,
            "account":   str(row.get("Account", "-")),
            "source":    str(row.get("Source", "-")),
            "brand":     str(row.get("Main keyword", "-")),
            "sentiment": str(row.get("Sentiment", "-")),
            "message":   str(msg_text)[:2000],
            "post_time": str(row.get("Post time", "")),
        })

    usable_total = len(messages_for_claude)

    # ข้อมูลว่าง/น้อยผิดปกติ → ไม่วิเคราะห์ ส่งสัญญาณให้แจ้งทีม verify (ไม่ใช่บอกลูกค้าว่า No Crisis)
    if usable_total < MIN_MESSAGES:
        return {
            "date":         (now_th() - timedelta(days=1)).strftime("%d %b %Y"),
            "total":        usable_total, "usable_total": usable_total, "raw_rows": total,
            "neg_ze":       neg_ze, "pos_ze": pos_ze,
            "crisis_count": 0, "crisis_rows": [], "all_crisis": [], "brand_counts": {},
            "low_data":     True,   # ไม่ต้อง retry — re-export ได้ข้อมูลเดิม
            "summary":      f"ดึงข้อมูลได้แต่มีข้อความที่ใช้ได้เพียง {usable_total} รายการ (ปกติ 80-120)",
            "error":        f"ข้อมูลน้อยผิดปกติ: มีข้อความที่มีเนื้อหาจริงเพียง {usable_total} รายการ (raw {total} แถว, ปกติ 80-120) — อาจเป็นปัญหา export/ZE หรือวันนั้นไม่มีข้อมูลจริง โปรดเข้า ZE ตรวจสอบเองว่ามี crisis หรือไม่",
        }

    print(f"  → Sending {usable_total} messages to Claude for analysis...")
    claude_result = claude_analyze(messages_for_claude)

    # แนบลิงก์โพสต์จริงกลับเข้าแต่ละ crisis item (join ด้วย id)
    for it in claude_result.get("crisis_items", []):
        try:
            it["link"] = id_to_link.get(int(it.get("id")), "")
        except (ValueError, TypeError):
            it["link"] = ""

    # เรียง crisis ตามความรุนแรง high→medium→low (กัน high หลุดท้ายแถว)
    sev_rank = {"high": 0, "medium": 1, "low": 2}
    items = sorted(claude_result.get("crisis_items", []),
                   key=lambda r: sev_rank.get(str(r.get("severity", "low")).lower(), 3))

    return {
        "date":         (now_th() - timedelta(days=1)).strftime("%d %b %Y"),
        "total":        usable_total,
        "usable_total": usable_total,
        "raw_rows":     total,
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

    prompt = f"""คุณเป็น Social Media Crisis Analyst สำหรับแบรนด์ความงามและสุขภาพในไทย

แบรนด์ที่ต้องติดตาม (alert เฉพาะที่เกี่ยวกับแบรนด์เหล่านี้):
- Xeomin = โบทูลินั่มท็อกซิน (เรียกว่า "โบ", ซีโอมิน, Xeomin)
- Belotero / Belotero Revive = ฟิลเลอร์ HA (เบลโลเทโร่, เบโลเทโร, เบลโลเทโร่ รีไวฟ์)
- Ulthera / Ulthera Prime = HIFU ยกกระชับ (อัลเทอร่า, อัลเธอร่า, อัลทีร่า, ไฮฟู่/HIFU ยกกระชับ, "เครื่องคลื่นเสียงยกกระชับ")
- Radiesse = ฟิลเลอร์/คอลลาเจนสติมูเลเตอร์ CaHA (เรดิเอส, เรดิแอส, เรดิเอซ)

**กฎข้อแรกสำคัญที่สุด — flag เฉพาะข้อความที่มี "สัญญาณลบ/อันตราย/เสี่ยง" จริงเท่านั้น**
ถ้าข้อความเป็นรีวิวบวก ชม พอใจ เชียร์แบรนด์ เปรียบเทียบว่าแบรนด์คุ้ม/ดีกว่า หรือเป็นกลาง → **ไม่ใช่ crisis** ไม่ต้อง flag เด็ดขาด
- อย่าเชื่อ ZE_sentiment ถ้าเนื้อหาจริงเป็นบวก (ZE tag ผิดได้) — เช่น "เอาเงินไปทำ ulthera ดีกว่ามั้ย" = ชม Ulthera ว่าคุ้ม = ไม่ใช่ crisis แม้ ZE=Negative
- การมีลิงก์/รูป/คลิป เพียงอย่างเดียว **ไม่ใช่** สัญญาณ crisis — โพสต์รีวิวบวกที่แนบรูปก่อน-หลัง = ไม่ใช่ crisis

ข้อ 1-5 ด้านล่างเป็นวิธี "จับสัญญาณลบที่ซ่อนอยู่" ใช้เฉพาะเมื่อมีเค้าลางลบ/อันตรายอยู่แล้ว ห้ามใช้ flag โพสต์บวก:
1. อ้างถึงแบรนด์แบบทับศัพท์ไทย/สะกดผิด/ชื่อเล่น/generic ("โบ", "ฟิลเลอร์ใต้ตา", "ทำไฮฟู่") + มีบริบทลบ → flag (ระบุ brand ว่า "ต้องตรวจสอบ" ถ้าไม่ชัด)
2. โทนประชดประชันที่ซ่อนเรื่องแย่ ("ขอบคุณคลินิกที่ทำให้ได้นอนรพ.")
3. romanized Thai หรือศัพท์การแพทย์อังกฤษเชิงลบ (necrosis, vascular occlusion, blindness, lumps, botched)
4. คนอื่นเล่าแทนถึงเรื่องแย่ ("แม่ไปฉีดมาแล้วหน้า...", "พาญาติไปแก้")
5. ชี้ไปรูป/คลิป/คอมเมนต์ **พร้อมเค้าลางลบ** ("ดูความเสียหายในคลิป", "อ่านที่คนเตือนในคอมเมนต์") — ไม่ใช่รูปรีวิวบวกทั่วไป

นิยาม crisis (severity = high ถ้าอันตรายถึงชีวิต/ตา/ถาวร หรือมีแนวโน้มไวรัล):
- อาการไม่พึงประสงค์: เนื้อตาย/หน้าเน่า/ผิวซีดเป็นลายๆ/ปวดมาก (vascular occlusion), ตาบอด/ตามัว, โบกระจาย→หายใจ/กลืน/พูด/เคี้ยวลำบาก (botulism), แพ้รุนแรง/ช็อก (anaphylaxis), แผลไหม้/เนื้อยุบถาวร (HIFU), ติดเชื้อ/เป็นหนอง, หน้าเบี้ยว/ปากเบี้ยว, เป็นก้อน/ฟิลเลอร์ไหล (ในเชิงผู้เสียหาย ไม่ใช่โพสต์ความรู้/เตือนเชิงวิชาการ)
- ของปลอม/หิ้ว/ไม่มี อย., หมอเถื่อน/พยาบาลฉีดเอง/คลินิกในคอนโด, ฉีดเด็ก/คนท้อง
- ร้องเรียน/แจ้งความ/ฟ้อง/อย./สคบ./สั่งปิด/เรียกคืนสินค้า
- คนถามหา "ฉีดสลาย/hyalase/เลาะฟิลเลอร์/หาหมอแก้เคสพัง" (= มีเคสเสียหายมาก่อน)
- สัญญาณไวรัล/รวมพลัง: โหนกระแส, สายไหมต้องรอด, รวมตัวผู้เสียหาย, ล่ารายชื่อ, ทัวร์ลง, แฉ
- ภาวะซึมเศร้า/ทำร้ายตัวเองหลังหน้าพัง

ไม่ใช่ crisis (ห้าม flag): ไม่เกี่ยวกับสินค้าข้างต้น, ข่าวสุขภาพทั่วไป, โปรโมชั่น/ขาย, **รีวิวบวก/ชม/พอใจ**, **เปรียบเทียบที่เชียร์แบรนด์**, **โพสต์รีวิวที่แนบรูปก่อน-หลังแบบบวก**, โพสต์ให้ความรู้/เตือนเชิงป้องกันทั่วไปจากคลินิก/หมอ (เว้นแต่มีเคสผู้เสียหายจริง)

หลักการ: พลาด crisis จริงไม่ได้ก็จริง แต่ "crisis จริง" ต้องมีสัญญาณลบ/อันตราย — อย่าเปลี่ยนรีวิวบวกหรือโพสต์เป็นกลางให้กลายเป็น crisis

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
    def esc(s):  # กันอักขระพิเศษ (&, <, >) ทำ Paragraph XML พัง
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    if crisis and result.get("crisis_rows"):
        story.append(Paragraph("รายละเอียด (เรียงตามความรุนแรง)", head_style))
        severity_map = {"high": "สูง", "medium": "กลาง", "low": "ต่ำ"}
        severity_color = {"high": "#c0392b", "medium": "#e67e22", "low": "#f1c40f"}

        for i, r in enumerate(result["crisis_rows"], 1):
            sev = str(r.get("severity", "low")).lower()
            sev_th = severity_map.get(sev, sev)
            sev_c  = colors.HexColor(severity_color.get(sev, "#f1c40f"))

            row_data = [[
                Paragraph(f"<b>[{i}] @{esc(r.get('account','-'))}</b> ({esc(r.get('source','-'))}) — แบรนด์: {esc(r.get('brand','-'))}", normal_style),
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
            story.append(Paragraph(f"เหตุผล: {esc(r.get('reason',''))}", small_style))
            story.append(Paragraph(f"\"{esc(r.get('message_preview','')[:120])}\"", small_style))
            if r.get("link"):
                link_style = ParagraphStyle("lnk", fontName=thai_font, fontSize=9, textColor=colors.HexColor("#1a5fb4"), leading=12)
                story.append(Paragraph(f'🔗 <link href="{esc(r["link"])}">ดูโพสต์ต้นฉบับ</link>', link_style))
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
        # วิเคราะห์/ดึงข้อมูลพัง — ส่งหา ADMIN (ทีมที่เข้า ZE ได้) ไม่ส่งหาลูกค้า
        subject = f"[⚠️ Monitor มีปัญหา] ต้องตรวจสอบเอง — {date}"
        body = f"""[แจ้งทีม NativeJump — ไม่ใช่รายงานลูกค้า]
รายงานประจำวัน: {date}
================================================
สถานะ: ⚠️ ระบบลองใหม่อัตโนมัติแล้วแต่ยังไม่สำเร็จ

ปัญหา:
{err}

หมายเหตุ: วันนี้ระบบ "ยังไม่ได้" ยืนยันว่าไม่มี crisis
รบกวนทีมเข้า Zocial Eye ตรวจสอบด้วยตนเอง แล้วแจ้งลูกค้าหากพบสิ่งผิดปกติ:
https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_ID}/all/message

ภาพรวมเท่าที่ได้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
================================================
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""
    elif crisis:
        # แยกระดับ: มี high/medium = alert จริง | มีแค่ low = แค่ให้ตรวจสอบ (กัน alarm fatigue)
        all_items = result.get("all_crisis", crisis_rows)
        sevs    = [str(r.get("severity", "")).lower() for r in all_items]
        n_high  = sevs.count("high")
        n_med   = sevs.count("medium")
        n_low   = sevs.count("low")
        n_hm    = n_high + n_med

        shown = len(crisis_rows)
        more_note = f"\n(แสดง {shown} จากทั้งหมด {crisis_count} รายการ — ดูครบในไฟล์แนบ)" if crisis_count > shown else ""

        if n_hm > 0:
            subject     = f"[CRISIS ALERT] พบเรื่องสำคัญ {n_hm} รายการ (สูง {n_high}/กลาง {n_med}) — {date}"
            status_line = f"CRISIS DETECTED — พบระดับสูง {n_high}, ระดับกลาง {n_med}" + (f", ระดับต่ำ(ต้องตรวจสอบ) {n_low}" if n_low else "")
        else:
            subject     = f"[ตรวจสอบ] พบ {n_low} รายการระดับต่ำที่ควรดู — {date}"
            status_line = f"ไม่พบระดับสูง/กลาง — มี {n_low} รายการระดับต่ำที่ควรตรวจสอบ"

        hits_txt = "\n".join([
            f"  [{i+1}] @{r.get('account','-')} ({r.get('source','-')}) "
            f"| แบรนด์: {r.get('brand','-')} | ระดับ: {severity_map.get(str(r.get('severity','')).lower(),'?')}\n"
            f"       เหตุผล: {r.get('reason','')}\n"
            f"       \"{r.get('message_preview', r.get('message',''))[:120]}\"\n"
            + (f"       ลิงก์: {r.get('link')}\n" if r.get('link') else "")
            for i, r in enumerate(crisis_rows)
        ])
        brands_txt = "\n".join([f"  - {b}: {c} ข้อความ" for b, c in sorted(brand_counts.items(), key=lambda x: -x[1])])
        body = f"""รายงานประจำวัน: {date}
================================================
สถานะ: {status_line}

สรุปผลการวิเคราะห์:
{claude_summary}

ภาพรวมวันนี้
  - ข้อความทั้งหมด:      {total} รายการ
  - ZE ระบุ Negative:   {neg_ze} รายการ
  - พบที่ต้องดู:         {crisis_count} รายการ (สูง {n_high} / กลาง {n_med} / ต่ำ {n_low})

แบรนด์ที่ถูกพูดถึง:
{brands_txt}

รายละเอียด (เรียงตามความรุนแรง):{more_note}
{hits_txt}
================================================
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

    # error → ส่งเฉพาะทีม NativeJump | สำเร็จ → ส่งครบทั้งลูกค้า + ทีม
    recipients = ADMIN_RECIPIENTS if err else REPORT_RECIPIENTS

    if not GMAIL_USER or not GMAIL_APP_PASS:
        path = f"/tmp/crisis_report_{datetime.now():%Y%m%d}.txt"
        open(path, "w").write(f"To: {', '.join(recipients)}\nSubject: {subject}\n\n{body}")
        print(f"  ⚠ No Gmail config — saved to {path}")
        return

    msg = MIMEMultipart()
    msg["Subject"]  = subject
    msg["From"]     = GMAIL_USER
    msg["To"]       = ", ".join(recipients)
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
            smtp.sendmail(GMAIL_USER, recipients, msg.as_string())
        print(f"  ✓ Summary sent to {len(recipients)} recipients: {', '.join(recipients)}")
    except Exception as e:
        print(f"  ✗ Email error: {e}")
        raise   # ส่งอีเมลไม่ได้ = ต้องรู้ ห้ามกลืน error เงียบ

# ─── Main ─────────────────────────────────────────────────────────────────────
async def run_pipeline(report_date: str):
    """รัน pipeline 1 รอบ — สำเร็จ = ส่งรายงานให้ลูกค้า | พังที่ไหน = raise (ให้ main ลองใหม่/แจ้ง admin)"""
    triggered_at = datetime.utcnow()

    # 1. Trigger export
    total = await trigger_export()
    if total == '0':
        # แบรนด์เหล่านี้มีคนพูดถึงทุกวัน → 0 = ผิดปกติ (น่าจะ scrape/timezone พลาด) ให้ลองใหม่
        raise RuntimeError("Zocial Eye อ่านได้ 0 ข้อความ (ผิดปกติสำหรับแบรนด์เหล่านี้) — อาจดึงข้อมูลผิดพลาด")

    # 2. Fetch Excel from email
    xlsx_path = fetch_excel_from_email(triggered_at)
    if not xlsx_path:
        raise RuntimeError("ดึงไฟล์ Excel จากอีเมลไม่สำเร็จภายในเวลาที่กำหนด")

    # 3. Analyze
    print("  → Analyzing Excel...")
    result = analyze_excel(xlsx_path)
    print(f"  → Total: {result['total']} | ZE Negative: {result['neg_ze']} | Crisis hits: {result['crisis_count']}")

    # ข้อมูลน้อย/ว่างผิดปกติ → แจ้งทีม verify เลย ไม่ retry (re-export ได้ข้อมูลเดิม) ไม่ส่งลูกค้า
    if result.get("low_data"):
        print(f"  ⚠ Low data ({result['total']} usable) — alerting team to verify, not sending client report")
        send_summary(result)
        return

    if result.get("error"):
        raise RuntimeError(result["error"])

    # 4. Generate PDF + 5. ส่งรายงานให้ลูกค้า
    print("  → Generating PDF report...")
    pdf_path = create_pdf_report(result)
    send_summary(result, xlsx_path, pdf_path)


async def main():
    print(f"[{now_th():%H:%M}] Zocial Eye Crisis Monitor starting...")
    report_date = (now_th() - timedelta(days=1)).strftime("%d %b %Y")

    last_err = None
    for attempt in range(1, PIPELINE_RETRIES + 1):
        try:
            await run_pipeline(report_date)
            print(f"[{now_th():%H:%M}] Done (attempt {attempt}).")
            return
        except Exception as e:
            last_err = e
            print(f"  ✗ pipeline attempt {attempt}/{PIPELINE_RETRIES} failed: {e}")
            if attempt < PIPELINE_RETRIES:
                print("  → ลองใหม่อัตโนมัติใน 60 วินาที...")
                await asyncio.sleep(60)

    # ลองครบทุกรอบแล้วยังพัง → แจ้ง admin (ไม่ใช่ลูกค้า) ห้ามเงียบ
    print(f"  ✗✗ ALL {PIPELINE_RETRIES} ATTEMPTS FAILED:\n{traceback.format_exc()}")
    try:
        send_summary({"date": report_date, "total": "?", "neg_ze": "?",
                      "crisis_count": 0, "crisis_rows": [], "brand_counts": {},
                      "error": f"{last_err}"})
    except Exception as e2:
        print(f"  ✗✗ Even the admin alert email failed: {e2}")
    sys.exit(1)   # GitHub Action ขึ้นแดง

if __name__ == "__main__":
    asyncio.run(main())
