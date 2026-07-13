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
CAMPAIGN_ID     = os.environ.get("CAMPAIGN_ID",         "93082")   # backward-compat (= brand)
# 2 แคมเปญ: แบรนด์ (Merz) + generic หัตถการ — รวมในอีเมลเดียว 2 section
CAMPAIGN_BRAND   = os.environ.get("CAMPAIGN_BRAND",   "93082")    # เรื่องที่ 1: แบรนด์เรา
CAMPAIGN_GENERIC = os.environ.get("CAMPAIGN_GENERIC", "104883")   # เรื่องที่ 2: ข่าวหมวดฟิลเลอร์/หัตถการ
EXPORT_EMAIL    = os.environ.get("EXPORT_EMAIL",        "kanthorn@nativejump.co")
# ผู้รับรายงานประจำวันที่สำเร็จ — ลูกค้า Merz + ทีม NativeJump (เป็น config ไม่ใช่ความลับ)
REPORT_RECIPIENTS = [
    "kamolrat.p@merz.com",
    "sarun.chompaisal@merz.com",
    "maytita.t@merz.com",
    "kanthorn@nativejump.co",
    "varithorn@nativejump.co",
    "nawarat@nativejump.co",
    "chaithawat@nativejump.co",
]
# ผู้รับแจ้งเตือน error เท่านั้น — เฉพาะทีม NativeJump ที่เข้า ZE ได้ (ไม่ส่งหาลูกค้า Merz)
ADMIN_RECIPIENTS = [
    "kanthorn@nativejump.co",
    "varithorn@nativejump.co",
    "nawarat@nativejump.co",
    "chaithawat@nativejump.co",
]
GMAIL_USER      = os.environ.get("GMAIL_USER",          "")
GMAIL_APP_PASS  = os.environ.get("GMAIL_APP_PASSWORD",  "")

ANTHROPIC_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
IMAP_HOST       = "imap.gmail.com"
IMAP_MAX_WAIT   = 8    # นาที รอ Excel email (เผื่อ retry 2 รอบให้พอใน 25 นาที)
IMAP_POLL_SEC   = 30   # วินาที poll แต่ละครั้ง
PIPELINE_RETRIES = 2   # ลองรันทั้ง pipeline กี่รอบก่อนแจ้ง error
MIN_MESSAGES     = 5    # ถ้าข้อความที่ใช้ได้น้อยกว่านี้ = ข้อมูลผิดปกติ แจ้งทีม verify (ปกติ 80-120/วัน)
MSG_TRUNC        = 450  # ตัดข้อความต่อ row (พอตัดสิน crisis + คุม token/rate limit)
LLM_CHUNK_CHARS  = 11000 # งบตัวอักษรข้อความต่อ 1 chunk (~5k tokens + instruction ~3.5k < 10k/นาที)
LLM_MIN_INTERVAL = 65   # วินาที เว้นระหว่างเรียก Claude (org limit 10k input tokens/นาที)
_last_llm_call   = [0.0]  # throttle state (เวลาเรียกครั้งล่าสุด)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def yesterday_str():
    return (now_th() - timedelta(days=1)).strftime("%-d %b %Y")

def all_messages_url(campaign_id):
    d = yesterday_str()
    return (
        f"https://zocialeye.wisesight.com/campaigns/{campaign_id}/all/message"
        f"?start={d.replace(' ', '+')}&end={d.replace(' ', '+')}&action=filter"
    )

# ─── Step 1: Playwright — login + trigger export ───────────────────────────────
async def trigger_export(campaign_id=CAMPAIGN_BRAND):
    print(f"  → Logging in to Zocial Eye (campaign {campaign_id})...")
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
        await page.goto(all_messages_url(campaign_id))
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
def fetch_excel_from_email(triggered_at: datetime, campaign_id=CAMPAIGN_BRAND) -> str | None:
    print(f"  → Waiting for Excel email for campaign {campaign_id} (up to {IMAP_MAX_WAIT} min)...")
    deadline = time.time() + IMAP_MAX_WAIT * 60
    # ไฟล์ export มีชื่อ ZE_all_message_on_<campaign>(...) → match เฉพาะแคมเปญนี้
    url_re = re.compile(rf'https://downloads\.zocialeye\.com/[^\s\'"<>]*on_{campaign_id}\([^\s\'"<>]+\.xlsx')

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

                url_match = url_re.search(body)
                if url_match:
                    url = url_match.group(0)
                    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    urllib.request.urlretrieve(url, tmp.name)
                    print(f"  → Excel downloaded (campaign {campaign_id}): {url}")
                    return tmp.name

            print(f"  → Not yet, retrying in {IMAP_POLL_SEC}s...")
            time.sleep(IMAP_POLL_SEC)

    print("  ✗ Timed out waiting for Excel email")
    return None

# ─── Step 3: Claude วิเคราะห์ทุก row ─────────────────────────────────────────
def analyze_excel(xlsx_path: str, scope: str = "brand") -> dict:
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
            "message":   str(msg_text)[:MSG_TRUNC],
            "post_time": str(row.get("Post time", "")),
        })

    usable_total = len(messages_for_claude)
    base = {
        "date":  (now_th() - timedelta(days=1)).strftime("%d %b %Y"),
        "scope": scope, "total": usable_total, "usable_total": usable_total, "raw_rows": total,
        "neg_ze": neg_ze, "pos_ze": pos_ze,
    }

    # แบรนด์: ข้อมูลน้อยผิดปกติ = ปัญหา (แจ้งทีม verify) | generic: น้อยได้ตามปกติ ไม่ใช่ error
    if scope == "brand" and usable_total < MIN_MESSAGES:
        return {**base, "crisis_count": 0, "crisis_rows": [], "all_crisis": [], "brand_counts": {},
                "low_data": True,
                "summary": f"ดึงข้อมูลแบรนด์ได้แต่มีข้อความที่ใช้ได้เพียง {usable_total} รายการ (ปกติ 80-120)",
                "error":   f"ข้อมูลแบรนด์น้อยผิดปกติ: มีข้อความจริงเพียง {usable_total} รายการ (raw {total} แถว) — อาจเป็นปัญหา export/ZE quota เต็ม โปรดเข้า ZE ตรวจสอบเอง"}
    if usable_total == 0:
        return {**base, "crisis_count": 0, "crisis_rows": [], "all_crisis": [], "brand_counts": {},
                "summary": "ไม่พบข้อความในช่วงเวลานี้"}

    print(f"  → Sending {usable_total} messages to Claude ({scope}) for analysis...")
    claude_result = claude_analyze(messages_for_claude, scope)

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
        **base,
        "crisis_count": claude_result["crisis_count"],
        "crisis_rows":  items[:15],          # โชว์ได้ถึง 15 (เรียงตามรุนแรง)
        "all_crisis":   items,               # เก็บครบไว้แนบไฟล์
        "summary":      claude_result["summary"],
        "brand_counts": claude_result["brand_counts"],
        "error":        claude_result.get("error"),
    }


# ใช้ tool-use บังคับ structured output → การันตี JSON ถูกต้องเสมอ (เลิกพึ่ง json.loads เอง)
CRISIS_TOOL = {
    "name": "report_crisis",
    "description": "รายงานผลวิเคราะห์ crisis จากข้อความที่ให้มา",
    "input_schema": {
        "type": "object",
        "properties": {
            "crisis_items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "integer"},
                        "account": {"type": "string"},
                        "source": {"type": "string"},
                        "brand": {"type": "string"},
                        "reason": {"type": "string"},
                        "severity": {"type": "string", "enum": ["low", "medium", "high"]},
                        "message_preview": {"type": "string"},
                    },
                    "required": ["id", "reason", "severity"],
                },
            },
            "summary": {"type": "string"},
        },
        "required": ["crisis_items", "summary"],
    },
}

def _call_claude(client, prompt, label=""):
    """เรียก Claude แบบ tool-use (structured output) + throttle + retry 3 ครั้ง. คืน dict หรือ raise."""
    last_err = None
    for attempt in range(3):
        wait = LLM_MIN_INTERVAL - (time.time() - _last_llm_call[0])
        if wait > 0:
            time.sleep(wait)
        _last_llm_call[0] = time.time()
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-6", max_tokens=8000,
                tools=[CRISIS_TOOL],
                tool_choice={"type": "tool", "name": "report_crisis"},
                messages=[{"role": "user", "content": prompt}])
            for block in resp.content:
                if getattr(block, "type", None) == "tool_use":
                    return block.input   # dict ที่ valid ตาม schema แล้ว
            raise ValueError("no tool_use block in response")
        except Exception as e:
            last_err = e
            print(f"  ✗ claude [{label}] attempt {attempt+1}/3: {e}")
            time.sleep(5)
    raise RuntimeError(str(last_err))


def claude_analyze(messages: list, scope: str = "brand") -> dict:
    """แบ่ง message เป็น chunk แล้ววิเคราะห์ทีละชุด (กัน rate limit) — รวมผลทุก chunk"""
    if not ANTHROPIC_KEY:
        return {"crisis_count": 0, "crisis_items": [], "summary": "No API key", "brand_counts": {}}

    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    def _fmt(m):
        return f"[{m['id']}] @{m['account']} ({m['source']}) brand={m['brand']} ZE_sentiment={m['sentiment']}\n{m['message']}"

    # แบ่ง messages เป็น chunk ตามงบตัวอักษร (กัน rate limit 10k input tokens/นาที)
    chunks, cur, cur_len = [], [], 0
    for m in messages:
        L = len(_fmt(m))
        if cur and cur_len + L > LLM_CHUNK_CHARS:
            chunks.append(cur); cur, cur_len = [], 0
        cur.append(m); cur_len += L
    if cur:
        chunks.append(cur)

    if scope == "generic":
        role = ('คุณเป็น Social Media Crisis Analyst เฝ้าระวัง "ข่าว/ประเด็นหมวดหัตถการความงาม" '
                '(ฟิลเลอร์/โบท็อก/ไฮฟู่/ยกกระชับ/ฉีดสารเติมเต็ม) ในภาพรวมอุตสาหกรรมไทย\n\n'
                'หน้าที่: สรุปข่าว/โพสต์หมวดนี้ที่ "มีปัญหา/วิกฤต" — เช่น เสียชีวิต/ตาบอด/หน้าพังจากฟิลเลอร์, '
                'จับของเถื่อน/คลินิกเถื่อน, ฟ้องร้อง/อย.จับ, ข่าวไวรัลเชิงลบ — แม้ไม่เอ่ยชื่อ Xeomin/Belotero/Ulthera/Radiesse ก็ให้ flag '
                '(ถ้าพาดพิงแบรนด์เหล่านี้ด้วย ให้ระบุใน field brand)')
        brand_framing = "แบรนด์ของลูกค้า (ใช้อ้างอิง — ถ้าเคสพาดพิงให้ระบุชื่อใน brand แต่ไม่จำเป็นต้องเกี่ยวแบรนด์นี้ถึงจะ flag):"
    else:
        role = "คุณเป็น Social Media Crisis Analyst สำหรับแบรนด์ความงามและสุขภาพในไทย"
        brand_framing = "แบรนด์ที่ต้องติดตาม (alert เฉพาะที่เกี่ยวกับแบรนด์เหล่านี้) — รวมชื่อเรียก/สะกดที่ลูกค้าใช้:"

    def make_prompt(msgs_text):
        return f"""{role}

{brand_framing}
- Xeomin = โบทูลินั่มท็อกซิน — Xeomin, ซีโอมิน, โอมิน, "โบ", "โบเยอรมัน" (โบสัญชาติเยอรมัน)
- Belotero / Belotero Revive = ฟิลเลอร์ HA — Belotero, เบลโลเทโร่, เบโลเทโร, เบโล, เบโลรีไวฟ์, เบลโลเทโร่ รีไวฟ์, ฟิลเลอร์เบโล
- Ulthera / Ulthera Prime = HIFU ยกกระชับ — Ulthera, อัลเทอร่า, อัลเทอรา, อัลเธอร่า, อัลทีร่า, อัลเทอร่าไพรม์, ไฮฟู่/HIFU ยกกระชับ, "เครื่องคลื่นเสียงยกกระชับ"
- Radiesse = ฟิลเลอร์/คอลลาเจนสติมูเลเตอร์ CaHA — Radiesse, เรดิเอส, เรเดียส, เรดิแอส, เรดิเอซ, R+, R Plus

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

เกณฑ์การให้ระดับความรุนแรง (severity) — อิงคำที่พบ โดยคำต้องอยู่ในบริบทเกี่ยวกับแบรนด์/การฉีดเท่านั้น (ไม่ใช่ความหมายทั่วไป เช่น "ก้อนเมฆ", "ไตวาย" ในข่าวสุขภาพ):

🔴 high — crisis ชัดเจน:
- ดำเนินคดี/ทางการ: ฟ้องร้อง, แจ้งความ, ร้องเรียน, อย., สคบ., สั่งปิด, เรียกคืนสินค้า
- ของปลอม/ผิดกฎหมาย: ของปลอม, ปลอม, ก็อป, ไม่แท้, หิ้ว, ลักลอบ, ยาปลอม
- หลอกลวง: หลอกลวง, โกง, ตุ๋น, หลอก
- บาดเจ็บ/อาการรุนแรง: หน้าเสีย, หน้าพัง, แพ้, บวม, อักเสบ, ติดเชื้อ, เนื้อตาย, ตาบอด, ฉีดพลาด
- ร้ายแรงสุด: ตาย, อันตราย, เสียโฉม, พิการ
- ภาวะฉุกเฉินทางการแพทย์ (high เสมอ): vascular occlusion (ผิวซีดเป็นลายๆ/ปวดมาก/หน้าเน่า), โบกระจาย→หายใจ/กลืนลำบาก, แพ้รุนแรง/ช็อก
- สัญญาณไวรัล/รวมพลัง: โหนกระแส, สายไหมต้องรอด, รวมตัวผู้เสียหาย, ล่ารายชื่อ, ทัวร์ลง, แฉ

🟠 medium — ต้องจับตา:
- ไม่ได้ผล: ไม่เห็นผล, ไม่เวิร์ค, เสียเงินฟรี, เสียดายเงิน
- อาการหลังฉีด: เป็นก้อน/nodule, ก้อนแข็ง, เป็นไต, ไหล, ย้อย, เบี้ยว/หน้าเบี้ยว, ไม่เท่ากัน
- ผู้ให้บริการเถื่อน: หมอเถื่อน, คลินิกเถื่อน, ไม่มีใบอนุญาต, พยาบาลฉีดเอง, ฉีดเด็ก/คนท้อง
- เจ็บ/ช้ำ: เจ็บมาก, ช้ำ, เขียว
- ถามหาวิธีแก้ (= มีเคสพังมาก่อน): ฉีดสลาย, hyalase, เลาะฟิลเลอร์, หาหมอแก้

🟡 low — negative ทั่วไป:
- ราคา/ความคุ้ม: แพง, ไม่คุ้ม, ผิดหวัง, เสียใจ
- ความคงทน: อยู่ได้ไม่นาน, สลายเร็ว, ไม่ติด

ไม่ใช่ crisis (ห้าม flag): ไม่เกี่ยวกับสินค้าข้างต้น, ข่าวสุขภาพทั่วไป, โปรโมชั่น/ขาย, **รีวิวบวก/ชม/พอใจ**, **เปรียบเทียบที่เชียร์แบรนด์**, **โพสต์รีวิวที่แนบรูปก่อน-หลังแบบบวก**, โพสต์ให้ความรู้/เตือนเชิงป้องกันทั่วไปจากคลินิก/หมอ (เว้นแต่มีเคสผู้เสียหายจริง)

หลักการ: พลาด crisis จริงไม่ได้ก็จริง แต่ "crisis จริง" ต้องมีสัญญาณลบ/อันตราย — อย่าเปลี่ยนรีวิวบวกหรือโพสต์เป็นกลางให้กลายเป็น crisis

ข้อความทั้งหมด:
{msgs_text}

รายงานผลผ่านเครื่องมือ report_crisis: ใส่ทุกเคสที่เป็น crisis ใน crisis_items (id, account, source, brand, reason, severity=low/medium/high, message_preview=120 ตัวอักษรแรก) และ summary สรุป 2-3 ประโยค ถ้าไม่มี crisis ให้ crisis_items เป็น [] และ summary บอกว่าไม่พบ"""

    # วิเคราะห์ทีละ chunk (throttle กัน rate limit) แล้วรวมผล
    print(f"  → {scope}: {len(messages)} msgs ใน {len(chunks)} chunk")
    all_items, brand_counts, failed = [], {}, 0
    for ci, chunk in enumerate(chunks, 1):
        prompt = make_prompt("\n".join(_fmt(m) for m in chunk))
        try:
            res = _call_claude(client, prompt, f"{scope} {ci}/{len(chunks)}")
        except Exception as e:
            # chunk เดียวพัง → ข้าม เก็บ chunk อื่นไว้ (ไม่ซ่อนทั้ง section)
            failed += 1
            print(f"  ⚠ chunk {ci}/{len(chunks)} ข้าม: {e}")
            continue
        all_items.extend(res.get("crisis_items", []) or [])
        for b, c in (res.get("brand_counts") or {}).items():
            try:
                brand_counts[b] = brand_counts.get(b, 0) + int(c)
            except (ValueError, TypeError):
                pass

    # สรุปจากผลจริง (ไม่เอา summary ของแต่ละ chunk มาต่อ เพราะจะขัดกันเอง)
    if not all_items:
        summary = "ตรวจสอบข้อความทั้งหมดแล้ว ไม่พบประเด็นที่ต้องติดตาม"
    else:
        nh = sum(1 for it in all_items if str(it.get("severity","")).lower() == "high")
        nm = sum(1 for it in all_items if str(it.get("severity","")).lower() == "medium")
        nl = sum(1 for it in all_items if str(it.get("severity","")).lower() == "low")
        tops = [str(it.get("reason","")).strip() for it in all_items
                if str(it.get("severity","")).lower() in ("high","medium")][:3]
        summary = f"พบ {len(all_items)} รายการที่ต้องดู (สูง {nh}/กลาง {nm}/ต่ำ {nl})"
        if tops:
            summary += " — เด่น: " + " / ".join(tops)
        summary = summary[:600]

    out = {"crisis_count": len(all_items), "crisis_items": all_items,
           "summary": summary, "brand_counts": brand_counts}
    if failed:
        out["partial"] = failed   # บาง chunk ข้าม — แจ้งทีม แต่ยังโชว์ที่เจอ
        if failed == len(chunks):
            # พังทุก chunk = วิเคราะห์ไม่ได้เลย → error (ห้ามรายงาน 0 มั่ว)
            out["error"] = f"วิเคราะห์ {scope} ไม่สำเร็จทุก chunk ({failed}/{len(chunks)}) — ต้องตรวจสอบเอง"
    return out

# ─── Step 3b: สร้าง PDF รายงาน ────────────────────────────────────────────────
def create_pdf_report(result: dict, section_title: str = "Zocial Eye Crisis Monitor") -> str:
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
    story.append(Paragraph(section_title, title_style))
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


# ─── Step 4: ส่งอีเมลสรุป (2 section: แบรนด์ + หมวดทั่วไป) ───────────────────────
def _sev_counts(result):
    items = result.get("all_crisis", []) or result.get("crisis_rows", []) or []
    s = [str(r.get("severity", "")).lower() for r in items]
    return s.count("high"), s.count("medium"), s.count("low")

def _section_text(result, heading, client_safe=False):
    if result is None:
        return f"{heading}\n  (ไม่ได้ดึงข้อมูลส่วนนี้)\n"
    if result.get("error"):
        if client_safe:   # ลูกค้าเห็น — ห้ามโชว์ error ดิบ/technical
            return f"{heading}\n  รอบนี้ยังประมวลผลส่วนนี้ไม่สำเร็จ — ทีม NativeJump กำลังตรวจสอบ\n"
        return f"{heading}\n  ⚠️ ดึง/วิเคราะห์ไม่สำเร็จ: {result['error']}\n"
    cc = result.get("crisis_count", 0)
    head = (f"{heading}\n"
            f"  ข้อความ {result.get('total',0)} | ZE Negative {result.get('neg_ze',0)} | พบที่ต้องดู {cc}")
    if cc == 0:
        return f"{head}\n  สถานะ: ไม่พบประเด็น\n  สรุป: {result.get('summary','-')}\n"
    nh, nm, nl = _sev_counts(result)
    rows = result.get("crisis_rows", [])
    more = f" (แสดง {len(rows)}/{cc} — ดูครบในไฟล์แนบ)" if cc > len(rows) else ""
    sevm = {"high": "สูง", "medium": "กลาง", "low": "ต่ำ"}
    hits = "\n".join(
        f"  [{i+1}] @{r.get('account','-')} ({r.get('source','-')}) | {r.get('brand','-')} | ระดับ {sevm.get(str(r.get('severity','')).lower(),'?')}\n"
        f"       เหตุผล: {r.get('reason','')}\n"
        f"       \"{r.get('message_preview','')[:120]}\"\n"
        + (f"       ลิงก์: {r.get('link')}\n" if r.get('link') else "")
        for i, r in enumerate(rows))
    return (f"{head} (สูง {nh}/กลาง {nm}/ต่ำ {nl})\n"
            f"  สรุป: {result.get('summary','-')}\n"
            f"  รายละเอียด (เรียงตามความรุนแรง):{more}\n{hits}")

def _combined_subject(b, g, date):
    bh, bm, bl = _sev_counts(b); gh, gm, gl = _sev_counts(g or {})
    bcc = b.get("crisis_count", 0); gcc = (g or {}).get("crisis_count", 0)
    if bh + bm > 0:   tag = "CRISIS ALERT"
    elif bl > 0:      tag = "ตรวจสอบ-แบรนด์"
    elif gh + gm > 0: tag = "เฝ้าระวัง-หมวดทั่วไป"
    elif gl > 0:      tag = "ตรวจสอบ-หมวดทั่วไป"
    else:             tag = "No Crisis"
    return f"[{tag}] Daily Monitor — แบรนด์ {bcc} · หมวดทั่วไป {gcc} — {date}"

def send_combined(brand_result: dict, generic_result: dict = None, attachments=None):
    """ส่งอีเมลเดียว 2 section. brand error/low_data → ส่งทีมเท่านั้น | ปกติ → ลูกค้า+ทีม"""
    date = brand_result.get("date") or (generic_result or {}).get("date", "")
    brand_err = bool(brand_result.get("error"))
    to_admin  = brand_err   # ปัญหาฝั่งแบรนด์ = ส่งทีมเท่านั้น (ลูกค้าไม่เห็น)
    recipients = ADMIN_RECIPIENTS if to_admin else REPORT_RECIPIENTS

    subject = (f"[⚠️ Monitor มีปัญหา] ต้องตรวจสอบเอง — {date}" if brand_err
               else _combined_subject(brand_result, generic_result, date))

    # ลูกค้าเห็น (ไม่ใช่อีเมลแจ้งทีม) → generic section ห้ามโชว์ error ดิบ
    client_safe = not to_admin
    s1 = _section_text(brand_result,   "【 เรื่องที่ 1 】 แบรนด์เรา (Xeomin / Belotero / Ulthera / Radiesse)")
    s2 = _section_text(generic_result, "【 เรื่องที่ 2 】 ข่าวหมวดฟิลเลอร์ / หัตถการทั่วไป", client_safe=client_safe)
    prefix = "[แจ้งทีม NativeJump — ไม่ใช่รายงานลูกค้า]\n" if to_admin else ""
    note   = (f"\nหมายเหตุ: ระบบลองใหม่อัตโนมัติแล้วแต่ยังไม่สำเร็จ — รบกวนเข้า ZE ตรวจสอบเอง\n"
              f"https://zocialeye.wisesight.com/campaigns/{CAMPAIGN_BRAND}/all/message\n" if brand_err else "")
    body = f"""{prefix}รายงานประจำวัน: {date}
================================================
{s1}
------------------------------------------------
{s2}
================================================{note}
ส่งโดย: Zocial Eye Crisis Monitor (อัตโนมัติ)"""

    if not GMAIL_USER or not GMAIL_APP_PASS:
        path = f"/tmp/crisis_report_{now_th():%Y%m%d}.txt"
        open(path, "w").write(f"To: {', '.join(recipients)}\nSubject: {subject}\n\n{body}")
        print(f"  ⚠ No Gmail config — saved to {path}")
        return

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for fpath, fname in (attachments or []):
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
        print(f"  ✓ Sent to {len(recipients)}: {', '.join(recipients)}")
    except Exception as e:
        print(f"  ✗ Email error: {e}")
        raise

def _send_admin_note(subject, body):
    """แจ้งทีม NativeJump เท่านั้น (เช่น generic section พัง) — ลูกค้าไม่เห็น"""
    if not GMAIL_USER or not GMAIL_APP_PASS:
        print("  ⚠ no gmail — admin note skipped"); return
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject; msg["From"] = GMAIL_USER; msg["To"] = ", ".join(ADMIN_RECIPIENTS)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PASS)
            smtp.sendmail(GMAIL_USER, ADMIN_RECIPIENTS, msg.as_string())
        print(f"  ✓ admin note sent to {len(ADMIN_RECIPIENTS)}")
    except Exception as e:
        print(f"  ✗ admin note failed: {e}")

# ─── Main ─────────────────────────────────────────────────────────────────────
async def run_pipeline(report_date: str):
    """ดึง 2 แคมเปญ (แบรนด์+generic) → วิเคราะห์ → ส่งอีเมลเดียว 2 section.
    ฝั่งแบรนด์ critical (พัง=raise/retry); ฝั่ง generic best-effort (พังไม่บล็อกลูกค้า)."""
    triggered_at = datetime.utcnow()
    fn_date = report_date.replace(' ', '_')

    # 1. Trigger ทั้ง 2 แคมเปญ (ยิงติดกัน อีเมลจะมาพร้อมๆ กัน)
    brand_total = await trigger_export(CAMPAIGN_BRAND)
    if brand_total == '0':
        raise RuntimeError("แคมเปญแบรนด์อ่านได้ 0 ข้อความ (ผิดปกติ) — อาจ ZE quota เต็ม")
    try:
        generic_total = await trigger_export(CAMPAIGN_GENERIC)
    except Exception as e:
        print(f"  ⚠ trigger generic ล้มเหลว (ไม่บล็อก): {e}")
        generic_total = '?'

    # 2. Fetch — แบรนด์ต้องได้ | generic best-effort
    brand_xlsx = fetch_excel_from_email(triggered_at, CAMPAIGN_BRAND)
    if not brand_xlsx:
        raise RuntimeError("ดึงไฟล์ Excel แบรนด์จากอีเมลไม่สำเร็จภายในเวลาที่กำหนด")
    generic_xlsx = None
    if generic_total != '0':
        try:
            generic_xlsx = fetch_excel_from_email(triggered_at, CAMPAIGN_GENERIC)
        except Exception as e:
            print(f"  ⚠ fetch generic ล้มเหลว (ไม่บล็อก): {e}")

    # 3. Analyze แบรนด์ (critical)
    print("  → Analyzing brand campaign...")
    brand_result = analyze_excel(brand_xlsx, scope="brand")
    print(f"  → [brand] total {brand_result['total']} | crisis {brand_result['crisis_count']}")
    if brand_result.get("low_data"):
        print("  ⚠ Low data ฝั่งแบรนด์ — แจ้งทีม verify (ไม่ส่งลูกค้า)")
        send_combined(brand_result, None); return
    if brand_result.get("error"):
        raise RuntimeError(brand_result["error"])

    # 3b. Analyze generic (best-effort)
    if generic_xlsx:
        print("  → Analyzing generic campaign...")
        generic_result = analyze_excel(generic_xlsx, scope="generic")
        print(f"  → [generic] total {generic_result['total']} | crisis {generic_result['crisis_count']}")
    else:
        generic_result = {"date": report_date, "scope": "generic", "total": generic_total,
                          "neg_ze": 0, "crisis_count": 0, "crisis_rows": [], "all_crisis": [],
                          "brand_counts": {}, "summary": "ไม่พบข่าวหมวดฟิลเลอร์/หัตถการทั่วไปในช่วงเวลานี้"}

    # 4. PDF + แนบไฟล์ + ส่ง
    attachments = []
    attachments.append((create_pdf_report(brand_result, "รายงานแบรนด์ (Merz)"), f"ZE_brand_{fn_date}.pdf"))
    attachments.append((brand_xlsx, f"ZE_brand_{fn_date}.xlsx"))
    if generic_result.get("crisis_count", 0) > 0 and not generic_result.get("error"):
        attachments.append((create_pdf_report(generic_result, "รายงานหมวดฟิลเลอร์ทั่วไป"), f"ZE_generic_{fn_date}.pdf"))
    if generic_xlsx:
        attachments.append((generic_xlsx, f"ZE_generic_{fn_date}.xlsx"))

    send_combined(brand_result, generic_result, attachments)

    # ประมวลผลไม่ครบบางส่วน (chunk ข้าม/พัง) = ลูกค้าเห็นเฉพาะที่สำเร็จ แต่ทีมต้องรู้ไปเช็ก
    issues = []
    if brand_result.get("partial"):
        issues.append(f"แบรนด์: ข้าม {brand_result['partial']} chunk — อาจพลาดบางเคส โปรดเช็ก ZE")
    if generic_result.get("error"):
        issues.append(f"หมวดทั่วไป: {generic_result['error']}")
    elif generic_result.get("partial"):
        issues.append(f"หมวดทั่วไป: ข้าม {generic_result['partial']} chunk")
    if issues:
        _send_admin_note(f"[⚠️ ทีม] ประมวลผลไม่ครบบางส่วน — {report_date}",
                         "อีเมลลูกค้าส่งแล้ว (แสดงเฉพาะส่วนที่วิเคราะห์สำเร็จ) แต่มีบางส่วนไม่ครบ:\n\n"
                         + "\n".join(issues) + "\n\nดู log ใน GitHub Actions สำหรับรายละเอียด")


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
        send_combined({"date": report_date, "total": "?", "neg_ze": "?",
                       "crisis_count": 0, "crisis_rows": [], "all_crisis": [], "brand_counts": {},
                       "error": f"{last_err}"}, None)
    except Exception as e2:
        print(f"  ✗✗ Even the admin alert email failed: {e2}")
    sys.exit(1)   # GitHub Action ขึ้นแดง

if __name__ == "__main__":
    asyncio.run(main())
