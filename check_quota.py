"""วินิจฉัยสถานะ ZE: จำนวนข้อความที่หน้าแคมเปญเห็น + ข้อความเตือนเรื่อง quota"""
import asyncio, os, json
from playwright.async_api import async_playwright

ZID  = os.environ.get("ZOCIAL_ID", "")
ZPW  = os.environ.get("ZOCIAL_PASS", "")
CAMP = os.environ.get("CAMPAIGN", "93082")
DAY  = os.environ.get("DAY", "27 Aug 2026")

async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=True)
        pg = await b.new_page()
        await pg.goto("https://zocialeye.wisesight.com/login", wait_until="domcontentloaded")
        await pg.wait_for_selector("input[name='username']", timeout=20000)
        await pg.fill("input[name='username']", ZID); await pg.fill("input[name='passwd']", ZPW)
        await pg.click("#btn-login")
        await pg.wait_for_function("() => !location.pathname.includes('/login')", timeout=30000)

        # 1) หน้ารวมแคมเปญ — มักโชว์ quota ที่ใช้ไป
        await pg.goto("https://zocialeye.wisesight.com/campaigns", wait_until="domcontentloaded")
        await pg.wait_for_timeout(6000)
        info = await pg.evaluate("""() => {
            const t = document.body.innerText;
            const lines = t.split('\\n').map(s=>s.trim()).filter(Boolean);
            const q = lines.filter(s => /quota|โควต|message|ข้อความ|limit|expire|หมดอาย/i.test(s)).slice(0,25);
            return {quotaLines: q};
        }""")
        print("=== หน้า /campaigns : บรรทัดที่เกี่ยวกับ quota ===")
        for l in info["quotaLines"]: print("  ", l[:160])

        camps = await pg.evaluate("""() => {
            const out = [];
            document.querySelectorAll('a[href*="/campaigns/"]').forEach(a => {
                const m = (a.getAttribute('href')||'').match(/campaigns\\/(\\d+)/);
                if (!m) return;
                const row = a.closest('tr') || a.closest('[class*=card]') || a.parentElement;
                out.push({id: m[1], name: (a.innerText||'').trim().slice(0,50),
                          row: row ? (row.innerText||'').replace(/\\s+/g,' ').slice(0,150) : ''});
            });
            const seen = new Set();
            return out.filter(c => !seen.has(c.id) && seen.add(c.id));
        }""")
        print(f"\n=== แคมเปญที่เห็นในบัญชี ({len(camps)}) ===")
        for c in camps:
            mark = "  <<< ที่เราใช้" if c["id"] in ("93082","104883") else ""
            print(f"  [{c['id']}] {c['name']}{mark}")
            if c["row"]: print(f"        {c['row'][:140]}")

        # 2) หน้า message ของแคมเปญ วันที่ระบุ
        d = DAY.replace(" ", "+")
        await pg.goto(f"https://zocialeye.wisesight.com/campaigns/{CAMP}/all/message?start={d}&end={d}&action=filter",
                      wait_until="domcontentloaded")
        await pg.wait_for_timeout(8000)
        res = await pg.evaluate("""() => {
            const badge = document.querySelector('.nav-tabs .active .badge, [class*="tab-active"] .badge');
            const modals = [...document.querySelectorAll('.modal')]
                .filter(m => getComputedStyle(m).display !== 'none')
                .map(m => ({id: m.id, text: (m.innerText||'').replace(/\\s+/g,' ').slice(0,220)}));
            const body = document.body.innerText;
            const warn = body.split('\\n').map(s=>s.trim())
                .filter(s => /quota|โควต|exceed|เกิน|limit|expire|หมดอาย|no data|ไม่พบ/i.test(s)).slice(0,15);
            return {badge: badge ? badge.innerText.trim() : 'ไม่เจอ badge', modals, warn};
        }""")
        print(f"\n=== แคมเปญ {CAMP} วันที่ {DAY} ===")
        print("  badge จำนวนข้อความ:", res["badge"])
        print("  modal ที่เปิดอยู่:", res["modals"] or "ไม่มี")
        print("  ข้อความเตือนในหน้า:")
        for w in res["warn"]: print("   -", w[:160])
        await b.close()

asyncio.run(main())
