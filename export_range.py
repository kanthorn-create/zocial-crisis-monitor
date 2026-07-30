"""One-off: สั่ง ZE export ช่วงวันที่กำหนดเอง (ส่งไฟล์เข้าอีเมล EXPORT_EMAIL)
ใช้ exportData() โดยตรงเหมือน monitor.py — env: CAMPAIGN, START_DATE, END_DATE (เช่น "11 Jul 2026")
"""
import asyncio, os
from playwright.async_api import async_playwright

ZOCIAL_ID    = os.environ.get("ZOCIAL_ID", "Nativejump01")
ZOCIAL_PASS  = os.environ.get("ZOCIAL_PASS", "Nativejump123")
CAMPAIGN     = os.environ.get("CAMPAIGN", "104883")
START_DATE   = os.environ.get("START_DATE", "11 Jul 2026")
END_DATE     = os.environ.get("END_DATE", "29 Jul 2026")
EXPORT_EMAIL = os.environ.get("EXPORT_EMAIL", "kanthorn@nativejump.co")


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        for attempt in range(3):
            try:
                await page.goto("https://zocialeye.wisesight.com/login", wait_until="domcontentloaded")
                await page.wait_for_selector("input[name='username']", timeout=20000)
                await page.fill("input[name='username']", ZOCIAL_ID)
                await page.fill("input[name='passwd']", ZOCIAL_PASS)
                await page.click("#btn-login")
                await page.wait_for_function("() => !location.pathname.includes('/login')", timeout=30000)
                break
            except Exception as e:
                print(f"login attempt {attempt+1}: {e}")
                await page.wait_for_timeout(3000)
        else:
            raise RuntimeError("login failed")

        url = (f"https://zocialeye.wisesight.com/campaigns/{CAMPAIGN}/all/message"
               f"?start={START_DATE.replace(' ', '+')}&end={END_DATE.replace(' ', '+')}&action=filter")
        await page.goto(url, wait_until="domcontentloaded")
        await page.wait_for_timeout(6000)

        has_fn = await page.evaluate("() => typeof exportData === 'function'")
        if not has_fn:
            raise RuntimeError("exportData not found")

        async with page.expect_response(lambda r: "exportdata" in r.url.lower(), timeout=30000) as ri:
            await page.evaluate("([ch, em]) => exportData(ch, em)", ["all", EXPORT_EMAIL])
        resp = await ri.value
        body = (await resp.text())[:300]
        print(f"exportData: HTTP {resp.status} | {body}")
        if resp.status != 200 or '"success":false' in body.replace(" ", ""):
            raise RuntimeError(f"export failed: {body}")
        print(f"OK — export {CAMPAIGN} {START_DATE} → {END_DATE} ส่งเข้า {EXPORT_EMAIL}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
