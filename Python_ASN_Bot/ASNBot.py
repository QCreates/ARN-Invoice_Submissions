import asyncio
import warnings
from playwright.async_api import async_playwright
import json
import requests

async def connect_browser():
    """Connect to an already running Chrome instance via CDP."""
    ws_url = "http://localhost:9222/json/version"

    try:
        response = requests.get(ws_url)
        response.raise_for_status()
        data = response.json()

        print(f"\n✅ Chrome DevTools Protocol detected at: {ws_url}")
        print(f"🔗 WebSocket Debugger URL: {data['webSocketDebuggerUrl']}")

        playwright = await async_playwright().start()
        browser = await playwright.chromium.connect_over_cdp(data["webSocketDebuggerUrl"])

        if not browser.contexts:
            print("⚠️ No browser contexts found! Creating a new one...")
            context = await browser.new_context()
        else:
            context = browser.contexts[0]

        target_page = None
        for page in context.pages:
            if "chrome://" not in page.url and "chatgpt.com" not in page.url:
                target_page = page
                break

        if not target_page:
            print("⚠️ No suitable open page found! Creating a new one...")
            target_page = await context.new_page()
            await target_page.goto("https://www.google.com")
            print("✅ New page created!")

        await target_page.bring_to_front()
        print(f"✅ Connected successfully to: {target_page.url}")

        # Confirm Playwright can interact with the page
        try:
            await target_page.evaluate("document.body.click()")
            print("✅ Permission check passed: Playwright can interact with the page.")
        except Exception as e:
            print(f"❌ Permission Error: {e}")

        return playwright, browser, target_page

    except requests.exceptions.RequestException as e:
        print(f"❌ Error connecting to Chrome DevTools: {e}")
        return None, None, None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return None, None, None

async def main():
    """Runs the script and ensures proper cleanup."""
    warnings.filterwarnings("ignore", category=ResourceWarning)  # Suppress asyncio resource warnings

    playwright, browser, page = await connect_browser()
    if not page:
        print("❌ No valid page found. Exiting.")
        return

    print("✅ Playwright is running. Press CTRL+C to stop.")

    try:
        while True:
            await asyncio.sleep(10)  # Keep script running
    except KeyboardInterrupt:
        print("\n🛑 Shutting down gracefully...")
        if browser:
            await browser.close()
        if playwright:
            await playwright.stop()
        print("✅ Playwright closed successfully.")

asyncio.run(main())
