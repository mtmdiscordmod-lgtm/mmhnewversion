from playwright.sync_api import sync_playwright
import os

def run_verification():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # 1. Verify Teacher Dashboard
        print("Verifying Teacher Dashboard...")
        cwd = os.getcwd()
        page.goto(f"file://{cwd}/Index.html")

        # Inject Mock
        page.add_script_tag(path=f"{cwd}/verification/mock_google.js")

        # Trigger init manually since window.onload fires before injection sometimes
        page.evaluate("window.onload()")

        # Click Song Manager
        page.get_by_text("🎵 SONG MANAGER").click()

        # Just manually call loadSongs to force refresh
        page.evaluate("loadSongs()")

        # Wait for card
        page.get_by_text("Mock Song").wait_for(timeout=5000)

        # Screenshot
        page.screenshot(path=f"{cwd}/verification/teacher_dashboard.png")
        print("Teacher Dashboard Screenshot saved.")

        # 2. Verify Student Portal
        print("Verifying Student Portal...")
        page.goto(f"file://{cwd}/studentrequest.html")

        # Inject Mock
        page.add_script_tag(path=f"{cwd}/verification/mock_google.js")

        # Trigger init
        page.evaluate("window.onload()")

        # Wait for form title specifically (avoiding strict mode violation with body text)
        # Using class selector to be specific
        page.locator(".title").get_by_text("SONG REQUEST").wait_for()
        page.get_by_placeholder("e.g. Daft Punk").fill("Test Artist")

        # Screenshot
        page.screenshot(path=f"{cwd}/verification/student_portal.png")
        print("Student Portal Screenshot saved.")

        browser.close()

if __name__ == "__main__":
    run_verification()
