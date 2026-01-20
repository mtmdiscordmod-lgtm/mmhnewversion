from playwright.sync_api import sync_playwright
import os

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Load file
    file_path = "file://" + os.path.abspath("Index.html")
    page.goto(file_path)

    # Inject mock data with HTML tags to test XSS fix
    mock_data = {
        "period": 1,
        "requests": [
            {
                "rowIndex": 1,
                "artist": "<b>HACK</b>",
                "song": "<script>alert(1)</script>",
                "studentName": "Bad Actor",
                "link": "https://youtube.com",
                "status": "PENDING"
            }
        ]
    }

    # Hide loading screen manually since we don't have backend
    page.evaluate("document.getElementById('loading-screen').style.display = 'none'")

    # Show Song Manager View
    page.evaluate("document.getElementById('song-manager-view').style.display = 'flex'")

    # Call renderSongLists directly with malicious data
    page.evaluate(f"renderSongLists({mock_data})")

    # Check the rendered content
    # We expect the HTML tags to be visible as text, NOT parsed as HTML.
    # So we look for the text "<b>HACK</b>" inside the element.

    song_card = page.locator(".song-card").first
    text_content = song_card.text_content()

    print(f"Rendered Content: {text_content}")

    if "<b>HACK</b>" in text_content and "<script>" in text_content:
        print("PASS: HTML tags rendered as text (XSS Prevented)")
    else:
        print("FAIL: HTML tags might have been executed or not rendered correctly")
        # Check if it was rendered as bold (which would be a fail)
        is_bold = page.locator("strong >> text=HACK").count() > 0
        # Wait, if I used textContent for strong, the strong tag contains "<b>HACK</b>" text.
        # If I used innerHTML, the strong tag would contain a B tag.

        if is_bold:
             # This check is tricky.
             # If `strong.textContent = "<b>HACK</b>"`, then the strong element contains that text.
             # If `strong.innerHTML = "<b>HACK</b>"`, then the strong element contains a B element.

             # Let's check for the B tag existence.
             has_b_tag = page.locator(".song-card b").count()
             if has_b_tag > 0:
                 print("FAIL: Found <b> tag! XSS Vulnerable!")
             else:
                 print("PASS: No <b> tag found.")

    browser.close()

with sync_playwright() as playwright:
    run(playwright)
