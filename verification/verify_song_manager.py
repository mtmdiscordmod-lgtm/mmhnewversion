from playwright.sync_api import sync_playwright
import os

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Load Index.html directly as file
    file_path = "file://" + os.path.abspath("Index.html")
    print(f"Loading {file_path}")
    page.goto(file_path)

    # 1. Verify Song Manager Nav Button
    nav_btn = page.locator("text=🎵 SONG MANAGER")
    nav_btn.wait_for()
    print("Nav Button Found")

    # 2. Click it
    nav_btn.click()

    # 3. Verify View Visibility
    view = page.locator("#song-manager-view")
    if view.is_visible():
        print("Song Manager View Visible")
    else:
        print("ERROR: Song Manager View NOT Visible")

    # 4. Verify Theme Pink (Check body class or bg layer)
    # The script adds 'theme-pink' to #bg-layer
    bg_layer = page.locator("#bg-layer")
    classes = bg_layer.get_attribute("class")
    if "theme-pink" in classes:
        print("Theme Pink Applied")
    else:
        print(f"ERROR: Theme Pink Missing. Classes: {classes}")

    # 5. Verify 2-Column Layout
    inbox = page.locator("#song-inbox")
    playlist = page.locator("#song-playlist")

    if inbox.is_visible() and playlist.is_visible():
        print("Inbox and Playlist Columns Found")

    # Screenshot
    page.screenshot(path="verification/song_manager_ui.png")
    print("Screenshot saved to verification/song_manager_ui.png")

    browser.close()

with sync_playwright() as playwright:
    run(playwright)
