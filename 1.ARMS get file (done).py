import re
import tkinter as tk
from tkinter import simpledialog
from playwright.sync_api import sync_playwright
import playwright

def sanitize_filename(filename: str) -> str:
    """
    Optional helper to remove characters invalid on Windows, macOS, etc.
    If your link texts contain ':' or other special characters,
    uncomment and use this to safely name files.
    """
    # return re.sub(r'[\\/*?:"<>|]', '-', filename)
    return filename  # Currently a no-op unless you uncomment the above.

def run(playwright):
    # 1) Launch the browser in non-headless mode and accept downloads.
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # 2) Navigate to the ARMS exported reports page (login screen).
    page.goto("https://arms3.onezero.com/view-exported-reports")

    # 3) Login credentials.
    username = "tongseng.wong@lirunex.com"
    password = "123Abc$$"

    # 4) Fill in the login form.
    page.get_by_role("textbox", name="Email").fill(username)
    page.get_by_role("textbox", name="Password").fill(password)
    
    # 5) Click the login button.
    page.get_by_role("button", name="Login").click()

    # 6) Pause to allow manual inspection in DevTools (resume by clicking the ▶ button in DevTools).
    page.pause()

    # 7) Click on "Reports".
    page.get_by_role("link", name="Reports", exact=True).click()

    # 8) Click on "Exported Reports" under #topNav.
    page.locator("#topNav").get_by_role("link", name="Exported Reports").click()

    # 9) Display a dialog for the user to enter a date.
    root = tk.Tk()
    root.withdraw()  # Hide the main window.
    date_str = simpledialog.askstring(
        title="Download Reports",
        prompt="Enter date (YYYY-MM-DD) to download reports:"
    )
    root.destroy()  # Close the temporary Tk window.

    # If the user cancels or doesn't enter a date, exit the script.
    if not date_str:
        print("❌ No date entered, script terminated.")
        browser.close()
        return

    # Portfolio list to match in the link text.
    portfolios = [
        "TL Martingale Portfolio",
        "US Martingale Portfolio (Server 1)",
        "US Martingale Portfolio (Server 2)",
        "369039422 Martingale Portfolio",
        "Portugal Martingale Portfolio",
        "SG Martingale Portfolio",
        "Ahmad Martingale Portfolio",
        "SG2 Martingale Portfolio",
        "Vietnam Martingale",
        "Crew Hedging",
        "369082364 Martingale",
        "Iran 2 Martingale",
        "B book exclude Autohedge",
        "369034123 Martingale Portfolio",
        "Rockgreen Martingale (Server 1)",
        "Rockgreen Martingale (Server 2)",
        "369102488 Martingale (server 2)",
        "369103139 Martingale (server 2)"
    ]

    weekly_meeting_tag = "(Weekly meeting 7 day)"

    # 10) Find and download files matching all conditions:
    #    - date_str in the link text
    #    - (Weekly meeting 7 day) in the link text
    #    - at least one of the portfolio names in the link text
    all_links = page.get_by_role("link").all()
    found_links = False
    for link in all_links:
        link_text = link.inner_text()

        # Condition 1: date_str is in the link text
        # Condition 2: weekly_meeting_tag is in the link text
        # Condition 3: any of the portfolio names is in the link text
        if (
            date_str in link_text
            and weekly_meeting_tag in link_text
            or any(portfolio in link_text for portfolio in portfolios)
        ):
            found_links = True
            print(f"Found link: {link_text}")
            with page.expect_download() as download_info:
                link.click()
            download = download_info.value

            # Optionally sanitize the filename to remove invalid characters.
            safe_name = sanitize_filename(link_text)
            download.save_as(safe_name)
            print(f"Downloaded: {safe_name}")
    
    if not found_links:
        print(f"No matching links found for date {date_str}")

    # 11) Wait for user input before closing.
    input("Press ENTER to exit and close the browser...")
    browser.close()

with sync_playwright() as pw:
    run(pw)
