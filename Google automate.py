
import random
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from playwright.sync_api import sync_playwright

def random_delay():
    delay = random.randint(1, 3)
    time.sleep(delay)
    return delay

# Define country-specific settings
COUNTRIES = [
    {"name": "Germany", "geolocation": {"latitude": 52.5200, "longitude": 13.4050}, "locale": "de-DE", "proxy": None},
]

def search_and_click(keyword, country):
    with sync_playwright() as p:
        # Configure browser with geolocation, locale, and optional proxy
        browser = p.chromium.launch(headless=False)
        context_args = {
            "geolocation": country["geolocation"],
            "locale": country["locale"],
            "permissions": ["geolocation"]
        }
        if country["proxy"]:
            context_args["proxy"] = {"server": country["proxy"]}

        context = browser.new_context(**context_args)
        page = context.new_page()
        
        # Navigate to Google
        page.goto("https://www.google.com/")
        page.wait_for_load_state("domcontentloaded")
        
        # Accept cookies if prompted
        try:
            page.locator("text=Accept").click()
        except Exception:
            print("No cookies banner found, continuing.")
        
        # Search for the keyword
        search_box = page.locator("textarea.gLFyf")
        search_box.fill(keyword)
        search_box.press("Enter")
        
        # Wait for results to load
        page.wait_for_selector("h3")
        links = page.query_selector_all('a:has(h3)')
        
        # Open first 5 links
        # for i, link in enumerate(links):
        for i, link in enumerate(links[:2]):
            href = link.get_attribute('href')
            if href:
                print(f"Opening link {i + 1} in {country['name']}: {href}")
                new_tab = context.new_page()
                new_tab.goto(href)
                random_delay()
                new_tab.close()
            else:
                print(f"Link {i + 1} in {country['name']} has no href attribute.")
        
        # Close context and browser
        context.close()
        browser.close()

def select_file():
    # Open file dialog to select the Excel file
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

if __name__ == "__main__":
    # Allow user to select the Excel file
    excel_file = select_file()
    
    if excel_file:
        # Read the keywords from the selected Excel file
        df = pd.read_excel(excel_file)
        
        # Assuming the Excel file has a column named 'Keywords'
        keywords = df['Keywords'].dropna().tolist()
        # keywords = df['A'].dropna().tolist()

        
        for keyword in keywords:
            for country in COUNTRIES:
                print(f"Searching for '{keyword}' from {country['name']}...")
                search_and_click(keyword, country)
    else:
        print("No file selected. Exiting.")
