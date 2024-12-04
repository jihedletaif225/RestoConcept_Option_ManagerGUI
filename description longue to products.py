

from playwright.async_api import async_playwright, Page
import logging
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import asyncio

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class RestoconceptAdmin:
    def __init__(self, username: str, password: str, excel_file: str):
        self.username = username
        self.password = password
        self.excel_file = excel_file

    async def login(self, page: Page) -> None:
        """
        Log in to the Restoconcept admin panel with robust error handling.
        """
        try:
            # Navigate to login page
            await page.goto("https://www.restoconcept.com/admin/logon.asp", wait_until="networkidle")
            await page.fill("#adminuser", self.username)
            await page.fill("#adminPass", self.password)
            
            # Click login and wait for navigation
            async with page.expect_navigation(wait_until="networkidle"):
                await page.click("#btn1")
            
            # Verify login success
            success_selectors = [
                'td[align="center"][style="background-color:#eeeeee"]:has-text("© Copyright 2024 - Restoconcept")',
                'a:has-text("Déconnexion")'
            ]
            for selector in success_selectors:
                if await page.is_visible(selector):
                    logger.info("Login successful")
                    return
            
            raise Exception("Login verification failed")
        
        except Exception as e:
            logger.error(f"Login failed: {str(e)}")
            raise

    async def edit_product(self, page: Page, product_id: str, description: str) -> None:
        """
        Edit a product by updating its description.
        """
        try:
            # Navigate to product edit page
            url = f"https://www.restoconcept.com/admin/SA_prod_edit.asp?action=edit&recid={product_id}"
            await page.goto(url, wait_until="networkidle")

            # Wait for iframe to load and switch context to iframe
            iframe_element = await page.query_selector('iframe#idContentoEdit2')
            iframe = await iframe_element.content_frame()

            # Delete the content in the body
            await iframe.fill('body[contenteditable="true"]', "")

            # Replace with new text
            await iframe.fill('body[contenteditable="true"]', description)

            # Click on the "Mettre à jour" button
            await page.click('button[style="font-family:arial; font-size:15px; cursor:pointer; background-color:#005c99; color:#fff; border:0; border-radius:3px; padding:3px 14px;"]')

            logger.info(f"Product {product_id} updated successfully.")
        
        except Exception as e:
            logger.error(f"Error during product edit for ID {product_id}: {str(e)}")
            raise

    async def run(self) -> None:
        # Read Excel file to get product IDs and descriptions
        data = pd.read_excel(self.excel_file)
        if "Product ID" not in data.columns or "SEO-Optimized Description" not in data.columns:
            logger.error("Excel file must contain 'Product ID' and 'Description' columns.")
            return

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)  # Change to True to run headless
            page = await browser.new_page()

            # Login to the admin panel
            await self.login(page)

            # Iterate over products and edit each
            for _, row in data.iterrows():
                product_id = str(row["Product ID"])
                description = str(row["SEO-Optimized Description"])
                logger.info(f"Updating product ID {product_id} with description: {description}")
                await self.edit_product(page, product_id, description)

            # Close the browser after the task
            await browser.close()

def select_excel_file() -> str:
    """
    Open a file dialog to select an Excel file.
    """
    Tk().withdraw()  # Hide the root Tk window
    file_path = askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")],
    )
    if not file_path:
        raise Exception("No file selected.")
    return file_path

# Usage
if __name__ == "__main__":
    excel_file_path = select_excel_file()
    admin = RestoconceptAdmin(
        username="",
        password="",
        excel_file=excel_file_path
    )
    asyncio.run(admin.run())
