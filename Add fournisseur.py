
import asyncio
import logging
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from playwright.async_api import async_playwright, Page
from typing import List, Dict

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class RestoconceptAdmin:
    def __init__(self, username: str, password: str, excel_file: str):
        """
        Initialize admin tool with credentials and Excel file path.
        
        :param username: Admin username
        :param password: Admin password
        :param excel_file: Path to Excel file with marque and fournisseur data
        """
        self.username = username
        self.password = password
        self.excel_file = excel_file
        self.process_data = self._load_excel_data()

    def _load_excel_data(self) -> List[Dict[str, str]]:
        """
        Load processing data from Excel file.
        
        :return: List of dictionaries with processing information
        """
        try:
            df = pd.read_excel(self.excel_file)
            # Ensure required columns exist
            required_columns = ['marque', 'fournisseur']
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Missing required column: {col}")
            
            # Convert to list of dictionaries, dropping rows with NaN
            return df.dropna(subset=required_columns).to_dict('records')
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}")
            return []

    async def login(self, page: Page) -> None:
        """
        Log in to the Restoconcept admin panel with robust error handling.
        
        :param page: Playwright Page object
        :raises Exception: If login fails
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

    async def process_marque(self, page: Page, marque: str) -> List[str]:
        """
        Extract edit links for products from a specific supplier.
        
        :param page: Playwright Page object
        :param marque: Supplier/Brand to process
        :return: List of product edit URLs
        """
        await page.goto("https://www.restoconcept.com/admin/SA_prod.asp", wait_until="networkidle")
        
        # Wait and select supplier
        await page.wait_for_selector('select[name="marque"]')
        await page.select_option('select[name="marque"]', marque)
        await page.click('button:has-text("Rechercher")')
        
        # Wait for results
        await page.wait_for_load_state("networkidle")

        all_edit_links = []
        base_url = "https://www.restoconcept.com/admin/"

        while True:
            # Extract edit links
            links = await page.locator('a:has-text("Editer")').all()
            current_page_links = [base_url + await link.get_attribute('href') for link in links]
            all_edit_links.extend(current_page_links)

            # Check for next page
            next_links = await page.locator('a:has-text("Suiv.")').all()
            if not next_links:
                break

            try:
                await next_links[0].click()
                await page.wait_for_load_state("networkidle")
            except Exception as e:
                logger.error(f"Error navigating to next page: {e}")
                break

        logger.info(f"Total product links found for {marque}: {len(all_edit_links)}")
        return all_edit_links

    async def process_produit(self, page: Page, url: str, fournisseur: str) -> None:
        """
        Process and update individual product details.
        
        :param page: Playwright Page object
        :param url: Product edit page URL
        :param fournisseur: Supplier ID to set
        """
        try:
            await page.goto(url, wait_until="networkidle")

            # Skip occasion products
            selected_option = await page.locator('select[name="photoplus"] option:checked').get_attribute("value")
            if selected_option == "occasion.jpg":
                logger.info(f"Skipping 'Occasion' product: {url}")
                return

            # Select fournisseur (supplier)
            await page.wait_for_selector('select[name="idf1"]')
            await page.select_option('select[name="idf1"]', fournisseur)

            # Update product details
            update_buttons = page.locator(
                'form:has(div:has-text("Fournisseurs")) button:has-text("Mettre à jour"), '
                'form:has(div:has-text("Fournisseurs")) button:has-text("Ajouter le fournisseur")'
            )

            if await update_buttons.is_visible():
                await update_buttons.click()
                await page.wait_for_load_state("networkidle")
                logger.info(f"Successfully processed product: {url}")
            else:
                logger.warning(f"No update button found for product: {url}")

        except Exception as e:
            logger.error(f"Error processing product {url}: {str(e)}")
            with open("failed_products.txt", "a") as failed_file:
                failed_file.write(f"{url}\n")


    async def run(self):
        """
        Main script execution with comprehensive error handling.
        """
        if not self.process_data:
            logger.error("No data to process from Excel file")
            return

        async with async_playwright() as p:
            try:
                browser = await p.chromium.launch(
                    headless=True,
                    args=['--no-sandbox', '--disable-setuid-sandbox']
                )
                context = await browser.new_context()
                page = await context.new_page()

                # Execute main workflow
                await self.login(page)
                
                # Process each marque and fournisseur combination
                for entry in self.process_data:
                    try:
                        marque = str(entry['marque'])
                        fournisseur = str(entry['fournisseur'])
                        
                        logger.info(f"Processing Marque: {marque}, Fournisseur: {fournisseur}")
                        
                        edit_links = await self.process_marque(page, marque)
                        
                        for link in edit_links:
                            try:
                                await self.process_produit(page, link, fournisseur)
                            except Exception as product_error:
                                logger.error(f"Skipping product due to error: {link}")

                    except Exception as marque_error:
                        logger.error(f"Error processing marque {marque}: {marque_error}")

            except Exception as e:
                logger.critical(f"Critical script error: {str(e)}")
            finally:
                await browser.close()

def browse_excel_file():
    """
    Open file dialog to browse for Excel file.
    
    :return: Path to selected Excel file
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
    )
    return file_path

def main():
    """
    Entry point for script execution.
    """
    # Prompt for Excel file
    excel_file = browse_excel_file()
    if not excel_file:
        print("No file selected. Exiting.")
        return

    # Credentials
    USERNAME = ""
    PASSWORD = ""
    
    # Create and run admin tool
    admin_tool = RestoconceptAdmin(USERNAME, PASSWORD, excel_file)
    asyncio.run(admin_tool.run())

if __name__ == "__main__":
    main()
