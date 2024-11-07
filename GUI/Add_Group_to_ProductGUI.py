




import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QFrame, QCheckBox, QProgressBar
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


class AutomationWorker(QThread):
    log_update = pyqtSignal(str)
    progress_update = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, username, password, product_ids, group_name, headless):
        super().__init__()
        self.username = username
        self.password = password
        self.product_ids = product_ids  # List of product IDs
        self.group_name = group_name
        self.headless = headless

    def run(self):
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            page = browser.new_page()

            try:
                self.login(page)
                for product_id in self.product_ids:  # Iterate over each product ID
                    self.add_product_to_group(page, product_id)
            except Exception as e:
                self.log_update.emit(f"An error occurred: {str(e)}")
            finally:
                browser.close()
                self.finished.emit()

    def login(self, page):
        self.log_update.emit("Attempting to log in...")
        self.progress_update.emit(10)
        page.goto("https://www.restoconcept.com/admin/logon.asp")
        page.fill("#adminuser", self.username)
        page.fill("#adminPass", self.password)
        page.click("#btn1")
        try:
            page.wait_for_selector(
                'td[align="center"][style="background-color:#eeeeee"]:has-text("© Copyright 2024 - Restoconcept")',
                timeout=3000
            )
            self.log_update.emit("Login successful.")
            self.progress_update.emit(40)
        except PlaywrightTimeoutError:
            raise Exception("Login failed. Please check your username and password.")

    def add_product_to_group(self, page, product_id):
        self.log_update.emit(f"Navigating to product page for ID: {product_id}")
        self.progress_update.emit(60)
        page.goto(f"https://www.restoconcept.com/admin/SA_prod_edit.asp?action=edit&recid={product_id}")

        self.log_update.emit(f"Checking for group: {self.group_name}")
        self.progress_update.emit(70)

        # Check if the option exists
        option_exists = page.evaluate(f"""
        () => {{
            const select = document.querySelector('select#idOptionGroup');
            if (!select) return false;
            return Array.from(select.options).some(option => option.text.includes('{self.group_name}'));
        }}
        """)

        if not option_exists:
            self.log_update.emit(f"Error: Group '{self.group_name}' not found in the dropdown for product ID {product_id}.")
            self.progress_update.emit(100)
            return

        self.log_update.emit(f"Selecting group: {self.group_name} for product ID {product_id}")
        self.progress_update.emit(80)
        page.select_option("select#idOptionGroup", label=self.group_name)

        self.log_update.emit(f"Clicking 'Add' button for product ID {product_id}")
        page.click("button[type='submit'][style='font-family:arial; font-size:14px; cursor:pointer; background-color:#005c99; color:#fff; border:0; border-radius:3px; padding:3px 14px;']:has-text('Ajouter')")

        self.log_update.emit(f"Added product {product_id} to group {self.group_name}")
        self.progress_update.emit(100)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Product Group Automation")
        self.setGeometry(100, 100, 600, 600)

        self.setup_ui()
        self.apply_styles()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # Title
        self.title_label = QLabel("Product Group Automation")
        self.title_label.setObjectName("titleLabel")
        main_layout.addWidget(self.title_label)

        # Frame for input fields
        input_frame = QFrame()
        input_layout = QVBoxLayout()
        input_frame.setLayout(input_layout)
        main_layout.addWidget(input_frame)

        # Input fields
        self.username_input = self.create_input_field("Username ", input_layout)
        self.password_input = self.create_input_field("Password ", input_layout, is_password=True)
        self.group_name_input = self.create_input_field("Group Name ", input_layout)

        # Products ID input
        self.products_id_input = QLineEdit()
        self.products_id_input.setPlaceholderText("Enter product IDs separated by commas")
        input_layout.addWidget(QLabel("Product IDs "))
        input_layout.addWidget(self.products_id_input)

        # Headless mode checkbox
        self.headless_checkbox = QCheckBox("Run in headless mode")
        self.headless_checkbox.setChecked(True)
        input_layout.addWidget(self.headless_checkbox)

        # Start button
        self.start_button = QPushButton("Start Automation")
        self.start_button.clicked.connect(self.start_automation)
        input_layout.addWidget(self.start_button)

        # Progress bar
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)

        # Log output
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        main_layout.addWidget(self.log_output)

    def create_input_field(self, label_text, layout, is_password=False):
        label = QLabel(label_text)
        input_field = QLineEdit()
        if is_password:
            input_field.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(label)
        layout.addWidget(input_field)
        return input_field

    def get_product_ids(self):
        # Split the input by commas and strip whitespace from each ID
        return [id.strip() for id in self.products_id_input.text().split(",") if id.strip()]

    def start_automation(self):
        # Collect the necessary parameters for the automation
        username = self.username_input.text()
        password = self.password_input.text()
        product_ids = self.get_product_ids()
        group_name = self.group_name_input.text()
        headless = self.headless_checkbox.isChecked()

        # Create the AutomationWorker thread with the collected parameters
        self.automation_worker = AutomationWorker(username, password, product_ids, group_name, headless)

        # Connect signals for logging, progress updates, and when the process finishes
        self.automation_worker.log_update.connect(self.log_message)
        self.automation_worker.progress_update.connect(self.update_progress_bar)
        self.automation_worker.finished.connect(self.on_automation_finished)

        # Start the automation thread
        self.automation_worker.start()

    def log_message(self, message):
        # Method to update the log display
        self.log_output.append(message)

    def update_progress_bar(self, value):
        # Method to update the progress bar
        self.progress_bar.setValue(value)

    def on_automation_finished(self):
        # Method to handle cleanup when automation finishes
        self.log_message("Automation completed.")
        self.progress_bar.setValue(100)

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f6f6f6;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
                color: #333333;
            }
            QFrame {
                background-color: #ffffff;
                border-radius: 10px;
                padding: 20px;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            }
            QLabel {
                font-size: 13px;
                line-height: 19px;
                color: #111;
                font-family: "Amazon Ember", Arial, sans-serif;
                padding-left: 2px;
                padding-bottom: 2px;
                font-weight: 500;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: bold;
                transition: background-color 0.3s ease;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 8px;
                padding: 8px;
            }
            QLineEdit:focus, QTextEdit:focus {
                border: 1px solid #3498db;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 8px;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 8px;
            }
            QTextEdit {
                background-color: #f6f6f6;
            }
        """)
        self.title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #3498db; margin-bottom: 20px;")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
