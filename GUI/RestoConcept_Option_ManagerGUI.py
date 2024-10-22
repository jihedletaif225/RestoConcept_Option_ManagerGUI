


import sys
import time
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                             QLineEdit, QTextEdit, QProgressBar, QMessageBox, QGridLayout, QFrame,
                             QListWidget, QListWidgetItem, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon

class PlaywrightWorker(QThread):
    progress_update = pyqtSignal(int)
    status_update = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, username, password, group_name, options, headless):
        super().__init__()
        self.username = username
        self.password = password
        self.group_name = group_name
        self.options = options
        self.headless = headless

    def run(self):
        with sync_playwright() as p:
            try:
                browser = p.chromium.launch(headless=self.headless)
                context = browser.new_context()
                page = context.new_page()

                if not self.login(page):
                    return

                if not self.navigate_to_option_group(page, self.group_name):
                    return

                total_options = len(self.options)
                for i, option_name in enumerate(self.options, 1):
                    if not self.add_option_to_group(page, option_name):
                        continue
                    progress = int((i / total_options) * 100)
                    self.progress_update.emit(progress)
                    self.status_update.emit(f"Added option: {option_name}")
                    time.sleep(1)

                self.status_update.emit("Process completed successfully.")
            except Exception as e:
                self.error_occurred.emit(f"An unexpected error occurred: {str(e)}")
            finally:
                if 'context' in locals():
                    context.close()
                if 'browser' in locals():
                    browser.close()

    def login(self, page):
        try:
            self.status_update.emit("Logging in...")
            page.goto("https://www.restoconcept.com/admin/logon.asp")
            page.fill("#adminuser", self.username)
            page.fill("#adminPass", self.password)
            page.click('input[type="image"][src="logon_cnx.jpg"]')
            
            try:
                page.wait_for_selector('td[align="center"][style="background-color:#eeeeee"]:has-text("© Copyright 2024 - Restoconcept")', timeout=5000)
                self.status_update.emit("Login successful.")
                return True
            except PlaywrightTimeoutError:
                self.error_occurred.emit("Login failed. Please check your username and password.")
                return False
        except Exception as e:
            self.error_occurred.emit(f"Login error: {str(e)}")
            return False

    def navigate_to_option_group(self, page, group_name):
        try:
            self.status_update.emit(f"Navigating to option group: {group_name}")
            page.goto("https://www.restoconcept.com/admin/options/optionsgroupslist.asp")
            page.fill("#psearch", group_name)
            page.click('input[type="image"][name="Submit"]')
            page.wait_for_load_state("networkidle")
            
            if page.locator('img[alt=" Ajouter/retirer des options "]').count() == 0:
                self.error_occurred.emit(f"Option group '{group_name}' not found. Please check the group name.")
                return False
            
            page.click('img[alt=" Ajouter/retirer des options "]')
            page.wait_for_load_state("networkidle")
            return True
        except Exception as e:
            self.error_occurred.emit(f"Error navigating to option group: {str(e)}")
            return False

    def add_option_to_group(self, page, option_name):
        try:
            self.status_update.emit(f"Adding option: {option_name}")
            page.fill('input[name="rch"]', option_name)
            page.click('input[type="image"][src="/admin/rechercher.gif"]')
            page.wait_for_load_state("networkidle")
            
            checkbox = page.locator('input[type="checkbox"][name="inclure0"]')
            if checkbox.is_visible():
                checkbox.check()
                page.click('input[type="image"][src="maj.gif"]')
                page.wait_for_load_state("networkidle")
                return True
            else:
                self.error_occurred.emit(f"Option '{option_name}' not found. Skipping this option.")
                return False
        except Exception as e:
            self.error_occurred.emit(f"Error adding option '{option_name}': {str(e)}")
            return False

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('RestoConcept Option Manager')
        self.setGeometry(100, 100, 800, 600)

        main_layout = QHBoxLayout()

        # Left panel for inputs
        left_panel = QFrame()
        left_layout = QVBoxLayout()

        title_label = QLabel('RestoConcept Option Manager')
        title_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(title_label)

        input_grid = QGridLayout()
        input_grid.addWidget(QLabel('Username:'), 0, 0)
        self.username_input = QLineEdit()
        input_grid.addWidget(self.username_input, 0, 1)
        input_grid.addWidget(QLabel('Password:'), 1, 0)
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        input_grid.addWidget(self.password_input, 1, 1)
        input_grid.addWidget(QLabel('Option Group:'), 2, 0)
        self.group_input = QLineEdit()
        input_grid.addWidget(self.group_input, 2, 1)
        left_layout.addLayout(input_grid)

        options_label = QLabel('Options:')
        left_layout.addWidget(options_label)

        self.options_input = QLineEdit()
        self.options_input.setPlaceholderText("Enter an option and press Enter")
        self.options_input.returnPressed.connect(self.add_option_to_list)
        left_layout.addWidget(self.options_input)

        self.options_list = QListWidget()
        left_layout.addWidget(self.options_list)

        button_layout = QHBoxLayout()
        self.add_button = QPushButton('Add Option')
        self.add_button.clicked.connect(self.add_option_to_list)
        button_layout.addWidget(self.add_button)
        self.remove_button = QPushButton('Remove Selected')
        self.remove_button.clicked.connect(self.remove_selected_option)
        button_layout.addWidget(self.remove_button)
        left_layout.addLayout(button_layout)

        self.headless_checkbox = QCheckBox('Run in headless mode')
        self.headless_checkbox.setChecked(True)
        left_layout.addWidget(self.headless_checkbox)

        self.start_button = QPushButton('Start Process')
        self.start_button.clicked.connect(self.start_process)
        left_layout.addWidget(self.start_button)

        left_panel.setLayout(left_layout)
        main_layout.addWidget(left_panel)

        # Right panel for status and progress
        right_panel = QFrame()
        right_layout = QVBoxLayout()

        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        right_layout.addWidget(self.status_text)

        self.progress_bar = QProgressBar()
        right_layout.addWidget(self.progress_bar)

        right_panel.setLayout(right_layout)
        main_layout.addWidget(right_panel)

        self.setLayout(main_layout)

        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                font-family: Arial, sans-serif;
                font-size: 14px;
                color: #1c1e21;
            }
            QFrame {
                background-color: #ffffff;
                border-radius: 8px;
                padding: 20px;
            }
            QPushButton {
                background-color: #1877f2;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #166fe5;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #dddfe2;
                border-radius: 6px;
                padding: 8px;
            }
            QLabel {
                color: #1c1e21;
                font-weight: bold;
            }
            QProgressBar {
                border: 1px solid #dddfe2;
                border-radius: 6px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #1877f2;
                border-radius: 6px;
            }
            QListWidget {
                border: 1px solid #dddfe2;
                border-radius: 6px;
            }
            QCheckBox {
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        
        title_label = self.findChild(QLabel, '')
        if title_label:
            title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #1877f2; margin-bottom: 20px;")

    def add_option_to_list(self):
        option = self.options_input.text().strip()
        if option:
            self.options_list.addItem(option)
            self.options_input.clear()

    def remove_selected_option(self):
        for item in self.options_list.selectedItems():
            self.options_list.takeItem(self.options_list.row(item))

    def start_process(self):
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        group_name = self.group_input.text().strip()
        options = [self.options_list.item(i).text() for i in range(self.options_list.count())]
        headless = self.headless_checkbox.isChecked()

        if not username or not password or not group_name or not options:
            QMessageBox.warning(self, 'Input Error', 'Please fill in all fields and add at least one option.')
            return

        self.worker = PlaywrightWorker(username, password, group_name, options, headless)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.status_update.connect(self.update_status)
        self.worker.error_occurred.connect(self.show_error)
        self.worker.start()

        self.start_button.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_status(self, status):
        self.status_text.append(status)
        self.status_text.verticalScrollBar().setValue(self.status_text.verticalScrollBar().maximum())
        if status == "Process completed successfully.":
            self.start_button.setEnabled(True)

    def show_error(self, error_message):
        QMessageBox.warning(self, 'Error', error_message)
        self.status_text.append(f"Error: {error_message}")
        self.start_button.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())