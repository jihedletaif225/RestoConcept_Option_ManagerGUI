


import sys
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                             QLineEdit, QFileDialog, QProgressBar, QCheckBox, QFrame, QMessageBox, QTextEdit)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QThread, pyqtSignal

import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QFrame, 
                             QLabel, QPushButton, QLineEdit, QCheckBox, QProgressBar, 
                             QTextEdit, QFileDialog, QMessageBox, QTabWidget, QScrollArea,
                             QSizePolicy)
from PyQt5.QtGui import QIcon, QFont, QColor, QPalette
from PyQt5.QtCore import Qt, QPropertyAnimation, QEasingCurve, QSize

import time


class OptionsUploaderThread(QThread):
    progress_update = pyqtSignal(int)
    status_update = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    log_update = pyqtSignal(str)

    def __init__(self, excel_file, username, password, headless):
        super().__init__()
        self.excel_file = excel_file
        self.username = username
        self.password = password
        self.headless = headless

    def run(self):
        options_df = pd.read_excel(self.excel_file)
        total_rows = len(options_df)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            context = browser.new_context()
            page = context.new_page()

            try:
                self.log_update.emit("Starting the upload process...")
                self.login(page)

                for index, row in options_df.iterrows():
                    self.status_update.emit(f"Processing option {index + 1} of {total_rows}")
                    self.log_update.emit(f"Processing option {index + 1} of {total_rows}")

                    try:
                        self.navigate_to_options_page(page)
                        self.fill_option_form(page, row)
                        self.submit_option(page)
                        self.handle_submission_result(page)
                    except Exception as e:
                        self.log_update.emit(f"Error processing option {index + 1}: {str(e)}")
                        continue

                    progress = int((index + 1) / total_rows * 100)
                    self.progress_update.emit(progress)

            except Exception as e:
                self.error_occurred.emit(f"An error occurred: {str(e)}")
                self.log_update.emit(f"Critical error: {str(e)}")
            finally:
                context.close()
                browser.close()

        self.status_update.emit("Upload process completed.")
        self.log_update.emit("Upload process completed. Check the log for details.")

    def login(self, page):
        self.log_update.emit("Attempting to log in...")
        page.goto("https://www.restoconcept.com/admin/logon.asp")
        page.fill("#adminuser", self.username)
        page.fill("#adminPass", self.password)
        page.click("#btn1")

        try:
            page.wait_for_selector('td[align="center"][style="background-color:#eeeeee"]:has-text("© Copyright 2024 - Restoconcept")', timeout=5000)
            self.log_update.emit("Login successful.")
        except PlaywrightTimeoutError:
            raise Exception("Login failed. Please check your username and password.")

    def navigate_to_options_page(self, page):
        page.goto("https://www.restoconcept.com/admin/options/optionslist.asp")
        page.click('a[href="/admin/SA_opt_edit.asp?action=add"]')

    def fill_option_form(self, page, row):
        page.fill("#optionDescrip", str(row['optionDescrip']))
        page.fill("#ref", str(row['ref']))
        page.fill("#pricetoadd", str(row['pricetoadd']))
        page.fill("#prixpublic", str(row['prixpublic']))
        page.select_option("#iddelai", str(row['iddelai']))

    def submit_option(self, page):
        page.click('button:has-text("Ajouter")')
        page.wait_for_load_state("networkidle")

    def handle_submission_result(self, page):
        if page.query_selector('text="Option déjà créée"'):
            self.log_update.emit("Product already exists. Skipping...")
        elif page.query_selector('text="Session expirée"'):
            self.log_update.emit("Session expired. Attempting to log in again...")
            self.login(page)
        elif page.query_selector('text="Option ajoutée avec succès"'):
            self.log_update.emit("Option added successfully.")
        else:
            self.log_update.emit("Unexpected result after submission. Please check manually.")


class OptionsUploaderGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.apply_styles()
        self.setup_animations()

    def initUI(self):
        self.setWindowTitle('RestoConcept Options Uploader')
        self.setGeometry(300, 300, 800, 600)
        self.setMinimumSize(600, 500)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        self.frame = QFrame()
        self.frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        frame_layout = QVBoxLayout()
        self.frame.setLayout(frame_layout)
        main_layout.addWidget(self.frame)

        self.title_label = QLabel('RestoConcept Options Uploader')
        self.title_label.setObjectName('title_label')
        self.title_label.setAlignment(Qt.AlignCenter)
        frame_layout.addWidget(self.title_label)

        # Create tabs
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName('tab_widget')
        frame_layout.addWidget(self.tab_widget)

        # Upload Tab
        upload_tab = QWidget()
        upload_layout = QVBoxLayout(upload_tab)
        self.tab_widget.addTab(upload_tab, "Upload")

        file_layout = QHBoxLayout()
        self.file_label = QLabel('Select Excel File:')
        file_layout.addWidget(self.file_label)
        self.file_button = QPushButton('Browse')
        self.file_button.setIcon(QIcon.fromTheme("document-open"))
        self.file_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_button)
        upload_layout.addLayout(file_layout)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Enter username")
        upload_layout.addWidget(QLabel('Username'))
        upload_layout.addWidget(self.username_input)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Enter password")
        upload_layout.addWidget(QLabel('Password'))
        upload_layout.addWidget(self.password_input)

        self.headless_checkbox = QCheckBox('Run in headless mode')
        self.headless_checkbox.setChecked(True)
        upload_layout.addWidget(self.headless_checkbox)

        self.upload_button = QPushButton('Upload Options')
        self.upload_button.clicked.connect(self.start_upload)
        upload_layout.addWidget(self.upload_button)

        self.progress_bar = QProgressBar()
        upload_layout.addWidget(self.progress_bar)

        self.status_label = QLabel('')
        self.status_label.setAlignment(Qt.AlignCenter)
        upload_layout.addWidget(self.status_label)

        # Log Tab
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        self.tab_widget.addTab(log_tab, "Log")

        self.log_textarea = QTextEdit()
        self.log_textarea.setReadOnly(True)
        log_layout.addWidget(self.log_textarea)

        self.setLayout(main_layout)

    def apply_styles(self):
        # Set the application style to Fusion for a more modern look
        QApplication.setStyle("Fusion")

        # Create a palette for the application
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.WindowText, QColor(50, 50, 50))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        palette.setColor(QPalette.ToolTipText, QColor(50, 50, 50))
        palette.setColor(QPalette.Text, QColor(50, 50, 50))
        palette.setColor(QPalette.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ButtonText, QColor(50, 50, 50))
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        QApplication.setPalette(palette)

        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 14px;
            }
            QFrame {
                background-color: #ffffff;
                border-radius: 10px;
                border: 1px solid #e0e0e0;
            }
            QLabel {
                color: #333333;
                font-weight: bold;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-weight: bold;
                transition: background-color 0.3s;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2a5d8f;
            }
            QLineEdit {
                border: 2px solid #e0e0e0;
                border-radius: 5px;
                padding: 8px;
                transition: border-color 0.3s;
            }
            QLineEdit:focus {
                border-color: #4a90e2;
            }
            QProgressBar {
                border: 1px solid #e0e0e0;
                border-radius: 5px;
                height: 25px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4a90e2;
                border-radius: 5px;
            }
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #f0f0f0;
                border: 1px solid #e0e0e0;
                border-bottom-color: #e0e0e0;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 12px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                border-bottom-color: #ffffff;
            }
        """)
        
        self.title_label.setStyleSheet("""
            font-size: 28px;
            font-weight: bold;
            color: #4a90e2;
            margin: 20px 0;
        """)

        self.log_textarea.setStyleSheet("""
            background-color: #f9f9f9;
            border: 1px solid #e0e0e0;
            border-radius: 5px;
            padding: 8px;
            font-family: 'Courier New', monospace;
            font-size: 12px;
        """)

    def setup_animations(self):
        self.button_animation = QPropertyAnimation(self.upload_button, b"geometry")
        self.button_animation.setDuration(100)
        self.button_animation.setEasingCurve(QEasingCurve.OutCubic)

    def browse_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.file_label.setText(f'Selected File: {file_name}')
            self.excel_file = file_name

    def start_upload(self):
        if not hasattr(self, 'excel_file'):
            self.show_error_message('Please select an Excel file first.')
            return

        username = self.username_input.text()
        password = self.password_input.text()

        if not username or not password:
            self.show_error_message('Please enter both username and password.')
            return

        headless = self.headless_checkbox.isChecked()

    # Create and start the upload thread
        self.upload_thread = OptionsUploaderThread(self.excel_file, username, password, headless)
        self.upload_thread.progress_update.connect(self.update_progress)
        self.upload_thread.status_update.connect(self.update_status)
        self.upload_thread.error_occurred.connect(self.show_error_message)
        self.upload_thread.log_update.connect(self.update_log)

        self.upload_thread.start()


    def simulate_upload(self):
        self.progress_bar.setValue(0)
        self.status_label.setText("Uploading...")
        self.upload_button.setEnabled(False)

        # Simulate progress (replace with actual upload logic)
        for i in range(101):
            self.progress_bar.setValue(i)
            QApplication.processEvents()
            time.sleep(0.05)

        self.status_label.setText("Upload complete!")
        self.upload_button.setEnabled(True)
        self.show_success_message("Options uploaded successfully!")

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_status(self, message):
        self.status_label.setText(message)
        if message == "Upload process completed.":
            self.upload_button.setEnabled(True)

    def update_log(self, message):
        self.log_textarea.append(message)

    def show_error_message(self, message):
        QMessageBox.critical(self, "Error", message)
        self.upload_button.setEnabled(True)
        self.status_label.setText("Upload failed. Please try again.")

    def show_success_message(self, message):
        QMessageBox.information(self, "Success", message)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Adjust layout or widget sizes here if needed

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_O and event.modifiers() & Qt.ControlModifier:
            self.browse_file()
        super().keyPressEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = OptionsUploaderGUI()
    ex.show()
    sys.exit(app.exec_())