import sys
import os
import pandas as pd
import smtplib
import time
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import urllib.parse 
from email.header import Header

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, QTextEdit,
    QProgressBar, QMessageBox, QListWidget, QListWidgetItem, QGroupBox,
    QSizePolicy, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtXml import QDomDocument

# --- Configuration ---
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# --- EmailSenderThread (No changes from the previous full code version) ---
class EmailSenderThread(QThread):
    progress_update = pyqtSignal(int, int, int, str, str)
    finished_signal = pyqtSignal(list)
    log_signal = pyqtSignal(str, str)

    def __init__(self, df_batch, email_column, sender_email, app_password,
                 subject_template, body_template_html, attachment_paths=None, parent=None):
        super().__init__(parent)
        self.df_batch = df_batch 
        self.email_column = email_column
        self.sender_email = sender_email
        self.app_password = app_password 
        self.subject_template = subject_template
        self.body_template_html = body_template_html 
        self.attachment_paths = attachment_paths if attachment_paths else []
        self.is_running = True
        self.batch_failed_data = []

    def _render_template(self, template_str, row_data):
        new_template_str = template_str
        for col_name_from_excel, value in row_data.items():
            escaped_col_name = re.escape(str(col_name_from_excel))
            str_value = str(value)
            pattern = r"\{\{\s*" + escaped_col_name + r"\s*\}\}"
            new_template_str = re.sub(pattern, str_value, new_template_str)
        new_template_str = re.sub(r"\{\{.*?\}\}", "[MISSING_DATA]", new_template_str)
        return new_template_str

    def run(self):
        sent_count = 0
        failed_count = 0
        total_emails = len(self.df_batch)
        self.batch_failed_data = []
        start_time = time.time()

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(self.sender_email, self.app_password)
        except smtplib.SMTPAuthenticationError:
            auth_fail_msg = "Gmail Authentication Failed. Check email/app password."
            self.log_signal.emit(auth_fail_msg, "error")
            self.progress_update.emit(0, 0, total_emails, "Authentication Failed", "Error")
            self.finished_signal.emit([(-1, "N/A", auth_fail_msg)])
            return
        except Exception as e:
            conn_err_msg = f"SMTP Connection Error: {str(e)}"
            self.log_signal.emit(conn_err_msg, "error")
            self.progress_update.emit(0, 0, total_emails, conn_err_msg, "Error")
            self.finished_signal.emit([(-1, "N/A", conn_err_msg)])
            return

        for original_df_index, row in self.df_batch.iterrows():
            if not self.is_running:
                break
            
            recipient_email = str(row.get(self.email_column, "")).strip()
            email_error_details = []

            if not recipient_email or "@" not in recipient_email:
                failed_count += 1
                self.batch_failed_data.append((original_df_index, recipient_email, "Invalid or missing email address in sheet"))
                self.progress_update.emit(sent_count, failed_count, total_emails, recipient_email, self._calculate_eta(start_time, sent_count + failed_count, total_emails))
                continue

            try:
                current_subject = self._render_template(self.subject_template, row)
                current_body_html = self._render_template(self.body_template_html, row)

                msg = MIMEMultipart()
                msg['From'] = self.sender_email
                msg['To'] = recipient_email
                msg['Subject'] = current_subject
                msg.attach(MIMEText(current_body_html, 'html', 'utf-8'))

                for path in self.attachment_paths:
                    if not os.path.exists(path):
                        attach_warn = f"Attachment not found: {os.path.basename(path)} for {recipient_email}"
                        self.log_signal.emit(attach_warn, "warning")
                        email_error_details.append(f"Skipped: {os.path.basename(path)}")
                        continue
                    
                    filename_unicode = os.path.basename(path)
                    try:
                        with open(path, "rb") as attachment_file:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment_file.read())
                        encoders.encode_base64(part)
                        
                        try:
                            h = Header(filename_unicode, 'utf-8')
                            filename_param_value = h.encode()
                        except Exception:
                            filename_param_value = filename_unicode.encode('ascii', 'replace').decode('ascii').replace('"', '_')
                            if not filename_param_value.strip() or filename_param_value == '?' * len(filename_unicode):
                                _, ext = os.path.splitext(filename_unicode)
                                filename_param_value = f"attachment{ext if ext else '.dat'}"
                        
                        filename_star_value = f"UTF-8''{urllib.parse.quote(filename_unicode, encoding='utf-8')}"
                        part.add_header('Content-Disposition', 
                                        'attachment', 
                                        filename=filename_param_value,
                                        **{'filename*': filename_star_value})
                        msg.attach(part)
                    except Exception as e_attach:
                        attach_err = f"Failed to attach '{filename_unicode}' for {recipient_email}: {e_attach}"
                        self.log_signal.emit(attach_err, "warning")
                        email_error_details.append(f"Failed attach: {filename_unicode}")
                
                server.sendmail(self.sender_email, recipient_email, msg.as_string())
                sent_count += 1
                log_status = recipient_email
                if email_error_details: log_status += f" (attach issues: {', '.join(email_error_details)})"
                self.progress_update.emit(sent_count, failed_count, total_emails, log_status, self._calculate_eta(start_time, sent_count + failed_count, total_emails))

            except Exception as e:
                failed_count += 1
                error_message = str(e)
                if email_error_details: error_message += f" (Attach issues: {', '.join(email_error_details)})"
                self.batch_failed_data.append((original_df_index, recipient_email, error_message))
                self.progress_update.emit(sent_count, failed_count, total_emails, f"{recipient_email} (Failed: {error_message[:30]}...)", self._calculate_eta(start_time, sent_count + failed_count, total_emails))
            
            time.sleep(0.1)

        try:
            server.quit()
        except Exception:
            pass
        self.finished_signal.emit(self.batch_failed_data)

    def _calculate_eta(self, start_time, processed_count, total_count):
        if processed_count == 0: return "Calculating..."
        elapsed_time = time.time() - start_time
        if elapsed_time == 0 or processed_count == 0 : return "Calculating..."
        time_per_email = elapsed_time / processed_count
        remaining_emails = total_count - processed_count
        eta_seconds = remaining_emails * time_per_email
        if eta_seconds < 0: eta_seconds = 0
        if eta_seconds < 60: return f"{int(eta_seconds)} sec"
        elif eta_seconds < 3600: return f"{int(eta_seconds / 60)} min"
        else: return f"{int(eta_seconds / 3600)} hr {int((eta_seconds % 3600) / 60)} min"

    def stop(self):
        self.is_running = False

class BulkEmailerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("G-Send: Bulk Gmail Sender")
        self.setMinimumSize(1100, 700) 
        self.setGeometry(100, 100, 1200, 750) 

        self.df = None
        self.email_sender_thread = None
        self.all_failed_data = []
        self.attachment_paths = []
        self.settings_verified_for_bulk = False
        self.is_sending_sample = False # Only one type of sample send now
        # self.sample_type = "" # No longer needed as there's only one way to test

        self.tooltips = {
            "excel_browse": "Click to select an Excel file (.xlsx, .xls) containing recipient data.",
            "email_column": "Select the column from your Excel sheet that contains the email addresses.",
            "sender_email": "Your full Gmail address (e.g., your.name@gmail.com).",
            "app_password": ("Your 16-character Gmail App Password. Spaces will be automatically removed. "
                             "Generate this from your Google Account settings if 2-Step Verification is ON. "
                             "DO NOT use your regular Gmail password here."),
            # "send_sample_smtp" tooltip removed as button is removed
            "add_attachment": "Add one or more files to be attached to every email sent.",
            "clear_attachments": "Remove all currently listed attachments.",
            "attachments_list": "List of files that will be attached to each email.",
            "subject_template": ("Subject line for your emails. Use {{ ColumnName }} to insert data from your Excel. "
                                 "Example: 'Invoice for {{ CompanyName }}' if 'CompanyName' is a column in your Excel."),
            "body_template": ("Main content of your email (HTML format). Use {{ ColumnName }} for personalization. "
                              "You can paste rich text (like tables from Word/Excel) or write HTML directly. "
                              "Example:\n<p>Dear {{ FirstName }},</p>\n<p>Your order <b>{{ OrderID }}</b> has shipped.</p>"),
            "send_sample_template": ("Sends a test email to YOUR Gmail address using current SMTP settings, "
                                   "the defined HTML subject/body templates, and attachments. "
                                   "This verifies all settings and allows bulk emailing upon success."), # Updated
            "send_bulk": ("Starts sending emails to all recipients in the loaded Excel sheet "
                          "using the verified settings and HTML templates. (Enabled after a successful Test Email Template send)"), # Updated
            "retry_failed": ("Attempts to resend emails only to those recipients who failed in the previous "
                             "bulk send attempt. (Enabled if there were failures and settings are verified)")
        }

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_app_layout = QHBoxLayout(central_widget)

        left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(left_panel_widget)
        left_panel_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        left_panel_widget.setMaximumWidth(450)


        file_group = QGroupBox("Excel File")
        file_group_layout = QVBoxLayout()
        file_layout = QHBoxLayout()
        self.file_path_label = QLabel("No Excel file selected.")
        file_layout.addWidget(self.file_path_label, 1)
        self.browse_button = QPushButton("Browse Excel File")
        self.browse_button.clicked.connect(self.browse_file)
        self.browse_button.setToolTip(self.tooltips["excel_browse"])
        file_layout.addWidget(self.browse_button)
        file_group_layout.addLayout(file_layout)
        email_col_layout = QHBoxLayout()
        email_col_label = QLabel("Email Column:")
        email_col_label.setToolTip(self.tooltips["email_column"])
        email_col_layout.addWidget(email_col_label)
        self.email_column_combo = QComboBox()
        self.email_column_combo.setPlaceholderText("Load Excel to see columns")
        self.email_column_combo.setToolTip(self.tooltips["email_column"])
        email_col_layout.addWidget(self.email_column_combo)
        file_group_layout.addLayout(email_col_layout)
        file_group.setLayout(file_group_layout)
        left_panel_layout.addWidget(file_group)

        creds_group = QGroupBox("Gmail SMTP Settings")
        creds_group_layout = QVBoxLayout()
        sender_email_label = QLabel("Your Gmail Address:")
        sender_email_label.setToolTip(self.tooltips["sender_email"])
        sender_email_layout = QHBoxLayout()
        sender_email_layout.addWidget(sender_email_label)
        self.sender_email_input = QLineEdit()
        self.sender_email_input.setPlaceholderText("your.email@gmail.com")
        self.sender_email_input.setToolTip(self.tooltips["sender_email"])
        self.sender_email_input.textChanged.connect(self.reset_settings_verification)
        sender_email_layout.addWidget(self.sender_email_input)
        creds_group_layout.addLayout(sender_email_layout)
        app_password_label = QLabel("Gmail App Password:")
        app_password_label.setToolTip(self.tooltips["app_password"])
        app_password_layout = QHBoxLayout()
        app_password_layout.addWidget(app_password_label)
        self.app_password_input = QLineEdit()
        self.app_password_input.setEchoMode(QLineEdit.Password)
        self.app_password_input.setToolTip(self.tooltips["app_password"])
        self.app_password_input.textChanged.connect(self.reset_settings_verification)
        app_password_layout.addWidget(self.app_password_input)
        creds_group_layout.addLayout(app_password_layout)
        creds_group.setLayout(creds_group_layout)
        left_panel_layout.addWidget(creds_group)


        action_buttons_group = QGroupBox("Actions")
        action_buttons_layout = QVBoxLayout()
        # "Verify SMTP Settings" button REMOVED
        # "Test Email Template" button is the primary verifier now
        self.send_template_test_button = QPushButton("Test Email & Verify Settings") # Renamed
        self.send_template_test_button.setToolTip(self.tooltips["send_sample_template"])
        self.send_template_test_button.clicked.connect(self.send_sample_mail_action) # No lambda needed if no extra args
        action_buttons_layout.addWidget(self.send_template_test_button)
        
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        action_buttons_layout.addWidget(separator)

        self.send_button = QPushButton("Send Bulk Emails") # Removed "2."
        self.send_button.setToolTip(self.tooltips["send_bulk"])
        self.send_button.clicked.connect(self.start_sending_emails)
        self.send_button.setEnabled(False)
        action_buttons_layout.addWidget(self.send_button)
        
        self.retry_button = QPushButton("Retry Failed Emails")
        self.retry_button.setToolTip(self.tooltips["retry_failed"])
        self.retry_button.clicked.connect(self.retry_failed_emails)
        self.retry_button.setEnabled(False)
        action_buttons_layout.addWidget(self.retry_button)
        action_buttons_group.setLayout(action_buttons_layout)
        left_panel_layout.addWidget(action_buttons_group)
        
        stats_group = QGroupBox("Live Statistics & Log")
        stats_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        stats_layout.addWidget(self.progress_bar)
        
        status_labels_widget = QWidget()
        status_labels_layout = QVBoxLayout(status_labels_widget)
        status_labels_layout.setContentsMargins(0,0,0,0)
        self.status_label = QLabel("Status: Idle. Please test email & verify settings.") # Updated text
        self.status_label.setWordWrap(True)
        status_labels_layout.addWidget(self.status_label)
        stat_numbers_layout = QHBoxLayout()
        self.sent_label = QLabel("Successfully Sent: 0")
        stat_numbers_layout.addWidget(self.sent_label)
        self.failed_label = QLabel("Failed to Send: 0")
        stat_numbers_layout.addWidget(self.failed_label)
        status_labels_layout.addLayout(stat_numbers_layout)
        self.eta_label = QLabel("Estimated Time of Completion: N/A")
        status_labels_layout.addWidget(self.eta_label)
        status_labels_widget.setFixedHeight(95) # Increased height for status area
        stats_layout.addWidget(status_labels_widget)

        stats_layout.addWidget(QLabel("Log:"))
        self.log_widget = QListWidget()
        stats_layout.addWidget(self.log_widget)
        stats_group.setLayout(stats_layout)
        left_panel_layout.addWidget(stats_group)
        left_panel_layout.addStretch(0)

        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        right_panel_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        content_group = QGroupBox("Email Template (HTML - use {{ ColumnName }} for placeholders)")
        content_layout = QVBoxLayout()
        subject_label = QLabel("Subject:")
        subject_label.setToolTip(self.tooltips["subject_template"])
        content_layout.addWidget(subject_label)
        self.subject_input = QLineEdit()
        self.subject_input.setText("Update for {{ Name }}")
        self.subject_input.setToolTip(self.tooltips["subject_template"])
        self.subject_input.textChanged.connect(self.reset_settings_verification)
        content_layout.addWidget(self.subject_input)
        body_label = QLabel("Body (HTML):")
        body_label.setToolTip(self.tooltips["body_template"])
        content_layout.addWidget(body_label)
        self.body_input = QTextEdit()
        self.body_input.setAcceptRichText(True)
        self.body_input.setPlaceholderText("Paste rich text (tables, formatted text) or type HTML directly.\nUse {{ Placeholder }} for dynamic content.")
        default_html_body = """<p>Dear {{ Name }},</p>
<p>This is a test email regarding your item: <b>{{ Item }}</b>.</p>
<p>Your reference is {{ RefID }}.</p>
<p>Best regards,<br>G-Send Team</p>"""
        self.body_input.setHtml(default_html_body)
        self.body_input.setToolTip(self.tooltips["body_template"])
        self.body_input.textChanged.connect(self.reset_settings_verification)
        content_layout.addWidget(self.body_input)
        # Test Email Template button moved to Actions group
        content_group.setLayout(content_layout)
        right_panel_layout.addWidget(content_group)

        attach_group = QGroupBox("Attachments (Applied to all emails)")
        attach_layout = QVBoxLayout()
        attach_buttons_layout = QHBoxLayout()
        self.add_attachment_button = QPushButton("Add Attachment(s)")
        self.add_attachment_button.setToolTip(self.tooltips["add_attachment"])
        self.add_attachment_button.clicked.connect(self.add_attachments)
        attach_buttons_layout.addWidget(self.add_attachment_button)
        self.clear_attachments_button = QPushButton("Clear Attachments")
        self.clear_attachments_button.setToolTip(self.tooltips["clear_attachments"])
        self.clear_attachments_button.clicked.connect(self.clear_attachments)
        attach_buttons_layout.addWidget(self.clear_attachments_button)
        attach_layout.addLayout(attach_buttons_layout)
        self.attachments_list_widget = QListWidget()
        self.attachments_list_widget.setFixedHeight(100)
        self.attachments_list_widget.setToolTip(self.tooltips["attachments_list"])
        attach_layout.addWidget(self.attachments_list_widget)
        attach_group.setLayout(attach_layout)
        right_panel_layout.addWidget(attach_group)

        main_app_layout.addWidget(left_panel_widget, 2) 
        main_app_layout.addWidget(right_panel_widget, 3)
        
        self.reset_settings_verification()

    def get_app_password(self):
        return self.app_password_input.text().replace(" ", "")

    def get_email_body_content(self):
        raw_html = self.body_input.toHtml()
        doc = QDomDocument()
        if doc.setContent(f"<html><body>{raw_html}</body></html>"):
            body_node = doc.documentElement().firstChildElement("body")
            if not body_node.isNull():
                cleaned_html = ""
                child = body_node.firstChild()
                while not child.isNull():
                    temp_doc_child = QDomDocument()
                    imported_node = temp_doc_child.importNode(child, True)
                    temp_doc_child.appendChild(imported_node)
                    cleaned_html += temp_doc_child.toString(-1)
                    child = child.nextSibling()
                if cleaned_html.strip():
                    return cleaned_html
        body_match = re.search(r"<body[^>]*>(.*?)</body>", raw_html, re.IGNORECASE | re.DOTALL)
        if body_match:
            self.log_message("Used regex fallback for HTML body cleaning.", "info")
            return body_match.group(1).strip()
        self.log_message("Warning: Could not cleanly extract body HTML. Sending raw editor output.", "warning")
        return raw_html

    def _prepare_and_start_sending(self, dataframe_to_send, is_sample_send=False, sample_recipient_email=None):
        sender_email = self.sender_email_input.text().strip()
        app_password = self.get_app_password()
        subject_template = self.subject_input.text()
        body_template_html = self.get_email_body_content()
        email_column = self.email_column_combo.currentText()

        if not is_sample_send:
            if not self.settings_verified_for_bulk:
                QMessageBox.warning(self, "Verification Required", "Please test email & verify settings first.")
                return False
            if self.df is None :
                 QMessageBox.warning(self, "Input Error", "Please load an Excel file first.")
                 return False
            if not email_column:
                QMessageBox.warning(self, "Input Error", "Please select the email column.")
                return False
            if dataframe_to_send.empty:
                QMessageBox.information(self, "No Data", "No emails to send in the current selection.")
                return False
        else: 
             if not sender_email or not app_password:
                QMessageBox.warning(self, "Input Error", "Gmail Address and App Password required for test send.")
                return False
             if not subject_template or len(body_template_html) < 50: # Basic check for "empty" HTML
                QMessageBox.warning(self, "Input Error", "Email Subject and Body template required for test send.")
                return False
        
        # self.send_smtp_verify_button.setEnabled(False) # Button removed
        self.send_template_test_button.setEnabled(False)
        self.send_button.setEnabled(False)
        self.retry_button.setEnabled(False)
        self.browse_button.setEnabled(False)

        self.status_label.setText(f"Status: Starting {'test ' if is_sample_send else ''}email process...")
        self.log_message(f"Starting {'test ' if is_sample_send else ''}email process...")
        self.reset_partial_stats_for_send() 
        if not dataframe_to_send.empty:
            self.progress_bar.setMaximum(len(dataframe_to_send))
        else:
            self.progress_bar.setMaximum(1) 
            if is_sample_send:
                self.progress_bar.setMaximum(1)
            else:
                self.progress_bar.setMaximum(100)
                self.progress_bar.setValue(0)

        self.email_sender_thread = EmailSenderThread(
            df_batch=dataframe_to_send,
            email_column=email_column if not is_sample_send else 'EmailTo',
            sender_email=sender_email,
            app_password=app_password,
            subject_template=subject_template,
            body_template_html=body_template_html,
            attachment_paths=self.attachment_paths
        )
        self.email_sender_thread.log_signal.connect(self.log_message)
        self.email_sender_thread.progress_update.connect(self.update_progress)
        self.email_sender_thread.finished_signal.connect(self.on_sending_finished)
        self.email_sender_thread.start()
        return True

    def send_sample_mail_action(self): # No longer needs sample_type
        sender_email = self.sender_email_input.text().strip()
        if not self.app_password_input.text():
             QMessageBox.warning(self, "Input Error", "App Password cannot be empty.")
             return
        
        sample_df_data = {
            str('EmailTo'): [sender_email], str('Name'): ['Valued Tester'],
            str('Item'): ['Test Item X'], str('ID'): ['TID-001'], str('RefID'): ['REF-XYZ']
        }
        if self.df is not None and not self.df.empty:
            first_row_data = self.df.iloc[0].to_dict()
            first_row_data[str('EmailTo')] = sender_email 
            for key, val in sample_df_data.items():
                if key not in first_row_data:
                    first_row_data[key] = val[0] if isinstance(val, list) else val
            sample_df = pd.DataFrame([first_row_data])
        else: 
            sample_df = pd.DataFrame(sample_df_data)
        
        self.is_sending_sample = True
        if not self._prepare_and_start_sending(sample_df, is_sample_send=True, sample_recipient_email=sender_email):
            # self.send_smtp_verify_button.setEnabled(True) # Button removed
            self.send_template_test_button.setEnabled(True)
            self.browse_button.setEnabled(True)
            self.is_sending_sample = False

    def start_sending_emails(self):
        if self.df is None:
            QMessageBox.warning(self, "Input Error", "Please load an Excel file first.")
            return
        self.all_failed_data = [] 
        self.retry_button.setEnabled(False)
        if not self._prepare_and_start_sending(self.df.copy()):
            # self.send_smtp_verify_button.setEnabled(True) # Button removed
            self.send_template_test_button.setEnabled(True)
            self.browse_button.setEnabled(True)
            self.send_button.setEnabled(self.settings_verified_for_bulk)

    def retry_failed_emails(self):
        if not self.all_failed_data:
            QMessageBox.information(self, "No Failures", "There are no emails marked as failed to retry.")
            return
        if self.df is None:
            QMessageBox.warning(self, "Error", "Original Excel data not loaded. Cannot retry.")
            return
        
        failed_indices = [item[0] for item in self.all_failed_data if isinstance(item, tuple) and len(item) > 0 and isinstance(item[0], int) and item[0] != -1]
        if not failed_indices:
            QMessageBox.information(self, "Info", "No valid failed items to retry.")
            return

        df_to_retry = self.df.loc[failed_indices].copy()
        self.all_failed_data = [] 
        if not self._prepare_and_start_sending(df_to_retry):
            # self.send_smtp_verify_button.setEnabled(True) # Button removed
            self.send_template_test_button.setEnabled(True)
            self.browse_button.setEnabled(True)
            self.retry_button.setEnabled(self.settings_verified_for_bulk and bool(self.all_failed_data))
        
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_path_label.setText(os.path.basename(file_path))
            self.log_message(f"Selected file: {file_path}")
            try:
                self.df = pd.read_excel(file_path)
                self.df.columns = [str(col).strip() for col in self.df.columns]
                self.email_column_combo.clear()
                self.email_column_combo.addItems(self.df.columns)
                self.log_message(f"Loaded {len(self.df)} rows. Columns: {', '.join(self.df.columns)}. Select email column.")
                common_email_cols = ['email', 'e-mail', 'email address']
                for i, col_name in enumerate(self.df.columns):
                    if col_name.lower() in common_email_cols:
                        self.email_column_combo.setCurrentIndex(i)
                        break
                self.reset_stats_for_new_file()
                self.reset_settings_verification()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")
                self.log_message(f"Error loading Excel: {e}", "error")
                self.df = None
                self.file_path_label.setText("No Excel file selected.")
                self.email_column_combo.clear()
                self.reset_settings_verification()

    def reset_stats_for_new_file(self):
        self.progress_bar.setMaximum(100) 
        self.progress_bar.setValue(0)    
        self.sent_label.setText("Successfully Sent: 0")
        self.failed_label.setText("Failed to Send: 0")
        self.eta_label.setText("Estimated Time of Completion: N/A")
        if hasattr(self, 'log_widget'):
            self.log_widget.clear()
        self.all_failed_data = []
        self.retry_button.setEnabled(False)

    def reset_partial_stats_for_send(self):
        self.progress_bar.setMaximum(100) 
        self.progress_bar.setValue(0)     
        self.sent_label.setText("Successfully Sent: 0")
        self.failed_label.setText("Failed to Send: 0")
        self.eta_label.setText("Estimated Time of Completion: N/A")

    def log_message(self, message, level="info"):
        item = QListWidgetItem(message)
        if level == "error": item.setForeground(Qt.red)
        elif level == "warning": item.setForeground(Qt.darkYellow)
        self.log_widget.addItem(item)
        self.log_widget.scrollToBottom()
        QApplication.processEvents()

    def update_progress(self, sent, failed, total, current_email_info, eta_str):
        if total > 0:
            self.progress_bar.setMaximum(total)
        else: 
            self.progress_bar.setMaximum(1)
        self.progress_bar.setValue(sent + failed)
        prefix = "(Test) " if self.is_sending_sample else "" # Changed from "Sample" to "Test"
        self.sent_label.setText(f"Successfully Sent: {prefix}{sent}")
        self.failed_label.setText(f"Failed to Send: {prefix}{failed}")
        self.status_label.setText(f"Status: {prefix}Processing {current_email_info} ({sent+failed}/{total})")
        self.eta_label.setText(f"ETA: {prefix}{eta_str}")

    def reset_settings_verification(self):
        self.settings_verified_for_bulk = False
        self.send_button.setEnabled(False)
        self.retry_button.setEnabled(False) 
        self.status_label.setText("Status: Settings changed. Please test email & verify settings.")
        if hasattr(self, 'log_widget'):
             self.log_message("Settings (SMTP, file, template, attachments) may have changed. Verification required for bulk send.", "warning")

    def add_attachments(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments", "", "All Files (*.*)")
        if files:
            changed = False
            for file_path in files:
                if file_path not in self.attachment_paths:
                    self.attachment_paths.append(file_path)
                    self.attachments_list_widget.addItem(os.path.basename(file_path))
                    changed = True
            if changed:
                self.log_message(f"Added attachment(s). Total: {len(self.attachment_paths)}.")
                self.reset_settings_verification()

    def clear_attachments(self):
        if not self.attachment_paths: return
        self.attachment_paths.clear()
        self.attachments_list_widget.clear()
        self.log_message("Cleared all attachments.")
        self.reset_settings_verification()
    
    def on_sending_finished(self, failed_data_from_thread):
        if self.is_sending_sample:
            self.handle_sample_mail_result(failed_data_from_thread) # No longer needs sample_type
        else:
            self.handle_bulk_mail_result(failed_data_from_thread)

        # self.send_smtp_verify_button.setEnabled(True) # Button removed
        self.send_template_test_button.setEnabled(True)
        self.browse_button.setEnabled(True)

    def handle_sample_mail_result(self, failed_data): # No longer needs sample_type
        self.is_sending_sample = False
        auth_error = any(item[0] == -1 for item in failed_data) if failed_data else False
        send_failure = any(item[0] != -1 for item in failed_data) if failed_data else False
        success = not auth_error and not send_failure

        if success:
            self.settings_verified_for_bulk = True 
            self.send_button.setEnabled(True)      
            self.retry_button.setEnabled(self.settings_verified_for_bulk and bool(self.all_failed_data))
            msg_title = "Test Email Success"
            msg_text = "Test email sent successfully! Settings verified for bulk send."
            self.status_label.setText(f"Status: {msg_title}. Ready for bulk send.")
            self.log_message("Test email successful. Settings verified.", "info")
            QMessageBox.information(self, msg_title, msg_text)
        else: 
            self.settings_verified_for_bulk = False 
            self.send_button.setEnabled(False)
            self.retry_button.setEnabled(False)
            error_msg_prefix = "Failed to send test email"
            error_msg = f"{error_msg_prefix}. "
            if auth_error and failed_data: error_msg += f"Authentication/Connection Error: {failed_data[0][2]}"
            elif send_failure and failed_data: error_msg += f"Could not send to {failed_data[0][1]}: {failed_data[0][2]}"
            elif not failed_data : error_msg += "No specific error data returned."
            else: error_msg += "Unknown error during test send."
            self.log_message(error_msg, "error")
            self.status_label.setText("Status: Test Email Failed. Bulk send disabled.")
            QMessageBox.critical(self, "Test Email Failed", error_msg + "\nPlease check settings and details.")
        
        if hasattr(self, 'email_sender_thread'):
             self.email_sender_thread = None

    def handle_bulk_mail_result(self, failed_data_from_thread):
        final_sent = 0
        final_failed = 0
        try: 
            sent_text_match = re.search(r'\d+', self.sent_label.text())
            if sent_text_match: final_sent = int(sent_text_match.group(0))
            failed_text_match = re.search(r'\d+', self.failed_label.text())
            if failed_text_match: final_failed = int(failed_text_match.group(0))
        except Exception:
            self.log_message("Warning: Could not accurately parse final sent/failed counts from UI labels.", "warning")
            if self.df and not self.df.empty:
                total_attempted_in_df = len(self.df) 
                actual_failures_from_thread = len([f for f in failed_data_from_thread if f[0] != -1])
                final_failed = actual_failures_from_thread
                final_sent = total_attempted_in_df - final_failed
            else: 
                 final_failed = len([f for f in failed_data_from_thread if f[0] != -1])
                 final_sent = 0

        self.status_label.setText(f"Status: Bulk Send Completed. Sent: {final_sent}, Failed: {final_failed}")
        self.log_message(f"Bulk email process finished. Successfully sent: {final_sent}, Failed: {final_failed}")

        for original_idx, email, reason in failed_data_from_thread:
            if original_idx != -1:
                if not any(existing_item[0] == original_idx for existing_item in self.all_failed_data):
                    self.all_failed_data.append((original_idx, email, reason))

        if self.all_failed_data:
            self.retry_button.setEnabled(self.settings_verified_for_bulk and bool(self.all_failed_data))
            self.log_message(f"{len(self.all_failed_data)} email(s) in retry queue.", "warning")
        else:
            self.retry_button.setEnabled(False)
            if final_sent > 0 and final_failed == 0 :
                 self.log_message("All emails in this batch sent successfully.", "info")
        
        self.send_button.setEnabled(self.settings_verified_for_bulk)
        if hasattr(self, 'email_sender_thread'):
            self.email_sender_thread = None

    def closeEvent(self, event):
        if hasattr(self, 'email_sender_thread') and self.email_sender_thread and self.email_sender_thread.isRunning():
            reply = QMessageBox.question(self, 'Confirm Exit',
                                         "Email sending is in progress. Are you sure you want to exit?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                if self.email_sender_thread:
                    self.email_sender_thread.stop()
                    self.email_sender_thread.wait(2000) 
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = BulkEmailerApp()
    main_win.show()
    sys.exit(app.exec_())