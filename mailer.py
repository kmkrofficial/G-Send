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

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, QTextEdit,
    QProgressBar, QMessageBox, QListWidget, QListWidgetItem, QGroupBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon

# --- Configuration ---
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# --- EmailSenderThread (No changes from previous version) ---
class EmailSenderThread(QThread):
    progress_update = pyqtSignal(int, int, int, str, str)
    finished_signal = pyqtSignal(list)
    log_signal = pyqtSignal(str, str)

    def __init__(self, df, email_column, sender_email, app_password,
                 subject_template, body_template, attachment_paths=None, parent=None):
        super().__init__(parent)
        self.df = df
        self.email_column = email_column
        self.sender_email = sender_email
        self.app_password = app_password
        self.subject_template = subject_template
        self.body_template = body_template
        self.attachment_paths = attachment_paths if attachment_paths else []
        self.is_running = True
        self.failed_emails_data = []

    def _render_template(self, template_str, row_data):
        new_template_str = template_str
        for col_name_from_excel, value in row_data.items():
            escaped_col_name = re.escape(str(col_name_from_excel))
            pattern = r"\{\{\s*" + escaped_col_name + r"\s*\}\}"
            new_template_str = re.sub(pattern, str(value), new_template_str)
        new_template_str = re.sub(r"\{\{.*?\}\}", "[MISSING_DATA]", new_template_str)
        return new_template_str

    def run(self):
        sent_count = 0
        failed_count = 0
        total_emails = len(self.df)
        self.failed_emails_data = []
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

        for original_df_index, row in self.df.iterrows():
            if not self.is_running: break
            recipient_email = str(row.get(self.email_column, "")).strip()
            email_error_details = []

            if not recipient_email or "@" not in recipient_email:
                failed_count += 1
                self.failed_emails_data.append((original_df_index, recipient_email, "Invalid or missing email address in sheet"))
                self.progress_update.emit(sent_count, failed_count, total_emails, recipient_email, self._calculate_eta(start_time, sent_count + failed_count, total_emails))
                continue

            try:
                current_subject = self._render_template(self.subject_template, row)
                current_body = self._render_template(self.body_template, row)

                msg = MIMEMultipart()
                msg['From'] = self.sender_email
                msg['To'] = recipient_email
                msg['Subject'] = current_subject
                msg.attach(MIMEText(current_body, 'plain'))

                for path in self.attachment_paths:
                    if not os.path.exists(path):
                        attach_warn = f"Attachment not found and skipped for {recipient_email}: {os.path.basename(path)}"
                        self.log_signal.emit(attach_warn, "warning")
                        email_error_details.append(f"Skipped attachment: {os.path.basename(path)} (not found)")
                        continue
                    try:
                        with open(path, "rb") as attachment_file:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment_file.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename=\"{os.path.basename(path)}\"")
                        msg.attach(part)
                    except Exception as e_attach:
                        attach_err = f"Failed to attach {os.path.basename(path)} for {recipient_email}: {e_attach}"
                        self.log_signal.emit(attach_err, "warning")
                        email_error_details.append(f"Failed to attach {os.path.basename(path)}")

                server.sendmail(self.sender_email, recipient_email, msg.as_string())
                sent_count += 1
                log_status = recipient_email
                if email_error_details:
                    log_status += f" (with attachment issues: {', '.join(email_error_details)})"
                self.progress_update.emit(sent_count, failed_count, total_emails, log_status, self._calculate_eta(start_time, sent_count + failed_count, total_emails))

            except Exception as e:
                failed_count += 1
                error_message = str(e)
                if email_error_details:
                    error_message += f" (Additional issues: {', '.join(email_error_details)})"
                self.failed_emails_data.append((original_df_index, recipient_email, error_message))
                self.progress_update.emit(sent_count, failed_count, total_emails, f"{recipient_email} (Failed: {error_message[:30]}...)", self._calculate_eta(start_time, sent_count + failed_count, total_emails))

            time.sleep(0.1)

        try:
            server.quit()
        except Exception:
            pass
        self.finished_signal.emit(self.failed_emails_data)

    def _calculate_eta(self, start_time, processed_count, total_count):
        if processed_count == 0: return "Calculating..."
        elapsed_time = time.time() - start_time
        if processed_count == 0: return "Calculating..."
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
        self.setWindowTitle("Bulk Email Sender for Gmail")
        self.setGeometry(100, 100, 850, 800) 

        self.df = None
        self.email_sender_thread = None
        self.all_failed_data = []
        self.attachment_paths = []
        self.settings_verified_for_bulk = False # Renamed for clarity
        self.is_sending_sample = False 
        self.sample_type = "" 

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.tooltips = {
            "excel_browse": "Click to select an Excel file (.xlsx, .xls) containing recipient data.",
            "email_column": "Select the column from your Excel sheet that contains the email addresses.",
            "sender_email": "Your full Gmail address (e.g., your.name@gmail.com).",
            "app_password": ("Your 16-character Gmail App Password. "
                             "Generate this from your Google Account settings if 2-Step Verification is ON. "
                             "DO NOT use your regular Gmail password here."),
            "send_sample_smtp": ("Sends a test email to YOUR Gmail address using current SMTP settings, "
                                 "template, and attachments. Verifies login and basic sending. "
                                 "A successful send here allows bulk emailing."),
            "add_attachment": "Add one or more files to be attached to every email sent.",
            "clear_attachments": "Remove all currently listed attachments.",
            "attachments_list": "List of files that will be attached to each email.",
            "subject_template": ("Subject line for your emails. Use {{ ColumnName }} to insert data from your Excel. "
                                 "Example: 'Invoice for {{ CompanyName }}' if 'CompanyName' is a column in your Excel."),
            "body_template": ("Main content of your email. Use {{ ColumnName }} for personalization. "
                              "Example:\nDear {{ FirstName }},\nYour order {{ OrderID }} has shipped."),
            "send_sample_template": ("Sends a test email to YOUR Gmail address using current SMTP settings, "
                                   "the defined subject/body templates, and attachments. "
                                   "Helps to preview the email content and test template rendering. "
                                   "A successful send here also allows bulk emailing."),
            "send_bulk": ("Starts sending emails to all recipients in the loaded Excel sheet "
                          "using the verified settings and templates. (Enabled after any successful test send)"),
            "retry_failed": ("Attempts to resend emails only to those recipients who failed in the previous "
                             "bulk send attempt. (Enabled if there were failures and settings are verified)")
        }
        # --- UI Setup (Identical to previous, just ensure tooltips are assigned) ---
        # File Selection
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
        main_layout.addWidget(file_group)

        # Gmail Credentials & Verification
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
        self.send_smtp_verify_button = QPushButton("1. Verify SMTP & Send Test Mail")
        self.send_smtp_verify_button.setToolTip(self.tooltips["send_sample_smtp"])
        self.send_smtp_verify_button.clicked.connect(lambda: self.send_sample_mail_action(sample_type="smtp_verify"))
        creds_group_layout.addWidget(self.send_smtp_verify_button)
        creds_group.setLayout(creds_group_layout)
        main_layout.addWidget(creds_group)

        # Attachments
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
        self.attachments_list_widget.setFixedHeight(80)
        self.attachments_list_widget.setToolTip(self.tooltips["attachments_list"])
        attach_layout.addWidget(self.attachments_list_widget)
        attach_group.setLayout(attach_layout)
        main_layout.addWidget(attach_group)

        # Email Content
        content_group = QGroupBox("Email Template (use {{ ColumnName }} for placeholders)")
        content_layout = QVBoxLayout()
        subject_label = QLabel("Subject:")
        subject_label.setToolTip(self.tooltips["subject_template"])
        content_layout.addWidget(subject_label)
        self.subject_input = QLineEdit()
        self.subject_input.setText("Update for {{ Name }}")
        self.subject_input.setToolTip(self.tooltips["subject_template"])
        self.subject_input.textChanged.connect(self.reset_settings_verification) # Added
        content_layout.addWidget(self.subject_input)
        body_label = QLabel("Body:")
        body_label.setToolTip(self.tooltips["body_template"])
        content_layout.addWidget(body_label)
        self.body_input = QTextEdit()
        self.body_input.setPlaceholderText("Hello {{ Name }},\n\nThis is an update regarding your account.\nYour ID is {{ ID }}.\n\nThanks,\nTeam")
        self.body_input.setText("Dear {{ Name }},\n\nThis is a test email regarding your item: {{ Item }}.\nYour reference is {{ RefID }}.\n\nBest regards,\nBulk Mailer")
        self.body_input.setToolTip(self.tooltips["body_template"])
        self.body_input.textChanged.connect(self.reset_settings_verification) # Added
        content_layout.addWidget(self.body_input)
        self.send_template_test_button = QPushButton("Test Email Template")
        self.send_template_test_button.setToolTip(self.tooltips["send_sample_template"])
        self.send_template_test_button.clicked.connect(lambda: self.send_sample_mail_action(sample_type="template_test"))
        content_layout.addWidget(self.send_template_test_button)
        content_group.setLayout(content_layout)
        main_layout.addWidget(content_group)

        # Action Buttons
        action_layout = QHBoxLayout()
        self.send_button = QPushButton("2. Send Bulk Emails")
        self.send_button.setToolTip(self.tooltips["send_bulk"])
        self.send_button.clicked.connect(self.start_sending_emails)
        self.send_button.setEnabled(False)
        action_layout.addWidget(self.send_button)
        self.retry_button = QPushButton("Retry Failed Emails")
        self.retry_button.setToolTip(self.tooltips["retry_failed"])
        self.retry_button.clicked.connect(self.retry_failed_emails)
        self.retry_button.setEnabled(False)
        action_layout.addWidget(self.retry_button)
        main_layout.addLayout(action_layout)

        # Statistics & Progress
        stats_group = QGroupBox("Live Statistics & Log")
        stats_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        stats_layout.addWidget(self.progress_bar)
        self.status_label = QLabel("Status: Idle. Please verify settings by sending a test mail.")
        stats_layout.addWidget(self.status_label)
        stat_numbers_layout = QHBoxLayout()
        self.sent_label = QLabel("Successfully Sent: 0")
        stat_numbers_layout.addWidget(self.sent_label)
        self.failed_label = QLabel("Failed to Send: 0")
        stat_numbers_layout.addWidget(self.failed_label)
        stats_layout.addLayout(stat_numbers_layout)
        self.eta_label = QLabel("Estimated Time of Completion: N/A")
        stats_layout.addWidget(self.eta_label)
        stats_layout.addWidget(QLabel("Log:"))
        self.log_widget = QListWidget()
        self.log_widget.setFixedHeight(100)
        stats_layout.addWidget(self.log_widget)
        stats_group.setLayout(stats_layout)
        main_layout.addWidget(stats_group)
        # --- End UI Setup ---
        
        self.reset_settings_verification() # Call once at init

    def browse_file(self):
        # ... (identical to previous version)
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
                self.reset_settings_verification() # File change resets verification
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")
                self.log_message(f"Error loading Excel: {e}", "error")
                self.df = None
                self.file_path_label.setText("No Excel file selected.")
                self.email_column_combo.clear()
                self.reset_settings_verification() # File load error also resets

    def reset_stats_for_new_file(self):
        # ... (identical to previous version)
        self.progress_bar.setValue(0)
        self.sent_label.setText("Successfully Sent: 0")
        self.failed_label.setText("Failed to Send: 0")
        self.eta_label.setText("Estimated Time of Completion: N/A")
        self.log_widget.clear()
        self.all_failed_data = []
        self.retry_button.setEnabled(False)


    def reset_partial_stats_for_send(self):
        # ... (identical to previous version)
        self.progress_bar.setValue(0)
        self.sent_label.setText("Successfully Sent: 0")
        self.failed_label.setText("Failed to Send: 0")
        self.eta_label.setText("Estimated Time of Completion: N/A")

    def log_message(self, message, level="info"):
        # ... (identical to previous version)
        item = QListWidgetItem(message)
        if level == "error": item.setForeground(Qt.red)
        elif level == "warning": item.setForeground(Qt.darkYellow)
        self.log_widget.addItem(item)
        self.log_widget.scrollToBottom()
        QApplication.processEvents()

    def update_progress(self, sent, failed, total, current_email_info, eta_str):
        # ... (identical to previous version)
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(sent + failed)
        prefix = "(Sample) " if self.is_sending_sample else ""
        self.sent_label.setText(f"Successfully Sent: {prefix}{sent}")
        self.failed_label.setText(f"Failed to Send: {prefix}{failed}")
        self.status_label.setText(f"Status: {prefix}Processing {current_email_info} ({sent+failed}/{total})")
        self.eta_label.setText(f"ETA: {prefix}{eta_str}")


    def reset_settings_verification(self): # MODIFIED
        self.settings_verified_for_bulk = False
        self.send_button.setEnabled(False)
        # Retry button also depends on this verification for safety
        self.retry_button.setEnabled(False) 
        self.status_label.setText("Status: Settings changed. Please verify by sending a test mail.")
        self.log_message("Settings (SMTP, file, template, attachments) may have changed. Verification required for bulk send.", "warning")

    def add_attachments(self):
        # ... (identical to previous version, calls reset_settings_verification)
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
        # ... (identical to previous version, calls reset_settings_verification)
        if not self.attachment_paths: return
        self.attachment_paths.clear()
        self.attachments_list_widget.clear()
        self.log_message("Cleared all attachments.")
        self.reset_settings_verification()

    def send_sample_mail_action(self, sample_type="smtp_verify"):
        # ... (identical to previous version in terms of how it initiates the send)
        sender_email = self.sender_email_input.text().strip()
        app_password = self.app_password_input.text()
        subject_template = self.subject_input.text()
        body_template = self.body_input.toPlainText()

        if not sender_email or not app_password:
            QMessageBox.warning(self, "Input Error", "Please enter your Gmail address and App Password.")
            return
        if not ("@" in sender_email and "." in sender_email):
            QMessageBox.warning(self, "Input Error", "Please enter a valid Gmail address.")
            return
        if not subject_template or not body_template:
            QMessageBox.warning(self, "Input Error", "Please provide a Subject and Body for the email template.")
            return

        self.is_sending_sample = True
        self.sample_type = sample_type 
        
        self.send_smtp_verify_button.setEnabled(False)
        self.send_template_test_button.setEnabled(False)
        self.send_button.setEnabled(False)
        self.retry_button.setEnabled(False)

        self.status_label.setText(f"Status: Sending {sample_type.replace('_', ' ')} email...")
        self.log_message(f"Attempting to send {sample_type.replace('_', ' ')} email to {sender_email}...")
        self.reset_partial_stats_for_send()

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

        self.email_sender_thread = EmailSenderThread(
            df=sample_df, email_column='EmailTo',
            sender_email=sender_email, app_password=app_password,
            subject_template=subject_template, body_template=body_template,
            attachment_paths=self.attachment_paths
        )
        self.email_sender_thread.log_signal.connect(self.log_message)
        self.email_sender_thread.progress_update.connect(self.update_progress)
        self.email_sender_thread.finished_signal.connect(self.on_sending_finished)
        self.email_sender_thread.start()

    def _start_email_thread(self, dataframe_to_send, email_column, is_retry=False):
        # ... (identical to previous version)
        sender_email = self.sender_email_input.text().strip()
        app_password = self.app_password_input.text()
        subject_template = self.subject_input.text()
        body_template = self.body_input.toPlainText()

        self.send_button.setEnabled(False)
        self.retry_button.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.send_smtp_verify_button.setEnabled(False)
        self.send_template_test_button.setEnabled(False)

        if is_retry:
            self.status_label.setText("Status: Starting retry...")
            self.log_message(f"Retrying {len(dataframe_to_send)} failed email(s)...")
        else:
            self.status_label.setText("Status: Starting bulk email process...")
            self.log_message(f"Starting bulk email process for {len(dataframe_to_send)} emails...")

        self.reset_partial_stats_for_send()

        self.email_sender_thread = EmailSenderThread(
            dataframe_to_send, email_column, sender_email, app_password,
            subject_template, body_template, self.attachment_paths
        )
        self.email_sender_thread.log_signal.connect(self.log_message)
        self.email_sender_thread.progress_update.connect(self.update_progress)
        self.email_sender_thread.finished_signal.connect(self.on_sending_finished)
        self.email_sender_thread.start()

    def on_sending_finished(self, failed_data_from_thread):
        # ... (identical to previous version)
        if self.is_sending_sample:
            self.handle_sample_mail_result(failed_data_from_thread, self.sample_type)
        else:
            self.handle_bulk_mail_result(failed_data_from_thread)

        self.send_smtp_verify_button.setEnabled(True)
        self.send_template_test_button.setEnabled(True)
        self.browse_button.setEnabled(True)


    def handle_sample_mail_result(self, failed_data, sample_type): # MODIFIED
        self.is_sending_sample = False
        
        auth_error = any(item[0] == -1 for item in failed_data)
        send_failure = any(item[0] != -1 for item in failed_data)
        success = not auth_error and not send_failure

        if success:
            self.settings_verified_for_bulk = True # Key change: ANY successful sample verifies settings
            self.send_button.setEnabled(True)       # Enable bulk send
            self.retry_button.setEnabled(bool(self.all_failed_data)) # Enable retry if there are past failures

            if sample_type == "smtp_verify":
                self.status_label.setText("Status: SMTP Verified. Ready for bulk send.")
                self.log_message("SMTP verification & test mail successful.", "info")
                QMessageBox.information(self, "SMTP Verification Success", "SMTP settings verified and test email sent successfully!")
            elif sample_type == "template_test":
                self.status_label.setText("Status: Template Test Successful. Settings verified. Ready for bulk send.")
                self.log_message("Template test email sent successfully. Settings also verified for bulk send.", "info")
                QMessageBox.information(self, "Template Test Success", "Test email with current template sent successfully! Settings verified.")
        else: # Failure
            self.settings_verified_for_bulk = False # Verification failed
            self.send_button.setEnabled(False)
            self.retry_button.setEnabled(False) # Disable retry if verification fails

            error_msg_prefix = "Failed to send SMTP verification" if sample_type == "smtp_verify" else "Failed to send template test"
            error_msg = f"{error_msg_prefix} email. "
            if auth_error: error_msg += f"Authentication/Connection Error: {failed_data[0][2]}"
            elif send_failure: error_msg += f"Could not send to {failed_data[0][1]}: {failed_data[0][2]}"
            else: error_msg += "Unknown error during sample send."

            self.log_message(error_msg, "error")
            self.status_label.setText(f"Status: {sample_type.replace('_', ' ').title()} Failed. Bulk send disabled.")
            QMessageBox.critical(self, "Sample Mail Failed", error_msg + "\nPlease check settings and details.")
        
        self.email_sender_thread = None


    def handle_bulk_mail_result(self, failed_data_from_thread):
        # ... (identical to previous version, but retry button enabling now also depends on self.settings_verified_for_bulk)
        try:
            final_sent = int(re.search(r'\d+', self.sent_label.text()).group(0))
            final_failed = int(re.search(r'\d+', self.failed_label.text()).group(0))
        except AttributeError:
            final_sent, final_failed = 0,0
            self.log_message("Warning: Could not parse sent/failed counts from labels.", "warning")

        self.status_label.setText(f"Status: Bulk Send Completed. Sent: {final_sent}, Failed: {final_failed}")
        self.log_message(f"Bulk email process finished. Successfully sent: {final_sent}, Failed: {final_failed}")

        for original_idx, email, reason in failed_data_from_thread:
            if original_idx != -1:
                if not any(existing_item[0] == original_idx for existing_item in self.all_failed_data):
                    self.all_failed_data.append((original_idx, email, reason))

        if self.all_failed_data:
            # Enable retry only if settings are currently considered verified AND there are failed items
            self.retry_button.setEnabled(self.settings_verified_for_bulk and bool(self.all_failed_data))
            self.log_message(f"{len(self.all_failed_data)} email(s) in retry queue.", "warning")
        else:
            self.retry_button.setEnabled(False)
            if final_sent > 0 and final_failed == 0 :
                 self.log_message("All emails in this batch sent successfully.", "info")

        self.send_button.setEnabled(self.settings_verified_for_bulk)
        self.email_sender_thread = None


    def start_sending_emails(self):
        if not self.settings_verified_for_bulk: # MODIFIED check
            QMessageBox.warning(self, "Verification Required", "Please send a successful test mail (either SMTP Verify or Template Test) first.")
            return
        # ... (rest of the method is identical)
        if self.df is None:
            QMessageBox.warning(self, "Input Error", "Please load an Excel file first.")
            return
        email_column = self.email_column_combo.currentText()
        if not email_column:
            QMessageBox.warning(self, "Input Error", "Please select the column containing email addresses.")
            return
        if self.df.empty:
            QMessageBox.information(self, "No Data", "The Excel sheet is empty or has no data to process.")
            return
        self.all_failed_data = []
        self.retry_button.setEnabled(False)
        self._start_email_thread(self.df.copy(), email_column, is_retry=False)


    def retry_failed_emails(self):
        if not self.settings_verified_for_bulk: # MODIFIED check
            QMessageBox.warning(self, "Verification Required", "Settings may have changed. Please send a successful test mail before retrying.")
            return
        # ... (rest of the method is identical)
        if not self.all_failed_data:
            QMessageBox.information(self, "No Failures", "There are no emails marked as failed to retry.")
            return
        if self.df is None:
            QMessageBox.warning(self, "Error", "Original Excel data not loaded. Cannot retry.")
            return
        email_column = self.email_column_combo.currentText()
        if not email_column:
            QMessageBox.warning(self, "Error", "Email column not selected. Cannot retry.")
            return
        failed_indices = [item[0] for item in self.all_failed_data]
        df_to_retry = self.df.loc[failed_indices].copy()
        self.all_failed_data = []
        self.retry_button.setEnabled(False)
        self._start_email_thread(df_to_retry, email_column, is_retry=True)

    def closeEvent(self, event):
        # ... (identical to previous version)
        if self.email_sender_thread and self.email_sender_thread.isRunning():
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