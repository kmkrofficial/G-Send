# G-Send: Bulk Gmail Sender

**G-Send** is a Python desktop application built with the PyQt5 framework that allows you to send personalized bulk emails using recipient data from an Excel sheet. It supports Gmail (via App Passwords for security), customizable email templates with Excel column placeholders, and multiple attachments. The application provides live sending statistics and a retry option for failed emails.

![G-Send Screenshot](placeholder_screenshot.png)
*(**Note**: Replace `placeholder_screenshot.png` with an actual screenshot of your application. You can take one and add it to your repository.)*

## Features

*   **Excel Integration:** Load recipient data directly from `.xlsx` or `.xls` files.
*   **Email Column Selection:** Choose which column in your Excel sheet contains the email addresses.
*   **Gmail Support:** Securely send emails via Gmail using App Passwords (2-Step Verification highly recommended).
*   **Customizable Templates:**
    *   Personalize email subjects and bodies using placeholders like `{{ ColumnName }}` that map to your Excel column headers.
    *   Whitespace around column names in placeholders (e.g., `{{  ColumnName  }}`) is handled.
*   **Multiple Attachments:** Attach one or more files to all outgoing emails.
*   **Live Statistics:**
    *   Number of successfully sent emails.
    *   Number of failed emails.
    *   Estimated time of completion (ETA).
*   **Retry Mechanism:** Option to retry sending emails only to recipients who failed in a previous attempt.
*   **Pre-Send Verification:**
    *   **SMTP Verification:** Send a test email to your own address to confirm SMTP settings and credentials.
    *   **Template Test:** Send a test email to your own address using the current subject/body templates and attachments to preview the final email.
    *   Bulk sending is enabled only after a successful test send from *either* test button.
*   **User-Friendly GUI:** Built with PyQt5 for a responsive desktop experience.
*   **Informative Tooltips:** Help popups for most fields and buttons.
*   **Logging:** View a live log of sending activities and errors.

## Prerequisites

Before you begin, ensure you have the following installed:

1.  **Python 3.7+:** Download Python
2.  **PIP:** (Usually comes with Python)
3.  **Git:** (Optional, for cloning the repository)

## Setup and Installation

1.  **Clone the Repository (Optional):**
    If you have Git, clone the repository:
    ```bash
    git clone https://github.com/your-username/g-send.git
    cd g-send
    ```
    Otherwise, download the source code ZIP and extract it.

2.  **Create a Virtual Environment (Recommended):**
    It's good practice to create a virtual environment for Python projects to manage dependencies.
    ```bash
    python -m venv venv
    ```
    Activate the virtual environment:
    *   On Windows:
        ```bash
        .\venv\Scripts\activate
        ```
    *   On macOS/Linux:
        ```bash
        source venv/bin/activate
        ```

3.  **Install Dependencies:**
    Install the required Python libraries using pip:
    ```bash
    pip install PyQt5 pandas openpyxl
    ```

4.  **Generate a Gmail App Password:**
    For security, this application uses Gmail App Passwords, not your regular Google account password. This is mandatory if you have 2-Step Verification enabled on your Gmail account (which is highly recommended).
    *   Go to your Google Account: [https://myaccount.google.com/](https://myaccount.google.com/)
    *   Navigate to **Security**.
    *   Under "Signing in to Google," find **"App passwords"**.
        *   If you don't see "App passwords," ensure 2-Step Verification is enabled for your account.
    *   Click **"Select app"** and choose **"Mail"**.
    *   Click **"Select device"** and choose **"Other (Custom name)"**.
    *   Give it a name (e.g., "G-Send Application").
    *   Click **"Generate"**.
    *   Google will display a **16-character password**. **Copy this password immediately and store it securely.** This is the password you will use in the application.

## How to Use G-Send

1.  **Run the Application:**
    Navigate to the project directory in your terminal (ensure your virtual environment is activated) and run:
    ```bash
    python bulk_mailer_app.py
    ```
    (Replace `bulk_mailer_app.py` with the actual name of your main Python script if different, e.g., `g_send_app.py`.)

2.  **Load Excel File:**
    *   Click the "**Browse Excel File**" button.
    *   Select your `.xlsx` or `.xls` file containing recipient data.
    *   **Excel Format:** Ensure your Excel file has clear column headers (e.g., `Name`, `Email`, `Product ID`, `City`). One column must contain the email addresses. Leading/trailing spaces in column headers will be automatically stripped by the application.

3.  **Select Email Column:**
    *   Once the Excel file is loaded, the "**Email Column**" dropdown will populate with your Excel's column headers.
    *   Select the column that contains the recipient email addresses.

4.  **Enter Gmail SMTP Settings:**
    *   **Your Gmail Address:** Enter your full Gmail address (e.g., `your.email@gmail.com`).
    *   **Gmail App Password:** Enter the 16-character **App Password** you generated in the setup steps. **Do not use your regular Gmail password.**

5.  **Add Attachments (Optional):**
    *   Click "**Add Attachment(s)**" to select one or more files to be attached to every email.
    *   Added attachments will appear in the list.
    *   Click "**Clear Attachments**" to remove all listed attachments.

6.  **Create Email Template:**
    *   **Subject:** Enter the subject line for your emails.
    *   **Body:** Enter the main content of your email.
    *   **Placeholders:** Use `{{ ColumnName }}` to insert data from your Excel sheet. Replace `ColumnName` with the exact (case-sensitive after stripping whitespace) column header from your Excel file.
        *   Example Subject: `Order Confirmation for {{ Name }}`
        *   Example Body:
            ```
            Dear {{ FirstName }},

            Thank you for your order #{{ OrderID }}.
            It will be shipped to {{ City }}.

            Best regards,
            Your Company
            ```

7.  **Verify Settings & Test Template:**
    *   **Button 1: "Verify SMTP & Send Test Mail"**:
        *   Click this button to send a test email to **your own Gmail address** (the one entered in "Your Gmail Address").
        *   This test uses your current SMTP credentials, the defined email template (subject/body), and any added attachments.
    *   **Button: "Test Email Template"** (below the email body input):
        *   Click this button to send another test email to **your own Gmail address**.
        *   This test also uses current SMTP credentials, template, and attachments. It's useful for quickly previewing how your email content (with placeholders) will look.

    **Important:** You must get a "successful" message from at least **one** of these test sends before the "Send Bulk Emails" button is enabled. A successful test indicates your SMTP settings and template configuration are likely correct. If you change SMTP settings, the Excel file, attachments, or the email template, you'll need to re-verify by sending another successful test mail.

8.  **Send Bulk Emails:**
    *   Once settings are successfully verified (via either test button), the "**2. Send Bulk Emails**" button will be enabled.
    *   Click it to start sending emails to all recipients listed in your Excel sheet.
    *   Observe the live statistics (Sent, Failed, ETA) and the log for progress and any errors.

9.  **Retry Failed Emails:**
    *   If any emails fail during a bulk send, the "**Retry Failed Emails**" button will become active after the process completes (and if your settings are still considered verified).
    *   Click this button to attempt sending emails only to those recipients that failed in the previous attempt.

## Code Explanation

The application consists of two main classes:

*   **`EmailSenderThread(QThread)`:**
    *   Manages the email sending process in a separate thread to keep the GUI responsive.
    *   Connects to Gmail's SMTP server using `smtplib`.
    *   Handles TLS encryption.
    *   Renders email templates by replacing `{{ ColumnName }}` placeholders with data from each row of the Excel sheet.
    *   Attaches files to emails.
    *   Emits signals (`progress_update`, `finished_signal`, `log_signal`) to update the GUI with statistics, completion status, and log messages.
    *   Includes basic error handling for SMTP connection, authentication, and individual email sending.

*   **`BulkEmailerApp(QMainWindow)` (or `GSendApp` if you rename the class):**
    *   Sets up the main application window and all UI elements (input fields, buttons, labels, lists, progress bar) using PyQt5.
    *   Handles user interactions:
        *   Browsing and loading Excel files using `pandas`. (Column headers are stripped of whitespace).
        *   Managing attachment lists.
        *   Initiating test sends and bulk sends.
    *   Manages the state of `settings_verified_for_bulk` which controls whether bulk operations are allowed.
    *   Displays live statistics and log messages received from `EmailSenderThread`.
    *   Provides tooltips for UI elements to guide the user.
    *   Includes logic for retrying failed emails.
    *   Handles graceful exit if the application is closed during an active sending process.

## Building the Executable (EXE for Windows)

You can package G-Send into a standalone executable using **PyInstaller**.

1.  **Install PyInstaller:**
    (If not already done, from your activated virtual environment)
    ```bash
    pip install pyinstaller
    ```

2.  **Create a `.spec` File:**
    Navigate to your project's root directory in the terminal and run:
    ```bash
    pyi-makespec --name GSend --windowed bulk_mailer_app.py
    ```
    (Replace `bulk_mailer_app.py` with your main script's name if different, e.g., `g_send_app.py`.)
    This creates `GSend.spec`.

3.  **Edit `GSend.spec`:**
    Open `GSend.spec` and modify the `Analysis` section, particularly `hiddenimports`, to ensure all necessary modules are included. A good starting point for `hiddenimports` would be:
    ```python
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.offsets',
        'openpyxl',
        'PyQt5.sip',
        # Add more if PyInstaller misses them during build
    ],
    ```
    You can also specify an application icon in the `EXE` section:
    ```python
    exe = EXE(
        # ... other arguments ...
        console=False,  # Important for GUI apps
        icon='path/to/your/gsend_icon.ico' # Optional: provide an .ico file
    )
    ```

4.  **Build the Executable:**
    Run PyInstaller with the spec file:
    ```bash
    pyinstaller GSend.spec
    ```
    For a single-file executable (larger initial startup time, but easier to distribute):
    ```bash
    pyinstaller --onefile GSend.spec
    ```

5.  **Find Your Executable:**
    The executable will be in the `dist` folder (`dist/GSend/GSend.exe` or `dist/GSend.exe` if using `--onefile`).

    *Refer to the PyInstaller documentation for more advanced configurations and troubleshooting.*

## Troubleshooting

*   **Authentication Failed:**
    *   Ensure you are using a **16-character App Password**, not your regular Gmail password.
    *   Double-check that 2-Step Verification is enabled on your Google Account.
    *   Verify your Gmail address is typed correctly.
    *   Check your Gmail for any security alerts from Google regarding blocked sign-in attempts.
*   **`ModuleNotFoundError` after building .exe:** Add the missing module to `hiddenimports` in your `.spec` file and rebuild.
*   **Emails not sending / Placeholders not working:**
    *   Ensure the `{{ ColumnName }}` in your templates exactly matches the column headers in your Excel file (case-sensitive, after stripping leading/trailing whitespace).
    *   Use the "Test Email Template" button to preview. `[MISSING_DATA]` will appear if a placeholder isn't found in the (dummy or actual) data row used for the test.
*   **Attachment Issues:** Ensure the file paths for attachments are correct and the files exist. The log will show warnings for skipped or failed attachments.

## Contributing

Contributions, issues, and feature requests are welcome! Please feel free to:
*   Open an issue to report bugs or suggest features.
*   Fork the repository and submit a pull request with your improvements.

## License

This project is open-source. You can specify a license if you wish (e.g., MIT License). If not specified, it typically falls under standard copyright.