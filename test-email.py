import smtplib

# --- Configuration ---
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587 # Port for TLS/STARTTLS

def verify_gmail_credentials(email_address, app_password):
    """
    Attempts to log in to Gmail's SMTP server using the provided credentials.
    Returns True if successful, False otherwise, along with an error message.
    """
    try:
        print(f"Attempting to connect to {SMTP_SERVER} on port {SMTP_PORT}...")
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.set_debuglevel(0) # Set to 1 for more verbose output from smtplib

        print("Sending EHLO...")
        server.ehlo()

        print("Starting TLS...")
        server.starttls()

        print("Sending EHLO again (post-TLS)...")
        server.ehlo()

        print(f"Attempting to login with username: {email_address}...")
        server.login(email_address, app_password)

        print("Login successful!")
        server.quit()
        return True, "Authentication successful."

    except smtplib.SMTPAuthenticationError as e:
        error_message = f"SMTP Authentication Error: {e.code} - {e.smtp_error.decode() if e.smtp_error else 'No specific error message.'}"
        print(f"ERROR: {error_message}")
        return False, error_message
    except smtplib.SMTPServerDisconnected:
        error_message = "SMTPServerDisconnected: Server unexpectedly disconnected. Check network or firewall."
        print(f"ERROR: {error_message}")
        return False, error_message
    except smtplib.SMTPConnectError as e:
        error_message = f"SMTPConnectError: Could not connect to server. {e}"
        print(f"ERROR: {error_message}")
        return False, error_message
    except ConnectionRefusedError:
        error_message = "ConnectionRefusedError: The server at {SMTP_SERVER}:{SMTP_PORT} refused the connection. Check server address, port, and firewall."
        print(f"ERROR: {error_message}")
        return False, error_message
    except Exception as e:
        error_message = f"An unexpected error occurred: {str(e)}"
        print(f"ERROR: {error_message}")
        return False, error_message
    finally:
        # Ensure server connection is closed if it was established
        if 'server' in locals() and server.sock:
            try:
                server.quit()
            except Exception:
                pass # Ignore errors on quit if connection was already problematic

if __name__ == "__main__":
    print("--- Gmail SMTP Authentication Test ---")
    user_email = input("Enter your Gmail address: ").strip()
    user_app_password = input("Enter your Gmail App Password (16 characters, no spaces): ").strip()

    if not user_email or "@" not in user_email:
        print("Invalid email address format.")
    elif not user_app_password or len(user_app_password) != 16 or " " in user_app_password:
        print("App Password should be 16 characters long and contain no spaces.")
    else:
        print("\nVerifying...")
        success, message = verify_gmail_credentials(user_email, user_app_password)

        print("\n--- Result ---")
        if success:
            print("✅ Authentication SUCCESSFUL!")
            print("Your Gmail address and App Password are correct and can connect to Gmail's SMTP server.")
        else:
            print("❌ Authentication FAILED.")
            print(f"   Reason: {message}")
            print("\nPlease double-check:")
            print("  1. You are using an APP PASSWORD (not your regular Google password).")
            print("  2. The App Password was copied correctly (16 characters, no spaces).")
            print("  3. Your Gmail address is typed correctly.")
            print("  4. 2-Step Verification is enabled on your Google Account.")
            print("  5. Check for Google security alerts or try the 'Unlock Captcha' link if problems persist.")