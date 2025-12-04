
import smtplib
from smtplib import SMTPAuthenticationError, SMTPException
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import json
import logging
from datetime import datetime
import pandas as pd  # Added to read Excel and convert to HTML

# Set up logging
log_dir = "logs"
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(log_dir, f"{datetime.now().strftime('%Y%m%d')}.log")),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

# Load configuration (from config.json)
try:
    with open("config.json", "r") as f:
        config = json.load(f)
    SMTP_SERVER = config.get("smtp_server", "smtp.gmail.com")
    SMTP_PORT = config.get("smtp_port", 587)
    SENDER_EMAIL = config.get("sender_email", "your_email@gmail.com")
    SENDER_PASSWORD = config.get("sender_password", "your_app_password")
    # Optional admin / alert settings (safe to leave empty)
    ADMIN_ALERT_RECIPIENTS = config.get("admin_email", []) or []
    ALERT_EMAIL = config.get("alert_email", "")  # optional: backup/alert sender email
    ALERT_PASSWORD = config.get("alert_password", "")  # optional: backup/alert password
    ALERT_SMTP_SERVER = config.get("alert_smtp_server", SMTP_SERVER)
    ALERT_SMTP_PORT = config.get("alert_smtp_port", SMTP_PORT)
    ALERT_WEBHOOK_URL = config.get("alert_webhook_url", "")
except Exception as e:
    logger.error(f"Error loading config.json: {str(e)}")
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = "your_email@gmail.com"
    SENDER_PASSWORD = "your_app_password"
    ADMIN_ALERT_RECIPIENTS = []
    ALERT_EMAIL = ""
    ALERT_PASSWORD = ""
    ALERT_SMTP_SERVER = SMTP_SERVER
    ALERT_SMTP_PORT = SMTP_PORT
    ALERT_WEBHOOK_URL = ""

# Helper to connect & login to SMTP
def _smtp_connect_and_login(host, port, username, password, use_starttls=True, timeout=30):
    server = smtplib.SMTP(host, port, timeout=timeout)
    server.ehlo()
    if use_starttls:
        server.starttls()
        server.ehlo()
    if username:
        server.login(username, password)
    return server

# Notification function for SMTP auth failure (non-blocking, safe)
def notify_admin_bad_credentials(error_msg):
    """
    Attempt to notify admin about SMTP authentication failure.
    Preference:
      1) Use ALERT_EMAIL + ALERT_PASSWORD if provided (separate SMTP credentials)
      2) Use ALERT_WEBHOOK_URL if provided
      3) Else, log & print a clear CRITICAL message
    """
    ts = datetime.utcnow().isoformat() + "Z"
    subject = f"URGENT: SMTP Authentication Failure - Manager Working Report ({ts})"
    body = (
        "Hi Team,\n\n"
        "The Manager Working Report automation detected an SMTP authentication failure while attempting to send zone reports.\n\n"
        f"Error: {error_msg}\n\n"
        "Action required: Please verify the primary SMTP credentials (sender_email / sender_password) in config.json."
        " If you want automatic email alerts when primary SMTP fails, add alert_email and alert_password to config.json.\n\n"
        f"Timestamp (UTC): {ts}\n\n"
        "Regards,\nAutomation"
    )

    # If there are no configured admin recipients, log and print and return
    if not ADMIN_ALERT_RECIPIENTS:
        logger.error("No admin_email configured in config.json; cannot notify admin by email. Auth error: %s", error_msg)
        print("CRITICAL: SMTP auth failed and no admin_email configured. See logs for details.")
        print(body)
        return False

    # 1) Try alert SMTP if configured
    if ALERT_EMAIL and ALERT_PASSWORD:
        try:
            msg = MIMEMultipart()
            msg["From"] = ALERT_EMAIL
            msg["To"] = ", ".join(ADMIN_ALERT_RECIPIENTS)
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            server = _smtp_connect_and_login(ALERT_SMTP_SERVER, ALERT_SMTP_PORT, ALERT_EMAIL, ALERT_PASSWORD)
            server.sendmail(ALERT_EMAIL, ADMIN_ALERT_RECIPIENTS, msg.as_string())
            server.quit()
            logger.info("Admin notified about SMTP auth failure via alert SMTP.")
            return True
        except SMTPException as e:
            logger.error("Alert SMTP failed: %s", e)
        except Exception as e:
            logger.error("Unexpected error while using alert SMTP: %s", e)

    # 2) Try webhook if provided
    if ALERT_WEBHOOK_URL:
        try:
            import requests
            payload = {"text": f"*{subject}*\n```\n{body}\n```"}
            resp = requests.post(ALERT_WEBHOOK_URL, json=payload, timeout=10)
            if resp.status_code in (200, 204):
                logger.info("Admin notified about SMTP auth failure via webhook.")
                return True
            else:
                logger.error("Webhook notify failed (%s): %s", resp.status_code, getattr(resp, "text", ""))
        except Exception as e:
            logger.error("Webhook notification error: %s", e)

    # 3) Fallback: log + print
    logger.error("Failed to send admin notification via alert SMTP/webhook. Please update config.json. Auth error: %s", error_msg)
    print("CRITICAL: SMTP authentication failed. Update sender_password in config.json (or configure alert_email/alert_password).")
    print(body)
    return False

# Function to send email with Excel as body and attachment
def send_email(zone_name, filename, to_recipients, cc_recipients=None):
    if not to_recipients:
        logger.warning(f"No 'To' email recipients for Zone: {zone_name}. Skipping email.")
        return

    cc_recipients = cc_recipients or []

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(to_recipients)
    if cc_recipients:
        msg["Cc"] = ", ".join(cc_recipients)

    subject_date = os.path.basename(filename).split('_')[-1].replace('.xlsx', '')
    msg["Subject"] = f"Wipro Lighting <> {zone_name} Manager Working Report - {subject_date}"

    # Read Excel and convert to HTML table
    try:
        df = pd.read_excel(filename)
        df = df.fillna("")  # Replace NaN with blank
        html_table = df.to_html(index=False, border=0)
        body_html = f"""
        <html>
        <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                font-size: 14px;
                color: #333;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
                font-size: 13px;
                min-width: 800px;
                border: 1px solid #ddd;
            }}
            th {{
                background-color: #f2f2f2;
                color: #000000;
                text-align: left;
                padding: 10px;
                border: 1px solid #ddd;
            }}
            td {{
                padding: 8px;
                border: 1px solid #ddd;
                vertical-align: top;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
        </style>
        </head>
        <body>
        <p>Dear Team,</p>
        <p>Please find below the joint working report for <b>{zone_name}</b>:</p>
        {html_table}
        </body>
        </html>
        """

        msg.attach(MIMEText(body_html, "html"))
    except Exception as e:
        logger.error(f"Error reading Excel file for HTML body: {str(e)}")
        # proceed to try attach and send even if HTML body creation fails

    # Attach Excel file
    try:
        with open(filename, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="xlsx")
            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(filename)
            )
            msg.attach(attachment)
    except Exception as e:
        logger.error(f"Error attaching file {filename}: {str(e)}")
        return

    # Send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            all_recipients = to_recipients + cc_recipients
            server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        logger.info(f"Email sent for Zone: {zone_name} to {', '.join(to_recipients)}{', cc: ' + ', '.join(cc_recipients) if cc_recipients else ''}")
    except SMTPAuthenticationError as e:
        logger.error(f"SMTP AUTH ERROR for Zone {zone_name}: {str(e)}")
        # Trigger admin alert (won't use primary broken creds)
        notify_admin_bad_credentials(str(e))
    except Exception as e:
        logger.error(f"Error sending email for Zone: {zone_name}: {str(e)}")

# Function to send failure notification
def send_failure_email(subject, message, to_recipients):
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(to_recipients)
    msg["Subject"] = subject

    msg.attach(MIMEText(message, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_recipients, msg.as_string())
        logger.info(f"Failure email sent to {', '.join(to_recipients)}")
    except SMTPAuthenticationError as e:
        logger.error(f"Primary SMTP auth error while sending failure email: {str(e)}")
        # If primary fails for consolidated failure email, try alerting admin via alert SMTP/webhook
        notify_admin_bad_credentials(str(e))
    except Exception as e:
        logger.error(f"Error sending failure email: {str(e)}")
