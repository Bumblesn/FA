import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import json
import logging
from datetime import datetime

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

# Load configuration
try:
    with open("config.json", "r") as f:
        config = json.load(f)
    SMTP_SERVER = config.get("smtp_server", "smtp.gmail.com")
    SMTP_PORT = config.get("smtp_port", 587)
    SENDER_EMAIL = config.get("sender_email", "your_email@gmail.com")
    SENDER_PASSWORD = config.get("sender_password", "your_app_password")
except Exception as e:
    logger.error(f"Error loading config.json: {str(e)}")
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = "your_email@gmail.com"
    SENDER_PASSWORD = "your_app_password"

# Function to send email with attachment
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
    
    msg["Subject"] = f"Wipro Lighting <> {zone_name} Manager Working Report - {os.path.basename(filename).split('_')[-1].replace('.xlsx', '')}"
    
    body = f"""
    Dear Team,
    
    Please find attached the joint working report {os.path.basename(filename).split('_')[-1].replace('.xlsx', '')}.
    
    Regards,
    Vivek
    """
    msg.attach(MIMEText(body, "plain"))
    
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
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            all_recipients = to_recipients + cc_recipients
            server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        logger.info(f"Email sent for Zone: {zone_name} to {', '.join(to_recipients)}{', cc: ' + ', '.join(cc_recipients) if cc_recipients else ''}")
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
    except Exception as e:
        logger.error(f"Error sending failure email: {str(e)}")