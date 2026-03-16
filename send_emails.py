import smtplib
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
    # Read Excel and convert to HTML table
    try:
        # df = pd.read_excel(filename)
        # df = df.fillna("")  # Replace NaN with blank
        # html_table = df.to_html(index=False, border=0)

        df = pd.read_excel(filename)
        df = df.fillna("")

        headers = df.columns.tolist()
        header_html = "".join(
            f'<th style="background-color:#f2f2f2;color:#000;text-align:left;padding:10px;border:1px solid #ddd;">{h}</th>'
            for h in headers
        )
        rows_html = ""
        for i, (_, row) in enumerate(df.iterrows()):
            region_val = str(row.iloc[2]) if len(row) > 2 else ""
            is_subtotal = region_val.endswith(" Total")
            if is_subtotal:
                row_style = 'background-color:#707070;border-top:2px solid #555;border-bottom:2px solid #555;'
                cells = ""
                for j, v in enumerate(row):
                    if j == 2:
                        cells += f'<td style="padding:10px;border:1px solid #888;font-weight:bold;font-style:italic;color:#ffffff;white-space:nowrap;">{v}</td>'
                    elif str(v) != "" and str(v) != "0":
                        cells += f'<td style="padding:10px;border:1px solid #888;font-weight:bold;color:#ffffff;text-align:right;">{v}</td>'
                    else:
                        cells += f'<td style="padding:10px;border:1px solid #888;color:#ffffff;"></td>'
            else:
                row_bg = '#f9f9f9' if i % 2 == 0 else '#ffffff'
                row_style = f'background-color:{row_bg};'
                cells = "".join(
                    f'<td style="padding:8px;border:1px solid #ddd;vertical-align:top;color:#333;">{v}</td>'
                    for v in row
                )
            rows_html += f'<tr style="{row_style}">{cells}</tr>\n'
        html_table = f"""
        <table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:13px;min-width:800px;border:1px solid #ddd;">
            <thead><tr>{header_html}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        """
        body_html = f"""
        <html>
        <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                font-size: 14px;
                color: #333;
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
        return


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
