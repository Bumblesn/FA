import schedule
import time
import subprocess
import json
import logging
from datetime import datetime
from send_emails import send_failure_email
import os
import pytz
LOCAL_TZ = pytz.timezone("Asia/Kolkata")

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
    max_retries = config.get("scheduler_max_retries", 5)
    retry_delay = config.get("scheduler_retry_delay", 300)
    admin_email = config.get("admin_email", "")
except Exception as e:
    logger.error(f"Error loading config.json: {str(e)}")
    max_retries = 5
    retry_delay = 300
    admin_email = ""

def run_zone_reports():
    # Skip Sunday (weekday 6)
    # if datetime.now().weekday() == 6:
    if datetime.now(LOCAL_TZ).weekday() == 6:
        logger.info("Today is Sunday. Skipping report generation.")
        return
    
    execution_status = "Failed"
    last_attempt = max_retries
    
    for attempt in range(1, max_retries + 1):
        logger.info(f"Attempt {attempt}/{max_retries}: Running zone_wise_reports.py...")
        try:
            result = subprocess.run(["python", "zone_wise_reports.py"], capture_output=True, text=True)
            logger.info("Script output: %s", result.stdout)
            if result.stderr:
                logger.error("Script errors: %s", result.stderr)
            
            if result.returncode == 0:
                logger.info(f"Script completed successfully on attempt {attempt}.")
                execution_status = "Success"
                last_attempt = attempt
                break
            else:
                logger.error(f"Script failed with exit code {result.returncode}.")
        
        except Exception as e:
            logger.error(f"Exception running script on attempt {attempt}/{max_retries}: {str(e)}")
        
        if attempt < max_retries:
            logger.info(f"Retrying in {retry_delay // 60} minutes...")
            time.sleep(retry_delay)
    
    if execution_status == "Failed":
        logger.error(f"Failed to run zone_wise_reports.py after {max_retries} attempts.")
        if admin_email:
            try:
                send_failure_email("Scheduler Failure", f"Failed to run zone_wise_reports.py after {max_retries} attempts.", [admin_email])
                logger.info("Sent failure notification to admin.")
            except Exception as e:
                logger.error(f"Failed to send failure email: {str(e)}")
    
    # Log last execution time and status
    #logger.info(f"Last execution completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {execution_status} on attempt {last_attempt}")
    logger.info(f"Last execution completed at {datetime.now(LOCAL_TZ).strftime('%Y-%m-%d %H:%M:%S IST')} - {execution_status} on attempt {last_attempt}")

# Schedule the script
# ## schedule.every().day.at(config.get("schedule_time", "23:38")).do(run_zone_reports)
local_run_time = config.get("schedule_time", "08:00")
h, m = map(int, local_run_time.split(":"))
local_dt = LOCAL_TZ.localize(datetime.now().replace(hour=h, minute=m, second=0))
utc_run_time = local_dt.astimezone(pytz.utc).strftime("%H:%M")

schedule.every().day.at(utc_run_time).do(run_zone_reports)
logger.info(f"Scheduler set for {local_run_time} IST ({utc_run_time} UTC)")

logger.info("Scheduler started. Waiting for daily execution (skipping Sundays)...")

# Keep the scheduler running
while True:
    schedule.run_pending()
    time.sleep(60)
    

run_zone_reports()
