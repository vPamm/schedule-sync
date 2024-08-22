import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import json
import os
from datetime import datetime, timedelta
import requests
import re
import uuid
from icalendar import Calendar, Event, vDatetime
from pytz import timezone
import random
import string

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("schedule_sync.log"),
        logging.StreamHandler()
    ]
)

# Constants
MICROSOFT_LOGIN_URL = ""
KRONOS_URL = ""
USERNAME = ""
PASSWORD = ""
TOTP_SECRET = ""
NEXTCLOUD_WEBDAV_URL = ""
NEXTCLOUD_USERNAME = ""
NEXTCLOUD_PASSWORD = ""

def capture_screenshot(driver, name):
    """Capture a screenshot for debugging."""
    screenshot_dir = "screenshots"
    os.makedirs(screenshot_dir, exist_ok=True)
    screenshot_path = os.path.join(screenshot_dir, f"{name}.png")
    driver.save_screenshot(screenshot_path)
    logging.info(f"Screenshot saved to {screenshot_path}")

def log_page_details(driver):
    """Log current page details for debugging."""
    logging.info(f"Current URL: {driver.current_url}")
    logging.info(f"Page Title: {driver.title}")
    capture_screenshot(driver, "current_page")

def safe_click(driver, by, value, retries=5):
    for attempt in range(retries):
        try:
            logging.info(f"Attempting to click element with {by}='{value}', attempt {attempt + 1} of {retries}")
            element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((by, value))
            )
            driver.execute_script("arguments[0].click();", element)
            logging.info(f"Successfully clicked element with {by}='{value}'")
            return
        except Exception as e:
            logging.warning(f"Retrying click due to: {e} (Attempt {attempt + 1} of {retries})")
            time.sleep(10)  # Increased wait between retries
            logging.info(f"Current URL: {driver.current_url}")
    raise Exception(f"Failed to click element with {by}='{value}' after {retries} retries")

def login_to_microsoft(driver):
    logging.info("Starting Microsoft login process")
    driver.get(MICROSOFT_LOGIN_URL)
    time.sleep(5)  # Allow the page to load fully
    
    # Enter the username and submit
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "loginfmt"))).send_keys(USERNAME)
    safe_click(driver, By.ID, "idSIButton9")

    # Wait for the password field to be present and visible before inputting password
    password_field = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.NAME, "passwd"))
    )
    
    # Explicitly focus on the password field and input the password
    password_field.click()
    password_field.send_keys(PASSWORD)
    
    logging.info("Password entered successfully")
    safe_click(driver, By.ID, "idSIButton9")
    
    # Generate the TOTP
    totp = pyotp.TOTP(TOTP_SECRET)
    code = totp.now()
    
    # Enter the TOTP code
    totp_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "otc")))
    totp_field.send_keys(code)
    
    logging.info("TOTP code entered successfully")
    
    # Wait for and click the submit button
    safe_click(driver, By.ID, "idSubmit_SAOTCC_Continue")
    
    # Wait for possible redirection to Kronos
    try:
        WebDriverWait(driver, 10).until(EC.url_contains(KRONOS_URL))
        logging.info("Successfully redirected to Kronos")
    except Exception as e:
        logging.warning(f"Redirection to Kronos failed: {e}")
        # Fallback: Manually navigate to the Kronos URL
        logging.info("Attempting manual navigation to Kronos")
        driver.get(KRONOS_URL)
        time.sleep(10)  # Give the page extra time to load
        log_page_details(driver)
        # Verify if we are on the Kronos site
        if KRONOS_URL in driver.current_url:
            logging.info("Manual navigation to Kronos successful")
        else:
            logging.error("Manual navigation to Kronos failed. Current URL: " + driver.current_url)
            capture_screenshot(driver, "failed_navigation")

def scrape_schedule(driver):
    logging.info("Starting to scrape the schedule")
    driver.get(KRONOS_URL)
    time.sleep(10)

    schedule_days = WebDriverWait(driver, 40).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.listItem"))
    )
    logging.info(f"Located {len(schedule_days)} schedule day elements.")

    schedule_data = []
    for day in schedule_days:
        day_date = day.get_attribute("datetime").split(' GMT')[0]
        
        shift_wrappers = day.find_elements(By.CSS_SELECTOR, "div.scheduleEntity.interactive.shift-wrapper")
        
        for shift in shift_wrappers:
            time_range = shift.find_element(By.CSS_SELECTOR, "time.label").text
            time_range_cleaned = re.sub(r'\s*\[.*?\]', '', time_range)  # Remove anything in square brackets
            shift_details = shift.find_element(By.CSS_SELECTOR, "div.details").text
            schedule_data.append({
                "date": day_date,
                "time_range": time_range_cleaned,
                "details": shift_details
            })
    logging.info(f"Schedule data scraped: {schedule_data}")
    
    return schedule_data

def create_icalendar_event(date_str, time_range, details):
    """Create an iCalendar event."""
    event = Event()

    date_part, time_part = date_str.split(' '), time_range.split('-')
    
    date_formatted = f"{date_part[0]} {date_part[1]} {date_part[2]} {date_part[3]}"
    
    start_time_str = f"{date_formatted} {time_part[0].strip()}"
    end_time_str = f"{date_formatted} {time_part[1].strip()}"

    # Assuming the times are in local time, e.g., US/Eastern
    local_tz = timezone('US/Eastern')
    start_time = local_tz.localize(datetime.strptime(start_time_str, "%a %b %d %Y %I:%M %p"))
    end_time = local_tz.localize(datetime.strptime(end_time_str, "%a %b %d %Y %I:%M %p"))
    
    timestamp = int(time.time())
    unique_uid = f"{uuid.uuid4()}-{timestamp}-{uuid.uuid4()}@mydomain.com"
    logging.info(f"Generated UID: {unique_uid} for event on {date_str}")

    event.add('summary', details)
    event.add('dtstart', vDatetime(start_time))
    event.add('dtend', vDatetime(end_time))
    event.add('uid', unique_uid)  # Ensure unique UID for each event
    event.add('dtstamp', vDatetime(datetime.utcnow()))  # Add timestamp for event creation
    
    return event

def create_individual_ics_files(schedule_data):
    """Generate individual iCalendar (.ics) files for each event in schedule data."""
    cal_base = Calendar()
    cal_base.add('prodid', '-//My Calendar Application//EN')
    cal_base.add('version', '2.0')
    cal_base.add('calscale', 'GREGORIAN')

    ics_filenames = []

    for entry in schedule_data:
        event = create_icalendar_event(entry['date'], entry['time_range'], entry['details'])
        cal = Calendar()
        for component in cal_base.subcomponents:
            cal.add_component(component)
        cal.add_component(event)
        
        uid = event.get('uid')
        ics_filename = f"{uid}.ics"
        ics_filepath = os.path.join('individual_events', ics_filename)

        with open(ics_filepath, 'wb') as f:
            f.write(cal.to_ical())
        
        ics_filenames.append(ics_filepath)
        logging.info(f"Generated iCalendar file: {ics_filepath}")
    
    return ics_filenames

def upload_to_nextcloud_individual_files(ics_filenames):
    """Upload individual .ics files to Nextcloud."""
    for ics_file in ics_filenames:
        with open(ics_file, 'rb') as f:
            response = requests.put(
                NEXTCLOUD_WEBDAV_URL + os.path.basename(ics_file),
                data=f,
                auth=(NEXTCLOUD_USERNAME, NEXTCLOUD_PASSWORD),
                headers={"Content-Type": "text/calendar"}
            )

        if response.status_code == 201:
            logging.info(f"Successfully uploaded {ics_file} to Nextcloud")
        else:
            logging.error(f"Failed to upload {ics_file} to Nextcloud: {response.status_code} - {response.text}")

def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run Chrome in headless mode
    options.add_argument("--no-sandbox")  # Bypass OS security model, useful in Docker
    options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
    options.add_argument("--disable-gpu")  # Disable GPU acceleration
    options.add_argument("--window-size=1920,1080")  # Set window size to avoid rendering issues

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    
    try:
        login_to_microsoft(driver)
        schedule_data = scrape_schedule(driver)
        ics_filenames = create_individual_ics_files(schedule_data)
        upload_to_nextcloud_individual_files(ics_filenames)
        
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    
    finally:
        driver.quit()
        logging.info("Browser closed")

if __name__ == "__main__":
    # Ensure the directory exists
    os.makedirs('individual_events', exist_ok=True)
    main()