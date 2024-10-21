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
RADICALE_WEBDAV_URL = ""
RADICALE_USERNAME = ""
RADICALE_PASSWORD = ""

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
    
    try:
        # Navigate to the Kronos schedule page
        driver.get(KRONOS_URL)
        logging.info("Navigated to Kronos schedule page.")
        log_page_details(driver)  # Log current page details and capture a screenshot

        # Wait for the schedule list to be present
        schedule_list = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.ID, "my-schedule-list"))
        )
        logging.info("Schedule list found.")

        # Locate all day elements (li elements with class 'withDivider')
        schedule_days = schedule_list.find_elements(By.CSS_SELECTOR, "li.withDivider")
        logging.info(f"Located {len(schedule_days)} schedule day elements.")

        schedule_data = []

        # Iterate over each day to find shifts
        for day in schedule_days:
            try:
                day_date = day.get_attribute("datetime")
                logging.info(f"Processing schedule for date: {day_date}")

                # Locate shift elements within the day
                shift_wrappers = day.find_elements(By.CSS_SELECTOR, "div.scheduleEntityWrapper")
                logging.info(f"Found {len(shift_wrappers)} shifts for date: {day_date}")

                for shift in shift_wrappers:
                    time_range = shift.find_element(By.CSS_SELECTOR, "time.label").text
                    time_range_cleaned = re.sub(r'\s*\[.*?\]', '', time_range)  # Clean time range
                    shift_details = shift.find_element(By.CSS_SELECTOR, "div.details").text
                    logging.info(f"Shift details: {time_range_cleaned}, {shift_details}")

                    schedule_data.append({
                        "date": day_date,
                        "time_range": time_range_cleaned,
                        "details": shift_details
                    })
            except Exception as e:
                logging.error(f"Error scraping shifts for date {day_date}: {str(e)}")
                logging.error(f"Full traceback: {traceback.format_exc()}")
                capture_screenshot(driver, f"error_scraping_shifts_{day_date}")

        logging.info(f"Schedule data successfully scraped: {schedule_data}")
        return schedule_data

    except TimeoutException as e:
        logging.error("Timeout while waiting for schedule elements.")
        capture_screenshot(driver, "timeout_error")
        # Save the page source for inspection
        with open("page_source_error.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        logging.error(f"Page source saved to page_source_error.html")
        logging.error(f"Full traceback: {traceback.format_exc()}")
        return []

    except Exception as e:
        logging.error(f"An error occurred while scraping the schedule: {str(e)}")
        logging.error(f"Full traceback: {traceback.format_exc()}")
        capture_screenshot(driver, "scrape_error")
        return []
    
    return schedule_data

def retrieve_existing_events():
    """Retrieve all existing events from the Radicale calendar."""
    response = requests.get(RADICALE_WEBDAV_URL, auth=(RADICALE_USERNAME, RADICALE_PASSWORD))
    
    if response.status_code in [200, 207]:  # Handle both 200 OK and 207 Multi-Status
        existing_events = {}
        if response.status_code == 200:
            logging.info("200 OK response received. Parsing calendar data.")
            cal = Calendar.from_ical(response.content)
            for component in cal.walk():
                if component.name == "VEVENT":
                    dtstart = component.get('dtstart').dt
                    dtend = component.get('dtend').dt
                    summary = component.get('summary')
                    event_key = (dtstart, dtend, summary)
                    existing_events[event_key] = component.to_ical().decode('utf-8')
        else:
            logging.info("207 Multi-Status response received. Parsing WebDAV data.")
            tree = response.content.decode('utf-8')
            events = re.findall(r'<href>(.*?)</href>', tree)
            for event in events:
                event_url = event.strip()
                event_response = requests.get(RADICALE_WEBDAV_URL + event_url, auth=(RADICALE_USERNAME, RADICALE_PASSWORD))
                if event_response.status_code == 200:
                    existing_event = Calendar.from_ical(event_response.content)
                    for component in existing_event.walk('VEVENT'):
                        dtstart = component.get('dtstart').dt
                        dtend = component.get('dtend').dt
                        summary = component.get('summary')
                        event_key = (dtstart, dtend, summary)
                        existing_events[event_key] = component.to_ical().decode('utf-8')
        logging.info(f"Retrieved {len(existing_events)} existing events from Radicale.")
        return existing_events
    else:
        logging.error(f"Failed to retrieve existing events: {response.status_code} - {response.text}")
        return {}

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
    
    unique_uid = f"{uuid.uuid4()}@mydomain.com"
    logging.info(f"Generated UID: {unique_uid} for event on {date_str}")

    event.add('summary', details)
    event.add('dtstart', vDatetime(start_time))
    event.add('dtend', vDatetime(end_time))
    event.add('uid', unique_uid)  # Ensure unique UID for each event
    event.add('dtstamp', vDatetime(datetime.utcnow()))  # Add timestamp for event creation
    
    return event

def create_individual_ics_files(schedule_data):
    """Generate individual iCalendar (.ics) files for each event in schedule data."""
    ics_filenames = []

    for entry in schedule_data:
        event = create_icalendar_event(entry['date'], entry['time_range'], entry['details'])
        cal = Calendar()
        cal.add_component(event)
        
        uid = event.get('uid')
        ics_filename = f"{uid}.ics"
        ics_filepath = os.path.join('individual_events', ics_filename)

        with open(ics_filepath, 'wb') as f:
            f.write(cal.to_ical())
        
        ics_filenames.append(ics_filepath)
        logging.info(f"Generated iCalendar file: {ics_filepath}")
    
    return ics_filenames

def compare_and_handle_existing(events_to_upload, existing_events):
    """Compare newly generated events with existing ones on Radicale and delete the local file if a match is found."""
    for event_filename in events_to_upload:
        with open(event_filename, 'rb') as f:
            new_event = Calendar.from_ical(f.read())
        
        for new_event_component in new_event.walk('VEVENT'):
            new_event_dtstart = new_event_component.get('dtstart').dt
            new_event_dtend = new_event_component.get('dtend').dt
            new_event_summary = new_event_component.get('summary')

            logging.info(f"Checking new event with start time {new_event_dtstart} and end time {new_event_dtend} against existing events on Radicale.")

            for event_key in existing_events.keys():
                existing_event_dtstart, existing_event_dtend, existing_event_summary = event_key

                if (new_event_dtstart == existing_event_dtstart and
                    new_event_dtend == existing_event_dtend and
                    new_event_summary == existing_event_summary):
                    
                    logging.info(f"Event with start time {new_event_dtstart} and end time {new_event_dtend} already exists on Radicale. Deleting local file: {event_filename}")
                    os.remove(event_filename)
                    break  # No need to check further if we found a match
            else:
                logging.info(f"No matching event found on Radicale for event starting at {new_event_dtstart}.")

def upload_to_radicale_individual_files(ics_filenames):
    """Upload individual .ics files to Radicale."""
    for ics_file in ics_filenames:
        if os.path.exists(ics_file):  # Check if the file still exists after comparison
            with open(ics_file, 'rb') as f:
                response = requests.put(
                    RADICALE_WEBDAV_URL + os.path.basename(ics_file),
                    data=f,
                    auth=(RADICALE_USERNAME, RADICALE_PASSWORD),
                    headers={"Content-Type": "text/calendar"}
                )

            if response.status_code == 201:
                logging.info(f"Successfully uploaded {ics_file} to Radicale")
            else:
                logging.error(f"Failed to upload {ics_file} to Radicale: {response.status_code} - {response.text}")

def main():
    # Setup the WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run Chrome in headless mode
    options.add_argument("--no-sandbox")  # Bypass OS security model, useful in Docker
    options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
    options.add_argument("--disable-gpu")  # Disable GPU acceleration
    options.add_argument("--window-size=1920,1080")  # Set window size to avoid rendering issues

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    
    try:
        # Retrieve existing events from Radicale
        existing_events = retrieve_existing_events()        
        
        login_to_microsoft(driver)
        schedule_data = scrape_schedule(driver)

        # Ensure the directory exists
        os.makedirs('individual_events', exist_ok=True)

        # Generate new .ics files
        ics_filenames = create_individual_ics_files(schedule_data)

        # Compare and handle existing events on Radicale
        compare_and_handle_existing(ics_filenames, existing_events)

        # Upload new events to Radicale
        upload_to_radicale_individual_files(ics_filenames)

    except Exception as e:
        logging.error(f"An error occurred: {e}")
    
    finally:
        driver.quit()
        logging.info("Browser closed")

if __name__ == "__main__":
    main()
