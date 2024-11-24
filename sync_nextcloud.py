import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import os
import re
from datetime import datetime
from icalendar import Calendar, Event, vDatetime
from pytz import timezone
import uuid
from caldav import DAVClient
import requests
import urllib3
import traceback

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
NEXTCLOUD_URL = ""
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
    logging.info(f"Microsoft Login URL: {MICROSOFT_LOGIN_URL}")
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
    logging.info(f"Kronos URL: {KRONOS_URL}")

    try:
        # Navigate to the Kronos schedule page
        driver.get(KRONOS_URL)
        logging.info("Navigated to Kronos schedule page.")
        log_page_details(driver)  # Log current page details and capture a screenshot

        # Locate all day elements (li elements with class 'withDivider')
        schedule_days = WebDriverWait(driver, 40).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.withDivider"))
        )
        logging.info(f"Located {len(schedule_days)} schedule day elements.")

        schedule_data = []
        seen_shifts = set()

        # Iterate over each day to find shifts
        for day in schedule_days:
            try:
                day_date = day.get_attribute("datetime")
                logging.info(f"Processing schedule for date: {day_date}")

                # Locate shift elements within the day
                shift_wrappers = day.find_elements(By.CSS_SELECTOR, "div.scheduleEntityWrapper, div.shiftPosition")
                logging.info(f"Found {len(shift_wrappers)} shifts for date: {day_date}")

                for shift in shift_wrappers:
                    try:
                        # Attempt to find the time element
                        try:
                            time_element = shift.find_element(By.CSS_SELECTOR, "p.props, time.label")
                            time_range = time_element.text

                            # Use regex to extract start time, end time, and shift length
                            match = re.search(
                                r'(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)\s*(\d+\.\d+)?',
                                time_range
                            )
                            if match:
                                start_time_str = match.group(1)
                                end_time_str = match.group(2)
                                shift_length = match.group(3)
                            else:
                                logging.warning(f"Could not parse time range: {time_range}")
                                continue  # Skip this shift
                        except NoSuchElementException:
                            logging.warning(f"No time element found for shift on date {day_date}. Skipping shift.")
                            continue  # Skip this shift if no time element is found

                        # Attempt to find the shift details
                        try:
                            shift_details = shift.find_element(By.CSS_SELECTOR, "p.label").text
                        except NoSuchElementException:
                            shift_details = "No details available"

                        # Create a unique key for the shift
                        shift_key = (day_date, start_time_str, end_time_str, shift_details, shift_length)

                        # Check if the shift has already been processed
                        if shift_key in seen_shifts:
                            logging.info(f"Duplicate shift found for date {day_date}: {start_time_str}-{end_time_str}, {shift_details}. Skipping.")
                            continue  # Skip adding this shift as it's a duplicate
                        else:
                            seen_shifts.add(shift_key)  # Add the shift to the set of seen shifts

                        logging.info(f"Shift details: {start_time_str}-{end_time_str} ({shift_length} hrs), {shift_details}")

                        schedule_data.append({
                            "date": day_date,
                            "start_time": start_time_str,
                            "end_time": end_time_str,
                            "details": shift_details,
                            "shift_length": shift_length
                        })
                    except Exception as e:
                        logging.error(f"Error scraping shift details for date {day_date}: {str(e)}")
                        logging.error(f"Full traceback: {traceback.format_exc()}")
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
        logging.error(f"An error occurred: {str(e)}")
        logging.error(f"Full traceback: {traceback.format_exc()}")
        return []

def retrieve_existing_events():
    """Retrieve all existing events from the 'personal' calendar in Nextcloud using CalDAV."""
    client = DAVClient(NEXTCLOUD_URL, username=NEXTCLOUD_USERNAME, password=NEXTCLOUD_PASSWORD)
    principal = client.principal()
    calendars = principal.calendars()

    logging.info("Available calendars:")
    for cal in calendars:
        logging.info(f"- {cal.name}")
    
    existing_events = {}
    for calendar in calendars:
        if calendar.name.lower() == 'personal'.lower():
            logging.info(f"Retrieving events from calendar: {calendar.name}")
            events = calendar.events()
            logging.info(f"Found {len(events)} events in 'personal' calendar.")
            for event in events:
                try:
                    ical = Calendar.from_ical(event.data)
                    for component in ical.walk('VEVENT'):
                        dtstart = component.get('dtstart').dt
                        dtend = component.get('dtend').dt
                        summary = component.get('summary')
                        event_key = (dtstart, dtend, summary)
                        existing_events[event_key] = component.to_ical().decode('utf-8')
                        logging.info(f"Event found: {summary} from {dtstart} to {dtend}")
                except Exception as e:
                    logging.error(f"Error parsing event: {e}")
            break
    
    logging.info(f"Retrieved {len(existing_events)} existing events from the 'personal' calendar.")
    return existing_events

def create_icalendar_event(date_str, start_time_str, end_time_str, details, shift_length=None):
    """Create an iCalendar event."""
    event = Event()

    # Parse the date and times
    date_part = date_str.split(' ')
    date_formatted = f"{date_part[0]} {date_part[1]} {date_part[2]} {date_part[3]}"

    start_datetime_str = f"{date_formatted} {start_time_str}"
    end_datetime_str = f"{date_formatted} {end_time_str}"

    # Assuming times are in US/Eastern timezone
    local_tz = timezone('US/Eastern')
    start_time = local_tz.localize(datetime.strptime(start_datetime_str, "%a %b %d %Y %I:%M %p"))
    end_time = local_tz.localize(datetime.strptime(end_datetime_str, "%a %b %d %Y %I:%M %p"))

    # Generate a unique UID
    unique_uid = f"{uuid.uuid4()}@mydomain.com"
    logging.info(f"Generated UID: {unique_uid} for event on {date_str}")

    # Adjust details if no details are available
    if details == "No details available":
        if shift_length:
            details = f"{shift_length} hrs"
        else:
            details = ""

    # Build the event summary
    event_summary = f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
    if details:
        event_summary += f": {details}"

    event.add('summary', event_summary)
    logging.info(f"Event summary set to: {event_summary}")

    # Add event timing and other properties
    event.add('dtstart', vDatetime(start_time))
    event.add('dtend', vDatetime(end_time))
    event.add('uid', unique_uid)
    event.add('dtstamp', vDatetime(datetime.utcnow()))

    return event


def create_individual_ics_files(schedule_data):
    """Generate individual iCalendar (.ics) files for each event in schedule data."""
    ics_filenames = []

    for entry in schedule_data:
        try:
            event = create_icalendar_event(
                entry['date'],
                entry['start_time'],
                entry['end_time'],
                entry['details'],
                entry.get('shift_length')  # Pass shift_length if available
            )
            cal = Calendar()
            cal.add_component(event)

            uid = event.get('uid')
            ics_filename = f"{uid}.ics"
            ics_filepath = os.path.join('individual_events', ics_filename)

            with open(ics_filepath, 'wb') as f:
                f.write(cal.to_ical())

            ics_filenames.append(ics_filepath)
            logging.info(f"Generated iCalendar file: {ics_filepath}")
        except Exception as e:
            logging.error(f"Error creating iCalendar event for entry: {entry}")
            logging.error(f"Exception: {e}")
            logging.error(f"Full traceback: {traceback.format_exc()}")

    return ics_filenames

def compare_and_handle_existing(events_to_upload, existing_events):
    """Compare newly generated events with existing ones on Nextcloud and delete the local file if a match is found."""
    for event_filename in events_to_upload:
        with open(event_filename, 'rb') as f:
            new_event = Calendar.from_ical(f.read())
        
        for new_event_component in new_event.walk('VEVENT'):
            new_event_dtstart = new_event_component.get('dtstart').dt
            new_event_dtend = new_event_component.get('dtend').dt
            new_event_summary = new_event_component.get('summary')

            logging.info(f"Checking new event with start time {new_event_dtstart} and end time {new_event_dtend} against existing events on Nextcloud.")

            for event_key in existing_events.keys():
                existing_event_dtstart, existing_event_dtend, existing_event_summary = event_key

                if (new_event_dtstart == existing_event_dtstart and
                    new_event_dtend == existing_event_dtend and
                    new_event_summary == existing_event_summary):
                    
                    logging.info(f"Event with start time {new_event_dtstart} and end time {new_event_dtend} already exists on Nextcloud. Deleting local file: {event_filename}")
                    os.remove(event_filename)
                    break  # No need to check further if we found a match
            else:
                logging.info(f"No matching event found on Nextcloud for event starting at {new_event_dtstart}.")

def upload_to_nextcloud_individual_files(ics_filenames):
    """Upload individual .ics files to the 'personal' calendar on Nextcloud using CalDAV."""
    client = DAVClient(NEXTCLOUD_URL, username=NEXTCLOUD_USERNAME, password=NEXTCLOUD_PASSWORD)
    principal = client.principal()
    calendar = None

    # Debug: List available calendars during upload phase
    logging.info("Available calendars during upload:")
    for cal in principal.calendars():
        logging.info(f"- {cal.name}")
        if cal.name.lower() == 'personal'.lower():  # Case-insensitive comparison
            calendar = cal
            logging.info("Found the 'personal' calendar during upload.")
            break

    if not calendar:
        logging.error("The 'personal' calendar was not found during upload.")
        return

    for ics_file in ics_filenames:
        if os.path.exists(ics_file):  # Check if the file still exists after comparison
            with open(ics_file, 'r') as f:
                event_data = f.read()

            try:
                calendar.add_event(event_data)
                logging.info(f"Successfully uploaded {ics_file} to Nextcloud")
            except Exception as e:
                logging.error(f"Failed to upload {ics_file} to Nextcloud: {e}")

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
        # Retrieve existing events from Nextcloud
        existing_events = retrieve_existing_events()        
        
        login_to_microsoft(driver)
        schedule_data = scrape_schedule(driver)

        # Ensure the directory exists
        os.makedirs('individual_events', exist_ok=True)

        # Generate new .ics files
        ics_filenames = create_individual_ics_files(schedule_data)

        # Compare and handle existing events on Nextcloud
        compare_and_handle_existing(ics_filenames, existing_events)

        # Upload new events to Nextcloud
        upload_to_nextcloud_individual_files(ics_filenames)

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        logging.error(f"Full traceback: {traceback.format_exc()}")  # Added traceback

    finally:
        driver.quit()
        logging.info("Browser closed")

if __name__ == "__main__":
    main()
