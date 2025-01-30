import os
import json
import pickle
import logging
import requests
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import Error as GoogleAPIError
import tkinter as tk
from tkinter import ttk, messagebox
import keyring
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time

class GoogleAPIError(Exception):
    """Custom exception for Google API related errors."""
    pass

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('calendar_sync.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class CalendarSync:
    """
    A class to synchronize events between Google Calendar and Office 365 Calendar.
    
    This application provides a GUI interface for users to input their credentials
    and initiate the synchronization process. It handles authentication for both
    Google and Office 365 calendars and transfers events from Google Calendar to
    Office 365.
    """

    def __init__(self):
        """Initialize the CalendarSync application with necessary configurations."""
        self.SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
        self.credentials_dir = 'credentials'
        self.config_file = 'config.json'
        
        # Ensure credentials directory exists
        if not os.path.exists(self.credentials_dir):
            os.makedirs(self.credentials_dir)
            logger.info("Created credentials directory")

        # Load configuration if exists
        self.load_config()

    def load_config(self):
        """Load configuration from config file if it exists."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.gmail_email_default = config.get('gmail_email', '')
                    self.o365_email_default = config.get('o365_email', '')
                    logger.info("Configuration loaded successfully")
            else:
                self.gmail_email_default = ''
                self.o365_email_default = ''
                logger.warning("No configuration file found")
        except json.JSONDecodeError:
            logger.error("Error reading config file")
            self.gmail_email_default = ''
            self.o365_email_default = ''

    def setup_gui(self):
        """Set up the graphical user interface."""
        self.root = tk.Tk()
        self.root.title("Calendar Sync")
        self.root.geometry("400x300")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Gmail account
        ttk.Label(main_frame, text="Gmail Account:").grid(row=0, column=0, sticky=tk.W)
        self.gmail_email = ttk.Entry(main_frame, width=40)
        self.gmail_email.grid(row=0, column=1, pady=5)
        self.gmail_email.insert(0, self.gmail_email_default)

        # O365 account
        ttk.Label(main_frame, text="O365 Account:").grid(row=1, column=0, sticky=tk.W)
        self.o365_email = ttk.Entry(main_frame, width=40)
        self.o365_email.grid(row=1, column=1, pady=5)
        self.o365_email.insert(0, self.o365_email_default)
        
        # O365 password
        ttk.Label(main_frame, text="O365 Password:").grid(row=2, column=0, sticky=tk.W)
        self.o365_password = ttk.Entry(main_frame, width=40, show="*")
        self.o365_password.grid(row=2, column=1, pady=5)

        # Buttons
        ttk.Button(main_frame, text="Start Sync", command=self.start_sync).grid(row=3, column=1, pady=20)
        ttk.Button(main_frame, text="Save Credentials", command=self.save_credentials).grid(row=4, column=1)

        # Status label
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=20)

    def authenticate_google(self):
        """
        Authenticate with Google Calendar API.
        
        Returns:
            google.auth.credentials.Credentials: Authenticated Google Calendar service
            
        Raises:
            FileNotFoundError: If client_secrets.json is missing
            GoogleAPIError: If authentication fails
        """
        try:
            creds = None
            token_path = os.path.join(self.credentials_dir, 'gmail_token.pickle')
            
            if os.path.exists(token_path):
                with open(token_path, 'rb') as token:
                    creds = pickle.load(token)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    client_secrets_path = os.path.join(self.credentials_dir, 'client_secrets.json')
                    if not os.path.exists(client_secrets_path):
                        raise FileNotFoundError(
                            "client_secrets.json not found in credentials directory. "
                            "Please download it from Google Cloud Console."
                        )
                    
                    flow = InstalledAppFlow.from_client_secrets_file(
                        client_secrets_path, self.SCOPES)
                    creds = flow.run_local_server(port=0)

                with open(token_path, 'wb') as token:
                    pickle.dump(creds, token)

            logger.info("Google authentication successful")
            return build('calendar', 'v3', credentials=creds)
            
        except FileNotFoundError as e:
            logger.error(f"Authentication failed: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error during Google authentication: {str(e)}")
            raise GoogleAPIError(f"Google authentication failed: {str(e)}")

    def authenticate_o365(self):
        """
        Authenticate with Office 365 using Selenium.
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        driver = None
        try:
            driver = webdriver.Chrome()
            logger.info("Starting O365 authentication")
            
            # Navigate to O365 login
            driver.get('https://outlook.office365.com/calendar/view/month')
            
            # Wait for and fill email
            email_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "loginfmt"))
            )
            email_input.send_keys(self.o365_email.get())
            
            # Click Next
            next_button = driver.find_element(By.ID, "idSIButton9")
            next_button.click()
            
            # Wait for and fill password
            password_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "passwd"))
            )
            time.sleep(1)
            driver.execute_script("arguments[0].value = arguments[1]", password_input, self.o365_password.get())
            
            # Click Sign in
            sign_in_button = driver.find_element(By.ID, "idSIButton9")
            sign_in_button.click()
            
            # Wait for DUO prompt
            time.sleep(2)
            
            # Store session cookies
            cookies = driver.get_cookies()
            with open(os.path.join(self.credentials_dir, 'o365_cookies.json'), 'w') as f:
                json.dump(cookies, f)
            
            logger.info("O365 authentication successful")
            return True
            
        except TimeoutException:
            logger.error("O365 authentication timed out")
            messagebox.showerror("Error", "Authentication timed out. Please try again.")
            return False
        except WebDriverException as e:
            logger.error(f"Selenium WebDriver error: {str(e)}")
            messagebox.showerror("Error", f"Browser automation error: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during O365 authentication: {str(e)}")
            messagebox.showerror("Error", f"O365 authentication failed: {str(e)}")
            return False
        finally:
            if driver:
                driver.quit()

    def create_event_key(self, event, is_google=True):
        """
        Create a unique key for an event based on its properties.
        
        Args:
            event (dict): Event data
            is_google (bool): Whether the event is from Google Calendar
            
        Returns:
            str: Unique event identifier
        """
        if is_google:
            start_time = event['start'].get('dateTime', event['start'].get('date'))
            event_id = event.get('id', '')
        else:
            start_time = event['start'].get('dateTime', event['start'].get('date'))
            event_id = event.get('id', '')
        
        # Create a key that combines multiple identifiers
        return f"{event_id}_{start_time}"

    def get_existing_events(self, session):
        """
        Retrieve existing events from O365 calendar with detailed information.
        
        Args:
            session (requests.Session): Authenticated session object
            
        Returns:
            dict: Dictionary of existing events with composite keys
        """
        try:
            # Get events for the next 30 days
            now = datetime.utcnow()
            end_date = (now + timedelta(days=30)).isoformat() + 'Z'
            now = now.isoformat() + 'Z'
            
            response = session.get(
                'https://graph.microsoft.com/v1.0/me/events',
                params={
                    '$select': 'id,subject,start,end,body,location',
                    '$filter': f"start/dateTime ge '{now}' and start/dateTime le '{end_date}'"
                }
            )
            response.raise_for_status()
            
            events_dict = {}
            for event in response.json().get('value', []):
                key = self.create_event_key(event, is_google=False)
                events_dict[key] = event
            
            logger.info(f"Retrieved {len(events_dict)} existing O365 events")
            return events_dict
        except Exception as e:
            logger.error(f"Error retrieving O365 events: {str(e)}")
            return {}

    def sync_calendars(self, google_service):
        """
        Synchronize events from Google Calendar to Office 365.
        
        Args:
            google_service: Authenticated Google Calendar service
        """
        try:
            # Get Gmail calendar events for the next 30 days
            now = datetime.utcnow()
            end_date = (now + timedelta(days=30)).isoformat() + 'Z'
            now = now.isoformat() + 'Z'
            
            events_result = google_service.events().list(
                calendarId='primary',
                timeMin=now,
                timeMax=end_date,
                maxResults=250,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            events = events_result.get('items', [])
            
            if not events:
                logger.info("No upcoming events found in Google Calendar")
                return
            
            # Get authentication token from browser session
            driver = webdriver.Chrome()
            try:
                driver.get('https://outlook.office365.com')
                token = driver.execute_script('return window.localStorage.getItem("accessToken")')
            finally:
                driver.quit()

            # Create session with stored O365 cookies
            session = requests.Session()
            try:
                with open(os.path.join(self.credentials_dir, 'o365_cookies.json'), 'r') as f:
                    cookies = json.load(f)
                    for cookie in cookies:
                        session.cookies.set(cookie['name'], cookie['value'])
            except FileNotFoundError:
                logger.error("O365 session not found")
                raise Exception("O365 session not found. Please authenticate first.")

            # Get existing events to check for duplicates
            existing_events = self.get_existing_events(session)
            
            # Sync events
            for event in events:
                event_key = self.create_event_key(event, is_google=True)
                event_summary = event.get('summary', 'No Title')
                
                # Format event data for O365
                o365_event = {
                    'subject': event_summary,
                    'start': {
                        'dateTime': event['start'].get('dateTime', event['start'].get('date')),
                        'timeZone': event['start'].get('timeZone', 'UTC')
                    },
                    'end': {
                        'dateTime': event['end'].get('dateTime', event['end'].get('date')),
                        'timeZone': event['end'].get('timeZone', 'UTC')
                    }
                }
                
                # Add description if available
                if 'description' in event:
                    o365_event['body'] = {
                        'contentType': 'text',
                        'content': event['description']
                    }
                
                # Add location if available
                if 'location' in event:
                    o365_event['location'] = {
                        'displayName': event['location']
                    }

                # Create session with token
                session.headers.update({
                    'Authorization': f'Bearer {token}',
                    'Content-Type': 'application/json'
                })

                try:
                    if event_key in existing_events:
                        # Event exists - check if it needs updating
                        existing_event = existing_events[event_key]
                        needs_update = (
                            existing_event['subject'] != event_summary or
                            existing_event['start']['dateTime'] != o365_event['start']['dateTime'] or
                            existing_event['end']['dateTime'] != o365_event['end']['dateTime'] or
                            ('location' in o365_event and ('location' not in existing_event or 
                                existing_event['location']['displayName'] != o365_event['location']['displayName']))
                        )
                        
                        if needs_update:
                            # Update existing event
                            response = session.patch(
                                f"https://graph.microsoft.com/v1.0/me/events/{existing_event['id']}",
                                json=o365_event
                            )
                            response.raise_for_status()
                            logger.info(f"Updated existing event: {event_summary}")
                        else:
                            logger.info(f"Event unchanged, no update needed: {event_summary}")
                    else:
                        # Create new event
                        response = session.post(
                            'https://graph.microsoft.com/v1.0/me/events',
                            json=o365_event
                        )
                        response.raise_for_status()
                        logger.info(f"Created new event: {event_summary}")
                except Exception as e:
                    logger.error(f"Failed to sync event {event_summary}: {str(e)}")

        except Exception as e:
            logger.error(f"Error during calendar sync: {str(e)}")
            raise

    def save_credentials(self):
        """Save user credentials securely using keyring."""
        try:
            keyring.set_password("calendar_sync", "gmail_email", self.gmail_email.get())
            keyring.set_password("calendar_sync", "o365_email", self.o365_email.get())
            logger.info("Credentials saved successfully")
            self.status_label.config(text="Credentials saved successfully!")
        except Exception as e:
            logger.error(f"Error saving credentials: {str(e)}")
            self.status_label.config(text="Error saving credentials!")

    def start_sync(self):
        """Start the calendar synchronization process."""
        try:
            self.status_label.config(text="Starting sync process...")
            logger.info("Starting calendar sync")
            
            # Authenticate with Google
            google_service = self.authenticate_google()
            
            # Authenticate with O365
            if self.authenticate_o365():
                # Perform the sync
                self.sync_calendars(google_service)
                logger.info("Sync completed successfully")
                self.status_label.config(text="Sync completed successfully!")
            else:
                logger.error("O365 authentication failed")
                self.status_label.config(text="O365 authentication failed")
                
        except Exception as e:
            logger.error(f"Sync error: {str(e)}")
            self.status_label.config(text=f"Error: {str(e)}")

    def run(self):
        """Start the Calendar Sync application."""
        logger.info("Starting Calendar Sync application")
        self.setup_gui()
        self.root.mainloop()

if __name__ == "__main__":
    app = CalendarSync()
    app.run()