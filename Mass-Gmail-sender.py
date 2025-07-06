#!/usr/bin/env python3
"""
Gmail Email Sender v1.0

A comprehensive email automation tool that uses Selenium WebDriver to send emails
through Gmail accounts. Supports multiple accounts, bulk email sending, and
parallel processing for efficiency.

Features:
- Multiple Gmail account management
- Bulk email sending with HTML content
- Parallel processing for faster execution
- Error handling and logging
- Screenshot capture on errors
- Account verification and recovery

Author: [Your Name]
Date: 2025-07-06
Version: 1.0

Requirements:
- Python 3.8+
- Chrome browser
- Valid Gmail accounts
- Excel files with account and recipient data
"""

import re
import os
import sys
import time
import cutie
import pandas
import random
import string
import logging
import platform
import datetime
from seleniumbase import SB
from functools import partial
from multiprocessing import Lock
from configparser import RawConfigParser
from multiprocessing.pool import ThreadPool
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import (
    NoSuchAttributeException, StaleElementReferenceException,
    NoSuchElementException, TimeoutException, ElementClickInterceptedException, 
    WebDriverException, ElementNotInteractableException
)


####################################################################################
# ---------------------------------Settings----------------------------------------#
####################################################################################

# Set working directory to script location
if getattr(sys, "frozen", False):
    # If running as compiled executable
    os.chdir(os.path.dirname(os.path.abspath(sys.executable)))
else:
    # If running as Python script
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Constants and configuration
COMMON_EXCEPTIONS = (
    NoSuchElementException, TimeoutException, StaleElementReferenceException,
    ElementClickInterceptedException, WebDriverException, ElementNotInteractableException,
    NoSuchAttributeException
)
CLEAR = "cls" if platform.system().upper() == "WINDOWS" else "clear"
_LIMIT_TRACK_FILE = "daily_limit.csv"
_FOLDER_CONTAINING_ALL_PROFILES = "GmailProfiles"
_FOLDER_CONTAINING_SCREENSHOTS = "ErrorScreenshots"
_PORT = 9223  # Starting port for Chrome debugging
_CONFIG = RawConfigParser()
_CONFIG.read("settings.cfg")

# Logging configuration
LOGGER = logging.getLogger("gmail_email_sender")
logging.basicConfig(
    filename="gmail_email_sender.log", 
    filemode="w",
    format="%(asctime)s ==> %(message)s"
)
LOGGER.setLevel(logging.DEBUG)


####################################################################################
# ---------------------------------Decorators--------------------------------------#
####################################################################################

def exceptional_handler(func):
    """
    Decorator to handle common Selenium exceptions with retry logic.
    
    Args:
        func: Function to be decorated
        
    Returns:
        Wrapped function with exception handling and retry mechanism
    """
    def wrapper(*args, **kwargs):
        retry = kwargs.get("retry", 0)
        max_retries = kwargs.get("max_retries", 2)
        
        if retry >= max_retries:
            raise TimeoutException("Maximum retries reached!")
        
        try:
            # Remove retry parameters from kwargs before calling function
            if "retry" in kwargs:
                kwargs.pop("retry")
            if "max_retries" in kwargs:
                kwargs.pop("max_retries")
            return func(*args, **kwargs)
        except COMMON_EXCEPTIONS:
            time.sleep(5)
            return wrapper(retry=retry + 1, max_retries=max_retries, *args, **kwargs)
    return wrapper


def wait_until(condition_func):
    """
    Decorator to wait until a condition is met with customizable parameters.
    
    Args:
        condition_func: Function that returns True when condition is met
        
    Returns:
        Wrapped function with waiting logic
    """
    def wrapper(*args, **kwargs):
        # Execute pre-loop callback
        kwargs.get("before_loop", lambda: True)()
        
        max_attempts = kwargs.get("max_tries", -1)
        dots = 1
        attempt = 0
        not_completed = False
        
        while True:
            # Execute in-loop pre-condition callback
            kwargs.get("in_loop_before", lambda: True)()
            
            # Check condition
            if condition_func(*args):
                break
            
            # Check max attempts
            if max_attempts != -1 and attempt >= max_attempts:
                not_completed = True
                break
            
            attempt += 1
            
            # Display waiting message with animated dots
            print(f"{kwargs.get('message', 'Waiting')}{'.' * dots}", end="\r")
            dots = 1 if dots > 2 else dots + 1
            
            # Execute in-loop post-condition callback
            kwargs.get("in_loop_after", lambda: True)()
            
            time.sleep(kwargs.get("sleep", 0.5))
            print(" " * 100, end="\r")  # Clear line
        
        # Execute post-loop callback
        kwargs.get("after_loop", lambda: True)()
        return not not_completed
    return wrapper


####################################################################################
# ---------------------------------Browser Handler---------------------------------#
####################################################################################

class BrowserHandler:
    """
    Base class for handling Chrome browser automation using Selenium WebDriver.
    
    Provides common browser operations like element finding, clicking, typing,
    and browser lifecycle management.
    """

    def __init__(self, temp_profile: str = None, port: int = 9222) -> None:
        """
        Initialize browser handler with configuration.
        
        Args:
            temp_profile (str): Path to Chrome profile directory
            port (int): Port for Chrome remote debugging
        """
        self.platform = platform.system().upper()
        self.temp_profile = temp_profile
        self.driver = None
        self.wait = None
        self.seleniumbase_driver = None
        self.sb_init = None
        self.port = port

    def get_element(self, css_selector: str, by_clickable: bool = False, 
                   multiple: bool = False) -> WebElement | list:
        """
        Find web element(s) using CSS selector with wait conditions.
        
        Args:
            css_selector (str): CSS selector to find element
            by_clickable (bool): Wait for element to be clickable
            multiple (bool): Return all matching elements
            
        Returns:
            WebElement or list of WebElements
        """
        condition = ec.presence_of_all_elements_located if multiple else ec.presence_of_element_located
        if by_clickable:
            condition = ec.element_to_be_clickable
        return self.wait.until(condition((By.CSS_SELECTOR, css_selector)))

    def find_elements(self, css_selector: str, reference_element=None):
        """
        Find elements without waiting (immediate search).
        
        Args:
            css_selector (str): CSS selector to find elements
            reference_element: Element to search within (default: driver)
            
        Returns:
            List of WebElements
        """
        reference = self.driver if reference_element is None else reference_element
        return reference.find_elements(By.CSS_SELECTOR, css_selector)

    def get_element_by_text(self, text: str, css_selector: str = None, elements=None):
        """
        Find element by its text content.
        
        Args:
            text (str): Text to search for
            css_selector (str): CSS selector to get elements
            elements: Pre-found elements to search through
            
        Returns:
            WebElement containing the text or None
        """
        if elements is None:
            elements = self.get_element(css_selector, multiple=True)
        
        elements_text = self.get_text(element=elements, multiple=True)
        for element_text in elements_text:
            if text.lower() in element_text.lower():
                return elements[elements_text.index(element_text)]
        return None

    @exceptional_handler
    def write(self, css_selector: str, data: str, enter=False):
        """
        Type text into an input element.
        
        Args:
            css_selector (str): CSS selector for input element
            data (str): Text to type
            enter (bool): Press Enter after typing
        """
        input_el = self.get_element(css_selector, by_clickable=True)
        input_el.clear()
        input_el.send_keys(data)
        if enter:
            input_el.send_keys(Keys.ENTER)

    @exceptional_handler
    def click_element(self, css_selector: str = None, element=None, scroll=True):
        """
        Click on an element with optional scrolling.
        
        Args:
            css_selector (str): CSS selector for element to click
            element: WebElement to click directly
            scroll (bool): Scroll element into view before clicking
        """
        if css_selector is not None:
            element = self.get_element(css_selector, by_clickable=True)
        
        if scroll:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'})", element
            )
        
        time.sleep(1)
        element.click()

    @exceptional_handler
    def get_text(self, css_selector=None, element=None, multiple=False):
        """
        Get text content from element(s).
        
        Args:
            css_selector (str): CSS selector for element
            element: WebElement to get text from
            multiple (bool): Get text from multiple elements
            
        Returns:
            String or list of strings
        """
        if css_selector is not None:
            element = self.get_element(css_selector, multiple=multiple)
        
        if not multiple:
            return element.get_property("innerText")
        return [el.get_property("innerText") for el in element]

    def start_chrome(self, headless: bool = False, **kwargs) -> None:
        """
        Start Chrome browser with SeleniumBase.
        
        Args:
            headless (bool): Run browser in headless mode
            **kwargs: Additional arguments for browser configuration
        """
        kwargs["chromium_arg"] = f'{kwargs.get("chromium_arg", "")},' \
                                f'--remote-debugging-port={self.port}'.strip(",")
        
        sb_init = SB(
            uc=True, 
            headed=not headless,
            user_data_dir=self.temp_profile,
            headless=headless, 
            **kwargs
        )
        
        self.seleniumbase_driver = sb_init.__enter__()
        self.sb_init = sb_init
        self.driver = self.seleniumbase_driver.driver
        self.wait = WebDriverWait(self.driver, 40)
        self.driver.set_page_load_timeout(300)

    def kill_browser(self) -> None:
        """Close browser and clean up resources."""
        if not hasattr(self, "driver") or self.driver is None:
            return
        
        self.driver.quit()
        self.sb_init.__exit__(None, None, None)
        self.driver = None
        time.sleep(5)


############################################################################################
# -----------------------------------Gmail Account Handler--------------------------------#
############################################################################################

class GmailAccount(BrowserHandler):
    """
    Gmail account handler for email automation.
    
    Extends BrowserHandler to provide Gmail-specific functionality including
    login, account verification, and email sending capabilities.
    """

    def __init__(self, email, password, recovery_email, **kwargs):
        """
        Initialize Gmail account handler.
        
        Args:
            email (str): Gmail address
            password (str): Account password
            recovery_email (str): Recovery email for verification
            **kwargs: Additional browser configuration
        """
        # Set profile path based on email username
        kwargs["temp_profile"] = os.path.abspath(f'GmailProfiles/{email.split("@")[0]}')
        super().__init__(**kwargs)
        
        # Start browser and set window size
        self.start_chrome()
        self.driver.set_window_size(1280, 720)
        
        # Store account credentials
        self.email = email
        self.password = password
        self.recovery_email = recovery_email

    @wait_until
    def wait_until_loaded(self):
        """
        Wait until Gmail login page is loaded.
        
        Returns:
            bool: True if page is loaded, False otherwise
        """
        try:
            # Check for login identifier field
            if self.find_elements("[name=identifier]"):
                return True
            # Check if already logged in
            elif self.find_elements("a[href=personal-info]"):
                self.driver.execute_script("alert('Logged In')")
                return True
        except:
            pass
        return False
        
    @wait_until
    def wait_until_gmail_logged_in(self, recovery_email):
        """
        Wait until Gmail login process is complete, handling various scenarios.
        
        Args:
            recovery_email (str): Recovery email for account verification
            
        Returns:
            bool: True if login successful, False otherwise
        """
        try:
            # Check if account is disabled
            if (self.find_elements("#headingText") and 
                "Your account has been disabled" in self.get_text("#headingText")):
                self.driver.execute_script("alert('Account Disabled')")
                return True
            
            # Handle account verification challenge
            elif self.find_elements("input[name=challengeListId]"):
                self.click_element("section ul li:nth-child(3)")
                self.write("input[type=email]", recovery_email, enter=True)
            
            # Handle "Not now" dialog (English)
            elif self.find_elements("div[aria-live=polite]"):
                button_element = self.get_element_by_text("not now", "button")
                if button_element is not None:
                    self.click_element(element=button_element)
            
            # Handle "Not now" dialog (French)
            elif self.find_elements("div[aria-live=polite]"):
                button_element = self.get_element_by_text("Pas maintenant", "button")
                if button_element is not None:
                    self.click_element(element=button_element)
            
            # Check if successfully logged in
            elif self.find_elements("a[href=personal-info]"):
                return True
                
        except:
            return False

    def login_gmail(self):
        """
        Perform Gmail login process.
        
        Returns:
            bool: True if login successful, False otherwise
        """
        try:
            # Navigate to Google accounts page
            self.driver.get("https://accounts.google.com/")
            self.wait_until_loaded()
            
            # Handle alert if already logged in
            if ec.alert_is_present()(self.driver):
                self.driver.switch_to.alert.accept()
                return True
            
            # Enter email and password
            self.write("[name=identifier]", self.email, enter=True)
            self.write("input[type=password]", self.password, enter=True)
            
            # Wait for login completion
            self.wait_until_gmail_logged_in(
                self.recovery_email, 
                message="Waiting until gmail logged in"
            )
            
            # Check if account is disabled
            if ec.alert_is_present()(self.driver):
                self.driver.switch_to.alert.accept()
                logging_error_screenshot(f"Account {self.email} is disabled!", self.driver)
                return False
            
            return True
            
        except COMMON_EXCEPTIONS:
            logging_error_screenshot(f"Cannot login {self.email}!", self.driver)
            return False
    
    @wait_until
    def wait_until_email_sent(self):
        """
        Wait until email is sent successfully.
        
        Returns:
            bool: True if email sent, False otherwise
        """
        try:
            # Check for success alert
            alert = self.find_elements("div[role=alert]")
            if alert and ("sent" in self.get_text(element=alert[0])):
                return True
            
            # Click send button if still visible
            elif self.find_elements("div[role=button][aria-label*=Ctrl-Enter]"):
                self.click_element("div[role=button][aria-label*=Ctrl-Enter]")
                return False
                
        except COMMON_EXCEPTIONS:
            return False
                                                      
    def send_emails(self, to, subject):
        """
        Send email to specified recipient.
        
        Args:
            to (str): Recipient email address
            subject (str): Email subject line
        """
        try:
            LOGGER.info(f"Sending email to {to} from {self.email}")
            
            # Navigate to Gmail if not already there
            if "https://mail.google.com/mail/u/0/" not in self.driver.current_url:
                self.driver.get("https://mail.google.com/mail/u/0/")
            
            # Wait for Gmail to load
            wait_until(lambda: self.find_elements(
                "div[role=navigation] > div:first-child div[style*=user-select]"
            ))(message="Waiting until loaded")
            
            # Close any open dialogs
            while self.find_elements("div[role=dialog] td > img:last-child"):
                self.click_element("div[role=dialog] td > img:last-child")
            
            # Click compose button
            self.click_element("div[role=navigation] > div:first-child div[style*=user-select]")
            
            # Fill recipient field
            self.write("input[aria-haspopup=listbox]", to, enter=True)
            
            # Fill subject field
            self.write("input[name=subjectbox]", subject)
            
            # Fill email body with HTML content
            email_inp = self.get_element("div[role=textbox]")
            with open(_CONFIG["EMAIL_INFO"]["email_html_file"], "r", encoding="utf-8") as f:
                html_data = f.read()
            
            # Extract body content from HTML
            body = re.findall("(?<=<body>).*(?=</body>)", html_data)
            html_data = body[0] if body else html_data
            
            # Insert HTML content into email body
            self.driver.execute_script(
                "arguments[0].innerHTML = arguments[1]", 
                email_inp, 
                html_data
            )
            
            time.sleep(2)
            
            # Send email
            self.click_element("div[role=button][aria-label*=Ctrl-Enter]")
            
            # Wait for send button to disappear
            wait_until(lambda: not self.find_elements(
                "div[role=button][aria-label*=Ctrl-Enter]"
            ))()
            
            # Verify email was sent
            email_sent = self.wait_until_email_sent(
                message="Waiting until email sent", 
                max_tries=120
            )
            
            if not email_sent:
                raise TimeoutException("Email sent dialog cannot be detected.")
            
            LOGGER.info(f"Email sent to {to} from {self.email}")
            
        except COMMON_EXCEPTIONS:
            logging_error_screenshot(
                f"Error occurred while sending email to {to} from {self.email}", 
                self.driver
            )


#################################################################################################
# -----------------------------------Utility Functions-----------------------------------------#
#################################################################################################

def send_email(recipients, row):
    """
    Send emails to all recipients using specified Gmail account.
    
    Args:
        recipients (pandas.DataFrame): DataFrame containing recipient emails
        row (pandas.Series): Row containing Gmail account information
    """
    browser = add_gmail(row, False)
    if browser is None:
        return
    
    # Send email to each recipient
    recipients[_CONFIG["RECIPIENT"]["email_col"]].apply(
        lambda rec: browser.send_emails(rec, _CONFIG["EMAIL_INFO"]["email_subject"])
    )
    
    browser.kill_browser()


def add_gmail(row, close=True):
    """
    Create and login to Gmail account.
    
    Args:
        row (pandas.Series): Row containing account information
        close (bool): Whether to close browser after login
        
    Returns:
        GmailAccount: Logged in Gmail account or None if failed
    """
    global _PORT
    row = row[1]
    
    # Thread-safe port assignment
    lock.acquire()
    rec_col = _CONFIG["GMAIL_ACCOUNTS"]["recovery_email_col"]
    browser = GmailAccount(
        row[_CONFIG["GMAIL_ACCOUNTS"]["email_col"]], 
        row[_CONFIG["GMAIL_ACCOUNTS"]["pass_col"]], 
        row[rec_col] if rec_col else "",
        port=_PORT
    )
    _PORT += 1
    lock.release()
    
    # Attempt login
    logged_in = browser.login_gmail()
    
    if close:
        browser.kill_browser()
    
    if not logged_in:
        return None
    
    return browser


def parallel_browsing(func, op):
    """
    Execute function in parallel across multiple Gmail accounts.
    
    Args:
        func: Function to execute
        op (int): Operation type (1 for email sending, other for account management)
        
    Returns:
        Wrapped function for parallel execution
    """
    # Load Gmail accounts data
    file_data = pandas.read_excel(_CONFIG["GMAIL_ACCOUNTS"]["emails_excel_file"], index_col=False)
    end_ind = (int(_CONFIG["GMAIL_ACCOUNTS"]["end_index"]) 
              if int(_CONFIG["GMAIL_ACCOUNTS"]["end_index"]) != -1 
              else len(file_data))
    file_data = file_data.loc[int(_CONFIG["GMAIL_ACCOUNTS"]["start_index"]): end_ind]
    
    def wrapper(*args, **kwargs):
        # Create thread pool for parallel processing
        pool = ThreadPool(int(_CONFIG["BROWSER"]["parallel_browsers"]))
        
        if op == 1:  # Email sending operation
            # Load recipients data
            end_ind_r = (int(_CONFIG["RECIPIENT"]["end_index"]) 
                        if int(_CONFIG["RECIPIENT"]["end_index"]) != -1 
                        else len(recipients))
            recipients = pandas.read_excel(_CONFIG["RECIPIENT"]["recipient_emails_excel"], index_col=False)
            recipients = recipients.loc[int(_CONFIG["RECIPIENT"]["start_index"]): end_ind_r]
            
            # Execute function with recipients for each account
            pool.map(partial(func, recipients), file_data.iterrows())
        else:  # Account management operation
            # Execute function for each account
            pool.map(func, file_data.iterrows())
    
    return wrapper


def logging_error_screenshot(message, driver):
    """
    Log error message and take screenshot for debugging.
    
    Args:
        message (str): Error message to log
        driver: WebDriver instance for screenshot
    """
    # Generate unique filename
    image_name = datetime.datetime.now().strftime("%d%m%Y%H%M")
    for _ in range(5):
        image_name += random.choice(string.ascii_letters)
    
    # Save screenshot
    image_path = os.path.abspath(f"{_FOLDER_CONTAINING_SCREENSHOTS}/{image_name}.png")
    driver.save_screenshot(image_path)
    
    # Log error with screenshot path
    LOGGER.exception(f"{message} => file:///{image_path}")
    driver.refresh()


######################################################################################
# -----------------------------------Main Execution---------------------------------#
######################################################################################

if __name__ == "__main__":
    try:
        # Initialize global lock for thread safety
        lock = Lock()
        
        # Create necessary directories
        os.makedirs(_FOLDER_CONTAINING_ALL_PROFILES, exist_ok=True)
        os.makedirs(_FOLDER_CONTAINING_SCREENSHOTS, exist_ok=True)
        
        # Initialize daily limit tracking
        limit_reached_rec = pandas.DataFrame(columns=["Email"])
        if os.path.exists(_LIMIT_TRACK_FILE):
            limit_reached_rec = pandas.read_excel(_LIMIT_TRACK_FILE, index_col=False).fillna("")
        
        # Display menu and get user selection
        menu = {"Add Gmail": add_gmail, "Send Emails": send_email}
        menu_list = list(menu.keys())
        selected_option = cutie.select(menu_list)
        
        # Execute selected operation
        parallel_browsing(menu[menu_list[selected_option]], selected_option)()
        
    except Exception as e:
        LOGGER.exception("Critical Error Occurred!")
        print(f"Critical error: {e}")
        input("Press Enter to exit...")