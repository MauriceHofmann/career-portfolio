"""
File:        MBI_Data_Crawler.py
Author:      Maurice Hofmann
Created:     01.02.2024
License:     Public

Description:
    This script reads data from CSV files, validates email addresses,
    uses Selenium to access a web application to retrieve additional
    user information, and exports the collected data to an Excel file.
    Invalid or incomplete records are logged for review.

Requirements:
    - A CSV file containing employee email addresses and any additional
      relevant information.
    - Access to the web application used to query user information.
    - Python libraries: pandas, selenium, openpyxl, colorlog, validators, requests.

Required Libraries:
    - pandas: For reading and processing CSV files and exporting data to Excel.
    - selenium: For web automation and retrieving user information.
    - openpyxl: For creating and editing Excel files.
    - requests: For API queries using a session cookie.
    - colorlog & logging: For colored log output in the console.
    - validators: For validating email addresses.
    - pathlib: For cross-platform file path handling.
    - csv: For automatic detection of CSV file delimiters.

Usage:
    1. Ensure that all required libraries are installed.
    2. Place the CSV file in the desired directory and set its path
       in the variable `FILE`.
    3. Run the script, e.g. `python MBI_Data_Crawler.py`.
    4. After processing, the collected employee data will be saved
       as an Excel file in the user's Downloads folder.
    5. All errors will be logged in a separate log file in the Downloads folder.
"""

# --------------------------------
#           Imports
# --------------------------------
import csv
import sys
import time
import ctypes
import logging
import colorlog
import requests
import validators
import pandas as pd
from pathlib import Path
from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# --------------------------------
#       Class Definition
# --------------------------------
class Employee:
    """Class representing an employee with core attributes.

    Attributes:
        userid (str): The unique user ID of the employee.
        mail (str): The employee's email address.
        plant (str): The plant or location where the employee works.
        costcenter (str): The cost center associated with the employee.
    """

    def __init__(self, userid=None, mail=None, plant=None, costcenter=None):
        self.userid = userid
        self.mail = mail
        self.plant = plant
        self.costcenter = costcenter


# --------------------------------
#           Variables
# --------------------------------
FILE = ""
CERTIFICATE = "t"

error_sum = pd.DataFrame(columns=['Error', 'User Data'])
employee_information_list = []


# --------------------------------
#           Functions
# --------------------------------
def configure_logger():
    """Configures the logger with colored output for console messages.

    The logger uses colorlog to display messages with colors:
    - Green for INFO
    - Yellow for WARNING
    - Red for ERROR
    """
    handler = colorlog.StreamHandler()
    handler.setFormatter(
        colorlog.ColoredFormatter(
            log_colors={
                'INFO': 'green',
                'WARNING': 'yellow',
                'ERROR': 'red',
            },
        )
    )

    logger = logging.getLogger()
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


def read_file(filepath: str):
    """Reads a CSV file and returns its content as a pandas DataFrame.

    Args:
        filepath (str): Path to the CSV file to read.

    Returns:
        pandas.DataFrame: DataFrame containing the file content.

    Raises:
        SystemExit: If the provided file type is not supported.
    """
    filetype = str(filepath).split('.')[-1].lower()

    if filetype == "csv":
        with open(FILE, 'r') as file:
            dialect = csv.Sniffer().sniff(file.readline())
            delimiter = dialect.delimiter

        users_df = pd.read_csv(filepath, delimiter=delimiter, header=0)
        users_df.columns = [col.lower() for col in users_df.columns]

    else:
        sys.exit(f"Ungültiger Dateityp '{filetype}'. Unterstütztes Format nur 'csv'.")
    
    return users_df


def login():
    """Logs into the web application and retrieves a session cookie.

    Uses Selenium Edge WebDriver to perform login and waits for successful redirection.

    Returns:
        tuple: (webdriver.Edge instance, cookie string)
    """
    cookie_value = ""

    driver = webdriver.Edge()
    driver.get("")

    WebDriverWait(driver, 180).until(EC.url_to_be(""))

    cookie_value = driver.get_cookie('TEX')['value']

    return driver, str(cookie_value)


def check_email_validity(mail: str):
    """Checks whether an email address is syntactically valid.

    Args:
        mail (str): The email address to validate.

    Returns:
        bool: True if the email is valid, otherwise False.
    """
    if validators.email(mail) is True:
        return True
    else:
        return False


def fetch_data(data: str, driver: webdriver, cookie: str):
    """Fetches employee information from a web source and stores it.

    Steps:
    1. Validates the email format.
    2. Uses Selenium to search for the user in a web application.
    3. Retrieves user ID from HTML.
    4. Calls an API endpoint with the session cookie.
    5. Parses and stores the data into an Employee object.

    Args:
        data (str): Single record (row) from user data DataFrame.
        driver (webdriver.Edge): Active Selenium WebDriver session.
        cookie (str): Authentication cookie for API requests.
    """
    mail = data[1]['email']
    mail_parts = mail.split('.')

    if mail_parts[-1] == 'io':
        mail = mail.replace("mercedes-benz.io", "mercedes-benz.com")

    email_validity = check_email_validity(mail)

    if not email_validity:
        error_message = "Invalid Mail"
        logging.error(f"{mail}\t {error_message}")
        error_sum.loc[len(error_sum)] = [error_message, list(data[1])]
        return

    driver.set_window_position(-10000, 0)

    try:
        driver.get(f"")
        wait = WebDriverWait(driver, 2)
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "Person_image__e2BGp")))
    except:
        error_message = "Mail not found"
        logging.error(f"{mail}\t {error_message}")
        error_sum.loc[len(error_sum)] = [error_message, list(data[1])]
        return

    html_code = driver.page_source
    pos_start = html_code.find('class="Person_container__utK-R containerLight" id="') + 51
    pos_end = pos_start + 7
    userid = html_code[pos_start:pos_end]

    ENDPOINT = f""
    HEADER = {'Cookie': f'TEX={cookie}'}

    try:
        response = requests.get(ENDPOINT, headers=HEADER, verify=CERTIFICATE)
    except:
        error_message = "Unable to load user"
        logging.error(f"{mail}\t {error_message}")
        error_sum.loc[len(error_sum)] = [error_message, list(data[1])]
        return

    if response.status_code == 200:
        employee_data = response.json()
        employee = Employee()

        try:
            employee.userid = employee_data['persons'][userid]['uid']
            employee.mail = employee_data['persons'][userid]['mail']
            employee.plant = employee_data['persons'][userid]['plant']
            employee.costcenter = employee_data['persons'][userid]['costCenter']

            employee_information_list.append(employee)
            logging.info(f"{mail}\t wurde erfolgreich exportiert")

        except:
            error_message = "Requested data not found"
            logging.error(f"{mail}\t {error_message}")
            error_sum.loc[len(error_sum)] = [error_message, list(data[1])]


def create_exel(employee_list: list):
    """Creates an Excel file from the collected employee data.

    Args:
        employee_list (list): List of Employee objects to be exported.

    Output:
        An Excel file saved in the user's Downloads folder.
    """
    employee_df = pd.DataFrame([vars(employee) for employee in employee_list])
    downloadpath = str(Path.home() / "Downloads")
    employee_df.to_excel(downloadpath + f"/Employee_Export_{time.strftime('%Y_%m_%d-%H-%M-%S')}.xlsx")


def error_handling(errors: pd.DataFrame):
    """Logs all encountered errors to a file.

    Args:
        errors (pandas.DataFrame): DataFrame containing all error messages and associated data.

    Output:
        A `.log` file in the user's Downloads folder.
    """
    downloadpath = str(Path.home() / "Downloads")
    filename = downloadpath + f"/Employee_Export_Errors_{time.strftime('%Y_%m_%d-%H-%M-%S')}.log"
    errors.to_csv(filename, sep='\t', index=False)