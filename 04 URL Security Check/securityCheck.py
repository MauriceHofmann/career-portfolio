"""
Datei:      Security_URL_Check.py
Autor:      Maurice Hofmann
Erstellt:   08.02.2024
Lizenz:     public

Beschreibung:
    Dieses Skript überprüft die Sicherheit von URLs, die aus CSV-, TXT- oder Excel-Dateien
    geladen werden. Es validiert die URLs, führt Sicherheitschecks mithilfe von testssl.sh
    durch, erstellt einen ausführlichen Sicherheitsbericht in Excel und versendet
    optional eine E-Mail mit den Ergebnissen.

Voraussetzungen:
    - Eine Datei mit URLs im CSV-, TXT- oder Excel-Format.
    - Windows Subsystem for Linux (WSL) oder ein Linux-/Mac-System für die Ausführung von testssl.sh.
    - Zugriff auf den SMTP-Server für den Versand der Ergebnisse per E-Mail.

Benötigte Bibliotheken:
    - pandas: Zum Einlesen von Dateien und Erstellen von Reports.
    - numpy: Für die Aufteilung von Daten für Multithreading.
    - openpyxl: Zum Erstellen und Bearbeiten von Excel-Dateien.
    - validators: Zur Validierung von URLs und E-Mail-Adressen.
    - subprocess: Für die Ausführung externer Befehle (testssl.sh).
    - smtplib & email.message: Zum Versand von E-Mail-Reports.
    - threading & multiprocessing: Für parallele Bearbeitung der URLs.
    - shutil, os, sys, time, json, platform: Für allgemeine System- und Dateifunktionen.

Verwendung:
    1. Lege die Datei mit URLs in das Verzeichnis "Files_Security_Check".
    2. Passe ggf. die Konfiguration in "config.JSON" an (checkAll, receiver_mail).
    3. Starte das Skript: `python Security_URL_Check.py`.
    4. Nach Abschluss:
       - Erfolgreich geprüfte URLs werden in einem Excel-Report gespeichert.
       - Fehlerhafte oder ungültige URLs werden separat protokolliert.
       - Optional wird der Bericht per E-Mail versendet, wenn eine gültige Empfängeradresse angegeben ist.
       - Die Originaldatei wird in das Verzeichnis "Archive/" verschoben.
"""

# --------------------------------
#           Imports
# --------------------------------
import os
import re
import sys
import time
import json
import shutil
import random
import smtplib
import platform
import threading
import subprocess
import validators
import numpy as np
import pandas as pd
import multiprocessing
from pathlib import Path
from datetime import date
from openpyxl import load_workbook
from email.message import EmailMessage


logfile_directory = f"Logfiles/Logfiles_{date.today()}"

urls = []
invalid_url = pd.DataFrame(columns=['URL', 'Status'])
security_check_failed = pd.DataFrame(columns=['URL', 'Status'])
security_check_successful = pd.DataFrame(columns=['URL', 'Status', 'Overall Grade','Grade cap reasons', 'Grade warning'])


# --------------------------------
#           Functions
# --------------------------------
def read_file(filepath: str):
    """Reads a file from a given file path and returns its content as a DataFrame.

    Args:
        filepath (str): The file path to the file to be read.

    Returns:
        DataFrame: A DataFrame containing the content of the read file.

    Raises:
        SystemExit: If the file type is not supported (only 'xlsx', 'txt', and 'csv' are supported).

    """
    
    filetype = str(filepath).split('.')[-1].lower()

    if filetype == "csv":
        url_df = pd.read_csv(filepath, header=None)

    elif filetype == "txt":
        url_df = pd.read_table(filepath, header=None)

    elif filetype == "xlsx" or filetype == "xls":
        url_df = pd.read_excel(filepath, header=None)

    else:
        sys.exit(f"Ungültiger Dateityp '{filetype}'. Unterstützte Formate sind nur 'xlsx', 'txt' und 'csv'.")
    
    return url_df



def archive_file(filepath: str):
    """Archives a file by moving it to the 'Archive' directory with a timestamped filename.

    Args:
        filepath (str): The path to the file to be archived.

    Returns:
        None
    """

    if not os.path.exists("Archive/"):
        os.makedirs("Archive/")

    filename = os.path.basename(filepath)
    destination_path = f"Archive/{time.strftime('%Y_%m_%d-%H-%M')}_{filename}"

    shutil.move(filepath, destination_path)
    


def sample_urls(urls: list):
    """Selects a random sample of URLs from the given list

    Args:
        urls (list): A list of URLs from which to sample

    Returns:
        _type_: A randomly selected sample of URLs from the input list
    """

    sample_size = max(1, len(urls) // 10)
    sample_list = random.sample(urls, sample_size)

    return sample_list




def check_url_validity(url: str):
    """Checks the validity of a given URL using the validators library.

    Args:
        url (str): The URL to be checked

    Returns:
        bool: True if the URL is valid, False otherwise

    Notes:
    - Uses the validators.url() function from the validators library to check URL validity
    - If the URL is valid, return True
    - If the URL is invalid, appends it to the 'invalid_url' list and returns False
    """

    if validators.url(url) == True:
        return True
    else:
        invalid_url.loc[len(invalid_url)] = url, "Invalid URL"
        return False



def check_email_validity(email: str):
    """Checks the validity of a given email using the validators library.

    Args:
        email (str): The email to be checked

    Returns:
        bool: True if the email is valid, False otherwise

    Notes:
    - Uses the validators.email() function from the validators library to check email validity
    - If the email is valid, return True
    - If the email is invalid, return False
    """

    if validators.email(email) == True:
        return True
    else:
        return False




def check_url_security(url: str):
    """Conducts a security check on the specified URL using the testssl.sh tool

    Args:
        url (str): The URL to be subjected to the security check

    Returns:
        bool: True if the check process is error-free, False otherwise

    Notes:
    - Executes the testssl.sh tool using the Windows Subsystem for Linux (WSL) to perform the security check
    - Checks if a folder for the log files exists, If not than create
    - Writes the output ofthe security check to a log file located 'Logfiles_{date.today()' directory with the name '{url}.log'
    - Parses the output to extract information such as the overall grade, grade cap reasons and grade warnings
    - If the security check is error-free (return code is 0), updates the 'security_check_successful' list and returns True
    - If the security check fails (return code is not 0), appends the URL to the 'security_check_failed' list and returns False
    """

    check_started = False

    overall_grade = ""
    grade_cap_reasons = ""
    grade_warning = ""
    
    i = 0

    wsl_command = f'./testssl.sh/testssl.sh --warnings off {url}'
    
    if platform.system() == "Windows":
        command = ['wsl', 'bash', '-c', wsl_command]
    else:
        command = ["sh", "-c", wsl_command]

    with open (f"{logfile_directory}/{url.split('//')[1]}.log", 'w') as logfile: 
        process = subprocess.Popen(command, stdout = subprocess.PIPE, stderr = subprocess.STDOUT, universal_newlines = True)  
        for line in process.stdout:

            ansicode_pattern = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')
            line_value = ansicode_pattern.sub('', line)

            sys.stdout.write(line)
            logfile.write(line_value)


            if "Start" in line and url.split('//')[1] in line:
                print("CHECK", check_started)
                if check_started == True:
                    print(i)
                    i += 1
                    return
                
                check_started = True

    
            if "Overall Grade" in line_value:
                overall_grade = str(line.split("Overall Grade")[-1].strip())
                overall_grade = re.sub(r'\x1b\[[0-9;]*m', '', overall_grade)
            elif "Grade cap reasons" in line:
                grade_cap_reasons = str(line.split("Grade cap reasons")[-1].strip())
                grade_cap_reasons = re.sub(r'\x1b\[[0-9;]*m', '', grade_cap_reasons)
            elif "Grade warning" in line:
                grade_warning = str(line.split("Grade warning")[-1].strip())
                grade_warning = re.sub(r'\x1b\[[0-9;]*m', '', grade_warning)
    
    return_code = process.wait()
    logfile.close()

    if return_code != 0:
        security_check_failed.loc[len(security_check_failed)] = url
        print("LOG", url)
        return False
    else:
        security_check_successful.loc[len(security_check_successful)] = url, "Success", overall_grade, grade_cap_reasons, grade_warning
        print("LOG", url)
        return True
    


def create_report(security_check_successful: pd.DataFrame, invalid_url: pd.DataFrame, security_check_failed: list):
    """Creates a security report based on the successful runs of URL security check

    Args:
        security_check_successful (list): A list containing the successful runs of URL security check

    Returns:
        str: The filename of the generated security report

    Notes:
    - Copies the Excel template named 'URL_Report_Template.xlsx to create the current report
    - The report filename is generated based on the current date and saved as 'Security_Check_{current.date}.xlsx'
    - Loads the newly created report file using openpyxl
    - Accesses the first worksheet of the report file
    - Appends each row form the 'security_check_successful' list to the worksheet
    - Saves the modified report file
    """

    # Copy Excel Template for current report
    template_filename = "Template/URL_Report_Template.xlsx"
    report_filename = f"{logfile_directory}/Security_Check_{time.strftime('%Y_%m_%d-%H-%M-%S')}.xlsx"

    shutil.copy(template_filename, report_filename)

    report_file = load_workbook(report_filename)
    report_spreadsheet = report_file.worksheets[0]

    # Status: Invalid
    for index, row in invalid_url.iterrows():
        report_spreadsheet.append(row.tolist())

    # Status: Failed
    for index, row in security_check_failed.iterrows():
        report_spreadsheet.append(row.tolist())

    # Status: Success
    for index, row in security_check_successful.iterrows():
        report_spreadsheet.append(row.tolist())

    report_file.save(report_filename) 

    return report_filename
    


def send_report(report_filename:str, receiver_mail: str): 
    """Sends an email report summarizing the results of the Security URL Check

    Args:
        report_filename (str): The filename of the generated security report to be attached
        recever_mail (str): The mail address of the receiver

    Notes:
    - Generates an HTML email body with a summary of key values and information about the security verification process
    - Uses the Outlook application to create a new email
    - Sets the email subject to include the date of the secuity verification results
    - Attaches the generated security report to the email
    - Sends the email to the recipient specified in the configuration file
    """

    REPORT_SUBJECT = f"[Automatic Security Log] Summary of URL verification results on {date.today()}"
    REPORT_MESSAGE = f"""
    <html>
        <body>
        <p>
            Dear User,<br><br>
            Thank your for your interest in the Security URL Check. We would like to inform you about the results of the autotmatic security verification process you initiated on {date.today()}.<br>
            Here is a brief summary of the key values:<br>
        </p> 

        <table border="2">
            <tr>
                <th width="250">Status</th>
                <th width="100">Occurrences</th>
            </tr>

            <tr>
                <td>Invalid URLs</td>
                <td align="center">{len(invalid_url)}</td>
            </tr>
            <tr>
                <td>Check process failed</td>
                <td align="center">{len(security_check_failed)}</td>
            </tr>
            <tr>
                <td>Grade is ok</td>
                <td align="center">{(security_check_successful["Overall Grade"].isin(["A+", "A"])).sum()}</td>
            </tr>
            <tr>
                <td>Grade is not ok</td>
                <td align="center">{(security_check_successful["Overall Grade"].isin(["A+", "A"]) == False).sum()}</td>
            </tr>

            <tfoot>
                <td bgcolor="lightgrey"><strong>Checked URL</strong></td>
                <td bgcolor="lightgrey" align="center"><strong>{len(urls)}</strong></td>
            </tfoot>
        </table>
        
        <p>
            For a detailed overview of the security results, please refer to the attachment of the email. There, your will find a comprehensive analysis of the checked URLs and their security assessments.<br>
            Additionally, detailed log files for each URL verification can be accessed <a href=''>here</a>.<br>
            Please note that this email is generated automatically. If you have any further questions or concerns, feel free to reach out to us.<br><br>
            Thank you for using our Security URL Check. <br><br>
            Best regards, <br>
            XXXX
        </p>     
    </body>
    </html>
    """

    server = smtplib.SMTP('', 25)
    server.ehlo()
    server.starttls()
    server.ehlo()

    message = EmailMessage()
    message.set_content(REPORT_MESSAGE, subtype="html")
    message['Subject'] = REPORT_SUBJECT

    if len(sys.argv) > 1:
        sender_mail = sys.argv[1]
    else:
        sender_mail = ""

    with open(report_filename, 'rb') as attachment:
        attachment_data = attachment.read()
        message.add_attachment(attachment_data, maintype='application', subtype='octet-stream', filename=report_filename)    
    
    server.sendmail(sender_mail, receiver_mail, message.as_string())
    
    server.quit()
    


def thread_work(data: list):
    for url in data:

        print("URL:", url)

        url_validity = check_url_validity(url)
        if url_validity is False:
            continue 

        process_status = check_url_security(url)
        if process_status is False:
            continue 




# --------------------------------
#           Main-Guard
# --------------------------------
if __name__ == "__main__":

    files = [file for file in os.listdir("Files_Security_Check") if file != ".gitignore"]
    
    print(files)
    
    if len(files) == 1:
        filepath = f"./Files_Security_Check/{files[0]}"
    else:
        sys.exit("Keine Datei gefunden")

    with open("config.JSON") as config_json:
        config = json.load(config_json)
        check_all = config["checkAll"] 

    urls = read_file(filepath)

    if check_all is False:
        urls = sample_urls(urls.values.tolist())


    if not os.path.exists(logfile_directory):
        os.makedirs(logfile_directory)


    if multiprocessing.cpu_count() > 1:
        threads = multiprocessing.cpu_count() -1
    else:
        threads = multiprocessing.cpu_count()

    print("THREADS", threads)

    threads = 1

    thread_list = []
    data_parts = np.array_split(urls, threads)

    print(data_parts)

    for i, part in enumerate(data_parts):
        if len(part) > 0:

            thread = threading.Thread(target=thread_work, args=(list(part[0]),))
            thread.start()

            thread_list.append(thread)

    # Wait if all threads finished
    for thread in thread_list:
        thread.join()


    report_filename = create_report(security_check_successful, invalid_url, security_check_failed)

    report = open(report_filename)
    report_filepath= os.path.realpath(report.name)

    
    with open("config.JSON") as config_json:
        config = json.load(config_json)
        receiver_mail = config["receiver_mail"] 
    
    email_validity = check_email_validity(receiver_mail)

    if receiver_mail !=  "" and email_validity == True:
        # Send report via mail
        send_report(report_filepath, receiver_mail)

    archive_file(filepath)
