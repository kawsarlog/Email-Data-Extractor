import imaplib
import traceback
import re
import os
import datetime
import email
from bs4 import BeautifulSoup
from email.header import decode_header
from openpyxl import load_workbook, Workbook

email_folder_name = 'Neue Leads'
unique_id_txt_filename = 'unique_ids.txt'

# IMAP configuration
imap_server = "imaps.udag.de"
username = "example@email.com"
password = "YourPassword123"
port = 993

# create filename for this run. Unique date and time for each file.
scrapeDateTime = datetime.datetime.now().strftime("%B_%d_%Y_%H%M")
excel_sheet_name = 'Timm email Data_' + scrapeDateTime + '.xlsx'

# Text File Create if not exist
if not os.path.isfile(unique_id_txt_filename):
    with open(unique_id_txt_filename, 'w') as file:
        pass

# Read Text File
unique_ids = open(unique_id_txt_filename, encoding='utf-8').read().splitlines()

# Write unique_ids in unique_id_txt_filename
def unique_id_txt(unique_id_txt_filename, unique_id):
    with open(unique_id_txt_filename, 'a', encoding='utf-8') as file:
        file.write(unique_id + '\n')

# Write Headline, Data and create a new excel sheet
def xl_write(sheet_name, Datas):
    wb = Workbook()
    ws = wb.active
    
    # Write Headlines
    headlines = Datas[0]
    ws.append(headlines)
    
    # Write Data on Excel sheet
    for Data in Datas[1:]:
        ws.append(Data)
        
    # Saveb Excel Sheet
    wb.save(sheet_name)

def get_all_emails_list(imap_server, port, username, password, email_folder_name):

    print(f'STEP 01: Connecting On Email Server at [{email_folder_name}]')

    # Connect to the IMAP server
    mail = imaplib.IMAP4_SSL(imap_server, port)

    # Login to the mailbox
    mail.login(username, password)

    # Select the mailbox (in this case, the inbox)
    mail.select(f'"INBOX.{email_folder_name}"')

    # Search for all emails in the inbox
    status, messages = mail.search(None, "ALL")

    # Get the list of email IDs
    email_ids = messages[0].split()

    print(f">>Connecting on Email Server status: [{status}]")
    print('STEP 02: Sucessfully Retrive email_ids')
    
    
    # Create a list to store email information
    all_emails = []
    
    print('STEP 03: Iterate through email IDs')
    print('----------------------------------')
    
    # Iterate through email IDs
    for sl, email_id in enumerate(email_ids[::-1]):
        # Fetch the email
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        # Extract the email content
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        # Store email information in a dictionary
        email_info = {
            "Subject": email_message["Subject"],
            "From": email.utils.parseaddr(email_message["From"]),
            "Date": email_message["Date"],
            "Content": 'Lorem ipsum'
        }

        # If the email is multipart, store the text content
        if email_message.is_multipart():
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":

                    try: email_info["Content"] = part.get_payload(decode=True).decode("utf-8")
                    except: email_info["Content"] = part.get_payload(decode=False)
        
        #if data is not available get from HTML
        if 'e-mail:' not in email_info["Content"].lower():
            for part in email_message.walk():
                if part.get_content_type()=='text/html':
                    
                    try: html_content = part.get_payload(decode=True).decode("utf-8")
                    except: html_content = part.get_payload(decode='False')
                    soup = BeautifulSoup(html_content, 'html.parser')
                    email_info["Content"] = soup.getText()
        
        
        print(f"{sl+1}>> From: {email.utils.parseaddr(email_message['From'])} and Date: {email_message['Date']}")
        
        # Append the email information to the list
        all_emails.append(email_info)

        
        
    print('----------------------------------')
    print(f"Total [{len(all_emails)}] Email Founds in '{email_folder_name}'")
    print('STEP 04: Logout from the mailbox')
    mail.logout()
    
    return all_emails

def get_Anrede(email_details):
    regax_eles = [
        'Anrede:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Anrede..[\r|\n]*(.*?).[\r|\n]*Vorname',
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_Vorname(email_details):
    regax_eles = [
        r'Vorname:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Vorname..[\r|\n]*(.*?)[\r|\n]',
        r'Vorname:.?[\r|\n]*(.*?)[\r|\n]',
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_Nachname(email_details):
    regax_eles = [
        r'Nachname:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Nachname..[\r|\n]*(.*?)[\r|\n]',
        r'Nachname:[\r|\n]*(.*?)[\r|\n]',

        
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_email(email_details):
    regax_eles = [
        r'E-Mail:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'E-Mail:?.?[\r|\n]*.?<mailTo:(.*?)>',
        r'E-Mail:[\r|\n]*(.*?)[\r|\n]',        
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0].lower().split('>')[0]
            
            if 'mailto:' in output_data:
                return output_data.split('mailto:')[-1].strip(' ').strip('\n').strip('\t').strip('>')
            
            
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_Telefon(email_details):
    regax_eles = [
        r'Telefon:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Telefon:?.?[\r|\n]*(.*?)[\r|\n]',
        r'Telefon:?.?\t[\r|\n]*(.*?)[\r|\n]',
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_Ihre_aktuelle_Beschaftigung(email_details):
    regax_eles = [
        r'Ihre aktuelle Beschäftigung:.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Ihre aktuelle Beschäftigung:[\r|\n]*(.*?)[\r|\n]',
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data

def get_Beziehen_Sie_weitere_andere_Leistungen(email_details):
    regax_eles = [
        r'Beziehen Sie weitere/andere Leistungen\?.?\(Krankengeld,.?etc...\).?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Beziehen Sie weitere/andere Leistungen\?.?\(Krankengeld,.?etc...\)[\r|\n]*(.*?)[\r|\n]',
        
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data


def get_Wie_sin_Sie_auf_uns_aufmerksam_geworden(email_details):
    regax_eles = [
        r'Wie sind Sie auf uns aufmerksam geworden\?.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Wie sind Sie auf uns aufmerksam geworden\?[\r|\n]*(.*?)[\r|\n]',        
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data


def get_Wie_bist_du_auf_uns_aufmerksam_geworden(email_details):
    regax_eles = [
        r'Wie bist du auf uns aufmerksam geworden\?.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Wie bist du auf uns aufmerksam geworden\?[\r|\n]*(.*?)\r',
        
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t')
        except:
            pass
    return output_data


def get_Fragen(email_details):
    regax_eles = [
        r'Fragen.?.?.?Wünsche:?.?[\r|\n|\t]*(.*?)[\r|\n|\t]',
        r'Fragen.?/.?Wünsche:?-?.?[\r|\n]*-? ?(.*?).?[\r|\n]',
        r'Fragen.?/.?Wünsche:?.?[\r|\n]*(.*?)[\r|\n]*',  
        r'Fragen.?/.?Wünsche:[\r|\n]*(.*?)[\r|\n]*',
    ]
    
    output_data = ''
    for regax_ele in regax_eles:
        try:
            output_data = re.findall(regax_ele, email_details)[0]
            return output_data.strip(' ').strip('\n').strip('\t').strip('-')
        except:
            pass
    return output_data

def strip_unique_id_email(unique_id_email_text):
    strip_strs = ['-', '.', ' ', "'", '"', ',', '+', ')', '(', '_', '$', '%', '\n', '\t', '\r', '@']
    output_text = unique_id_email_text
    
    for strip_str in strip_strs:
        output_text = output_text.replace(strip_str, '')
    
    return output_text

# Get all_emails
all_emails = get_all_emails_list(imap_server, port, username, password, email_folder_name)

# Debugging Single list rows
# unique_ids = []
# all_emails_single = [all_emails[-1]]
# re.findall(r'Fragen / Wünsche::?.?[\r|\n|\t]*(.*?)[\r|\n|\t]', all_emails[5]['Content'])[0]
# all_emails[5]['Content']
# print(all_emails[5]['Content'])

# Excel Sheet Headlines
headlines = ['Anrede', 'Vorname', 'Nachname', 'email', 'Telefon', 'Ihre_aktuelle_Beschaftigung', 'Beziehen_Sie_weitere_andere_Leistungen', 'Wie_sin_Sie_auf_uns_aufmerksam_geworden', 'Wie_bist_du_auf_uns_aufmerksam_geworden', 'Fragen']

# Whole Data List
write_datas = [headlines, ]

print('STEP 05: Print all email information')
print()
print()
for idx, email_info in enumerate(all_emails, start=1):
    print('___________________________________________')
    email_info_Subject = email_info['Subject']
    email_info_From = email_info['From']
    email_info_Date = email_info['Date']
    
    print(f"Subject: {email_info_Subject}")
    print(f"From: {email_info_From}")
    print(f"Date: {email_info_Date}")
    
    unique_id_email_text = str(email_info_Subject)[-1] + str(email_info_From) + str(email_info_Date)
    unique_id = strip_unique_id_email(unique_id_email_text)
    
    if unique_id in unique_ids:
        print('              Record Already Exist [X]              ')
        continue
    
    if "Content" in email_info:
        print('              New Record Found [OK]              ')
        
        email_details = email_info["Content"]

        Anrede = get_Anrede(email_details)
        if len(Anrede) < 2: get_Anrede(email_details.replace('\t', ''))
        print(f"Anrede: {Anrede}")

        Vorname = get_Vorname(email_details)
        if len(Vorname) < 2: get_Vorname(email_details.replace('\t', ''))
        print(f"Vorname: {Vorname}")

        Nachname = get_Nachname(email_details)
        if len(Nachname) < 2: get_Nachname(email_details.replace('\t', ''))
        print(f"Nachname: {Nachname}")
        
        email = get_email(email_details)
        if len(email) < 2: get_email(email_details.replace('\t', ''))
        print(f"email: {email}")
        
        Telefon = get_Telefon(email_details)
        if len(email) < 2: get_email(email_details.replace('\t', ''))
        print(f"Telefon: {Telefon}")
        
        Ihre_aktuelle_Beschaftigung = get_Ihre_aktuelle_Beschaftigung(email_details)
        print(f"Ihre_aktuelle_Beschaftigung: {Ihre_aktuelle_Beschaftigung}")

        Beziehen_Sie_weitere_andere_Leistungen = get_Beziehen_Sie_weitere_andere_Leistungen(email_details)
        print(f"Beziehen_Sie_weitere_andere_Leistungen: {Beziehen_Sie_weitere_andere_Leistungen}")
        
        Wie_sin_Sie_auf_uns_aufmerksam_geworden = get_Wie_sin_Sie_auf_uns_aufmerksam_geworden(email_details)
        print(f"Wie_sin_Sie_auf_uns_aufmerksam_geworden: {Wie_sin_Sie_auf_uns_aufmerksam_geworden}")
        
        Wie_bist_du_auf_uns_aufmerksam_geworden = get_Wie_bist_du_auf_uns_aufmerksam_geworden(email_details)
        print(f"Wie_bist_du_auf_uns_aufmerksam_geworden: {Wie_bist_du_auf_uns_aufmerksam_geworden}")

        Fragen = get_Fragen(email_details)
        if len(Fragen) < 2: get_Fragen(email_details.replace('\t', ''))
        print(f"Fragen: {Fragen}")
        
        # ---re.findall(r'Fragen.?/.?Wünsche:?-?.?[\r|\n]*-? ?(.*?).?[\r|\n]', email_details)[0]
        # ---email_details
        
        single_write_data = Anrede, Vorname, Nachname, email, Telefon, Ihre_aktuelle_Beschaftigung, Beziehen_Sie_weitere_andere_Leistungen, Wie_sin_Sie_auf_uns_aufmerksam_geworden, Wie_bist_du_auf_uns_aufmerksam_geworden, Fragen
        write_datas.append(single_write_data)
        
        # Added unique_ids List
        unique_ids.append(unique_id)
        
        # Added unique_ids in Text File
        unique_id_txt(unique_id_txt_filename, unique_id)

print(f'STEP 06: Write Data on excel Sheet: [{excel_sheet_name}]')
xl_write(excel_sheet_name, write_datas)
print('--------------------')
print('Execution Compleded')
print('--------------------')
