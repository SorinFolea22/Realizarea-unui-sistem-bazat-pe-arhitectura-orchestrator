import imaplib
import email
import os
from email.header import decode_header
import docx
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from config import EMAIL_USERNAME, EMAIL_PASSWORD

# Detaliile pentru conectarea la serverul de email
IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Configurație Google Sheets
SHEET_NAME = 'Cerere Concediu'  # Numele fișierului Google Sheet
CREDENTIALS_FILE = '.json'  # Calea către fișierul de acreditări JSON

TOTAL_ZILE_CONCEDIU = 21  

# Conectare la serverul de email
def connect_to_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        mail.select("inbox")
        print("Connected to email server.")
        return mail
    except Exception as e:
        print(f"Failed to connect to email server: {e}")
        return None

# Conectare la Google Sheets
def connect_to_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open(SHEET_NAME).sheet1
        print("Connected to Google Sheets.")
        return sheet
    except Exception as e:
        print(f"Failed to connect to Google Sheets: {e}")
        return None

# Verifică și adaugă capul de tabel dacă lipsește
def ensure_sheet_headers(sheet):
    headers = ['Nume', 'Perioada', 'Data înregistrării', 'Număr zile concediu', 'Număr zile concediu rămase']
    existing_headers = sheet.row_values(1)
    if existing_headers != headers:
        sheet.insert_row(headers, 1)
        print("Headers added to Google Sheets.")

# Trimite emailul de răspuns utilizând SMTP simplu
def send_email_response(to_address, subject, body):
    try:
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = EMAIL_USERNAME
        msg['To'] = to_address

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USERNAME, to_address, msg.as_string())
        print(f"Email sent to {to_address}.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Verifică dacă emailul are atașamente și descarcă fișierul Word dacă îndeplinește condițiile
def check_attachments(msg):
    print("Checking for attachments...")
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue
        file_name = part.get_filename()
        if file_name:
            decoded_name = decode_header(file_name)[0][0]
            if isinstance(decoded_name, bytes):
                decoded_name = decoded_name.decode()
            print(f"Decoded attachment name: {decoded_name}")
            if "Cerere Concediu" in decoded_name and decoded_name.endswith(".docx"):
                file_path = os.path.join(os.getcwd(), decoded_name)
                with open(file_path, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Downloaded attachment to: {file_path}")
                return file_path
    return None

# Funcția pentru calcularea numărului de zile lucrătoare
def calculate_workdays(start_date, end_date):
    num_workdays = 0
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5:  # Luni (0) până Vineri (4)
            num_workdays += 1
        current_date += timedelta(days=1)
    return num_workdays

# Funcția pentru calcularea zilelor de concediu rămase
def calculate_remaining_days(sheet, name, current_days):
    records = sheet.get_all_records()
    total_days_taken = sum(record['Număr zile concediu'] for record in records if record['Nume'] == name)
    remaining_days = TOTAL_ZILE_CONCEDIU - (total_days_taken + current_days)
    return max(remaining_days, 0)

# Extrage numele, perioada, data înregistrării și numărul de zile de concediu din fișierul Word
def extract_details_from_docx(file_path):
    try:
        print(f"Extracting details from {file_path}")
        doc = docx.Document(file_path)
        details = {'Nume': '', 'Perioada': '', 'Data înregistrării': '', 'Număr zile concediu': '', 'Număr zile concediu rămase': ''}
        full_text = "\n".join([para.text for para in doc.paragraphs])
        print(f"Full text: {full_text}")

        # Folosim regex pentru a extrage numele, perioada și data înregistrării
        name_match = re.search(r'Subsemnatul\(a\)\s+([A-Z ]+)', full_text)
        period_match = re.search(r'începând cu data de (\d{2}\.\d{2}\.\d{4})\s+și pâna la data de (\d{2}\.\d{2}\.\d{4})', full_text)
        date_match = re.search(r'Data\s*:\s*(\d{2}\.\d{2}\.\d{4})', full_text)

        if name_match:
            details['Nume'] = name_match.group(1).strip()
        if period_match:
            start_date = datetime.strptime(period_match.group(1), '%d.%m.%Y')
            end_date = datetime.strptime(period_match.group(2), '%d.%m.%Y')
            details['Perioada'] = f"{period_match.group(1)} - {period_match.group(2)}"
            details['Număr zile concediu'] = calculate_workdays(start_date, end_date)  # Calculăm numărul de zile de concediu
        if date_match:
            details['Data înregistrării'] = date_match.group(1).strip()

        print(f"Extracted details: {details}")
        return details if details['Nume'] and details['Perioada'] and details['Data înregistrării'] and details['Număr zile concediu'] else None
    except Exception as e:
        print(f"Failed to extract details from {file_path}: {e}")
        return None

# Salvează detaliile într-un Google Sheet și trimite emailul de răspuns
def save_to_google_sheets(details, sheet, email_address):
    try:
        print("Saving details to Google Sheets.")
        remaining_days = calculate_remaining_days(sheet, details['Nume'], details['Număr zile concediu'])
        if remaining_days > 0:
            details['Număr zile concediu rămase'] = remaining_days
            existing_data = sheet.get_all_values()
            next_row = len(existing_data) + 1
            sheet.update(f'A{next_row}', [[details['Nume'], details['Perioada'], details['Data înregistrării'], details['Număr zile concediu'], details['Număr zile concediu rămase']]])
            print("Details saved to Google Sheets.")
            # Trimite email de confirmare
            subject = "Cerere concediu înregistrată"
            body = f"Cererea dumneavoastră a fost înregistrată, mai aveți un număr de {remaining_days} zile de concediu."
        else:
            # Trimite email de refuz
                        subject = "Cerere concediu respinsă"
                        body = "Cererea dumneavoastră nu a fost înregistrată deoarece nu mai aveți zile de concediu suficiente."
        
        send_email_response(email_address, subject, body)
    except Exception as e:
        print(f"Failed to save details to Google Sheets: {e}")

# Funcția principală care verifică emailurile necitite și extrage detaliile necesare
def check_unread_emails_and_extract_details():
    mail = connect_to_email()
    sheet = connect_to_google_sheets()
    if not mail or not sheet:
        return

    try:
        ensure_sheet_headers(sheet)
        status, messages = mail.search(None, '(UNSEEN)')
        if status != 'OK':
            print("No unread emails found.")
            return

        mail_ids = messages[0].split()
        print(f"Found {len(mail_ids)} unread emails.")

        for mail_id in mail_ids:
            status, msg_data = mail.fetch(mail_id, "(RFC822)")
            if status != 'OK':
                print(f"Failed to fetch email with ID {mail_id}")
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    email_address = email.utils.parseaddr(msg.get('From'))[1]
                    print(f"Processing email ID {mail_id}")
                    file_path = check_attachments(msg)
                    if file_path:
                        details = extract_details_from_docx(file_path)
                        if details:
                            save_to_google_sheets(details, sheet, email_address)
                        else:
                            print("No details extracted from the document.")
                    else:
                        print("No valid attachment found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        mail.logout()

if __name__ == "__main__":
    check_unread_emails_and_extract_details()

