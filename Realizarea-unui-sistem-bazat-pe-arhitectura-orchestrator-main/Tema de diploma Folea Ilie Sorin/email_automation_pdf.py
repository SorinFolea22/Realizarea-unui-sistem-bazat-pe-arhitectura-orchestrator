import imaplib
import email
from email.header import decode_header
import os
import re
import pandas as pd
from PyPDF2 import PdfReader
import requests

# Email account credentials
username = ""
password = ""  

# Telegram bot credentials
telegram_bot_token = ''
telegram_chat_id = ''  

# Connect to the mail server
print("Connecting to the mail server...")
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)
mail.select("inbox")
print("Connected to the mail server.")

def download_attachments():
    # Search for unread emails
    print("Searching for unread emails...")
    status, messages = mail.search(None, 'UNSEEN')
    if status != "OK":
        print("Error searching for unread emails.")
        return

    email_ids = messages[0].split()
    if not email_ids:
        print("No unread emails found.")
        return

    print(f"Found {len(email_ids)} unread emails.")

    for email_id in email_ids:
        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        if status != "OK":
            continue

        attachment_downloaded = False  

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                if msg.is_multipart():
                    for part in msg.walk():
                        # Check if the part has an attachment
                        if part.get_content_maintype() == "multipart":
                            continue
                        if part.get("Content-Disposition") is None:
                            continue

                        filename = part.get_filename()
                        if filename and filename.endswith('.pdf'):
                            # Use raw string or double backslashes
                            filepath = os.path.join(r"C:/Users/Sorin/Desktop/Documente orchestrator task1", filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"Downloaded attachment: {filepath}")
                            attachment_downloaded = True
                            invoice_data = extract_invoice_data(filepath)
                            if invoice_data:
                                excel_file_path = r'C:/Users/Sorin/Desktop/facturi/facturi prelucrate/invoices.xlsx'
                                save_to_excel(invoice_data, excel_file_path)
                                send_file_via_telegram(excel_file_path)
                
        # Mark the email as read if any attachment was downloaded
        if attachment_downloaded:
            mail.store(email_id, '+FLAGS', '\\Seen')

def extract_invoice_data(pdf_path):
    invoice_data = {
        "Furnizor": {},
        "Cumparator": {},
        "Factura": {},
        "Produse": [],
        "Total": {}
    }
    try:
        with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            number_of_pages = len(reader.pages)
            text = ""
            for page_num in range(number_of_pages):
                page = reader.pages[page_num]
                text += page.extract_text()

            def extract_group(pattern, text):
                match = re.search(pattern, text)
                if match:
                    return match.group(1).strip()
                return None

            # Extract furnizor data
            invoice_data["Furnizor"]["Nume"] = extract_group(r'Furnizor\s*:\s*(.*?)\s*C\.I\.F\.', text)
            invoice_data["Furnizor"]["CIF"] = extract_group(r'C\.I\.F\.\s*:\s*(\d+)', text)
            invoice_data["Furnizor"]["Nr ord reg com"] = extract_group(r'Nr ord reg com\s*:\s*(.*?)\s*Sediul', text)
            invoice_data["Furnizor"]["Sediul"] = extract_group(r'Sediul\s*:\s*(.*?)\s*Judet', text)
            invoice_data["Furnizor"]["Judet"] = extract_group(r'Judet\s*:\s*(.*?)\s*Cont', text)
            invoice_data["Furnizor"]["Cont"] = extract_group(r'Cont\s*:\s*(.*?)\s*Banca', text)
            invoice_data["Furnizor"]["Banca"] = extract_group(r'Banca\s*:\s*(.*?)\s*Capital Social', text)
            invoice_data["Furnizor"]["Capital Social"] = extract_group(r'Capital Social\s*:\s*(.*?)\s*Punct de lucru', text)

            # Extract cumparator data
            invoice_data["Cumparator"]["Nume"] = extract_group(r'Cumparator\s*:\s*(.*?)\s*C\.I\.F\.', text)
            invoice_data["Cumparator"]["CIF"] = extract_group(r'C\.I\.F\.\s*:\s*(RO \d+)', text)
            invoice_data["Cumparator"]["Nr ord reg com"] = extract_group(r'Nr ord reg com\s*:\s*(.*?)\s*Sediul', text)
            invoice_data["Cumparator"]["Sediul"] = extract_group(r'Sediul\s*:\s*(.*?)\s*Judet', text)
            invoice_data["Cumparator"]["Judet"] = extract_group(r'Judet\s*:\s*(.*?)\s*Cont', text)
            invoice_data["Cumparator"]["Cont"] = extract_group(r'Cont\s*:\s*(.*?)\s*Banca', text)
            invoice_data["Cumparator"]["Banca"] = extract_group(r'Banca\s*:\s*(.*?)\s*FACTURA', text)

            # Extract factura data
            invoice_data["Factura"]["Serie"] = extract_group(r'SERIA\s*:\s*(.*?)\s', text)
            invoice_data["Factura"]["Numar"] = extract_group(r'NR. FACTURII\s*:\s*(.*?)\s', text)
            invoice_data["Factura"]["Data"] = extract_group(r'DATA \(zi/luna/an\)\s*:\s*(.*?)\s', text)

            # Extract product data
            product_pattern = re.compile(r'(\d+)\.\s+(.+?)\s+(\w+)\s+(\d+\.\d+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)')
            matches = product_pattern.findall(text)
            for match in matches:
                invoice_data["Produse"].append({
                    "Nr. crt.": match[0].strip(),
                    "Denumire": match[1].strip(),
                    "UM": match[2].strip(),
                    "Cantitate": match[3].strip(),
                    "Preț unitar (fără TVA)": match[4].strip(),
                    "Valoare (fără TVA)": match[5].strip(),
                    "Valoare TVA": match[6].strip()
                })

            # Extract total data
            subtotal_pattern = re.compile(r'TOTAL\s+(\d+\.\d{2})\s*')
            total_general_pattern = re.compile(r'TOTAL GENERAL\s+(\d+\.\d{2})\s*RON')

            match_subtotal = subtotal_pattern.search(text)
            if match_subtotal:
                invoice_data["Total"]["Subtotal (fără TVA)"] = match_subtotal.group(1)
            else:
                invoice_data["Total"]["Subtotal (fără TVA)"] = None

            match_total_general = total_general_pattern.search(text)
            if match_total_general:
                invoice_data["Total"]["Total de plată (incl. TVA)"] = match_total_general.group(1)
            else:
                invoice_data["Total"]["Total de plată (incl. TVA)"] = None

    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return invoice_data

def save_to_excel(data, excel_file):
    rows = []

    # Furnizor data
    rows.append(["Furnizor"])
    for key, value in data["Furnizor"].items():
        rows.append([key, value])
    rows.append([""])

    # Cumparator data
    rows.append(["Cumparator"])
    for key, value in data["Cumparator"].items():
        rows.append([key, value])
    rows.append([""])

    # Factura data
    rows.append(["Factura"])
    for key, value in data["Factura"].items():
        rows.append([key, value])
    rows.append([""])

    # Produse data
    rows.append(["Produse"])
    rows.append(["Nr. crt.", "Denumire", "UM", "Cantitate", "Preț unitar (fără TVA)", "Valoare (fără TVA)", "Valoare TVA"])
    
    total_valoare = 0
    total_valoare_tva = 0

    for product in data["Produse"]:
        valoare = float(product["Valoare (fără TVA)"].replace(',', ''))
        valoare_tva = float(product["Valoare TVA"].replace(',', ''))
        
        total_valoare += valoare
        total_valoare_tva += valoare_tva
        
        rows.append([product["Nr. crt."], product["Denumire"], product["UM"], product["Cantitate"], product["Preț unitar (fără TVA)"], product["Valoare (fără TVA)"], product["Valoare TVA"]])
    rows.append([""])

    # Total data
    rows.append(["Total", total_valoare])
    rows.append(["TVA", total_valoare_tva])
    rows.append(["Total general", total_valoare + total_valoare_tva])

    df = pd.DataFrame(rows)

    # Ensure the directory exists
    excel_directory = os.path.dirname(excel_file)
    if not os.path.exists(excel_directory):
        os.makedirs(excel_directory)

    df.to_excel(excel_file, index=False, header=False)
    print(f"Datele au fost încărcate în {excel_file}.")

def send_file_via_telegram(file_path):
    url = f"https://api.telegram.org/bot{telegram_bot_token}/sendDocument"
    with open(file_path, "rb") as file:
        response = requests.post(url, data={"chat_id": telegram_chat_id}, files={"document": file})
    if response.status_code == 200:
        print("Fișierul a fost trimis pe Telegram")
    else:
        print(f"Eroare la trimiterea fișierului pe Telegram: {response.json()}")

if __name__ == "__main__":
    download_attachments()
