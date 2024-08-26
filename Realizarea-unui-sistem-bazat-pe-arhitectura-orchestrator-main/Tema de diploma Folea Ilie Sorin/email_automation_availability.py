import imaplib
import email
from email.header import decode_header
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pypandoc
from datetime import datetime

# Email account credentials
username = ""
password = ""  

# Lista zilelor de sărbătoare din România
HOLIDAYS = [
    "2024-01-01", "2024-01-02",  # Anul Nou
    "2024-04-19", "2024-04-20",  # Vinerea Mare și Paștele Ortodox
    "2024-05-01",                # Ziua Muncii
    "2024-06-16", "2024-06-17",  # Rusaliile
    "2024-08-15",                # Adormirea Maicii Domnului
    "2024-11-30",                # Sfântul Andrei
    "2024-12-01",                # Ziua Națională a României
    "2024-12-25", "2024-12-26"   # Crăciunul
]

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
        print(f"Processing email ID: {email_id.decode()}")
        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        if status != "OK":
            print(f"Error fetching email ID: {email_id.decode()}")
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
                        if filename and filename.endswith('.docx'):
                            # Use raw string or double backslashes
                            filepath = os.path.join(r"C:/Users/Sorin/Desktop/Documente orchestrator task1", filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"Downloaded attachment: {filepath}")
                            attachment_downloaded = True
                            pdf_path = convert_docx_to_pdf(filepath)
                            send_email_with_attachment(msg["From"], "Converted PDF", "Here is the converted PDF file.", pdf_path)

        # Mark the email as read if any attachment was downloaded
        if attachment_downloaded:
            print(f"Marking email {email_id.decode()} as read")
            result = mail.store(email_id, '+FLAGS', '\\Seen')
            print(f"Result of marking as read: {result}")

        # Send an auto-reply if no document was downloaded and it's weekend or holiday
        if not attachment_downloaded:
            is_holiday, message_body = is_weekend_or_holiday()
            if is_holiday:
                send_auto_reply(msg["From"], message_body)

def convert_docx_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    try:
        pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path, extra_args=['--pdf-engine=xelatex'])
        print(f"Converted {docx_path} to {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")
    return pdf_path

def send_email_with_attachment(to_address, subject, body, attachment_path):
    from_address = "@gmail.com"
    password = "" 

    msg = MIMEMultipart()
    msg["From"] = from_address
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(attachment_path)}",
        )
        msg.attach(part)

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_address, msg.as_string())
        server.quit()
        print(f"Email sent to {to_address} with attachment {attachment_path}")
    except Exception as e:
        print(f"Error sending email to {to_address}: {e}")

def is_weekend_or_holiday():
    current_datetime = datetime.now()
    today = current_datetime.weekday()
    current_date_str = current_datetime.strftime("%Y-%m-%d")
    print(f"Today is {current_datetime}, which is a {'weekend' if today in [5, 6] else 'weekday'}.")
    print(f"Checking if {current_date_str} is a holiday: {'Yes' if current_date_str in HOLIDAYS else 'No'}")
    
    # Check if it's weekend or a holiday
    if today == 5 or today == 6:
        return True, "Mulțumesc pentru mesaj, momentan este weekend. O să revin cu un mesaj luni."
    elif current_date_str in HOLIDAYS:
        return True, "Mulțumesc pentru mesaj. O să revin cu un mesaj în următoarea zi lucrătoare."
    return False, ""

def send_auto_reply(to_address, message_body):
    from_address = "orchestrator223@gmail.com"
    subject = "Momentan indisponibil"

    msg = MIMEMultipart()
    msg["From"] = from_address
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(message_body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_address, msg.as_string())
        server.quit()
        print(f"Auto-reply sent to {to_address}")
    except Exception as e:
        print(f"Error sending auto-reply to {to_address}: {e}")

if __name__ == "__main__":
    download_attachments()
