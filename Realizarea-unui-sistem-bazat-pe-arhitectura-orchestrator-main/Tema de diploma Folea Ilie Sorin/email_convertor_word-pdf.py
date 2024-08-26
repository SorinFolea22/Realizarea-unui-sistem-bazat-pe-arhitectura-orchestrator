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

# Email account credentials
username = ""
password = ""

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

def convert_docx_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    try:
        pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path, extra_args=['--pdf-engine=xelatex'])
        print(f"Converted {docx_path} to {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")
    return pdf_path

def send_email_with_attachment(to_address, subject, body, attachment_path):
    from_address = ""
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

if __name__ == "__main__":
    download_attachments()
