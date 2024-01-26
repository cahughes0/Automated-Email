import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from dotenv import load_dotenv


PORT = 587
EMAIL_SERVER = "smtp.office365.com"

current_dir = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd()
envars = current_dir / ".env"
load_dotenv(envars)

sender_email = os.getenv("EMAIL")
password_email = os.getenv("PASSWORD")

# Google Sheets details
SHEET_ID = "<1TWCwKJzZtIe4zoxuxrZpfw2MexZ2Rxc-qjA7f_DUn9Q>"
SHEET_NAME = "exampleautomatedemail"
URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

class EmailConfig: #reusability
    def __init__(self, sender_email, password_email, email_server, port):
        self.sender_email = sender_email
        self.password_email = password_email
        self.email_server = email_server
        self.port = port



def read(url): #converts to dictionary format
    df = pd.read_csv(url)
    word_data = df.to_dict(orient='records')
    return word_data


def send_email(config, subject, receiver_email, name, invoice_no, date, amount): #designs email in HTML format
    html_template = f"""
    <html>
      <body>
        <p>Hello {name}!</p>
        <p>Thank you for your purchase of <strong>{amount}</strong> on {date}. We appreciate your support and hope you enjoy our product.</p>
        <p>Your Invoice:</p>
        <p>{invoice_no}</p>
      </body>
    </html>
    """

    msg = EmailMessage() #constructs email content
    msg["Subject"] = subject
    msg["From"] = formataddr(("Custom Sender", config.sender_email))
    msg["To"] = receiver_email
    msg["BCC"] = config.sender_email

    msg.add_alternative(html_template, subtype="html")

    with smtplib.SMTP(config.email_server, config.port) as server: #ensures connection with server during process
        server.starttls() #connection
        server.login(config.sender_email, config.password_email)
        server.sendmail(config.sender_email, receiver_email, msg.as_string())


def send_emails(config, data): #actually sends email to specific user
    for record in data:
        send_email(
            config,
            subject="Thank you for your purchase",
            name=record["customer_name"],
            receiver_email=record["email"],
            date=record["date"],
            invoice_no=record["invoice_no"],
            amount=record["amount"]
        )


if __name__ == '__main__':
    current_dir = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd() #directory
    envars = current_dir / ".env"
    load_dotenv(envars) #loads env variables (sensitive info)

    sender_email = os.getenv("EMAIL")
    password_email = os.getenv("PASSWORD")
    email_server = os.getenv("EMAIL_SERVER", "smtp.office365.com")
    port = int(os.getenv("PORT", 587))

    config = EmailConfig(sender_email, password_email, email_server, port)

    data = read(URL)
    send_emails(config, data)


