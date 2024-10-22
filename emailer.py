import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd

# Function to read email body from the file
def get_email_body(file_path):
    with open(file_path, 'r') as file:
        email_body = file.read()
    return email_body

# Function to send dynamic email
def send_email(to_email, subject, email_body, recipient_name, company_name):
    # Email credentials and settings
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    from_email = 'yourmail@gmail.com'  # Replace with your email
    password = 'xxxx xxxx xxxx xxxx'# Replace with your email password

    # Setup the MIME
    msg = MIMEMultipart('related')
    msg['From'] = from_email
    msg['To'] = str(to_email)  # Ensure to_email is a string
    msg['Subject'] = subject
    print(f"Receiver name: ", recipient_name)

    # Replace placeholders in the email body
    email_body = email_body.replace("Sonal", recipient_name).replace("Amazon", company_name)

    # Attach the email body
    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    msg_alternative.attach(MIMEText(email_body, 'html'))

    # Attach the image
    with open('./profile-pic (2).png', 'rb') as img_file:
        msg_image = MIMEImage(img_file.read())
        msg_image.add_header('Content-ID', '<image1>')
        msg.attach(msg_image)

    # Establishing a connection to the server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(from_email, password)

    # Sending the email
    server.send_message(msg)
    print(f"Email sent to {to_email}")

    # Terminating the session
    server.quit()

# Reading the Excel file for dynamic content (e.g., recipient email)
def get_recipients_from_excel(file_path):
    df = pd.read_excel(file_path)
    recipients = []
    for index, row in df.iterrows():
        email = row['Email']
        name = row['Name']
        company = row['Company']
        if pd.notna(email) and pd.notna(name) and pd.notna(company):
            recipients.append((str(email), str(name), str(company)))
    return recipients

if __name__ == '__main__':
    # Paths to the files
    email_body_file = './email_body.txt'
    excel_file = './hr_details.xlsx'

    # Getting email body
    email_body = get_email_body(email_body_file)

    # Extract recipient email, name, and company from Excel file
    recipients = get_recipients_from_excel(excel_file)

    # Subject of the email
    subject = "Application for SDE Position"

    # Sending the email to each recipient
    for recipient_email, recipient_name, company_name in recipients:
        send_email(recipient_email, subject, email_body, recipient_name, company_name)