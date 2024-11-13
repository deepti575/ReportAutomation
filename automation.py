import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import win32com.client as win32
import os

# Function to get user input for email details
def get_email_details():
    email_list = input("Enter the recipient email addresses (comma-separated): ").split(',')
    sender_email = input("Enter your email address: ")
    sender_password = input("Enter your app password: ")  # Use app password for security
    subject = input("Enter the subject of the email: ")
    body = input("Enter the body of the email: ")
    return email_list, sender_email, sender_password, subject, body

# Define the paths
base_directory = os.path.dirname(os.path.abspath(__file__))
report_directory = os.path.join(base_directory, "reports")
workbook_path = os.path.join(base_directory, "AmazonSalesProject.xlsm")
macro_name = "GenerateReport"

# Get email details from the user
email_list, sender_email, sender_password, subject, body = get_email_details()

# Function to get the latest file in the directory
def get_latest_file(directory):
    files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    latest_file = max(files, key=os.path.getctime)
    return latest_file

# Function to run the Excel macro
def run_excel_macro():
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(workbook_path)
    excel.Application.Run(macro_name)
    workbook.Close(SaveChanges=True)
    excel.Quit()

# Function to send an email with the report attachment
def send_email(file_path, email_list, sender_email, sender_password, subject, body):
    for recipient in email_list:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient.strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        attachment = open(file_path, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
        msg.attach(part)

        # Set up the SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)  # Login to the email account
        text = msg.as_string()
        server.sendmail(sender_email, recipient.strip(), text)  # Send the email
        server.quit()  # Close the connection

# Main function to run the macro and send the email every hour
def main():
    if not os.path.exists(report_directory):
        os.makedirs(report_directory)
    
    while True:
        run_excel_macro()
        latest_file = get_latest_file(report_directory)
        send_email(latest_file, email_list, sender_email, sender_password, subject, body)
        time.sleep(3600)  # Wait for 1 hour

if __name__ == "__main__":
    main()