import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import csv

# Email configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SENDER_EMAIL = 'Demo71405@gmail.com'
SENDER_PASSWORD = 'Demo@123'
print(type(SENDER_EMAIL))
print(type(SENDER_PASSWORD))

# Get today's date
today = datetime.now().strftime("%Y-%m-%d")

def check_authentication():
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure the connection
            server.login(SENDER_EMAIL, SENDER_PASSWORD)  # Attempt to log in
            print("Authentication successful!")
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Please check your username and password.")
    except Exception as e:
        print(f"Error during authentication: {e}")


def send_email(recipient_email, recipient_name):
    subject = "Happy Birthday!"
    body = f"Dear {recipient_name},\n\nHappy Birthday! Wishing you a wonderful day filled with joy and happiness.\n\nBest wishes,\nYour Company"
    
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    # Send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure the connection
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
            print(f"Email sent to {recipient_name} at {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_name}: {e}")

# Read employee data from CSV file
if __name__ == '__main__':
    check_authentication()  # Check authentication before proceeding

    # Read employee data from CSV file
    employees = []
    csv_file_path = r"D:\USER PROFILE DATA\Desktop\Project-2\Data\Employees_data.csv"

    try:
        with open(csv_file_path, newline='', encoding='ISO-8859-1') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                employees.append(row)
    except Exception as e:
        print(f"Error reading CSV file: {e}")

    # Check for birthdays today and send emails
    found_birthday = False  # Flag to check if any birthday is found
    for employee in employees:
        print(f"Checking {employee['Name']} with DOB {employee['DOB']} against {today}")
        if employee["DOB"] == today:
            found_birthday = True
            send_email(employee["Gmail"], employee["Name"])

    if not found_birthday:
        print("No birthdays found for today.")
