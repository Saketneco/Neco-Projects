#!/usr/bin/env python3

import pymysql
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

class DatabaseConnectionError(Exception):
    """Custom exception for database connection errors."""
    pass

class BirthdayNotifier:
    def __init__(self, db_config, email_config):
        self.db_config = db_config
        self.email_config = email_config
        self.connection = None

    def connect_to_db(self):
        try:
            print("Connecting to the database...")
            self.connection = pymysql.connect(**self.db_config)
            print("Connection successful.")
        except pymysql.MySQLError as e:
            print(f"Error code: {e.args[0]}, Error message: {e.args[1]}")
            raise DatabaseConnectionError(f"Error connecting to the database: {e}")

    def get_birthday_data(self):
        cursor = None
        try:
            cursor = self.connection.cursor()
            formatted_date = datetime.now().strftime('%m-%d')
            month, day = formatted_date.split('-')

            query = """
                SELECT NAME, DEPARTMENT, DESIGNATION, DOB
                FROM emp_name
                WHERE DAY(DOB) = %s AND  MONTH(DOB) = %s  
            """
            cursor.execute(query, (day, month))
            results = cursor.fetchall()

            #print(results)
            if not results:
                print("No birthday records found for today.")
            else:
                print(f"Found {len(results)} birthday(s) for today.")

            return results
        except Exception as e:
            print(f"Unexpected error: {e}")
            raise DatabaseConnectionError(f"Unexpected error: {e}")
        finally:
            if cursor:
                cursor.close()
                print("Cursor closed.")

    def send_birthday_email(self, birthday_data):
        if not birthday_data:
            print("No birthdays found for today.")
            return

        formatted_today = datetime.now().strftime("%B %d, %Y")
        subject = "🎉 Happy Birthday to Our Team Members! 🎉"
        
        birthday_cards = self.generate_birthday_cards(birthday_data)

        body = f"""
        <div style="font-family: 'Arial', sans-serif; margin: 0; padding: 20px; background-color: #2C3E50; color: #ECF0F1; text-align: center; max-width: 800px; margin: auto;">
            <h1 style="color: #FFD700; font-size: 2.5em; margin-bottom: 20px;">🎉 Happy Birthday! 🎉</h1>
            <p style="font-size: 1.5em; color: #ECF0F1;">Today is <strong style="color: #FFD700;">{formatted_today}</strong> and we celebrate:</p>
            <div style="display: flex; flex-wrap: wrap; justify-content: center;">{birthday_cards}</div>
            <p style="font-size: 1.5em; margin-top: 20px; color: #ECF0F1;">Wishing everyone a wonderful day filled with love and happiness!</p>
            <p style="font-size: 1.5em; font-weight: bold; color: #FFD700;">💖 NECO FAMILY 💖</p>
        </div>
        """

        self.send_email(subject, body)

    def generate_birthday_cards(self, birthday_data):
        birthday_cards = ""
        for index, person in enumerate(birthday_data):
            card_html = f"""
                <div style="margin: 10px; padding: 10px; background-color: rgba(255, 255, 255, 0.2); border-radius: 10px; width: 180px; text-align: center;">
                    <h3 style="color: #FFD700;">Shri {person[0]}</h3>
                    <p style="color: #ECF0F1;">{person[1]} - {person[2]}</p>
                </div>
            """
            birthday_cards += card_html
            
            # Insert a new row every 5 cards
            if (index + 1) % 5 == 0 and index + 1 != len(birthday_data):
                birthday_cards += "</div><div style='display: flex; flex-wrap: wrap; justify-content: center;'>"

        return birthday_cards

    def send_email(self, subject, body):
        msg = MIMEMultipart()
        msg['From'] = self.email_config['from']
        msg['To'] = self.email_config['to']
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        try:
            with smtplib.SMTP(self.email_config['smtp_server'], self.email_config['smtp_port']) as server:
                server.starttls()
                server.login(self.email_config['from'], self.email_config['password'])
                server.send_message(msg)
            print(f"Birthday email sent to {msg['To']}")
        except Exception as e:
            print(f"Failed to send email: {e}")

    def close_connection(self):
        if self.connection:
            self.connection.close()
            print("Database connection closed.")

if __name__ == "__main__":
    db_config = {
        'user': 'root',
        'password': '',  
        'host': 'localhost',
        'port': 3306,
        'database': 'employee_database',
    }

    email_config = {
        'from': "saket.verma@necoindia.com",
        'to': "saket.verma@necoindia.com",
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'password': "pxou ynvf ricv vvei"  
    }

    notifier = BirthdayNotifier(db_config, email_config)

    try:
        notifier.connect_to_db()
        birthday_data = notifier.get_birthday_data()
        notifier.send_birthday_email(birthday_data)
    except DatabaseConnectionError as e:
        print(e)
    finally:
        notifier.close_connection()
