import oracledb
from datetime import datetime, timedelta
import random
import json

import Bmail_hindu_calender as hindu

class DatabaseConnectionError(Exception):
    """Custom exception for database connection errors."""
    pass

class BirthdayNotifier:
    def __init__(self, db_config, email_configs):
        self.db_config = db_config
        self.email_configs = email_configs
        self.connection = None
        
        current_date = datetime.now()

        # Use timedelta to adjust the date (e.g., adding 5 days to the current date)
        adjusted_date = current_date + timedelta(days=1)
        self.adjusted_date = adjusted_date

    def connect_to_db(self):
        try:
            print("Connecting to the database...")
            self.connection = oracledb.connect(
                user=self.db_config['user'],
                password=self.db_config['password'],
                host=self.db_config['host'],
                port=self.db_config["port"],
                service_name=self.db_config["service_name"]
                
            )
            print("Connection successful.")
        except oracledb.DatabaseError as e:
            error, = e.args
            print(f"Error code: {error.code}, Error message: {error.message}")
            raise DatabaseConnectionError(f"Error connecting to the database: {e}")

    def get_birthday_data(self):
        cursor = None
        try:
            for email_config, query in (self.email_configs).items():
                
                cursor = self.connection.cursor()
                

                # Format the adjusted date as 'DD-MMM' (e.g., '09-NOV')
                formatted_date = self.adjusted_date.strftime('%d-%b').upper()
                day, month = formatted_date.split('-')

                cursor.execute(query, {'month': month, 'day': day})
                results = cursor.fetchall()

                if not results:
                    print("No birthday records found for today.")
                else:
                    print(f"Found {len(results)} birthday(s) for today.")
                
                # Save the generated HTML to a file
                self.generate_birthday_html(results,email_config)

                
        except Exception as e:
            print(f"Unexpected error: {e}")
            raise DatabaseConnectionError(f"Unexpected error: {e}")
        finally:
            if cursor:
                cursor.close()
                print("Cursor closed.")
                
    def get_random_dark_background_color(self):
        dark_colors_hex = [
            "#2c4850", "#2c5046", "#27381f", "#2d2e19", "#3b2c20",
            "#2a3439", "#0b0b36", "#011c01", "#250533", "#2C3E50"
        ]
        return random.choice(dark_colors_hex)

    def generate_birthday_html(self, birthday_data, email_config):
        if not birthday_data:
            print("No birthdays found for today.")
            return
        
        today_date = self.adjusted_date.strftime("%d/%m/%Y")
        month_tithi = self.date_month_hindu_calender(today_date)
        month_tithi = month_tithi
        
        formatted_today = self.adjusted_date.strftime("%B %d, %Y")
        subject = "🎉 Happy Birthday to Our Team Members! 🎉"
        
        background_color = self.get_random_dark_background_color()
        
        birthday_cards = self.generate_birthday_cards(birthday_data)

        body = f"""
        <div style="font-family: 'Arial', sans-serif; margin: 0; padding: 5vw; 
            background-color: {background_color};
            filter: blur(20px); z-index: -1;
            color: #ECF0F1; text-align: center; 
            max-width: 100%; margin: auto; 
            background-size: contain; 
            background-position: center; 
            background-repeat: no-repeat;
            word-wrap: break-word; word-break: break-word;"> 
            
    <h1 style="color: #FFD700; font-size: calc(1vw + 1.3em); margin-bottom: 5vw; word-wrap: break-word; word-break: break-word;"> 🎉 Wishing a Happy Birthday to Each Beloved Member of Our Family! 🎉 </h1>
    <p style="font-size: calc(1vw + 0.7em); color: #ECF0F1; word-wrap: break-word; word-break: break-word;"> On <strong style="color: #FFD700;">{formatted_today}</strong> <span style="color: #FF8000;">({month_tithi})</span> and we celebrate: </p>
    <div style="display: flex; flex-wrap: wrap; justify-content: flex-start; align-items: center; text-align: center; margin-right: 20px; word-wrap: break-word; word-break: break-word;">
        {birthday_cards}
    </div>
    <p style="font-size: calc(1vw + 0.7em); margin-top: 5vw; color: #ECF0F1; word-wrap: break-word; word-break: break-word;"> Wishing everyone a wonderful day filled with love and happiness! </p>
    <p style="font-size: calc(1vw + 0.7em); font-weight: bold; color: #FFD700; word-wrap: break-word; word-break: break-word;"> 💖 NECO FAMILY 💖 </p>
</div>
        """
        
        # Save to a file for testing
        self.save_html_to_file(body)

    def generate_birthday_cards(self, birthday_data):
        birthday_cards = ""
        for index, person in enumerate(birthday_data):
            
            DOB_of_person = person[3].strftime("%d/%m/%Y")
            month_tithi = self.date_month_hindu_calender(DOB_of_person)
            month_tithi = month_tithi.upper()
            
            card_html = f"""
                <div style="position: relative; margin: 10px; padding: 5px; background-color: rgba(255, 255, 255, 0.2); border-radius: 10px; width: calc(31%); height: auto; text-align: center; box-sizing: border-box; word-wrap: break-word; word-break: break-word;">
    <h3 style="color: #FFD700; margin: 0; padding: 0; font-size: calc(0.5vw + 0.7em);">SHRI {person[0].replace("SHRI", "")}</h3>
    <p style="color: #ECF0F1; font-size: calc(0.5vw + 0.6em);">
        <strong style="color: #FF8000; overflow-wrap: break-word;">({month_tithi})</strong><br><br>
        <span style="color: #ECF0F1;">{person[2]}</span><br><br>
        <span style="color: #ECF0F1;">{person[1]}</span><br>
    </p>
</div>
            """
            birthday_cards += card_html
            
            if (index + 1) % 3 == 0 and index + 1 != len(birthday_data):
                birthday_cards += "</div><div style='display: flex; flex-wrap: wrap; justify-content: flex-start; align-items: center; text-align: center; margin-right: 20px;'>"
    
        return birthday_cards             

    def save_html_to_file(self, body):
        with open('birthday_notifications.html', 'w', encoding='utf-8') as f:
            f.write(body)
        print("HTML saved as 'birthday_notifications.html'. Open this file to test the output.")


    def date_month_hindu_calender(self, dob):
        url = hindu.generate_url(dob)
        month, tithi = hindu.fetch_panchang_data(url)
        formated_date = month +", "+tithi
        return formated_date

    def close_connection(self):
        if self.connection:
            self.connection.close()
            print("Database connection closed.")

if __name__ == "__main__":
    db_config = {
        'user': 'jnl',
        'password': 'jnl123',
        'host' : '80.0.1.81', 
        'port': 1521, 
        'service_name' : "jnil"  
    }
    
    email_configs = {}

    Query_1 = """
                SELECT NAME, DEPARTMENT, DESIGNATION, DOB, MAS_GRA_CD
                FROM V_EMPDATA
                WHERE TO_CHAR(DOB, 'MON') = :month AND TO_CHAR(DOB, 'DD') = :day
                ORDER By MAS_GRA_CD DESC
            """

    Query_2 = """
                SELECT NAME, DEPARTMENT, DESIGNATION, DOB, MAS_GRA_CD
                FROM ngpmast
                WHERE TO_CHAR(DOB, 'MON') = :month AND TO_CHAR(DOB, 'DD') = :day
                ORDER By MAS_GRA_CD DESC
            """    

    email_configs = {
                    f'{Query_1}' : f'{Query_1}', 
                    f'{Query_2}' : f'{Query_2}'
                   }

    notifier = BirthdayNotifier(db_config, email_configs)

    try:
        notifier.connect_to_db()
        notifier.get_birthday_data()
    except DatabaseConnectionError as e:
        print(e)
    finally:
        notifier.close_connection()
