import oracledb
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime ,timedelta
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
        adjusted_date = current_date - timedelta(days=3)
        print(adjusted_date.strftime("%d/%m/%Y"))
        self.adjusted_date = adjusted_date

    def connect_to_db(self):
        try:
            print("Connecting to the database...")
            self.connection = oracledb.connect(
                user=self.db_config['user'],
                password=self.db_config['password'],
                host = self.db_config['host'],
                port = self.db_config["port"],
                service_name = self.db_config["service_name"]
                
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
                #  day= str(int(day)+1)
                #  print(day,type(day))

                # query = """
                #     SELECT NAME, DEPARTMENT, DESIGNATION, DOB, MAS_GRA_CD
                #     FROM ngpmast
                #     WHERE TO_CHAR(DOB, 'MON') = :month AND TO_CHAR(DOB, 'DD') = :day
                #     ORDER By MAS_GRA_CD DESC
                # """
                #print(query)
                cursor.execute(query, {'month': month, 'day': day})
                results = cursor.fetchall()

                if not results:
                    print("No birthday records found for today.")
                else:
                    print(f"Found {len(results)} birthday(s) for today.")
                
                # for date in results:
                #     print(date)
                self.send_birthday_email(results,email_config)

                
        except Exception as e:
            print(f"Unexpected error: {e}")
            raise DatabaseConnectionError(f"Unexpected error: {e}")
        finally:
            if cursor:
                cursor.close()
                print("Cursor closed.")
                
                
    # Function to randomly pick a background color
    def get_random_dark_background_color(self):
        return random.choice(dark_colors_hex)
    

    def send_birthday_email(self, birthday_data ,email_config):
        if not birthday_data:
            print("No birthdays found for today.")
            return
        
        # for row in birthday_data:
        #     print(row[3].strftime("%d/%m/%Y"))
            
        # formatted_today = self.date_month_hindu_calender()
        # print(formatted_today)
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
            word-wrap: break-word; word-break: break-word;"> <!-- Added here for overall container -->
            
    <h1 style="color: #FFD700; font-size: calc(1vw + 1.3em); margin-bottom: 5vw; word-wrap: break-word; word-break: break-word;"> 🎉 Wishing a Happy Birthday to Each Beloved Member of Our Family! 🎉 </h1>
    <p style="font-size: calc(1vw + 0.7em); color: #ECF0F1; word-wrap: break-word; word-break: break-word;"> On <strong style="color: #FFD700;">{formatted_today}</strong> <span style="color: #FF8000;">({month_tithi})</span> and we celebrate: </p>
    <div style="display: flex; flex-wrap: wrap; justify-content: flex-start; align-items: center; text-align: center; margin-right: 20px; word-wrap: break-word; word-break: break-word;">
        {birthday_cards}
    </div>
    <p style="font-size: calc(1vw + 0.7em); margin-top: 5vw; color: #ECF0F1; word-wrap: break-word; word-break: break-word;"> Wishing everyone a wonderful day filled with love and happiness! </p>
    <p style="font-size: calc(1vw + 0.7em); font-weight: bold; color: #FFD700; word-wrap: break-word; word-break: break-word;"> 💖 NECO FAMILY 💖 </p>
</div>

        """

        self.send_email(subject, body,email_config)

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
    
    <!-- Button styled to stay within the box -->
    <div style="display: flex; justify-content: center; align-items: center; margin-top: auto;">
    <a href="mailto:?subject=Happy Birthday,{person[0].replace('SHRI','')}&body=Dear%20{person[0].replace('SHRI','')},%0D%0A%0D%0A🎉%20Wishing%20you%20a%20wonderful%20birthday!🎂%0D%0A%0D%0A%20I%20hope%20this%20year%20brings%20you%20good%20health,%20joy,%20and%20everything%20you%20wish%20for.%20May%20your%20special%20day%20be%20filled%20with%20love,%20laughter,%20and%20memories%20to%20treasure.I%E2%80%99m%20truly%20thankful%20to%20have%20you%20in%20my%20life%20and%20for%20the%20friendship%20we%20share.%20Enjoy%20every%20moment%20today,%20and%20may%20this%20year%20bring%20you%20even%20more%20success%20and%20happiness.%20Take%20care%20and%20celebrate%20your%20day%20in%20the%20most%20special%20way.%20Happy%20birthday%20again%2C%20and%20here's%20to%20another%20great%20year%20ahead."
        style="display: inline-block; padding: 10px 10px; background-color: #0066CC; color: white; font-size: calc(0.5vw + 0.6em); text-decoration: none; border-radius: 6px; font-weight: bold; font-style: italic; font-family: cursive; transition: background-color 0.3s ease, transform 0.3s ease; text-align: center; margin-top: auto;">
        <span style="font-family: Arial, sans-serif;">Wish Birthday! 🎉</span>
    </a>
    </div>
</div>



            """
            birthday_cards += card_html
            
            if (index + 1) % 3 == 0 and index + 1 != len(birthday_data):
                birthday_cards += "</div><div style='display: flex; flex-wrap: wrap; justify-content: flex-start; align-items: center; text-align: center; margin-right: 20px;'>"
    
        return birthday_cards             

    def send_email(self, subject, body,email_config):
        
        email_config = email_config.replace("'",'"')
        email_config = json.loads(email_config)
        # print(email_config)
        # for col ,val in email_config.items():
        #     print(val)
        msg = MIMEMultipart()
        msg['From'] = email_config['from']
        msg['To'] = email_config['to']
       # msg['cc'] = email_config['cc']
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        

        try:
            with smtplib.SMTP(email_config['smtp_server'], email_config['smtp_port']) as server:
                server.starttls()
                server.login(email_config['from'], email_config['password'])
                server.send_message(msg)
            print(f"Birthday email sent to {msg['To']}")
        except Exception as e:
            print(f"Failed to send email: {e}")
            
    def date_month_hindu_calender(self,dob):
        #date = self.adjusted_date.strftime("%d/%m/%Y")       
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
    
    dark_colors_hex = [
     "#2c4850", 
     "#2c5046",  
     "#27381f",  
     "#2d2e19",  
     "#3b2c20",  
     "#2a3439",  # Gunmetal
     "#0b0b36",  # Midnight Blue
     "#011c01",  # Dark Green
     "#250533",   # Dark Violet (Still a bit bold, but kept for its dark tone)
     "#2C3E50",
     '#370617',
]
    email_configs ={}

    Query_1 = """
                SELECT NAME, DEPARTMENT, DESIGNATION, DOB, MAS_GRA_CD
                FROM V_EMPDATA
                WHERE TO_CHAR(DOB, 'MON') = :month AND TO_CHAR(DOB, 'DD') = :day
                ORDER By MAS_GRA_CD DESC
            """

    email_config_1 = {
        'from': "myneco@necoindia.com",
        'to' : "saket.verma@necoindia.com",
       # 'to': "alok.sharma@necoindia.com",
       # 'to': "everyone.spd@necoindia.com", 
       # 'to' : "everyone.ho@necoindia.com",
       # 'cc': "sukanta.nayak@necoindia.com",
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'password': "ukzu matf aupi kesv"  
    }
    
    Query_2 = """
                SELECT NAME, DEPARTMENT, DESIGNATION, DOB, MAS_GRA_CD
                FROM ngpmast
                WHERE TO_CHAR(DOB, 'MON') = :month AND TO_CHAR(DOB, 'DD') = :day
                ORDER By MAS_GRA_CD DESC
            """    
            
    email_config_2 = {
        'from': "myneco@necoindia.com",
        'to' : "saketverma1911@gmail.com",
       # 'to' : "saket.verma@necoindia.com",
       # 'to': "alok.sharma@necoindia.com",
       # 'to': "everyone.spd@necoindia.com", 
       # 'to' : "everyone.ho@necoindia.com",
       # 'cc': "sukanta.nayak@necoindia.com",
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'password': "ukzu matf aupi kesv"  
    }
    
    
    email_configs = {
                    f'{email_config_1}' : f'{Query_1}', 
                    f'{email_config_2}' : f'{Query_2}'
                   }

    notifier = BirthdayNotifier(db_config, email_configs)

    try:
        notifier.connect_to_db()
        

        hindu.generate_url
            
        notifier.get_birthday_data()
        # for row in birthday_data:
        #     print(row)
        #notifier.send_birthday_email(birthday_data)
    except DatabaseConnectionError as e:
        print(e)
    finally:
        notifier.close_connection()
