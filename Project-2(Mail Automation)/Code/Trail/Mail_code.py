import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

def get_birthday_data(file_path):

    df = pd.read_excel(file_path)
    today = datetime.now().strftime("%m/%d/%y")
    birthday_data = df[df['DOB'].dt.strftime("%m/%d/%y") == today]
    
    return birthday_data[['Name', 'Gmail', 'Designation', 'Photos Link']]

def send_birthday_email(birthday_data):
    if not birthday_data.empty:
        formatted_today = datetime.now().strftime("%B %d, %Y")
        subject = "🎉 Happy Birthday to Our Team Members! 🎉"

        body = f"""
            <div style="font-family: 'Arial', sans-serif; margin: 0; padding: 20px; background-color: #2C3E50; color: #ECF0F1; text-align: center;">
                <h1 style="color: #FFD700; font-size: 2.5em; margin-bottom: 20px;">🎉 Happy Birthday! 🎉</h1>
                <p style="font-size: 1.5em; color: #ECF0F1;">Today is <strong style="color: #FFD700;">{formatted_today}</strong> and we celebrate:</p>
                
                <div style="display: flex; flex-wrap: wrap; justify-content: center;">
                    {''.join(f"""
                        <div style="margin: 10px; padding: 10px; background-color: rgba(255, 255, 255, 0.2); border-radius: 10px; width: 180px; text-align: center;">
                            <img src="{row['Photos Link']}" alt="{row['Name']}" style="width: 100%; height: auto; border-radius: 10px;">
                            <h3 style="color: #FFD700;">Shri {row['Name']}</h3>
                            <p style="color: #ECF0F1;">{row['Designation']}</p>
                        </div>
                    """ for _, row in birthday_data.iterrows())}
                </div>

                <p style="font-size: 1.5em; margin-top: 20px; color: #ECF0F1;">Wishing everyone a wonderful day filled with love and happiness!</p>
                <p style="font-size: 1.5em; font-weight: bold; color: #FFD700;">💖 NECO FAMILY 💖</p>
            </div>
        """

        msg = MIMEMultipart()
        msg['From'] = "saket.verma@necoindia.com"  
        msg['To'] = "saket.verma@necoindia.com" 
       # msg['Cc'] = "shaji.thomas@necoindia.com, alok.pandey@necoindia.com"
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'html'))

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login("saket.verma@necoindia.com", "pxou ynvf ricv vvei")  
                server.send_message(msg)
            print(f"Birthday email sent to {msg['To']}")
        except Exception as e:
            print(f"Failed to send email: {e}")
    else:
        print("No birthdays found for today.")

if __name__ == "__main__":

    file_path = "D:\\USER PROFILE DATA\\Downloads\\Employees_data.xlsx"
    birthday_data = get_birthday_data(file_path)
    send_birthday_email(birthday_data)


 


