import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
from bs4 import BeautifulSoup
import random
from datetime import datetime

# Function to fetch thought of the day and author names (English and Hindi)
def fetch_thought_of_the_day():
    url = 'https://www.shabdkosh.com/quote-of-the-day/english-hindi/'
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract the English quote and author name
        english_blockquote = soup.find_all('blockquote')[0]
        english_quote = english_blockquote.find('span', style="font-size:1.5em;")
        english_author = english_blockquote.text.split('”')[1].strip()

        # Extract the Hindi quote and author name
        hindi_blockquote = soup.find_all('blockquote')[1]
        hindi_quote = hindi_blockquote.find('span', style="font-size:1.5em;")
        hindi_author = hindi_blockquote.text.split('”')[1].strip()

        if english_quote and hindi_quote:
            return (
                english_quote.text.strip(), english_author,
                hindi_quote.text.strip(), hindi_author
            )
        else:
            return None, None, None, None
    else:
        return None, None, None, None

# List of soft pastel background colors
soft_colors = [
    "#A7C7E7",  # Soft Blue
    "#B8E5D5",  # Light Mint
    "#F8D0B6",  # Pale Peach
    "#E1D6F2",  # Soft Lavender
    "#F1C4C4",  # Light Coral
    "#F7D0D0",  # Baby Pink
    "#F8F2B8",  # Soft Yellow
    "#B0C6E3",  # Powder Blue
    "#E8D6F7",  # Light Lavender
    "#F2E1E1",  # Misty Rose
    "#D4E9D4",  # Soft Green
    "#B7D8E4",  # Light Sky Blue
    "#B8E1E0",  # Pale Turquoise
    "#FAD0C4",  # Peach Cream
    "#FFF1E0"   # Soft Ivory
]

# Function to randomly pick a background color
def get_random_soft_background_color():
    return random.choice(soft_colors)

# Function to send email using provided email data
def send_email(email_data):
    from_email = email_data['from_email']
    from_password = email_data['from_password']

    # SMTP server configuration (for Gmail)
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # Create the MIME object
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(email_data['to_email'])
    msg['Subject'] = email_data['subject']
    msg.attach(MIMEText(email_data['body'], 'html'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, from_password)
            server.sendmail(from_email, email_data['to_email'], msg.as_string())
            print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email. Error: {e}")

# Main function to fetch quotes, format the email, and send it
def main():
    english_quote, english_author, hindi_quote, hindi_author = fetch_thought_of_the_day()

    if english_quote and hindi_quote:
        # Get a random background color for the email body
        background_color = get_random_soft_background_color()

        # Set static text color
        text_color = "#333333"  # Dark gray for readability
        quote_color = "#2C3E50"  # Darker blue for quote text
        author_color = "#2980B9"  # Blue for author text
        header_color = "#3498DB"  # Light blue for header

        # Prepare the HTML content for the email body with dynamic background
        html_content = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Thought of the Day</title>
            <style>
                /* Inlining styles */
                body {{
                    font-family: 'Arial', sans-serif;
                    color: {text_color};
                    margin: 0;
                    padding: 0;
                }}
                .container {{
                    width: 100%;
                    max-width: 650px;
                    margin: 0 auto;
                    padding: 40px;
                    background-color: #ffffff;
                    box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.1);
                    border-radius: 10px;
                    border-top: 8px solid {header_color};
                }}
                h1 {{
                    text-align: center;
                    color: {header_color};
                    font-size: 32px;
                    margin-bottom: 30px;
                    font-weight: bold;
                }}
                .quote-container {{
                    margin-bottom: 20px;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
                }}
                .quote {{
                    font-size: 1.6em;
                    font-style: italic;
                    color: {quote_color};
                    margin-bottom: 10px;
                    line-height: 1.5;
                }}
                .author {{
                    font-size: 1.1em;
                    font-weight: bold;
                    color: {author_color};
                    text-align: right;
                    margin-top: 15px;
                }}
                footer {{
                    text-align: center;
                    margin-top: 30px;
                    font-size: 13px;
                    color: #aaa;
                }}
                @media (max-width: 600px) {{
                    .container {{
                        padding: 20px;
                    }}
                    h1 {{
                        font-size: 24px;
                    }}
                    .quote {{
                        font-size: 1.4em;
                    }}
                    .author {{
                        font-size: 1em;
                    }}
                }}
            </style>
        </head>
        <body>
            
                <h1>Thought of the Day</h1>

                <!-- English Quote -->
                <div class="quote-container" style="background-color: {background_color};">
                    <p class="quote">“{english_quote}”</p>
                    <p class="author">- {english_author}</p>
                </div>

                <!-- Hindi Quote -->
                <div class="quote-container" style="background-color: {background_color};">
                    <p class="quote">“{hindi_quote}”</p>
                    <p class="author">- {hindi_author}</p>
                </div>

                <footer>
                    <p>&copy; 2024 Thought of the Day. All Rights Reserved.</p>
                </footer>
            
        </body>
        </html>
        """

        # Create the email data dictionary
        email_data = {
            'from_email': 'myneco@necoindia.com',  # Replace with your Gmail address
            'from_password': 'ukzu matf aupi kesv',  # Replace with your Gmail App Password
            'to_email': ['everyone.spd@necoindia.com', 'everyone.ho@necoindia.com'],  # Correct formatting
            'subject': 'Thought of the Day',
            'body': html_content
        }

        # Send the email with the data from the dictionary
        send_email(email_data)
    else:
        print("Failed to fetch Thought of the Day. Please check the website.")

if __name__ == "__main__":
    main()
