birthday_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Birthday Wish</title>
    <style>
        body, html {{
            height: 100%;
            margin: 0;
            font-family: 'Arial', sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            background-color: #f0f8ff;
        }}
        .container {{
            background-color: rgba(255, 250, 240, 0.9);
            padding: 40px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
            max-width: 600px;
        }}
        h1 {{
            color: #FF6F61;
        }}
        .message {{
            font-size: 1.2em;
            color: #333;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🎉 Happy Birthday, {name}! 🎉</h1>
        <p>Today is {date} and we celebrate:</p>
        <ul>
            {names}
        </ul>
        <p class="message">Wishing you a joyful day filled with love and happiness!</p>
        <p>💖 NECO FAMILY 💖</p>
    </div>
</body>
</html>
"""

# Function to create the birthday wish
def create_birthday_wish(name, date, celebrants):
    names_list = ''.join([f"<li><strong>{person}</strong></li>" for person in celebrants])
    return birthday_html.format(name=name, date=date, names=names_list)

# Example usage
name = "Alice"
date = "October 1, 2024"
celebrants = ["Shri Bob", "Shri Charlie", "Shri Dana"]

# Create the HTML content
html_content = create_birthday_wish(name, date, celebrants)

# Save to an HTML file with UTF-8 encoding
with open('birthday_wish.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("Birthday wish HTML file created!")
