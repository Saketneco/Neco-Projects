import imaplib
import email
import os

# Configuration
EMAIL = 'saket.verma@necoindia.com'
PASSWORD = 'Saketneco'  # Use app password if 2FA is enabled
IMAP_SERVER = 'imap.gmail.com'
DOWNLOAD_FOLDER = os.path.expanduser('~/Desktop/')  # Downloads to the Desktop

def connect_to_email():
    """Connect to the Gmail server and return the connection object."""
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    return mail

def download_attachments(mail, sender_email):
    """Download attachments from unread emails from a specific sender."""
    mail.select('inbox')  # Select the inbox
    # Search for all emails that are unseen (unread)
    status, messages = mail.search(None, 'UNSEEN')
    email_ids = messages[0].split()

    for email_id in email_ids:
        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])

        # Check if the email is from the specified sender
        if msg['from'] == sender_email:
            # Check for attachments
            if msg.is_multipart():
                for part in msg.walk():
                    # If part is an attachment
                    if part.get_content_disposition() == 'attachment':
                        filename = part.get_filename()
                        if filename:
                            # Create a full path to save the attachment
                            filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                            with open(filepath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            print(f'Downloaded: {filename}')

def main():
    # Create download folder if it doesn't exist
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    # Connect to email and download attachments from the specified sender
    mail = connect_to_email()
    download_attachments(mail, 'krishnendu.ray@necoindia.com')
    mail.logout()

if __name__ == "__main__":
    main()
