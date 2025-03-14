import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os



def send_email(sender_email, sender_password, subject, body, file_path, recipient_emails = []):
    try:
        for recipient_email in recipient_emails:
            # Set up the MIME message
            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = recipient_email
            message['Subject'] = subject

            # Attach the email body
            message.attach(MIMEText(body, 'plain'))

            # Attach the file if specified
            if file_path and os.path.exists(file_path):
                with open(file_path, 'rb') as file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={os.path.basename(file_path)}'
                    )
                    message.attach(part)

            # Connect to the SMTP server and send the email
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()  # Secure the connection
                server.login(sender_email, sender_password)
                server.send_message(message)
                print("Email sent successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

