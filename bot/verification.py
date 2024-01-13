import smtplib
from email.mime.text import MIMEText

def send_verification_email(email, verification_code):
    # Set up your email configuration
    smtp_server = 'smtp.gmail.com'
    port = 587
    sender_email = 'alizaaamm@gmail.com'
    password = 'uwbc mrcf yhui ektn'

    # Compose the email
    subject = 'Verification Code for Your Bot'
    body = f'Your verification code is: {verification_code}'
    message = MIMEText(body)
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = email

    # Connect to the SMTP server and send the email
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, [email], message.as_string())

