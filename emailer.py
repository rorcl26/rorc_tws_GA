import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from datetime import datetime

class Emailer:
    def __init__(self, sender_email, app_password):
        self.sender_email = sender_email
        self.app_password = app_password

    def send_email(self, recipient_emails, subject, html_content):
        if not self.sender_email or not self.app_password:
            print("Error: Gmail credentials not found.")
            return False
            
        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = self.sender_email
            
            if isinstance(recipient_emails, list):
                msg['To'] = ", ".join(recipient_emails)
                rcpt = recipient_emails
            else:
                msg['To'] = recipient_emails
                rcpt = [r.strip() for r in recipient_emails.split(',')]
                
            part2 = MIMEText(html_content, 'html')
            msg.attach(part2)
            
            print(f"Connecting to SMTP server...")
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            server.login(self.sender_email, self.app_password)
            server.sendmail(self.sender_email, rcpt, msg.as_string())
            server.quit()
            
            print(f"Email sent successfully to {msg['To']}")
            return True
        except Exception as e:
            print(f"Failed to send email: {str(e)}")
            return False

if __name__ == "__main__":
    # Test block
    from dotenv import load_dotenv
    load_dotenv()
    
    sender = os.getenv("GMAIL_USER")
    pwd = os.getenv("GMAIL_APP_PASSWORD")
    recipients = os.getenv("MAIL_RECIPIENT", "test@example.com")
    
    emailer = Emailer(sender, pwd)
    html = "<h1>Test Email</h1><p>This is a test from Taiwan Stock Analyzer.</p>"
    # emailer.send_email(recipients, "TW Stock Analysis Test", html)
