import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import glob
from datetime import datetime

def send_report_email(sender_email, sender_password, receiver_email, attachment_path, subject=None):
    if subject is None:
        subject = f"Daily Lending Report - {datetime.now().strftime('%Y-%m-%d')}"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    body = "Attached is the daily lending portfolio report.\n"
    body += "This report was generated automatically via GitHub Actions."
    msg.attach(MIMEText(body, 'plain'))

    if not os.path.exists(attachment_path):
        print(f"Error: File not found at {attachment_path}")
        return False

    try:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f"attachment; filename= {os.path.basename(attachment_path)}",
            )
            msg.attach(part)
    except Exception as e:
        print(f"Error attaching file: {e}")
        return False

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        print("Email sent successfully!")
        return True
    except Exception as e:
        print(f"Email sending failed: {e}")
        return False

if __name__ == "__main__":
    SENDER_EMAIL = os.getenv("SENDER_EMAIL")
    SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
    RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

    if not all([SENDER_EMAIL, SENDER_PASSWORD, RECEIVER_EMAIL]):
        print("Error: Missing email environment variables.")
        print("Set SENDER_EMAIL, SENDER_PASSWORD, RECEIVER_EMAIL")
        exit(1)

    reports_dir = 'reports'
    if not os.path.exists(reports_dir):
        print("Error: reports folder not found.")
        exit(1)

    all_files = glob.glob(os.path.join(reports_dir, '*.xlsx'))
    excel_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

    if not excel_files:
        print("Error: No Excel files found.")
        exit(1)

    latest_file = max(excel_files, key=os.path.getctime)
    print(f"Sending file: {latest_file}")

    success = send_report_email(SENDER_EMAIL, SENDER_PASSWORD, RECEIVER_EMAIL, latest_file)

    if success:
        print("Email sent successfully.")
    else:
        print("Email sending failed.")
        exit(1)