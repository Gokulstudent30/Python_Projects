import pandas as pd
import smtplib
import schedule
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------- CONFIG ----------
EXCEL_FILE = "/home/gokul/PycharmProjects/PythonProject/Project_1_Mail/Mail_data.xlsx"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

SENDER_EMAIL = "gokul.sakthivelrajan@tessolve.com"
SENDER_PASSWORD = "vyxmbygjplrspwzb"    # Outlook App Password
RECEIVER_EMAIL = "gokul.sakthivelrajan@tessolve.com"

SUBJECT = "Weekly Report"
# -----------------------------------------------


# ---------- Convert multi-line text → bullet points ----------
def to_bullets(text):
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    if len(lines) <= 1:
        return text.strip()

    return "\n    • " + "\n    • ".join(lines)


# ---------- Main mail sending function ----------
def send_weekly_mail():
    print("Fetching data and sending weekly report...")

    # Read Excel
    df = pd.read_excel(EXCEL_FILE)

    # Extract values
    meetings = to_bullets(df["data"][0])
    highlights = to_bullets(df["data"][1])
    lowlights = to_bullets(df["data"][2])
    improvements = to_bullets(df["data"][3])

    # Build email body
    body = f"""
Hi Kamal!

Greetings for the day!!
Please find the progress for this week.
Meetings:
    {meetings}

Highlights:
    {highlights}

Lowlights:
    {lowlights}

Improvements:
    {improvements}

Regards,
Gokul Sakthivelrajan | T1897
"""

    # Prepare email
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(body, "plain"))

    # Send email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())

        print("Weekly report sent successfully!")

    except Exception as e:
        print("Error sending email:", e)
if __name__ == "__main__":
    send_weekly_mail()
