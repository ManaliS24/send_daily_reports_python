import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import time

def send_email(subject, body, to_email):
    from_email = "XXXX@gmail.com" #protecting email-id for security. Use your own
    password = '*****' #use application specific password
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = sm.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(from_email, password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print("Message Sent")

    except Exception as e:
        print(f"Error occured: {str(e)}")


def job():
    subject = 'daily report'
    body = 'This is daily report from XXXX@gmail.com'
    data=pd.read_excel("subscriber.xlsx")
    email_col=data.get("Email")
    list_of_emails=list(email_col)
    to_email = ", ".join(list_of_emails)
    print(to_email)
    send_email(subject, body, to_email)

# schedule.every().hour.at(":28").do(job)
schedule.every().day.at("08:15").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)