import win32com.client as win32
from jinja2 import Template
import pandas as pd
import os
import time
from dotenv import load_dotenv
import datetime


load_dotenv()
CC_EMAIL = os.getenv("CC_EMAIL")
SEND_FOLDER = os.getenv("SEND_FOLDER")
RESUME = os.getenv("RESUME")
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
root_folder = namespace.GetDefaultFolder(6).Parent
send_folder = root_folder.Folders.Item(SEND_FOLDER)
inbox = root_folder.Folders.Item("Inbox")

def main():

    with open("./templates/subject.txt", "r") as f:
        subject_template = Template(f.read())
    
    with open("./templates/connection_email.html", 'r') as f:
        html_template = Template(f.read())

    df = pd.read_csv("./data/dummy_data.csv")

    resume_path = f"./attachments/{RESUME}"

    for _, row in df.iterrows():
        data = row.to_dict()
        subject = subject_template.render(**data)
        body = html_template.render(**data)
        
        mail = outlook.CreateItem(0)
        mail.To = data["email"]
        mail.CC = CC_EMAIL
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Attachments.Add(os.path.abspath(resume_path))
        mail.SaveSentMessageFolder = send_folder 


        try:
            mail.Send()
            print(f"Email sent successfully to {data['email']}")
        except Exception as e:
            print(f"Failed to send email to {data['email']}: {str(e)}")

        time.sleep(3)


def delete_failed_email(hours=3):
    past_time = (
        datetime.datetime.now() - 
        datetime.timedelta(days=2)
    ).strftime('%m/%d/%Y %H:%M %p')
    restriction = f"[ReceivedTime] >= '{past_time}'"
    recent_items = inbox.Items.Restrict(restriction)
    recent_items.Sort("[ReceivedTime]", True)

    count = 0
    for item in recent_items:
        subject = item.Subject.lower()
        
        if "undeliverable" in subject:
            item.Delete()
            count += 1
    
    print(f"{count} Undelivered messages deleted")



if __name__ == "__main__":
    main()
    time.sleep(5)
    delete_failed_email(hours=3)
