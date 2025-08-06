import pandas as pd
import win32com.client as win32
import os

df = pd.read_excel(r'Documents\Email List.xlsx')
outlook = win32.Dispatch('outlook.application')

for index, row in df.iterrows():
    recipient = row['Email']
    name = row['Name']
    attachment_path = row['Attachment Path']
    cc = row['CC']

    # Fix for attachment path error
    if not isinstance(attachment_path, str) or not os.path.isfile(attachment_path):
        print(f"Attachment not found or invalid for {name}: {attachment_path}")
        continue

    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.CC = f"address1@email.com; ...; {cc}"
    mail.SentOnBehalfOfName = "myemail@mail.com"
    mail.Subject = "Email Subject Here"
    mail.GetInspector
    signature = mail.HTMLBody
    mail.HTMLBody = f"""<p>Dear {name},</p>
<p>Thank you for trying out my program! Attached, you'll find whatever attachment you linked in the spreadhseet, have a good one.</p>
<p>Thank you,</p>
{signature}
"""
    mail.Attachments.Add(attachment_path)
    mail.Display()
