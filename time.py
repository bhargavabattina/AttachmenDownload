import win32com.client
import os
import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # 6 refers to the index of the inbox folder

# Loop through all the unread mails in the inbox
for mail in inbox.Items.Restrict("[Unread]=True"):
    if mail.Attachments.Count > 0: # Check if the mail has attachments
        for attachment in mail.Attachments:
            # Check if there is a folder with current date
            today = datetime.date.today().strftime('%Y-%m-%d')
            folder_path = os.path.join(os.getcwd(), today)
            if not os.path.exists(folder_path):
                os.mkdir(folder_path)

            # Download and move the attachment to the folder
            attachment.SaveAsFile(os.path.join(folder_path, attachment.FileName))