import win32com.client
import os
import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # 6 refers to the inbox folder

unread_msgs = inbox.Items.Restrict("[Unread]=True") # get only unread messages
for msg in unread_msgs:
    if msg.Attachments.Count > 0:
        # get the date of the email and create a folder with that name
        email_date = datetime.datetime.strptime(str(msg.ReceivedTime), '%m %d %Y %H:%M:%S.SSSZZZ')
        folder_name = email_date.strftime("%Y-%m-%d") # format folder name as YYYY-MM-DD
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        # save all attachments to the newly created folder
        for attachment in msg.Attachments:
            attachment.SaveAsFile(os.path.join(folder_name, attachment.FileName))