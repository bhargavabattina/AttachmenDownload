import os
import win32com.client

# Create a new instance of the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application")

# Get the default inbox folder
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)

# Define the file format you want to download, e.g., PDF
file_format = ".pdf","xlsx"

# Get all the unread emails in the inbox folder
unread_emails = inbox.Items.Restrict("[Unread]=True")

# Loop through all the unread emails
for email in unread_emails:
	message_date = email.ReceivedTime.date()
	folder_name = message_date.strftime("%Y-%m-%d")
	# Define the output directory for the attachment
	output_directory = os.path.join(os.getcwd(), f"C:\\Users\\bharu\\OneDrive\\Desktop\\new\\{folder_name}")
	os.makedirs(output_directory, exist_ok=True)

	# Loop through all the attachments in the email
	for attachment in email.Attachments:

		# Check if the attachment is of the specified file format
		if attachment.FileName.endswith(file_format):

			# Save the attachment to the output directory
			attachment.SaveAsFile(os.path.join(output_directory, attachment.FileName))

			# Mark the email as read
			email.UnRead = True
			email.Save()