import win32com

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items

for message in messages:
    attachments = message.attachments
    for attachment in attachments:
        print(attachment.name)