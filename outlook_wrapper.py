import win32com.client
import os
from tkinter import *
import tempfile

mapi = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

n_msgs = 5
allowed_ext = ['.doc', '.docx', '.xls', '.xlsx', '.odt', '.ods', '.pdf', '.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif']


def get_temp_path(a):
    ext = '.' + a.FileName.rsplit('.',1)[1]
    print(ext)
    if ext.lower() in allowed_ext:
        fd, outpath = tempfile.mkstemp(ext)
        os.close(fd)
        a.SaveAsFile(outpath)
    else: return
    return outpath


inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort('ReceivedTime', True)
root = Tk()
root.title(f'Последние {n_msgs} писем')
for message in list(messages)[:n_msgs]:
    # message.PrintOut() Быстрая печать
    attachments = message.attachments
    if not attachments:
        continue
    Label(root, text=f'{message.sender}, {message.Subject}').pack()
    for attachment in attachments:
        Button(root, text=f'{attachment.FileName}, {round(attachment.Size / 1024 / 1024, 2)}, Mb',
               command=lambda a=attachment: os.startfile(get_temp_path(a))).pack()
