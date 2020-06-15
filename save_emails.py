import os
import settings
import win32com.client

from datetime import datetime, timedelta


class MailBox:

    default_delta = timedelta(weeks=40)

    def __init__(self, name, messages, delta=None):
        self.name = name
        self.messages = messages

        if delta:
            self.delta = delta
        else:
            self.delta = self.default_delta


mailboxes = []

# Get Inbox
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.getDefaultFolder(6)
mailboxes.append(MailBox(inbox.Name, inbox.Items))

# Get Sent Folder
sent = outlook.getDefaultFolder(5)
mailboxes.append(MailBox(sent.Name, sent.Items))

# Get Custom Subfolders
folders = inbox.Folders
for folder in folders:
    mailboxes.append(MailBox(folder.Name, folder.Items, delta=timedelta(weeks=16)))

now = datetime.now()

forbidden_chars = [':', '=>', '/', '?', '->', '"', '”', '!', '*', '(', ')', '|', '#']

for mailbox in mailboxes:

    start = now - mailbox.delta

    for message in mailbox.messages:
        print(message.subject)
        try:
            received_time = datetime.strptime(str(message.receivedTime).strip('+00:00'), '%Y-%m-%d %H:%M:%S.%f')
            print(received_time)

            # If the email is before the earliest date to keep emails (in start)
            if received_time < start:
                year = str(received_time.year)
                month = str(received_time.month).zfill(2)

                # Make sure folder, year, and month directories are created
                folder_path = '%s%s' % (settings.archive_folder, mailbox.name)
                year_path = '%s\\%s' % (folder_path, year)
                month_path = '%s\\%s' % (year_path, month)
                if not os.path.isdir(folder_path):
                    os.mkdir(folder_path)
                if not os.path.isdir(year_path):
                    os.mkdir(year_path)
                if not os.path.isdir(month_path):
                    os.mkdir(month_path)

                message_subject = message.subject.replace(' ', '_')
                for char in forbidden_chars:
                    message_subject = message_subject.replace(char, '')

                message_path = '%s\\%s-%s.msg' \
                               % (month_path,
                                  received_time.strftime('%y-%m-%d-%H-%M'),
                                  message_subject
                                  )
                print(message_path)
                message.saveAs(message_path)
                message.delete()
        except (ValueError, AttributeError):
            pass
