import os
import settings
import win32com.client

from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.getDefaultFolder(6)

messages = inbox.Items

now = datetime.now()
start = now - timedelta(weeks=40)
print(start)
forbidden_chars = [':', '=>', '/', '?', '->', '"', '”', '!', '*', '(', ')', '|', '#']

for message in messages:
    print(message.subject)
    try:
        received_time = datetime.strptime(str(message.receivedTime).strip('+00:00'), '%Y-%m-%d %H:%M:%S.%f')
        print(received_time)

        # If the email is before the earliest date to keep emails (in start)
        if received_time < start:
            year = str(received_time.year)
            month = str(received_time.month).zfill(2)

            # Make sure year and month directories are created
            year_path = '%s%s' % (settings.archive_folder, year)
            month_path = '%s\\%s' % (year_path, month)
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
