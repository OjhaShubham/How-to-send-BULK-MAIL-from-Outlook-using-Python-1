# How-to-send-BULK-MAIL-from-Outlook-using-Python-1
How to send BULK MAIL from Outlook using Python # 1

So, you want to send bulk mail from Outlook? Technically, you cannot send bulk mail from Outlook because of the sending limits, but there are some methods that you can use to work around them. The most challenging limit is the 30 emails per minute limit. From my experience, you can send more than 30 per minute through your Outbox, but they will not be delivered to the recipient at a rate faster than 30 per minute. Also, you want to avoid pushing thousands of emails to your Outbox at once because it can cause Outlook to become unresponsive and crash... potentially eating away all your time savings by having to restart. I've been able to successfully push about 100 emails per minute and keep everything running smoothly. However, for this example, I'm going to push the emails through the Outbox at a rate of 30 per minute, which is the same rate at which they will be delivered.


import csv
from time import sleep
import win32com.client as client

# create template for message body
template = "{}, please submit your time as soon as possible!"

# open distribution list
with open('people.csv', 'r', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

# chunk distribution list into blocks of 30
chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]

# create outlook instance
outlook = client.Dispatch('Outlook.Application')

# iterate through chunks and send mail
for chunk in chunks:
    # iterate through each recipient in chunk and send mail
    for name, address in chunk:
        message = outlook.CreateItem(0)
        message.To = address
        message.Subject = "Your time entry is past due!"
        message.Body = template.format(name)
        message.Send()

    # wait 60 seconds before sending next chunk
    sleep(60)
