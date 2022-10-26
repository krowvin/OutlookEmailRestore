import win32com.client as client
from pywintypes import com_error
import time
# startup outlook instance
outlook = client.Dispatch('Outlook.Application')

# get namespace so that we can access folders
namespace = outlook.GetNameSpace('MAPI')

# Select your inbox - Change to your Email Adress or Primary Folder name in Outlook
account = namespace.Folders['myusername@mydomain.com']
# get the inbox folder, specifically
deleted_items = account.Folders['Deleted Items']
print(f"Attempting to restore: {len(deleted_items.Items)}...")
print("You will want to run this a few times to be sure it got what it could!")
time.sleep(4)
inbox = account.Folders['Inbox']
# Loop the deleted items
for i in deleted_items.Items:
    try: 
        # Attempt to move the item
        try:
            print(f'[{i.SentOn.strftime("%d-%m-%y %H:%M")}] Restored: { i.Move(inbox)}')
        except AttributeError:
            print(f"[Unknown] Restored: {i.Move(inbox)}")
        time.sleep(0.05)
    except com_error as err:
        # If it fails (due to certificate) skip it!
        err = str(err).replace("\n", "\n\t")
        print(f"[ERROR] SKIPPING {i.subject}\n\t{err}")

