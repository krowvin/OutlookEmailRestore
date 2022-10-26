# Restore Outlook Emails
* Restore vast amounts of emails while sipping coffee/tea ☕
* Skip over emails with invalid certificates - then restore those few emails manually
  * Error: `Your Digital ID name cannot be found by the underlying security 
system.`
  - This error is caused by switching to a new laptop and trying to bulk restore emails. Your new laptop will not have the previous TPM chip and for some reason the bulk restore cannot solve this
* 

This script allows you to restore deleted emails from your Deleted Items folder using Python. It will skip over these digital ID errors. When it's done the only items left in your inbox will be the emails with certificates. You can then manually restore these one at a time (instead of bulk/having to restore everything)

# Requirments
* No libraries Require
* Python3
* System with your Outlook email on it
