# Fabric Python Helper
Collection of useful python code for automation in Microsoft Fabric.

## GraphAPI
Requires an app registration with delegated User.Read, Mail.ReadWrite, Mail.ReadWrite.Shared, Mail.Send and offline_access scopes. The public client flow must also be enabled.
Requires an azure key vault which the notebook owner has permissions to update and read secrets on.

In Fabric notebook run:
```
from fabric_python_helper import graph_api as gr
```
### Emails
Create an instance for an email account:
```
tennant_id = "23...."
client_id = "00df..."
akv_url = "https://name.vault.azure.net/"
refresh_secret_name = "Secret Name"

email_account = gr.Emails(tennant_id, client_id, akv_url, refresh_secret_name)

```
#### Initial Authentication

On first run you will need to authenticate, this is using device flow authentication:

```
email_account.get_initial_tokens()
```

It will give you a link and device code which you need to follow and authenticate with.

#### Connect to Graph API
```
email_account.connect()
```
This retrieves an access token for the current session (lasts an hour) and replaces the refresh token in the key vault. Currently unclear as to whether this resets the expiry period.

#### Find the message id (currently returns first email only)
```
message_id = email_account.search_message_by_subject_and_sender("Attachment Test", "email@address.com")
```

#### Get list of attachment ids:
```
attachments = email_account.get_attachment_ids_and_names(message_id)
```

#### Get attachments:
(You can loop through the list of attachments if required)
```
attachment = email_account.download_attachment(message_id = message_id, attachment_id = attachments[0][0], is_binary=False, encoding="utf-8")
```

#### Save to onelake:
Assuming it's a text file eg you're here because you're grabbing a csv from an email you can just use 

```
mssparkutils.fs.put(file_path, attachment, overwrite=True)
```
If you returned a binary other methods would be required.

#### Delete the email:

```
email_account.delete_email(message_id)
```

#### Send an email

```
email_addresses = ["email1@address.co.uk", "email2@address.co.uk"]

subject = "Test Email"

content = """
Hi,
I'm the content of the email.
I'm on my own line.

    From Ben               
"""

email_account.send_email(subject, content, email_addresses)
```

## Building a Wheel File
In order to build I had to add to create a pip.ini file in `C:\ProgramData\pip\`
<br>
Which contained:
```
[global]
trusted-host = files.pythonhosted.org
               pypi.org
```
You can then build with:
- `pip install build`
- `python -m build`