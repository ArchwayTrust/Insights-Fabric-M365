# Fabric Python Helper
Collection of useful python code for automation in Microsoft Fabric.

## GraphAPI
Authenticates as an app registration (service principal), using either a certificate held in Azure Key Vault (recommended) or a client secret. See `fabric_python_helper/auth.py` for what the certificate and key vault need.

The app needs access to the mailbox. Use Exchange RBAC for Applications to limit it to specific mailboxes, assigning `Application Mail.ReadWrite` (search, download attachments, delete) and `Application Mail.Send` (send emails). Don't grant Mail permissions as Entra ID API permissions, as those apply to every mailbox in the tenant.

In Fabric notebook run:
```
from fabric_python_helper import graph_api as gr
from fabric_python_helper.auth import KeyVaultCertificateCredential
```
### Emails
Create an instance for a mailbox, authenticating with a certificate:
```
tenant_id = "23...."
client_id = "00df..."
akv_url = "https://name.vault.azure.net/"
certificate_name = "Certificate Name"

credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, certificate_name)
email_account = gr.Emails(credential, "mailbox@address.com")
```

Or with a client secret:
```
from azure.identity import ClientSecretCredential

client_secret = notebookutils.credentials.getSecret(akv_url, "Secret Name")
credential = ClientSecretCredential(tenant_id, client_id, client_secret)
email_account = gr.Emails(credential, "mailbox@address.com")
```

Each method also takes an optional `shared_mailbox_email` to use a different mailbox.

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