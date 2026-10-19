# Fabric Python Helper

Python helpers for automation in Microsoft Fabric notebooks.

- **`fabric_python_helper.auth`** gets app-only access tokens for Microsoft APIs (Microsoft Graph, Power BI, Fabric, OneLake and others) as a service principal, using a certificate held in Azure Key Vault. Key Vault signs each request, so the certificate's private key never leaves the key vault.
- **`fabric_python_helper.graph_api`** finds emails in a mailbox, downloads their attachments, deletes them and sends emails, using Microsoft Graph.

The package is built for Fabric notebooks. It uses `notebookutils` to call Key Vault as the identity running the notebook, so it won't run outside Fabric (or Synapse).

## Contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [Authentication](#authentication)
  - [How certificate authentication works](#how-certificate-authentication-works)
  - [Setting up a certificate](#setting-up-a-certificate)
  - [Using the credential](#using-the-credential)
  - [Scopes](#scopes)
  - [Using a client secret instead](#using-a-client-secret-instead)
  - [Renewing the certificate](#renewing-the-certificate)
- [Emails](#emails)
  - [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox)
  - [Creating an Emails instance](#creating-an-emails-instance)
  - [Finding an email](#finding-an-email)
  - [Downloading attachments](#downloading-attachments)
  - [Deleting an email](#deleting-an-email)
  - [Sending an email](#sending-an-email)
  - [Full example](#full-example)
- [Troubleshooting](#troubleshooting)
- [Upgrading from 0.3.x](#upgrading-from-03x)
- [Building and publishing](#building-and-publishing)
- [Licence](#licence)

## Installation

The package is published on PyPI as [`fabric_python_helper`](https://pypi.org/project/fabric-python-helper/). Its dependencies, `azure-core` and `requests`, are already in the Fabric runtime.

**In a Fabric environment** (use this for scheduled notebooks and pipelines):

1. Open the environment your notebooks use.
2. Under **Public libraries**, select **Add from PyPI**, enter `fabric_python_helper` and choose a version, for example `0.4.0`.
3. Select **Publish**. Notebooks pick up the new version when their next session starts.

**In a single notebook session** (useful for testing a new version):

```python
%pip install fabric_python_helper==0.4.0
```

Always pin a version. 0.4.0 has breaking changes from 0.3.x, see [Upgrading from 0.3.x](#upgrading-from-03x).

## Quick start

Read the latest report email from a mailbox, save its attachment to a lakehouse and delete the email:

```python
from fabric_python_helper.auth import KeyVaultCertificateCredential
from fabric_python_helper.graph_api import Emails

credential = KeyVaultCertificateCredential(
    tenant_id="<tenant-id>",
    client_id="<app-registration-client-id>",
    akv_url="https://<key-vault-name>.vault.azure.net/",
    certificate_name="<certificate-name>",
)

mailbox = Emails(credential, "reports@contoso.com")

message_id = mailbox.search_message_by_subject_and_sender("Daily Sales Report", "noreply@supplier.com")
attachment_id, file_name = mailbox.get_attachment_ids_and_names(message_id)[0]
csv_text = mailbox.download_attachment(message_id, attachment_id)

lakehouse = "abfss://<workspace>@onelake.dfs.fabric.microsoft.com/<lakehouse>.Lakehouse"
notebookutils.fs.put(f"{lakehouse}/Files/Reports/{file_name}", csv_text, overwrite=True)

mailbox.delete_email(message_id)
```

Before this works you need to [set up a certificate](#setting-up-a-certificate) and [give the app access to the mailbox](#giving-the-app-access-to-a-mailbox).

## Authentication

### How certificate authentication works

`KeyVaultCertificateCredential` authenticates as an app registration (service principal) using the OAuth 2.0 client credentials flow with a certificate:

1. It reads the public part of the certificate from Key Vault.
2. It builds a client assertion: a JSON Web Token (JWT) that is valid for 10 minutes, identifies the app and names the certificate by its SHA-256 thumbprint.
3. Key Vault signs the assertion with the certificate's private key (RSASSA-PSS, `PS256`). The private key never leaves the key vault.
4. The signed assertion is exchanged with Microsoft Entra ID for an access token.
5. Access tokens are cached per scope and renewed when they have less than 5 minutes left, so the credential can be used for long-running notebooks.

Two identities are involved, and each needs different permissions:

| Identity | What it is | What it needs |
| --- | --- | --- |
| The identity running the notebook | You, when running interactively. When a pipeline runs the notebook, whatever the notebook activity runs as, for example the workspace identity. | Permission to read the certificate and sign with its key in Key Vault. |
| The app registration | The service principal the access tokens are issued to. | Access to the API being called, for example access to a mailbox in Exchange. |

Anyone who can sign with the key can act as the app, so treat the Key Vault **Sign** permission as you would the app's password.

### Setting up a certificate

These steps use the Azure CLI, for example in Azure Cloud Shell. Each step also describes the portal equivalent.

**1. Create an app registration**

In the Microsoft Entra admin center, go to **App registrations** > **New registration**. Choose **Single tenant** and leave the redirect URI empty. Note the **Application (client) ID** and **Directory (tenant) ID**.

Only add API permissions if the API you're calling needs them, see [Scopes](#scopes). For mailboxes, don't add any: access is granted in Exchange instead.

**2. Create the certificate in Key Vault with a non-exportable private key**

Save this policy as `policy.json`:

```json
{
  "issuerParameters": { "name": "Self" },
  "keyProperties": { "exportable": false, "keyType": "RSA", "keySize": 2048, "reuseKey": false },
  "secretProperties": { "contentType": "application/x-pkcs12" },
  "x509CertificateProperties": {
    "subject": "CN=fabric-mailbox-app",
    "validityInMonths": 12,
    "keyUsage": ["digitalSignature"]
  },
  "lifetimeActions": [
    { "trigger": { "daysBeforeExpiry": 30 }, "action": { "actionType": "EmailContacts" } }
  ]
}
```

Then create the certificate:

```bash
az keyvault certificate create --vault-name <key-vault-name> --name <certificate-name> --policy @policy.json
```

`"exportable": false` means nobody, including Key Vault administrators, can download the private key. The lifetime action emails the vault's certificate contacts 30 days before expiry. Add yourself as a contact if you aren't one already:

```bash
az keyvault certificate contact add --vault-name <key-vault-name> --email <you@contoso.com>
```

In the portal: Key Vault > **Certificates** > **Generate/Import**, then under **Advanced Policy Configuration** set **Exportable Private Key** to **No**.

**3. Upload the public certificate to the app registration**

```bash
az keyvault certificate download --vault-name <key-vault-name> --name <certificate-name> --file cert.pem --encoding PEM
az ad app credential reset --id <client-id> --cert "@cert.pem" --append
```

`--append` keeps the app's existing credentials. In the portal: download the certificate in CER format from its version in Key Vault, then in the app registration go to **Certificates & secrets** > **Certificates** > **Upload certificate**.

**4. Let the identity running the notebook use the certificate**

With Azure role-based access control (the default for new key vaults), assign two roles scoped to just this certificate and its key. `<object-id>` is the object ID of the user, group or service principal that runs the notebooks. For a workspace identity, use the object ID of its service principal in Entra ID, which has the same name as the workspace.

```bash
vault_id=$(az keyvault show --name <key-vault-name> --query id --output tsv)

az role assignment create --assignee <object-id> --role "Key Vault Reader" \
    --scope "$vault_id/certificates/<certificate-name>"

az role assignment create --assignee <object-id> --role "Key Vault Crypto User" \
    --scope "$vault_id/keys/<certificate-name>"
```

If the key vault uses access policies instead:

```bash
az keyvault set-policy --name <key-vault-name> --object-id <object-id> \
    --certificate-permissions get --key-permissions sign
```

Avoid broader roles such as Key Vault Certificate User, Key Vault Secrets User or Key Vault Administrator. They aren't needed, and on an exportable certificate they allow reading the private key.

**5. Give the app access to the API you're calling**

See the [Scopes](#scopes) table, and [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox) for email.

### Using the credential

```python
from fabric_python_helper.auth import KeyVaultCertificateCredential, Scopes

credential = KeyVaultCertificateCredential(
    tenant_id="<tenant-id>",
    client_id="<client-id>",
    akv_url="https://<key-vault-name>.vault.azure.net/",
    certificate_name="<certificate-name>",
)
```

| Parameter | Required | Description |
| --- | --- | --- |
| `tenant_id` | Yes | Directory (tenant) ID. |
| `client_id` | Yes | Application (client) ID of the app registration. |
| `akv_url` | Yes | Key Vault URL, `https://<key-vault-name>.vault.azure.net/`. |
| `certificate_name` | Yes | Name of the certificate in Key Vault. |
| `scope` | No | Default scope for `get_access_token()`. Not needed if you always pass a scope, or only use the credential with semantic link or Azure SDK clients. |
| `certificate_version` | No | Certificate version to sign with. Defaults to the latest version. See [Renewing the certificate](#renewing-the-certificate). |

No calls are made until a token is requested.

**Calling an API directly**

`get_access_token(scope)` returns a token string. Call it before each request; it returns the cached token until it's close to expiring.

```python
import requests

token = credential.get_access_token(Scopes.FABRIC)
response = requests.get(
    "https://api.fabric.microsoft.com/v1/workspaces",
    headers={"Authorization": f"Bearer {token}"},
)
workspaces = response.json()["value"]
```

If you mostly call one API, set it as the default scope:

```python
graph = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, certificate_name, scope=Scopes.GRAPH)
token = graph.get_access_token()
```

**Semantic link (sempy)**

The credential implements the Azure SDK `TokenCredential` interface (`get_token`), so semantic link functions accept it through their `credential` parameter. This needs semantic link 0.12 or later. If your runtime has an older version, add `semantic-link` to the environment or run `%pip install -U semantic-link`.

```python
import sempy.fabric as fabric

fabric.refresh_dataset("Sales", workspace="Finance", credential=credential)
```

To use it for every semantic link call in a block:

```python
from fabric.analytics.environment.credentials import SetFabricAnalyticsDefaultTokenCredentials

with SetFabricAnalyticsDefaultTokenCredentials(credential):
    items = fabric.list_items(workspace="Finance")
```

Some semantic link functions don't support service principals, for example `execute_tmsl`, `evaluate_measure` and `read_table(mode="rest")`. Service principals also can't access anyone's "My workspace".

**Azure SDK clients**

Any Azure SDK client that takes a `credential` works too. For example, reading OneLake as the app:

```python
from azure.storage.filedatalake import DataLakeServiceClient

onelake = DataLakeServiceClient("https://onelake.dfs.fabric.microsoft.com", credential=credential)
workspace = onelake.get_file_system_client("<workspace-name>")

for path in workspace.get_paths("<lakehouse-name>.Lakehouse/Files"):
    print(path.name)
```

### Scopes

Tokens are requested with the client credentials flow, so a scope is always an API's resource URI followed by `/.default`. `/.default` means "every permission this app has been granted on that API". Apps can't request individual permissions such as `Mail.Read`.

Entra ID issues a token whether or not the app has access to the API. If the API then returns 401 or 403, check the last column.

| Constant | Scope | API | How to give the app access |
| --- | --- | --- | --- |
| `Scopes.GRAPH` | `https://graph.microsoft.com/.default` | Microsoft Graph | Application permissions on the app registration, with admin consent. For mailboxes, use [Exchange RBAC](#giving-the-app-access-to-a-mailbox) instead. |
| `Scopes.POWER_BI` | `https://analysis.windows.net/powerbi/api/.default` | Power BI REST API | Add the service principal to the workspace, and allow service principals to use the APIs in the Fabric admin portal (**Tenant settings** > **Developer settings**). |
| `Scopes.FABRIC` | `https://api.fabric.microsoft.com/.default` | Fabric REST API | As for Power BI. |
| `Scopes.STORAGE` | `https://storage.azure.com/.default` | Azure Storage and OneLake | OneLake: a workspace role or item permissions. Storage accounts: an Azure role such as Storage Blob Data Reader. |
| `Scopes.KEY_VAULT` | `https://vault.azure.net/.default` | Azure Key Vault | An Azure role or access policy on the key vault. |
| `Scopes.SQL` | `https://database.windows.net/.default` | Azure SQL, Fabric warehouses, SQL analytics endpoints and SQL databases | Fabric: a workspace role or item permissions. Azure SQL: `CREATE USER [<app-name>] FROM EXTERNAL PROVIDER`. |
| `Scopes.KUSTO` | `https://api.kusto.windows.net/.default` | Azure Data Explorer, Fabric Eventhouse and KQL databases | A workspace role, or a database role such as `.add database <db> viewers ('aadapp=<client-id>;<tenant-id>')`. |
| `Scopes.AZURE_MANAGEMENT` | `https://management.azure.com/.default` | Azure Resource Manager | An Azure role on the subscription, resource group or resource. |
| `Scopes.LOG_ANALYTICS` | `https://api.loganalytics.io/.default` | Azure Monitor Log Analytics | Log Analytics Reader on the Log Analytics workspace. |
| (none) | `https://<tenant>.sharepoint.com/.default` | SharePoint REST API | SharePoint application permissions, ideally `Sites.Selected`. SharePoint only allows app-only access with a certificate; client secrets are rejected. |
| (none) | `https://<org>.crm11.dynamics.com/.default` | Dataverse | Use your environment URL. Add the app as an application user in the Power Platform admin center. |

One credential can get tokens for any number of these APIs.

**When you don't need this package:** `notebookutils.credentials.getToken()` already returns tokens for the identity running the notebook, for the `storage`, `pbi`, `keyvault` and `kusto` audiences. Use this package when you need another API (such as Graph or SharePoint) or need to act as a separate app.

### Using a client secret instead

Certificates are recommended, but anything in this package that takes a credential also accepts `azure.identity.ClientSecretCredential`, which is already in the Fabric runtime:

```python
from azure.identity import ClientSecretCredential
from fabric_python_helper.graph_api import Emails

client_secret = notebookutils.credentials.getSecret("https://<key-vault-name>.vault.azure.net/", "<secret-name>")
credential = ClientSecretCredential("<tenant-id>", "<client-id>", client_secret)

mailbox = Emails(credential, "reports@contoso.com")
```

The identity running the notebook needs **Get** on the secret. Compared with a certificate:

- The secret is read into the notebook's memory, whereas the certificate's private key never leaves Key Vault.
- Client secrets expire after at most 2 years and have to be copied into Key Vault by hand when they're renewed.
- SharePoint's REST API rejects client secrets for app-only access.

### Renewing the certificate

By default the credential signs with the latest version of the certificate, and Entra ID only accepts certificates that have been uploaded to the app registration. When you renew:

1. Create a new version by running the same `az keyvault certificate create` command as before. From this moment the credential signs with the new version.
2. Straight away, download the new version and upload it with `az ad app credential reset ... --append`, as in step 3 of the setup.
3. Once notebooks are running successfully, remove the old certificate from the app registration.

To avoid any failures between steps 1 and 2, first pin the current version with `certificate_version="<current-version>"`, renew and upload, then remove the pin.

## Emails

`fabric_python_helper.graph_api.Emails` works with one mailbox at a time using Microsoft Graph, authenticating as the app rather than as a user.

### Giving the app access to a mailbox

Use [RBAC for Applications in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac) to give the app access to specific mailboxes only. Run these as an Exchange administrator (Exchange Administrator role, or the Organization Management role group), using the [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) PowerShell module:

```powershell
Connect-ExchangeOnline -UserPrincipalName <admin@contoso.com>

# From Entra ID > Enterprise applications (not App registrations)
$appId    = "<application-id>"
$objectId = "<object-id>"

# 1. Register the app's service principal with Exchange
New-ServicePrincipal -AppId $appId -ObjectId $objectId -DisplayName "<app-name>"

# 2. Create a scope containing only the mailbox. Matching on GUID keeps working if the address changes.
$mbx = Get-Mailbox -Identity "reports@contoso.com"
New-ManagementScope -Name "<app-name> mailboxes" -RecipientRestrictionFilter "GUID -eq '$($mbx.GUID)'"

# 3. Assign the role, limited to that scope
New-ManagementRoleAssignment -App $objectId -Role "Application Mail.ReadWrite" -CustomResourceScope "<app-name> mailboxes"

# Only if the app needs to send email
New-ManagementRoleAssignment -App $objectId -Role "Application Mail.Send" -CustomResourceScope "<app-name> mailboxes"

# 4. Check: the mailbox should show InScope = True, any other mailbox False
Test-ServicePrincipalAuthorization -Identity $objectId -Resource "reports@contoso.com" | Format-Table
Test-ServicePrincipalAuthorization -Identity $objectId -Resource "someone.else@contoso.com" | Format-Table
```

To cover several mailboxes, use a filter such as `MemberOfGroup -eq '<distinguished name of a mail-enabled security group>'` and add the mailboxes to the group.

`Test-ServicePrincipalAuthorization` reflects changes immediately, but Graph can take from 30 minutes to 2 hours to apply them.

**Don't** add `Mail.Read`, `Mail.ReadWrite` or `Mail.Send` as application permissions on the app registration. Those apply to every mailbox in the tenant, and they add to the Exchange role assignment rather than being limited by it.

Roles needed by each method:

| Method | Exchange role |
| --- | --- |
| `search_message_by_subject_and_sender`, `get_attachment_ids_and_names`, `download_attachment` | `Application Mail.Read` or `Application Mail.ReadWrite` |
| `delete_email` | `Application Mail.ReadWrite` |
| `send_email` | `Application Mail.Send` |

### Creating an Emails instance

```python
from fabric_python_helper.graph_api import Emails

mailbox = Emails(credential, "reports@contoso.com")
```

| Parameter | Description |
| --- | --- |
| `credential` | A `KeyVaultCertificateCredential`, a `ClientSecretCredential`, or any other credential with a `get_token` method. |
| `mailbox` | Email address of the mailbox to use. |

A token is requested from the credential for each call, and the credential caches it, so one instance can be used for the whole notebook.

Every method also takes an optional `shared_mailbox_email` argument to work with a different mailbox for that call. The app needs access to that mailbox too.

### Finding an email

```python
message_id = mailbox.search_message_by_subject_and_sender(
    subject="Daily Sales Report",
    sender_email="noreply@supplier.com",
)
```

| Parameter | Default | Description |
| --- | --- | --- |
| `subject` | | The exact subject to match. |
| `sender_email` | | The sender's email address. |
| `only_search_inbox` | `True` | Search only the Inbox. Set to `False` to search every folder. |
| `only_return_latest` | `True` | Return only the ID of the most recent match. Set to `False` to return every match. |
| `shared_mailbox_email` | `None` | Search a different mailbox. |

With `only_return_latest=True` it returns the message ID as a string. With `False` it returns a list of messages, newest first, each a dictionary with `id`, `subject`, `from`, `receivedDateTime`, `parentFolderId` and `hasAttachments`.

If nothing matches it raises an exception with the message `No matching messages found.` Use this to handle emails that don't always arrive:

```python
try:
    message_id = mailbox.search_message_by_subject_and_sender("Weekly Exceptions", "noreply@supplier.com")
except Exception as e:
    if str(e) != "No matching messages found.":
        raise
    message_id = None
    print("No exceptions report this week.")
```

Processing every matching email, oldest first:

```python
messages = mailbox.search_message_by_subject_and_sender(
    "Daily Sales Report", "noreply@supplier.com", only_return_latest=False
)

for message in reversed(messages):
    print(message["receivedDateTime"], message["id"])
```

### Downloading attachments

`get_attachment_ids_and_names` returns a list of `(attachment_id, file_name)` tuples:

```python
attachments = mailbox.get_attachment_ids_and_names(message_id)
# [('AAMkAD...', 'sales.csv'), ('AAMkAD...', 'logo.png')]
```

`download_attachment` returns the attachment's content. By default it's decoded as UTF-8 text, which suits CSV and text files:

```python
for attachment_id, file_name in attachments:
    if file_name.endswith(".csv"):
        csv_text = mailbox.download_attachment(message_id, attachment_id)
        notebookutils.fs.put(f"{lakehouse}/Files/Sales/{file_name}", csv_text, overwrite=True)
```

Use `encoding` for text in another encoding, for example `encoding="cp1252"` for some Windows exports.

For binary files such as Excel workbooks or PDFs, set `is_binary=True` to get bytes:

```python
import io
import pandas as pd

content = mailbox.download_attachment(message_id, attachment_id, is_binary=True)

# Read an Excel workbook directly
df = pd.read_excel(io.BytesIO(content))

# Or save it to the notebook's default lakehouse
with open(f"/lakehouse/default/Files/Sales/{file_name}", "wb") as f:
    f.write(content)
```

Only file attachments can be downloaded. Attached emails and links to cloud files raise an exception.

### Deleting an email

```python
mailbox.delete_email(message_id)
```

Returns `True`, or raises an exception if the email couldn't be deleted. Delete emails after processing them, so the next run doesn't pick them up again.

### Sending an email

```python
mailbox.send_email(
    subject="Sales load complete",
    content="Today's sales file was loaded successfully.\n\nRows loaded: 1,204",
    email_addresses=["finance-team@contoso.com", "someone@contoso.com"],
)
```

The email is sent from the mailbox as plain text and saved in its Sent Items. This needs the `Application Mail.Send` role. Use `shared_mailbox_email` to send from a different mailbox.

### Full example

A notebook that loads three daily reports from a shared mailbox, run from a pipeline. Missing reports are skipped rather than failing the run.

```python
# Parameters cell
TENANT_ID = "<tenant-id>"
CLIENT_ID = "<client-id>"
KEY_VAULT = "https://<key-vault-name>.vault.azure.net/"
CERTIFICATE_NAME = "<certificate-name>"
MAILBOX = "reports@contoso.com"
SENDER = "noreply@supplier.com"
LAKEHOUSE = "abfss://<workspace>@onelake.dfs.fabric.microsoft.com/<lakehouse>.Lakehouse"
```

```python
from fabric_python_helper.auth import KeyVaultCertificateCredential
from fabric_python_helper.graph_api import Emails

credential = KeyVaultCertificateCredential(TENANT_ID, CLIENT_ID, KEY_VAULT, CERTIFICATE_NAME)
mailbox = Emails(credential, MAILBOX)

reports = {
    "Daily Change Report": "Files/ServiceDesk/changes.csv",
    "Merged Request Report": "Files/ServiceDesk/merged.csv",
    "Deleted Tickets": "Files/ServiceDesk/deleted.csv",
}

for subject, path in reports.items():
    try:
        message_id = mailbox.search_message_by_subject_and_sender(subject, SENDER)
    except Exception as e:
        if str(e) != "No matching messages found.":
            raise
        print(f"No email for '{subject}', skipping.")
        continue

    attachment_id, file_name = mailbox.get_attachment_ids_and_names(message_id)[0]
    content = mailbox.download_attachment(message_id, attachment_id)
    notebookutils.fs.put(f"{LAKEHOUSE}/{path}", content, overwrite=True)
    mailbox.delete_email(message_id)
    print(f"Loaded '{subject}' ({file_name}).")
```

## Troubleshooting

| Error | Likely cause |
| --- | --- |
| `Failed to get certificate ... Status code: 403, Forbidden` | The identity running the notebook can't read the certificate. Assign Key Vault Reader on the certificate, or the certificate **Get** access policy. |
| `Failed to get certificate ... Status code: 404` | The certificate name or Key Vault URL is wrong. |
| `Failed to sign client assertion ... Forbidden ... keys/sign/action` | The identity running the notebook can't sign with the key. Assign Key Vault Crypto User on the key, or the key **Sign** access policy. |
| `AADSTS700027` (certificate not registered on application) | The certificate version used to sign hasn't been uploaded to the app registration. This usually happens after a renewal, see [Renewing the certificate](#renewing-the-certificate). |
| `AADSTS700016` (application not found in the directory) | The client ID or tenant ID is wrong. |
| `AADSTS7000215` (invalid client secret) | Client secret only: the secret is wrong or has expired. |
| Graph `403` with `ErrorAccessDenied` | The app has no Exchange role for this mailbox. Check `Test-ServicePrincipalAuthorization`, or wait up to 2 hours after a change. |
| Graph `404` for the mailbox | The mailbox address is wrong, or the mailbox doesn't exist. |
| `No matching messages found.` | No email in the Inbox matches the exact subject and sender. Check both, or try `only_search_inbox=False`. |
| `Attachment ... is a #microsoft.graph.itemAttachment` | The attachment is an attached email, not a file. |
| `ModuleNotFoundError: No module named 'notebookutils'` | The code is running outside Fabric. The package only works in Fabric (or Synapse) notebooks. |

## Upgrading from 0.3.x

0.4.0 removes the delegated (signed-in user) authentication, which stored a user's refresh token in Key Vault, and the `pbi_admin` module.

**Emails**

Before:

```python
email_account = gr.Emails(tennant_id, client_id, akv_url, refresh_secret_name)
email_account.connect()
```

After:

```python
from fabric_python_helper.auth import KeyVaultCertificateCredential

credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, certificate_name)
email_account = gr.Emails(credential, shared_mailbox)
```

The other method calls, including `shared_mailbox_email=`, are unchanged. `connect()`, `get_initial_tokens()` and `get_user()` have been removed.

Notebooks that didn't pass `shared_mailbox_email` read the signed-in user's own mailbox. They now need the mailbox passed explicitly, and the app needs access to it in Exchange.

Once every notebook has moved over, delete the refresh token secret from Key Vault and disable **Allow public client flows** on the old app registration (or delete it).

**pbi_admin**

| Removed | Use instead |
| --- | --- |
| `AccessTokens().get_token_as_fabric_notebook_owner()` | `notebookutils.credentials.getToken("pbi")` |
| `AccessTokens().get_token_as_service_principal(...)` | `KeyVaultCertificateCredential(...).get_access_token(Scopes.POWER_BI)`, or `ClientSecretCredential` |
| `SemanticModels(...).refresh_semantic_model()` | The pipeline **Semantic model refresh** activity, `sempy_labs.refresh_semantic_model()`, or `sempy.fabric.refresh_dataset()` |
| `Dataflows(...).refresh_dataflow()` | The pipeline **Dataflow** activity |

**Dependencies:** `msal` is no longer required, and `azure-core` is.

## Building and publishing

### Versioning

The version number is in two places, which must match:

- `setup.py`: `version="0.4.0"`
- `fabric_python_helper/__init__.py`: `__version__="0.4.0"`

Versions follow [semantic versioning](https://semver.org/). While the major version is 0, increase the minor version (0.4.0 to 0.5.0) for breaking changes, and the patch version (0.4.0 to 0.4.1) for fixes. PyPI never accepts the same version twice, even if it's deleted.

### Building a wheel locally

From the repository root, in PowerShell:

```powershell
python -m pip install --upgrade build twine
Remove-Item dist -Recurse -Force -ErrorAction SilentlyContinue
python -m build
python -m twine check dist/*
```

This creates two files in `dist/`:

- `fabric_python_helper-0.4.0-py3-none-any.whl`, the wheel
- `fabric_python_helper-0.4.0.tar.gz`, the source distribution

`twine check` confirms the package metadata and this README will render on PyPI. Clearing `dist/` first stops old versions being uploaded by mistake.

If pip can't verify PyPI's certificate on your network, create `C:\ProgramData\pip\pip.ini` containing:

```ini
[global]
trusted-host = files.pythonhosted.org
               pypi.org
```

### Testing a wheel in Fabric before publishing

Upload the `.whl` file to a lakehouse's `Files` folder, then in a notebook attached to that lakehouse:

```python
%pip install /lakehouse/default/Files/fabric_python_helper-0.4.0-py3-none-any.whl
```

Or add it to an environment under **Custom libraries** > **Upload**.

### Publishing to PyPI with a GitHub release

This is the normal route. The workflow in `.github/workflows/python_publish_package.yml` runs when a release is published. It builds the package, uploads it to PyPI, and attaches the wheel and source distribution to the release.

1. Update the version in both places, commit and push to `main`.
2. On GitHub, go to **Releases** > **Draft a new release**.
3. Create a tag matching the version, for example `v0.4.0`, targeting `main`.
4. Add release notes, then select **Publish release**.
5. Check the workflow run under **Actions**, and that the new version appears on [PyPI](https://pypi.org/project/fabric-python-helper/).

The workflow needs a repository secret called `PYPI_API_TOKEN`, containing a PyPI API token scoped to this project (PyPI > **Account settings** > **API tokens**).

The upload uses `--skip-existing`. If you forget to change the version, the workflow still succeeds but nothing new reaches PyPI. Always check the version on PyPI after releasing.

### Publishing to PyPI manually

After building and checking the wheel as above:

```powershell
python -m twine upload dist/*
```

When prompted, enter `__token__` as the username and your PyPI API token (starting `pypi-`) as the password.

To try a release on [TestPyPI](https://test.pypi.org/) first (this needs a separate TestPyPI account and token):

```powershell
python -m twine upload --repository testpypi dist/*
```

```python
%pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ fabric_python_helper==0.4.0
```

### After publishing

Update the version in each Fabric environment that uses the package (**Public libraries**), and publish the environment.

### Release checklist

1. Update the version in `setup.py` and `fabric_python_helper/__init__.py`.
2. Update this README, including [Upgrading](#upgrading-from-03x) for breaking changes.
3. Build, run `twine check`, and test the wheel in Fabric.
4. Commit and push to `main`.
5. Publish a GitHub release tagged `v<version>`.
6. Confirm the version on PyPI.
7. Update and publish the Fabric environments.

## Licence

BSD 3-Clause. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
