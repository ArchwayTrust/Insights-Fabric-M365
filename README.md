# alt-fabric-m365

Python helpers for automation in Microsoft Fabric notebooks, from Archway Learning Trust (ALT).

Up to 0.5.0 this package was called `fabric_python_helper`. If you're upgrading from a 0.x version, see [Upgrading to 1.0.0](#upgrading-to-100).

The package has a module for each Microsoft 365 service, plus `auth`, which they share:

- **`alt_fabric_m365.auth`** gets app-only access tokens for Microsoft 365 APIs, mainly Microsoft Graph, Exchange and SharePoint, as a service principal, using a certificate held in Azure Key Vault. Key Vault signs each request, so the certificate's private key never leaves the key vault.
- **`alt_fabric_m365.exchange`** finds emails in an Exchange Online mailbox, downloads their attachments, deletes them and sends emails, using Microsoft Graph.
- **`alt_fabric_m365.sharepoint`** manages SharePoint document libraries and their permissions through Entra ID security groups: creating libraries, keeping group membership in sync, navigation links and site sharing settings.

Use a Fabric [workspace identity](#use-a-workspace-identity-where-you-can) wherever Fabric supports one. This package is for Microsoft 365 work it can't do, such as Exchange mailboxes and SharePoint, see [Why Exchange and SharePoint need an app registration](#why-exchange-and-sharepoint-need-an-app-registration).

The package is built for Fabric notebooks. `KeyVaultCertificateCredential` uses `notebookutils` to call Key Vault as the identity running the notebook, so it only works in Fabric (or Synapse) notebooks. The package itself can be imported anywhere.

## Contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [Authentication](#authentication)
  - [Use a workspace identity where you can](#use-a-workspace-identity-where-you-can)
  - [Why Exchange and SharePoint need an app registration](#why-exchange-and-sharepoint-need-an-app-registration)
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
- [SharePoint libraries](#sharepoint-libraries)
  - [What it manages](#what-it-manages)
  - [Giving the app access](#giving-the-app-access)
  - [Limiting the app to chosen sites with Sites.Selected](#limiting-the-app-to-chosen-sites-with-sitesselected)
  - [Creating a SharePointLibraryManager](#creating-a-sharepointlibrarymanager)
  - [Operations](#operations)
  - [The action log](#the-action-log)
  - [Safety features](#safety-features)
  - [Using it in your own setting](#using-it-in-your-own-setting)
  - [Example notebook](#example-notebook)
- [Troubleshooting](#troubleshooting)
- [Upgrading to 1.0.0](#upgrading-to-100)
- [Building and publishing](#building-and-publishing)
- [Licence](#licence)

## Installation

The package is published on PyPI as [`alt-fabric-m365`](https://pypi.org/project/alt-fabric-m365/) and imported as `alt_fabric_m365`. Its dependencies, `azure-core`, `pandas` and `requests`, are already in the Fabric runtime.

**In a Fabric environment** (use this for scheduled notebooks and pipelines):

1. Open the environment your notebooks use.
2. Under **Public libraries**, select **Add from PyPI**, enter `alt-fabric-m365` and choose a version, for example `1.2.0`.
3. Select **Publish**. Notebooks pick up the new version when their next session starts.

**In a single notebook session** (useful for testing a new version):

```python
%pip install alt-fabric-m365==1.2.0
```

Always pin a version. 1.0.0 removes the old names that 0.x versions kept working with a warning, see [Upgrading to 1.0.0](#upgrading-to-100). 1.0.1 only updates the documentation. 1.0.2 stops `update_memberships` emptying an Auto group the app doesn't own, adopts ownerless Auto groups it has no members to add to, and makes dry runs of `create_libraries` report group ownership. 1.1.0 fixes problems found in a review of 1.0.2:

- A site, redirect or Key Vault URL with a backslash or a user name before its host is refused, as it could send a token to another host.
- Taking over an existing library keeps the permissions of its folders and files (`inheritance_broken` is logged), and a library gets its own permissions and groups before its display name is set.
- `retire_moved_libraries` leaves a library whose new site is the same site, spelt differently, and `unlink_libraries` skips a code whose groups are another library's (`group_name_collision`).
- A dry run of `update_memberships` reads each group's members, so it shows removals (`remove_from_group (dry run)`), `removals_skipped` and the guardrail. `reconcile_group (dry run)` is gone.
- A membership row with an unknown permission no longer removes anyone, and `skip` is now in `WARNING_ACTIONS`.
- Options must be of their documented type: a String pipeline parameter such as `"False"` raises `TypeError` rather than counting as true, so give those parameters the type Bool or Int. Blanks in `create_libraries`' flag columns take their defaults.
- `rename_libraries` no longer accepts `old_library_code`, and raises `ValueError` if given one. It never worked (the library was looked up by its new code), so no notebook could rely on it, and it's removed in a minor version.

1.2.0 adds `ensure_onedrive_shortcuts` and a run-wide removal limit, and fixes more problems from the same review:

- `update_memberships` adds up its removals across the whole call. If they'd remove more than `max_removals_per_run` members (default 100), none are made, and one `removal_guardrail_triggered` row with no library says so. If a normal run can remove more, such as at the start of a school year, raise the limit.
- `ensure_onedrive_shortcuts` makes OneDrive shortcuts that `update_memberships` couldn't, for example because the user's OneDrive didn't exist yet.
- A `library_code` SharePoint can't use as a library's URL raises `ValueError` before anything is changed: see the `library_code` row under [Input tables](#input-tables). Such codes never worked after the first run. Trim padded codes in the query that reads them.
- `retire_moved_libraries` only retires a library once its replacement works on the new site (`retire_deferred` until then).
- `create_libraries` logs `library_created` and `library_renamed`, and its dry run reports them and a new navigation label. Dry runs of `rename_libraries`, `set_library_metadata` and `grant_group_access` read what a live run reads, so they report missing libraries and groups, and grants already held.
- A 503 with `Retry-After` is waited out like a 429, and each attempt of a request gets a current token, so a long throttle doesn't end in a 401.
- A group created moments ago isn't created again while Graph's name lookups still miss it, and Graph's 400 for a new group is waited out like its 404.
- Cancelling a cell stops a parallel run starting queued libraries.

From 1.0.0, only a new major version (2.0.0) can break existing notebooks, see [Versioning](#versioning). The history of the 0.x versions is in the [0.7.2 README](https://github.com/ArchwayTrust/Insights-Fabric-M365/blob/v0.7.2/README.md#installation).

## Quick start

Read the latest report email from a mailbox, save its CSV attachment to a lakehouse and delete the email:

```python
from alt_fabric_m365.auth import KeyVaultCertificateCredential
from alt_fabric_m365.exchange import Emails

credential = KeyVaultCertificateCredential(
    tenant_id="<tenant-id>",
    client_id="<app-registration-client-id>",
    akv_url="https://<key-vault-name>.vault.azure.net/",
    certificate_name="<certificate-name>",
)

mailbox = Emails(credential, "reports@contoso.com")

message_id = mailbox.search_message_by_subject_and_sender("Daily Sales Report", "noreply@supplier.com")

csv_files = [(attachment_id, file_name)
             for attachment_id, file_name in mailbox.get_attachment_ids_and_names(message_id)
             if file_name.lower().endswith(".csv")]
if not csv_files:
    raise RuntimeError("The latest Daily Sales Report email has no CSV attachment.")

attachment_id, file_name = csv_files[0]
csv_text = mailbox.download_attachment(message_id, attachment_id)

lakehouse = "abfss://<workspace>@onelake.dfs.fabric.microsoft.com/<lakehouse>.Lakehouse"
notebookutils.fs.put(f"{lakehouse}/Files/Reports/{file_name}", csv_text, overwrite=True)

mailbox.delete_email(message_id)
```

Before this works you need to [set up a certificate](#setting-up-a-certificate) and [give the app access to the mailbox](#giving-the-app-access-to-a-mailbox).

## Authentication

### Use a workspace identity where you can

A [workspace identity](https://learn.microsoft.com/en-us/fabric/security/workspace-identity) is a service principal that Fabric creates for a workspace and manages for you. Fabric holds its credentials and rotates them itself, so there's no secret or certificate to store, renew or leak. When something in Fabric can authenticate as the workspace identity, use it rather than an app registration of your own, including the one this package uses.

A workspace identity can be used for:

| Use | How |
| --- | --- |
| OneLake shortcuts to Azure Data Lake Storage Gen2 | Choose **Workspace identity** as the shortcut's authentication method, and give the identity a role such as Storage Blob Data Reader on the storage account. |
| Firewall-enabled storage accounts | [Trusted workspace access](https://learn.microsoft.com/en-us/fabric/security/security-trusted-workspace-access) lets shortcuts, pipelines and semantic models reach storage that only allows selected networks. Needs a purchased Fabric capacity (F SKU). |
| Pipelines | Connections for the Copy, Lookup and GetMetadata activities, for sources such as ADLS Gen2, Blob Storage and Azure SQL. |
| Semantic models (import mode) | A cloud connection using the workspace identity, chosen in the model's **Gateway and cloud connections** settings. |
| Dataflows Gen2 (CI/CD) | The workspace identity as a connector's authentication method. |
| Reading SharePoint Online lists | The SharePoint Online list connector in Dataflows Gen2, pipelines (Copy and Lookup) and copy jobs. Read only. |
| Notebooks run from a pipeline | Set the Notebook activity's connection to the workspace identity. The notebook then runs as it, so `notebookutils.credentials.getToken()` returns its tokens for OneLake and Azure Storage (`storage`), Power BI and Fabric (`pbi`), Key Vault (`keyvault`) and Kusto (`kusto`), and `notebookutils.credentials.getSecret()` reads Key Vault as it. |
| Azure resources that accept Microsoft Entra ID | Give the identity a role on the resource, as for any service principal: Key Vault, storage accounts, Azure SQL and so on. Its service principal has the same name as the workspace. |

The [list of connectors that support it](https://learn.microsoft.com/en-us/fabric/data-factory/connector-overview) grows, so check it before setting up an app registration.

Things to know before relying on one:

- Only a workspace admin can create it, and not in a **My workspace**. Deleting the identity, or permanently deleting the workspace, breaks everything that uses it, and a deleted identity can't be restored.
- Admins, members and contributors of the workspace can use it in connections, and the identity that runs a pipeline using it must hold one of those roles. Treat those roles as access to everything the workspace identity can reach.
- Don't change its app registration or enterprise application in Entra ID, other than adding API permissions; Fabric may revert other changes, and they can stop it working.
- If a Conditional Access policy for workload identities covers all service principals, exclude workspace identities from it, or they won't work.
- It can't be used across tenants or with on-premises data gateway connections.

### Why Exchange and SharePoint need an app registration

This package's `Emails` and `SharePointLibraryManager` call Microsoft Graph and the SharePoint REST API. Those calls can't be made as a workspace identity, so they authenticate as an app registration you create, using `KeyVaultCertificateCredential`:

- **Your code can't get a Graph or SharePoint token as a workspace identity.** Fabric only uses the identity's tokens inside the features in the table above, and `notebookutils.credentials.getToken()` only supports the `storage`, `pbi`, `keyvault` and `kusto` audiences. Fabric keeps the identity's credentials, so you can't request a token for `https://graph.microsoft.com` or `https://<tenant>.sharepoint.com` with them yourself. Granting it Graph permissions or an Exchange role wouldn't help, as nothing can use them.
- **Fabric's connectors only cover reading data.** The Exchange Online connector only signs in as a user. The SharePoint Online list connector does accept the workspace identity, so use it if you only need to copy list data into Fabric. Neither can read and delete emails, send them, create libraries, manage groups or change permissions, which is what this package does.
- **SharePoint only accepts app-only tokens obtained with a certificate.** That needs a certificate registered on the app, which you can't add to a workspace identity, and client secrets are rejected.
- **The permissions stay separate from the workspace identity's.** Mailbox and SharePoint access are powerful, so it's better held by an app registration that does only this job, where it can be reviewed and revoked on its own. How far it reaches depends on what you grant:
  - Mail access is limited to the mailboxes in its Exchange role assignment (see [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox)).
  - SharePoint access is limited to chosen sites only with `Sites.Selected` (see [Limiting the app to chosen sites with Sites.Selected](#limiting-the-app-to-chosen-sites-with-sitesselected)). `Sites.FullControl.All` covers every site.
  - `SharePointLibraryManager` also needs `Group.ReadWrite.All`, `GroupMember.ReadWrite.All` and `User.Read.All`, which apply to the whole tenant (see [Giving the app access](#giving-the-app-access)).

  **Sign** on the key in Key Vault controls who can use the certificate this package signs with. It doesn't cover any other certificate or client secret on the app registration, which can also authenticate as the app, so remove those you don't need.

The two work together. Run the notebook as the workspace identity, and give the workspace identity permission to sign with the certificate in Key Vault. Key Vault is one of the audiences the workspace identity can get tokens for, so `KeyVaultCertificateCredential` uses it to sign as the app, and the app is what Exchange and SharePoint see. The workspace identity still does everything else in the notebook, such as writing to OneLake.

Anyone who can run code as the workspace identity can then act as the app too, so keep notebooks that use the app in a workspace whose admins, members and contributors you'd trust with its mailboxes and sites.

| Task | Authenticate as |
| --- | --- |
| Read or write OneLake, a lakehouse or Azure Storage | The workspace identity |
| Call the Power BI or Fabric REST APIs, or Kusto | The workspace identity |
| Copy SharePoint Online list data into Fabric | The workspace identity, through the SharePoint Online list connector |
| Read secrets or sign with keys in Key Vault | The workspace identity |
| Read, delete or send email (`Emails`) | An app registration, with `KeyVaultCertificateCredential` |
| Manage SharePoint libraries and their groups (`SharePointLibraryManager`) | An app registration, with `KeyVaultCertificateCredential` (a certificate is required) |
| Call any other Microsoft Graph or SharePoint API | An app registration, with `KeyVaultCertificateCredential` |

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

You can use ```code policy.json```

Then create the certificate:

```bash
az keyvault certificate create --vault-name <key-vault-name> --name <certificate-name> --policy "@policy.json"
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

See the [Scopes](#scopes) table, [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox) for email, and [Limiting the app to chosen sites with Sites.Selected](#limiting-the-app-to-chosen-sites-with-sitesselected) for SharePoint.

### Using the credential

```python
from alt_fabric_m365.auth import KeyVaultCertificateCredential, Scopes

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
| `akv_url` | Yes | Key Vault URL, `https://<key-vault-name>.vault.azure.net/`. Anything that isn't an `https://` Key Vault URL raises `ValueError`. Only the Azure public cloud is supported, so sovereign cloud vaults (such as `vault.azure.cn` or `vault.usgovcloudapi.net`) are rejected. |
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

If a client passes `claims` to `get_token` (for example after a Continuous Access Evaluation challenge), a new token is requested rather than the cached one returned. Asking for a token in a different `tenant_id` from the credential's raises `ValueError`.

### Scopes

If you've used Microsoft Graph as a signed-in user, you'll know scopes as permission names such as `User.Read.All` or `Mail.Read`. This package uses scopes differently. Here a scope picks **which API** a token is for. **What** the app can do in that API is set up separately, by an administrator.

#### What a scope is

Every scope has two parts: the API's resource URI, and a permission on that API.

```text
https://graph.microsoft.com/User.Read.All
└──────── resource ───────┘└─permission─┘
```

When you ask for `User.Read.All` on its own, Entra ID reads it as `https://graph.microsoft.com/User.Read.All`. Graph is the default resource when the prefix is left out, which is why the resource part is easy to miss.

#### Why every scope here ends in `/.default`

There are two ways to get a token:

- **Delegated.** A user signs in and the app acts on their behalf. The app asks for the permissions it wants (`User.Read.All Mail.Read`), the user or an administrator consents, and the token carries them in its `scp` claim. The app can never do more than the signed-in user could.
- **App-only** (the client credentials flow). There's no user. The app signs in as itself, here with a certificate, as its service principal.

`KeyVaultCertificateCredential` only does app-only. In that flow you can't ask for individual permissions, so the permission part of the scope must be the literal `.default`:

```text
https://graph.microsoft.com/.default
```

`.default` means "every application permission an administrator has already granted this app on this API". So the scope only says which API the token is for. That becomes the token's audience (its `aud` claim), and the API rejects any token whose audience isn't itself. A Graph token can't be used for SharePoint, and a SharePoint token can't be used for Graph.

| | Delegated | App-only (this package) |
| --- | --- | --- |
| Signs in as | A user | The app's service principal |
| You request | Permission names, such as `User.Read.All Mail.Read` | `{resource}/.default`, such as `https://graph.microsoft.com/.default` |
| What it can do is decided by | The scopes requested, consent, and the user's own access | What an administrator granted the app ahead of time |
| Permissions appear in the token as | The `scp` claim | The `roles` claim |

#### Where `User.Read.All` goes instead

Permissions move from request time to setup time. If the app needs `User.Read.All`, an administrator adds it to the app registration as an **application** permission (**API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**) and grants admin consent. Your code doesn't change: it still asks for `Scopes.GRAPH`, and the token that comes back lists `User.Read.All` in its `roles` claim.

Permissions on the app registration only apply to APIs that use them, mainly Microsoft Graph and SharePoint. Many of the APIs below ignore them, and access comes from somewhere else:

- **Role-based APIs** (Power BI, Fabric, OneLake, Azure Storage, Key Vault, SQL, Kusto, Azure Resource Manager, Log Analytics). Give the service principal a role on the thing it needs: a Fabric workspace role, an Azure role assignment, or a database user. The token's `roles` claim is empty, and the API looks up the service principal's roles itself on each request.
- **Exchange mailboxes.** Although the calls go through Graph, give the app an Exchange role limited to particular mailboxes, not Graph `Mail.*` permissions. See [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox).
- **SharePoint sites with `Sites.Selected`.** The permission on the app registration grants nothing on its own. Each site must also be granted to the app. See [Limiting the app to chosen sites with Sites.Selected](#limiting-the-app-to-chosen-sites-with-sitesselected).

In every case, **Entra ID issues a token whether or not the app has access.** A token just says "this is app X, for API Y". If access isn't set up, the API rejects the request with 401 or 403, so getting a token proves the certificate works, not that permissions are right. If you get a 401 or 403, check the last column of the table below.

#### Checking what a token contains

To see which API a token is for and which permissions it carries, decode its payload. This only reads the token; it doesn't check the signature:

```python
import base64
import json

def token_claims(token):
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))

claims = token_claims(credential.get_access_token(Scopes.GRAPH))
print(claims["aud"])            # https://graph.microsoft.com
print(claims.get("roles", []))  # for example ['Sites.Selected', 'Group.ReadWrite.All', 'User.Read.All']
```

A permission missing from `roles` hasn't been granted, or hasn't had admin consent. After granting one, new tokens pick it up. Tokens already issued keep their old `roles` until they expire, usually within an hour, so start a new notebook session or create a new credential to get a token with the change straight away.

#### Common scopes

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
| (none) | `https://<tenant>.sharepoint.com/.default` | SharePoint REST API | SharePoint application permissions, ideally [`Sites.Selected`](#limiting-the-app-to-chosen-sites-with-sitesselected) with each site granted to the app. SharePoint only allows app-only access with a certificate; client secrets are rejected. |
| (none) | `https://<org>.crm11.dynamics.com/.default` | Dataverse | Use your environment URL. Add the app as an application user in the Power Platform admin center. |

One credential can get tokens for any number of these APIs. Each token is cached per scope, so a notebook that calls Graph and SharePoint gets one token for each and reuses them until they're about to expire.

The `scope` you pass when creating the credential is only a default for `get_access_token()`. `get_token()` always takes the scope as an argument, which is how semantic link, Azure SDK clients and `SharePointLibraryManager` ask for the tokens they need.

#### Other things called scope

This README uses "scope" in two other senses that have nothing to do with tokens:

- **Azure role assignment scope**, as in `az role assignment create --scope ...` in [Setting up a certificate](#setting-up-a-certificate). This is the resource a role applies to, such as one key vault or one certificate.
- **Exchange management scope**, as in `New-ManagementScope` in [Giving the app access to a mailbox](#giving-the-app-access-to-a-mailbox). This is the set of mailboxes an Exchange role assignment covers.

Both limit where a permission applies. An OAuth scope only picks the API a token is for.

**When you don't need this package:** `notebookutils.credentials.getToken()` already returns tokens for the identity running the notebook, for the `storage`, `pbi`, `keyvault` and `kusto` audiences. For those APIs, prefer the workspace identity, see [Use a workspace identity where you can](#use-a-workspace-identity-where-you-can). Use this package when you need another API (such as Graph or SharePoint) or need to act as a separate app.

### Using a client secret instead

Certificates are recommended, but `Emails` and anything else in this package that takes a credential, except `SharePointLibraryManager`, also accepts `azure.identity.ClientSecretCredential`, which is already in the Fabric runtime:

```python
from azure.identity import ClientSecretCredential
from alt_fabric_m365.exchange import Emails

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

`alt_fabric_m365.exchange.Emails` works with mailboxes using Microsoft Graph, authenticating as the app rather than as a user.

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

# 3. Assign the role, limited to that scope. Use "Application Mail.Read" instead if the app never deletes emails.
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
from alt_fabric_m365.exchange import Emails

mailbox = Emails(credential, "reports@contoso.com")
```

| Parameter | Description |
| --- | --- |
| `credential` | A `KeyVaultCertificateCredential`, a `ClientSecretCredential`, or any other credential with a `get_token` method. |
| `mailbox` | Optional. Email address of the default mailbox to use. |

A token is requested from the credential for each call, and the credential caches it, so one instance can be used for the whole notebook.

Requests Graph throttles (`429`, or `503` with `Retry-After`), server errors and dropped connections are retried, honouring `Retry-After`, as the SharePoint manager's are. Sending an email is only retried when Graph certainly didn't send it, as a retry after a server error or timeout could send it twice.

Every method also takes an optional `shared_mailbox_email` argument to work with a different mailbox for that call. The app needs access to that mailbox too. If you always pass `shared_mailbox_email`, you can leave out `mailbox`, as in `Emails(credential)`. A call with neither raises `ValueError`.

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

If nothing matches it raises `MessageNotFoundError`, with the message `No matching messages found.` Use this to handle emails that don't always arrive:

```python
from alt_fabric_m365.exchange import MessageNotFoundError

try:
    message_id = mailbox.search_message_by_subject_and_sender("Weekly Exceptions", "noreply@supplier.com")
except MessageNotFoundError:
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

`get_attachment_ids_and_names` returns a list of `(attachment_id, file_name)` tuples, one for each file attached to the email:

```python
attachments = mailbox.get_attachment_ids_and_names(message_id)
# [('AAMkAD...', 'sales.csv'), ('AAMkAD...', 'summary.xlsx')]
```

Inline images, such as logos in an email signature, are left out unless you pass `include_inline=True`. Attached emails and links to cloud files are always left out, as they can't be downloaded. An empty list means the email has nothing to download, so check for it rather than taking the first item.

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

Only file attachments can be downloaded. Passing the ID of an attached email or a link to a cloud file (from somewhere other than `get_attachment_ids_and_names`) raises an exception.

### Deleting an email

```python
mailbox.delete_email(message_id)
```

Returns `True`, also when the email is already gone, or raises an exception if it couldn't be deleted. Delete emails after processing them, so the next run doesn't pick them up again.

### Sending an email

```python
mailbox.send_email(
    subject="Sales load complete",
    content="Today's sales file was loaded successfully.\n\nRows loaded: 1,204",
    email_addresses=["finance-team@contoso.com", "someone@contoso.com"],
)
```

The email is sent from the mailbox as plain text and saved in its Sent Items. `email_addresses` can also be one address as a string. This needs the `Application Mail.Send` role. Use `shared_mailbox_email` to send from a different mailbox.

### Full example

A notebook that loads three daily reports from a shared mailbox, run from a pipeline. Missing reports are skipped rather than failing the run, and an email without a CSV attachment is left in the mailbox for someone to look at.

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
from alt_fabric_m365.auth import KeyVaultCertificateCredential
from alt_fabric_m365.exchange import Emails, MessageNotFoundError

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
    except MessageNotFoundError:
        print(f"No email for '{subject}', skipping.")
        continue

    csv_files = [(attachment_id, file_name)
                 for attachment_id, file_name in mailbox.get_attachment_ids_and_names(message_id)
                 if file_name.lower().endswith(".csv")]
    if not csv_files:
        print(f"No CSV attachment on '{subject}', skipping.")
        continue

    attachment_id, file_name = csv_files[0]
    content = mailbox.download_attachment(message_id, attachment_id)
    notebookutils.fs.put(f"{LAKEHOUSE}/{path}", content, overwrite=True)
    mailbox.delete_email(message_id)
    print(f"Loaded '{subject}' ({file_name}).")
```

## SharePoint libraries

`alt_fabric_m365.sharepoint.SharePointLibraryManager` manages SharePoint document libraries and their permissions through Entra ID security groups. Each operation takes a pandas DataFrame describing the libraries or memberships you want, makes SharePoint match it, and returns an [action log](#the-action-log) of what it did.

### What it manages

For each library (a row with a `site_url`, a `library_code` and a `group_prefix`):

- **The library.** Created with the code as its URL slug (`/sites/<site>/<code>`), then given its display name. The slug never changes, so renaming a library only changes its display name. It is given its own permissions, not inherited from the site.
  - A document library that already has the code as its URL slug is taken over rather than created again. If it still inherits the site's permissions, it stops doing so, and the action log has an `inheritance_broken` row: site members then only have access through its groups. Folders and files in it that have their own permissions keep them.
- **Security groups**, named `{group_prefix}-{library_code}-{suffix}` (or `{group_prefix}-{site}-{library_code}-{suffix}` with `include_site_name=True`). Names are how the manager finds existing groups, so don't change the naming after rollout.
  - `Auto-Read`, `Auto-Contribute` and `Auto-Owner` are owned by the automation. Their membership is synced from your data and they are emptied when the library is unlinked.
  - `Manual-Read` and `Manual-Contribute` are optional, hand-curated groups. They are granted access like the Auto groups, but their membership is never changed.
  - Group names use the library **code**, so they stay the same when the library's display name changes. `create_libraries` updates the groups' descriptions, which name the library, to match. A library's code never changes, as it's also the library's URL, so neither do its group names.
  - The app registration is made **owner** of every group it creates, and an existing group is only granted access or synced if the app owns it (see [Safety features](#safety-features)). Pass `app_owned=False` for groups people own, such as a leadership team's groups owned by the headteacher.
- **Navigation** (optional). A link to the library in the site's quick launch, under a heading such as "Document Libraries".
- **Provenance** (optional). `source_system` and `source_id` stored in the library's property bag, invisible to users, so a library can be traced back to the record that created it. A whole number is stored without decimal places, even from a column pandas made float because it has a null (`42`, not `42.0`).
- **Site sharing** (optional). Turns off "Allow members to share", "Allow access requests" and "Allow members to invite others" on each site, so access only comes from the groups.

### Giving the app access

The app registration needs these application permissions, with admin consent:

| API | Permission | Used for |
| --- | --- | --- |
| Microsoft Graph | `Sites.Selected` with full control on each site (recommended), or `Sites.FullControl.All` | Finding sites, creating and renaming libraries |
| Microsoft Graph | `Group.ReadWrite.All` and `GroupMember.ReadWrite.All` | Creating and renaming groups, syncing membership |
| Microsoft Graph | `User.Read.All` | Resolving users by UPN, mail or proxy address |
| Microsoft Graph | `Files.ReadWrite.All` | Only for OneDrive shortcuts (`create_onedrive_shortcuts=True`, or `ensure_onedrive_shortcuts`) |
| SharePoint | `Sites.Selected` with full control on each site (recommended), or `Sites.FullControl.All` | Permissions, property bag, navigation and sharing settings |

`Sites.FullControl.All` gives the app full control of every site in the tenant, including every user's OneDrive. Use `Sites.Selected` instead, so the app can only touch the sites it manages, see [Limiting the app to chosen sites with Sites.Selected](#limiting-the-app-to-chosen-sites-with-sitesselected). The group and user permissions have no per-site equivalent and always apply to the whole tenant. So does `Files.ReadWrite.All`, which reaches files on every site as well as every OneDrive, so turning on OneDrive shortcuts gives up per-site isolation for files.

SharePoint only accepts app-only tokens obtained with a certificate, so use `KeyVaultCertificateCredential`, not a client secret. Set up the certificate as in [Setting up a certificate](#setting-up-a-certificate), uploading it to this app registration. The identity running the notebook (for a pipeline, whatever the notebook activity runs as) needs **Get** on the certificate and **Sign** on its key.

### Limiting the app to chosen sites with Sites.Selected

`Sites.Selected` is an application permission that grants nothing by itself. Admin consent for it lets SharePoint grant the app access one site at a time, and the app can only reach the sites it has been granted. Any other site returns 403, as if the app had no SharePoint permissions at all.

It works in three layers, and the app only has access to a site when all three are in place:

1. **Consent.** `Sites.Selected` is on the app registration, for each API the app calls, with admin consent. `SharePointLibraryManager` calls both Microsoft Graph and the SharePoint REST API, so it needs it on both.
2. **A grant on the site.** The app is given a role on each site, such as `fullcontrol`. The grant is stored on the site in SharePoint, not on the app registration, so you won't see it in Entra ID. One grant covers both Graph and the SharePoint REST API.
3. **The token.** The app's token lists `Sites.Selected` in its `roles` claim, which it does once step 1 is done, see [Checking what a token contains](#checking-what-a-token-contains).

This gives you two ways to remove access: remove one site's grant to block that site, or revoke the app's admin consent for `Sites.Selected` to block every site, see [Changing or removing a grant](#changing-or-removing-a-grant).

#### Roles

| Role | Allows |
| --- | --- |
| `read` | Reading lists, libraries, items and files |
| `write` | `read`, plus creating, changing and deleting items and files |
| `owner` | `write`, plus managing lists and libraries. PnP PowerShell calls this **Manage**. |
| `fullcontrol` | Everything, including permissions and site settings |

`SharePointLibraryManager` needs `fullcontrol`. It gives libraries their own permissions, grants groups access to them, and changes the site's sharing settings and members group, which the lower roles can't do. For your own code, grant the lowest role that works, for example `read` to read a list or `write` to upload files.

#### 1. Add the permission to the app registration

In the Microsoft Entra admin center, open the app registration and go to **API permissions**:

1. **Add a permission** > **Microsoft Graph** > **Application permissions**, select **Sites.Selected** and add it.
2. **Add a permission** > **SharePoint** > **Application permissions**, select **Sites.Selected** and add it.
3. Select **Grant admin consent for <your organisation>**. This needs a Privileged Role Administrator or Global Administrator.
4. If the app has `Sites.FullControl.All`, `Sites.ReadWrite.All`, `Sites.Read.All` or `Sites.Manage.All` under either API, revoke them. Permissions add up, so leaving one granted would keep the app's access to every site. Removing a permission from the configured list doesn't revoke its admin consent: the app keeps it, and the page then lists it under **Other permissions granted for <your organisation>**. For each one, select **...** > **Revoke admin consent** (for one still in the configured list, revoke it and then remove it). Check under **Enterprise applications** > the app > **Permissions** that none of them are left.

Then get a fresh token and check its `roles` claim lists `Sites.Selected` and none of the broader permissions, see [Checking what a token contains](#checking-what-a-token-contains). Tokens issued before the change keep their old permissions until they expire, usually within an hour.

With the Azure CLI. `00000003-0000-0000-c000-000000000000` is Microsoft Graph and `00000003-0000-0ff1-ce00-000000000000` is SharePoint, and each permission's ID is looked up by name:

```bash
for api in 00000003-0000-0000-c000-000000000000 00000003-0000-0ff1-ce00-000000000000; do
    role_id=$(az ad sp show --id $api --query "appRoles[?value=='Sites.Selected'].id" --output tsv)
    az ad app permission add --id <client-id> --api $api --api-permissions $role_id=Role
done

az ad app permission admin-consent --id <client-id>
```

#### 2. Grant the app each site

Granting a site needs `Sites.FullControl.All`, as a site grant can include full control. You don't need to give that to this app: a SharePoint Administrator does this once per site from their own session, using the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation) (the `Microsoft.Graph.Sites` module).

```powershell
Connect-MgGraph -Scopes "Sites.FullControl.All"

$appId   = "<client-id>"
$appName = "<app-name>"   # only a label stored with the grant; use the app registration's name

$sites = @(
    "contoso.sharepoint.com:/sites/hr",
    "contoso.sharepoint.com:/sites/finance"
)

foreach ($path in $sites) {
    $site = Get-MgSite -SiteId $path
    New-MgSitePermission -SiteId $site.Id -BodyParameter @{
        roles               = @("fullcontrol")
        grantedToIdentities = @(@{ application = @{ id = $appId; displayName = $appName } })
    }
}
```

Each site is given as `<hostname>:/sites/<name>`, the site URL with `:` after the hostname. `Connect-MgGraph -Scopes` is the delegated kind of scope from [Scopes](#scopes): the administrator signs in as themselves, and the session can only do what both the scope and their own role allow. The first time anyone uses it, Entra ID asks for consent to the Microsoft Graph Command Line Tools app having `Sites.FullControl.All`, which an administrator who can grant tenant-wide consent has to accept. Microsoft also says the person granting must have access to the site. If `New-MgSitePermission` returns 403, add yourself to the site as an owner, or as a site collection administrator in the SharePoint admin center, and try again.

If you use [PnP PowerShell](https://pnp.github.io/powershell/), this does the same:

```powershell
Grant-PnPEntraIDAppSitePermission -AppId <client-id> -DisplayName "<app-name>" -Permissions FullControl `
    -Site https://contoso.sharepoint.com/sites/hr
```

#### 3. Check the grant

List the apps granted a site, and their roles. Listing a site's permissions doesn't return their roles, so each one is fetched by ID:

```powershell
$site = Get-MgSite -SiteId "contoso.sharepoint.com:/sites/hr"

Get-MgSitePermission -SiteId $site.Id | ForEach-Object {
    $permission = Get-MgSitePermission -SiteId $site.Id -PermissionId $_.Id
    [pscustomobject]@{
        App   = $permission.GrantedToIdentitiesV2.Application.DisplayName -join ", "
        Roles = $permission.Roles -join ", "
        Id    = $permission.Id
    }
}
```

Then check from a notebook, as the app. A granted site returns 200, and any other site 403:

```python
import requests
from alt_fabric_m365.auth import Scopes

token = credential.get_access_token(Scopes.GRAPH)
for site in ["contoso.sharepoint.com:/sites/hr", "contoso.sharepoint.com:/sites/not-granted"]:
    response = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site}",
                            headers={"Authorization": f"Bearer {token}"})
    print(site, response.status_code)
```

A new grant usually works within a few minutes. Finally, run the manager with `dry_run=True`: sites the app can't read show up as `error` rows.

#### Changing or removing a grant

Use the `Id` from the check above:

```powershell
# Change the role
Update-MgSitePermission -SiteId $site.Id -PermissionId <id> -Roles @("write")

# Remove the app's access to the site
Remove-MgSitePermission -SiteId $site.Id -PermissionId <id>
```

To block every site, revoke the admin consent for `Sites.Selected` on both APIs, as in step 1. Removing it from the configured list on its own isn't enough. This isn't immediate: tokens the app already holds keep working until they expire, usually within an hour. Removing a site's grant is the quicker way to cut off one site.

#### Things to know

- **Grant a new site before the manager uses it.** When your data adds a library on a site the app hasn't been granted, that site gets one `error` row and its libraries are skipped. The rest of the run carries on. In the [recommended flow](#recommended-flow) the watermark isn't advanced, so once you've granted the site the next run picks them up.
- **Keep the old site's grant until retiring is done.** `retire_moved_libraries` can't retire a library on a site the app can no longer read: a 403 from the old site, or from the new one, is an `error` row, and the old library is left as it is. Only remove the grant after the old libraries have been retired. An old site that no longer exists (404) is `retire_skipped`, as there's nothing left to retire; a new site that doesn't exist yet defers the retire (`retire_deferred`), so the old library keeps its access.
- **A grant covers the whole site collection**, including any subsites, and doesn't change who else can use the site. It doesn't break permission inheritance on anything.
- **The app can't find sites it hasn't been granted.** It can only work with the site URLs you give it, which is all `SharePointLibraryManager` needs.
- **Other permissions still reach the whole tenant.** `Group.ReadWrite.All`, `GroupMember.ReadWrite.All` and `User.Read.All` aren't limited by `Sites.Selected`. Nor is `Files.ReadWrite.All`: as an application permission it can read and write files in every document library on every site, as well as every user's OneDrive. Adding it for `create_onedrive_shortcuts=True` therefore gives the app file access to sites it hasn't been granted, so only add it if you use shortcuts and accept that.
- **Lists, libraries and files can be granted too**, with `Lists.SelectedOperations.Selected`, `ListItems.SelectedOperations.Selected` and `Files.SelectedOperations.Selected`. `SharePointLibraryManager` creates libraries, so it needs a grant on the site. See Microsoft's [overview of Selected permissions](https://learn.microsoft.com/en-us/graph/permissions-selected-overview).

### Creating a SharePointLibraryManager

```python
from alt_fabric_m365.auth import KeyVaultCertificateCredential
from alt_fabric_m365.sharepoint import SharePointLibraryManager

credential = KeyVaultCertificateCredential(
    "<tenant-id>", "<client-id>", "https://<key-vault-name>.vault.azure.net/", "<certificate-name>",
)

mgr = SharePointLibraryManager(
    credential,
    sharepoint_hostname="contoso.sharepoint.com",
    dry_run=True,                   # log what would change, change nothing
    max_workers=4,                  # libraries processed in parallel (1: everything in turn)
    include_site_name=False,        # see group naming above; don't change after rollout
    add_to_navigation=True,
    navigation_label="Document Libraries",
    max_removals_per_group=10,      # see Safety features
    max_removals_per_run=100,
)
```

Create the credential without a scope: the manager asks it for a Graph token and a SharePoint token (`https://<sharepoint_hostname>/.default`), and the credential caches each until shortly before it expires. The constructor gets both tokens straight away, so wrong credentials or permissions fail here rather than part way through. `dry_run`, `max_workers`, `include_site_name`, `add_to_navigation`, `navigation_label`, `create_onedrive_shortcuts`, `max_removals_per_group`, `max_removals_per_run` and `allow_mass_removal` are defaults that each operation can override, for example `mgr.create_libraries(df, dry_run=False)`. Each must be of its documented type: a notebook or pipeline parameter of type String, such as `"False"`, raises `TypeError` rather than counting as true, so give those parameters the type Bool or Int.

### Operations

| Method | What it does | DataFrame columns |
| --- | --- | --- |
| `create_libraries(df)` | Ensures each library exists with its own permissions, its groups, metadata and navigation link. Safe to re-run: existing libraries and groups are reused, and a changed display name is applied. | `site_url`, `library_code`, `library_display_name`, `group_prefix`. Optional: `auto_groups` (default True), `manual_read`, `manual_contribute`, `app_owned` (default True), `source_system`, `source_id` |
| `update_memberships(df, libraries_df=None, nested_groups_df=None)` | Makes each library's Auto groups contain exactly the users (and nested groups) given. Users are matched by UPN, then mail, then proxy addresses, so aliased domains work. Libraries in `libraries_df` with no rows in `df` have their Auto groups emptied. | `site_url`, `library_code`, `user_principal_name`, `permission` (`read`, `contribute` or `owner`), `group_prefix` |
| `unlink_libraries(df)` | For libraries removed at source: renames to `{code}_UNLINKED`, empties the Auto groups, keeps the Manual groups, removes the navigation link. Nothing is deleted. A library already unlinked isn't renamed again (`already_unlinked`). A code whose group names are another library's (`group_name_collision`, see [Safety features](#safety-features)) is skipped, so its groups are never emptied. | `site_url`, `library_code`, `group_prefix` |
| `retire_moved_libraries(df)` | For a library whose site changed: renames the copy on the **old** site to `{code}_MOVED` and removes the groups' access to it. With `new_site_url`, waits until the library is working there (`retire_deferred` until then). A library already retired isn't renamed again (`already_retired`), and only groups that still have access are touched. | `site_url` (the old site), `library_code`, `group_prefix`. Optional: `new_site_url` |
| `rename_libraries(df)` | Sets new display names, and the navigation links' titles. A library that already has its name is left alone (`already_named`). Groups keep their names, as the code can't change. Up to 1.0.x it accepted `old_library_code` to rename the groups to a new code, but that never worked, so a row with one now raises `ValueError`. | `site_url`, `library_code`, `library_display_name` |
| `grant_group_access(df)` | Grants existing groups a permission level on existing libraries, for example a leadership group on every library. Additive: never removes access. | `site_url`, `library_code`, `group_name`, `permission` |
| `enforce_site_sharing_lockdown(site_urls)` | Turns off member sharing, access requests and member invites on each site, and checks the change stuck. A site listed more than once, even spelt differently, is checked once. | A list of site URLs (for one site, `[site_url]`) |
| `set_library_metadata(df)` / `get_library_metadata(df)` | Writes or reads `source_system` and `source_id` in each library's property bag. | `site_url`, `library_code`, and for set `source_system` and/or `source_id` |
| `ensure_onedrive_shortcuts(df)` | Not part of the usual flow. Creates any missing OneDrive shortcuts to libraries, such as ones that failed because the user's OneDrive didn't exist yet (`onedrive_shortcut_failed`), which `update_memberships` never retries. A shortcut to the library already there, even renamed, is left alone (`onedrive_shortcut_exists`). A shortcut to a folder in the library doesn't count. | `site_url`, `library_code`, `user_principal_name`: the members table works as it is |

Read-only helpers: `mgr.group_name(group_prefix, site_url, library_code, suffix)` returns a group name exactly as the operations build it, and `mgr.get_navigation_link(site_url, library_code)` and `mgr.get_site_sharing_settings(site_url)` read back what is on the site.

```python
import pandas as pd

site = "https://contoso.sharepoint.com/sites/Academy"

libraries = pd.DataFrame({
    "site_url": [site, site],
    "library_code": ["MATH", "ENG"],
    "library_display_name": ["Maths", "English"],
    "group_prefix": "ACA-Sharepoint",
    "manual_read": True,
})
results = mgr.create_libraries(libraries)

members = pd.DataFrame({
    "site_url": site,
    "library_code": ["MATH", "MATH", "ENG"],
    "user_principal_name": ["jo@contoso.com", "sam@contoso.com", "alex@contoso.com"],
    "permission": "contribute",
    "group_prefix": "ACA-Sharepoint",
})
results = mgr.update_memberships(members, libraries_df=libraries)
```

### The action log

Every operation returns a DataFrame with the columns `site_url`, `document_library`, `group`, `user`, `action`, `permission` and `status`, one row per thing it did or (in a dry run) would do. Problems are rows too, so one bad library, user or site never stops the others. Where nothing needed changing, the row says so (`already_member`, `already_assigned`, `metadata_unchanged`), and nothing is written. Two sets of action names help decide whether a run succeeded:

- `FATAL_ACTIONS`: `error`, `add_failed`, `remove_failed`. Something failed and should be retried, so fail the run.
- `WARNING_ACTIONS`: for example `user_not_found`, `library_not_found` and `removal_guardrail_triggered`. Problems with the data that failing the run wouldn't fix.

```python
from alt_fabric_m365.sharepoint import FATAL_ACTIONS, WARNING_ACTIONS

warnings = results[results["action"].isin(WARNING_ACTIONS)]
failures = results[results["action"].isin(FATAL_ACTIONS)]
if not failures.empty:
    raise RuntimeError(f"{len(failures)} failed action(s)")
```

### Safety features

- **Dry run.** With `dry_run=True` nothing is changed; the log shows what would be, with actions ending `(dry run)`. A dry run of `update_memberships` reads each group's current members, so it lists the members a live run would add and remove, and shows `removals_skipped` and `removal_guardrail_triggered` where a live run would. A dry run of `grant_group_access` reads the groups and the libraries' role assignments, so it reports missing groups and libraries, and grants already held, as a live run would.
- **Removal guardrail.** If syncing a group would remove more than `max_removals_per_group` members, its removals are skipped and logged as `removal_guardrail_triggered` (additions still happen). Re-run with `allow_mass_removal=True` once you've checked the data.
- **Run-wide removal limit.** Each group's limit alone wouldn't stop an empty or truncated members table, which would empty every small group. So `update_memberships` adds up the removals for the whole call before making any: if they'd remove more than `max_removals_per_run` members in total (default 100), none are made, and one `removal_guardrail_triggered` row, with no library, says how many there would have been. `allow_mass_removal=True` lifts this limit too.
- **Incomplete data never drives removals.** If any user (or nested group) for a role can't be resolved, removals of that kind from that role's group are skipped (`removals_skipped`), as the "stale" member may be the one that failed to resolve. A row whose `permission` isn't `read`, `contribute` or `owner` is logged as `skip`, and as it could be meant for any role, removals of that kind are skipped from all of the library's groups.
- **Manual groups are never synced or emptied,** and service principals and devices in Auto groups are left in place.
- **No guessing between duplicates.** If two Entra groups share a name the manager would use, it stops for that library rather than pick one. Library codes that would produce the same group names (such as `PE-KS3` and `PEKS3`, or codes differing only in case) are skipped with `group_name_collision`, including when the other library was created by an earlier run. A library is only ever found by its URL, so a code is never matched to another library that merely has that title (`library_name_conflict`). A library is given its own permissions and groups before its display name, so if another list already has that name, it keeps its code as its title (also `library_name_conflict`) rather than being left open to the whole site.
- **Groups must be the app's.** Group names are predictable, so a group someone else creates with the name a library would use could otherwise be handed that library's access. The app is made owner of the groups it creates, and before it grants or syncs an existing group it checks the app is an owner: a group with no owner (as earlier versions created them) has the app added and is logged as `group_owner_added`; a group owned only by others is skipped with `group_not_owned`. This includes an Auto group that `update_memberships` would empty, so a group the app doesn't own is never emptied. Dry runs of `create_libraries` and `update_memberships` read the owners and log what a live run would do. Groups are only checked when an operation uses them, so to adopt them all at once see [Adopting groups created by earlier versions](#adopting-groups-created-by-earlier-versions). The app's object ID is read from its Graph token, or can be passed as `service_principal_id`. Groups created with `app_owned=False` aren't owned or checked, and `grant_group_access` grants the groups it's given by name.
- **Creating a group, navigation link or OneDrive shortcut is never repeated** after a failure that may have succeeded (a 5xx or timeout), so a retry can't make a duplicate. The next run finds what was created. Graph's name lookups can miss a group for a short while after it's created, so a manager also remembers the groups it creates. Until the lookup finds one, later operations in the same run reuse it rather than create it again, and check its owner and read its members as they would any group's.
- **Only document libraries count as libraries.** An ordinary list whose URL ends in the code (`/Lists/ENG`) is never adopted, though its title still counts for `library_name_conflict`.
- **A moved library keeps working until its replacement works.** Given `new_site_url`, `retire_moved_libraries` leaves the old library alone (`retire_deferred`) until the library is working on the new site: it exists, has its own permissions, and gives at least one of its groups access. If `new_site_url` is the same site as `site_url`, spelt differently (such as in case, or with a trailing slash), the library hasn't moved, so it's left alone (`retire_skipped`).
- **Permission levels are found by type**, not by name, so sites in another language, or with renamed levels, work too.
- **Tokens only go to your tenant.** Every `site_url` must be `https://` on `sharepoint_hostname`, with no user name and no backslash before its path (which some URL parsers read as part of the host and others don't), and redirects to another host or to plain http are refused. A redirect that can't be followed is an error, never a success.
- **Throttling and transient errors are retried,** honouring `Retry-After`, with exponential backoff. Each attempt gets a current token from the credential, so a request retried for several minutes doesn't fail because its token expired while it waited.

### Adopting groups created by earlier versions

Groups created before 0.5.0, by earlier versions or by hand-written code they replaced, have no owner. The app is added as a group's owner the first time a live run uses it:

- **Auto groups** are checked by every `update_memberships` run, for every library you give it in `df` or `libraries_df`. In the [recommended flow](#recommended-flow) that's every live library, so nothing extra is needed.
- **Manual groups** are only checked by `create_libraries`. The recommended flow only gives it libraries that changed since the last run, so a library that hasn't changed never has its Manual groups checked.

To adopt them all, run the create pass once over every live library rather than only the changed ones. It is safe to re-run, so apart from ownership it only changes what had drifted from your data. Add this after step 3 of the [recommended flow](#recommended-flow), which provides `live` and `save`:

```python
# One-off: the create pass over every live library, not only the changed ones
results = mgr.create_libraries(live)
save("adopt_groups", results)
display(results[results["action"].isin(["group_owner_added", "group_owner_added (dry run)", "group_not_owned"])])
```

Run it once with `DRY_RUN = True`, then once with `DRY_RUN = False`, and remove it. The manager's default and the `dry_run` column `save` writes both come from `DRY_RUN`, so the audit table records what actually ran. `group_owner_added` rows (ending `(dry run)` in a dry run) are the groups adopted. `group_not_owned` rows are groups someone else owns: check who created each one, then add the app's service principal as a co-owner or delete the group.

Groups created with `app_owned=False`, such as leadership groups a person owns, aren't checked or adopted.

### Using it in your own setting

The manager only works with DataFrames, so it doesn't mind where the data comes from: a school's faculties, a company's departments, project teams or anything else that needs a library with its own access. This section suggests a set of tables to drive it from, the tables to record what it did, and the order to run the operations in. The [example notebook](#example-notebook) is one setting built this way.

#### Input tables

Keep the desired state in a few lakehouse tables, maintained by whatever feeds them (an MIS or HR extract, a hand-maintained table, a dataflow). Names and the extra columns used only by the flow are suggestions; the columns the manager reads are fixed.

**`sharepoint_libraries`**: one row per library, holding its current state.

| Column | Type | Required | Meaning |
| --- | --- | --- | --- |
| `library_id` | string | Yes | Stable ID from the source system. Stored on the library as `source_id`, so a library can be traced back to its record. |
| `site_url` | string | Yes | Full URL of the site the library lives on, `https://<tenant>.sharepoint.com/sites/<site>`. |
| `library_code` | string | Yes | Short code, used as the library's URL and in its group names (for example `MATH`). Letters, digits, spaces, `-`, `_` and `.` only, not starting or ending with a space or `.`, and not containing `_vti_` in any case, which SharePoint reserves: SharePoint changes other characters when it makes the URL, so any other code raises `ValueError` before anything is changed. Never change it after the library is created; rename with `library_display_name` instead. To recode a library, keep its row but set `deleted` (and `last_changed_ts`), and add a new row, with a new `library_id`, for the new code: the next run unlinks the old library and creates the new one. |
| `library_display_name` | string | Yes | Name shown in SharePoint (for example `Mathematics`). Can change; the next create pass applies it. |
| `group_prefix` | string | Yes | Start of the group names, for example `ACA-Sharepoint`. Use one prefix per site. |
| `auto_groups` | boolean | No (True) | Create the synced Auto-Read and Auto-Contribute groups. |
| `manual_read`, `manual_contribute` | boolean | No (False) | Create hand-curated Manual-Read / Manual-Contribute groups. |
| `app_owned` | boolean | No (True) | The app owns the groups. False for groups a person owns and curates. |
| `previous_site_url` | string | Flow only | The site the library was on before a move, or null. |
| `deleted` | boolean | Flow only | True once the library has been removed at source. Keep the row, so the unlink pass sees it. |
| `last_changed_ts` | timestamp | Flow only | When the row last changed. Drives the incremental passes. |

A blank (null) in `auto_groups`, `manual_read`, `manual_contribute` or `app_owned` takes the default shown. They can also hold 1 and 0, but anything else, such as the text `"False"`, raises `ValueError` before anything is changed.

**`sharepoint_library_members`**: one row per person per library. The full desired membership, not changes.

| Column | Type | Required | Meaning |
| --- | --- | --- | --- |
| `library_id` | string | Yes | Joins to `sharepoint_libraries`. |
| `user_principal_name` | string | Yes | The person's sign-in name or email. Aliases and secondary addresses are resolved. A blank or null one is reported as `user_not_found`, and nobody is removed from that permission's group, as for any user who can't be found. |
| `permission` | string | Yes | `read`, `contribute` or `owner`. |

**`sharepoint_library_nested_groups`** (optional): existing Entra ID security groups to nest in a library's Auto groups, such as an all-staff group.

| Column | Type | Required | Meaning |
| --- | --- | --- | --- |
| `library_id` | string | Yes | Joins to `sharepoint_libraries`. |
| `member_group` | string | Yes | The group's object ID, or its exact display name. |
| `permission` | string | Yes | `read`, `contribute` or `owner`. |

**`sharepoint_library_grants`** (optional): existing groups that should hold a permission on a library without being synced, such as a leadership team on every library.

| Column | Type | Required | Meaning |
| --- | --- | --- | --- |
| `library_id` | string | Yes | Joins to `sharepoint_libraries`. |
| `group_name` | string | Yes | The group's exact display name. |
| `permission` | string | Yes | `read`, `contribute` or `owner`. |

#### Output tables

**`sharepoint_action_log`**: every operation's action log, appended. The seven action log columns (`site_url`, `document_library`, `group`, `user`, `action`, `permission`, `status`, all strings), plus:

| Column | Type | Meaning |
| --- | --- | --- |
| `run_id` | string | Identifies the run, to find everything it did. |
| `operation` | string | Which operation logged the row, for example `update_memberships`. |
| `dry_run` | boolean | Whether the run was a dry run. |
| `logged_at` | timestamp | When the row was written. |

**`sharepoint_sync_state`**: the watermark for the incremental passes.

| Column | Type | Meaning |
| --- | --- | --- |
| `process` | string | Names the flow, for example `sharepoint_libraries`. Include the rollout scope in it if you filter the inputs, so a wider scope starts a fresh watermark. |
| `last_processed_ts` | timestamp | The newest `last_changed_ts` a clean run has processed. |
| `updated_at` | timestamp | When it was last advanced. |

Optionally, also write the `user_not_found` rows to a table of their own (`run_id`, `site_url`, `library_code`, `user_principal_name`, `reason`, `logged_at`), so whoever owns the source data can fix the addresses.

#### Recommended flow

Run the operations in this order, from a pipeline, with a dry run first whenever the inputs or the package version change:

1. **Read the libraries table, apply any scope, then read the watermark**, and the newest `last_changed_ts` as an upper bound. Only rows between the two are "changed" in this run. Rows that land while it runs are left for the next one.
2. **Lock down site sharing** on every live site (full pass). It's independent of the rest, and corrects settings changed by hand.
3. **Create** changed, not-deleted libraries (incremental). This also applies renames and new metadata. It runs before the grant and membership passes, which need the libraries to exist.
4. **Unlink** changed, deleted libraries (incremental).
5. **Retire** libraries whose site changed (incremental, on `previous_site_url`). This runs after the create pass, and with `new_site_url` it only retires a library once its replacement is working on the new site, so a failed create never leaves the library without working access (`retire_deferred` until then).
6. **Grant** the extra groups on every live library (full pass). Additive, so it also restores access removed by hand.
7. **Sync memberships** for every live library (full pass), passing all live libraries as `libraries_df` so a library whose last member left has its groups emptied.
8. **Save each operation's action log** straight after it runs, so a later failure doesn't lose it.
9. **Fail the run** if any row's action is in `FATAL_ACTIONS`. Show the `WARNING_ACTIONS` rows, but don't fail on them: they're data problems that failing can't fix.
10. **Advance the watermark** to the upper bound, only after a clean run that wasn't a dry run. A failed run leaves it alone, so its changes are retried.

Every operation is safe to re-run, so a failed run is fixed by running it again.

```python
# Parameters cell
DRY_RUN = True
SCHEMA = "<workspace>.<lakehouse>.<schema>"
TENANT_ID = "<tenant-id>"
CLIENT_ID = "<client-id>"
KEY_VAULT = "https://<key-vault-name>.vault.azure.net/"
CERTIFICATE_NAME = "<certificate-name>"
SHAREPOINT_HOSTNAME = "<tenant>.sharepoint.com"
```

```python
import uuid
from datetime import datetime, timezone

import pandas as pd

from alt_fabric_m365.auth import KeyVaultCertificateCredential
from alt_fabric_m365.sharepoint import FATAL_ACTIONS, WARNING_ACTIONS, SharePointLibraryManager

credential = KeyVaultCertificateCredential(TENANT_ID, CLIENT_ID, KEY_VAULT, CERTIFICATE_NAME)
mgr = SharePointLibraryManager(
    credential, SHAREPOINT_HOSTNAME,
    dry_run=DRY_RUN, max_workers=4, include_site_name=False, add_to_navigation=True,
)

RUN_ID = str(uuid.uuid4())
PROCESS = "sharepoint_libraries"
logs = []


def save(operation, log):
    """Step 8: keep the log for the gate, and append it to the audit table."""
    logs.append(log)
    if not log.empty:
        rows = log.assign(run_id=RUN_ID, operation=operation, dry_run=DRY_RUN, logged_at=datetime.now(timezone.utc))
        spark.createDataFrame(rows).write.format("delta").mode("append").saveAsTable(f"{SCHEMA}.sharepoint_action_log")


# Step 1: what changed since the last clean run, up to now
libraries = spark.table(f"{SCHEMA}.sharepoint_libraries").toPandas()
# Setting-specific: filter the rows in scope here (see "Adding setting-specific steps")

# The watermark is read after any scope is applied, as the scope is part of its key
state = spark.sql(f"SELECT last_processed_ts FROM {SCHEMA}.sharepoint_sync_state WHERE process = '{PROCESS}'").collect()
since = state[0]["last_processed_ts"] if state else datetime(1900, 1, 1)
until = libraries["last_changed_ts"].max() if not libraries.empty else since
changed = libraries[(libraries["last_changed_ts"] > since) & (libraries["last_changed_ts"] <= until)]
live = libraries[~libraries["deleted"]].assign(source_id=lambda df: df["library_id"])

save("site_sharing_lockdown", mgr.enforce_site_sharing_lockdown(live["site_url"]))                          # step 2
save("create_libraries", mgr.create_libraries(live[live["library_id"].isin(changed["library_id"])]))        # step 3
save("unlink_libraries", mgr.unlink_libraries(changed[changed["deleted"]]))                                  # step 4

# Step 5: new_site_url makes retire wait until the library exists on its new site (not for deleted ones)
moved = changed[changed["previous_site_url"].notna() & (changed["previous_site_url"] != changed["site_url"])]
save("retire_moved_libraries", mgr.retire_moved_libraries(moved.assign(
    site_url=moved["previous_site_url"], new_site_url=moved["site_url"].where(~moved["deleted"]),
)))

# Steps 6 and 7 read the optional tables too; leave out any you don't use
keys = live[["library_id", "site_url", "library_code", "group_prefix"]]
grants = spark.table(f"{SCHEMA}.sharepoint_library_grants").toPandas().merge(keys, on="library_id")
save("grant_group_access", mgr.grant_group_access(grants))                                                  # step 6

members = spark.table(f"{SCHEMA}.sharepoint_library_members").toPandas().merge(keys, on="library_id")
nested = spark.table(f"{SCHEMA}.sharepoint_library_nested_groups").toPandas().merge(keys, on="library_id")
save("update_memberships", mgr.update_memberships(members, libraries_df=keys, nested_groups_df=nested))      # step 7

# Setting-specific: extra passes go here, before the gate (see "Adding setting-specific steps")

# Step 9: fail on failed actions, show warnings
results = pd.concat(logs, ignore_index=True)
display(results[results["action"].isin(WARNING_ACTIONS)])
failed = results[results["action"].isin(FATAL_ACTIONS)]
if not failed.empty:
    raise RuntimeError(f"{len(failed)} failed action(s), see sharepoint_action_log where run_id = '{RUN_ID}'")

# Step 10: advance the watermark after a clean live run
if not DRY_RUN and not changed.empty:
    spark.sql(f"""
        MERGE INTO {SCHEMA}.sharepoint_sync_state AS t
        USING (SELECT '{PROCESS}' AS process, TIMESTAMP'{until:%Y-%m-%d %H:%M:%S.%f}' AS last_processed_ts) AS s
        ON t.process = s.process
        WHEN MATCHED THEN UPDATE SET t.last_processed_ts = s.last_processed_ts, t.updated_at = current_timestamp()
        WHEN NOT MATCHED THEN INSERT (process, last_processed_ts, updated_at)
            VALUES (s.process, s.last_processed_ts, current_timestamp())
    """)
```

#### Adding setting-specific steps

The skeleton is deliberately generic. Most settings add a few rules of their own, at the two places marked in it. These are the patterns the [example notebook](#example-notebook) uses for its school faculties; swap in your own names.

**Scope and exclusions**, straight after reading the libraries table. Everything after it works on the filtered rows, so a filtered-out library is neither created nor synced. Include the scope in the watermark's `process` key, so widening it starts a fresh watermark and the newly included rows are processed:

```python
# Roll out a few sites at a time, let the source opt rows out, and keep the
# leadership library's code (below) for that library alone
SITES_IN_SCOPE = {"https://<tenant>.sharepoint.com/sites/Academy"}
LEADERSHIP_CODE = "SLT"

libraries = libraries[
    libraries["site_url"].isin(SITES_IN_SCOPE)
    & ~libraries.get("exclude", pd.Series(False, index=libraries.index)).astype("boolean").fillna(False)
    & (libraries["library_code"].map(sanitise_for_group_name).str.upper() != LEADERSHIP_CODE)
]
PROCESS = f"sharepoint_libraries|{','.join(sorted(SITES_IN_SCOPE))}"
```

This needs `from alt_fabric_m365.sharepoint import sanitise_for_group_name`. The code check compares sanitised codes, because that's how they appear in group names: `S-L-T` would otherwise share the leadership library's groups. The example notebook does the same filtering in Spark SQL instead, in its `eligible()` predicate, and also reports the excluded rows as warnings so they don't go unnoticed.

**A leadership library on each site** (the example notebook's SLT library), before the gate. It's governed only by hand-curated Manual groups that a person (the headteacher) owns, so `app_owned=False`, and it has no Auto groups. A second pass grants those groups access to every library on their site:

```python
# One leadership library per site, with Manual groups only, owned by a person
leadership = live[["site_url", "group_prefix"]].drop_duplicates().assign(
    library_code=LEADERSHIP_CODE, library_display_name="Senior Leadership",
    auto_groups=False, manual_read=True, manual_contribute=True, app_owned=False,
)
save("create_leadership_libraries", mgr.create_libraries(leadership))

# Its groups get Contribute and Read on every library on their site. Names come
# from mgr.group_name, so they always match the groups create_libraries made.
leadership_grants = pd.DataFrame(
    [
        {
            "site_url": row.site_url,
            "library_code": row.library_code,
            "group_name": mgr.group_name(row.group_prefix, row.site_url, LEADERSHIP_CODE, suffix),
            "permission": permission,
        }
        for row in keys.itertuples(index=False)
        for suffix, permission in (("Manual-Contribute", "contribute"), ("Manual-Read", "read"))
    ],
    columns=["site_url", "library_code", "group_name", "permission"],
)
save("grant_leadership_access", mgr.grant_group_access(leadership_grants))
```

The leadership library isn't in `sharepoint_libraries`, so the membership sync never touches its groups, and the grant pass is additive, so a library that's later unlinked keeps the leadership access. If two academies could share a site, give each its own code or skip the shared site, as `create_libraries` keeps only the first of two rows for the same library.

**Other rules the example notebook adds**, which need no new operations:

- **Group prefix from the source.** It builds `group_prefix` as `academy_code || '-Sharepoint'` in SQL, so each academy's groups are named after it.
- **Nested groups from flags.** Instead of a nested-groups table, it builds the rows from two flag columns and a group column on the faculty list (`all_staff_read`, `all_staff_contribute`, `all_staff_security_group`), one row per flag that's set.
- **Unresolved users with context.** It joins the `user_not_found` rows back to the source table, so each carries the academy and subject, which helps whoever fixes the address in the source system.

Things to decide once, before the first live run, as changing them later orphans what was created:

- **Group naming.** `group_prefix` and `include_site_name` make up every group name, and names are how the manager finds its groups. Use one prefix per site, so two sites never share a library's groups.
- **Library codes.** A library's code is its URL and part of its group names. Treat it as permanent, and use `library_display_name` for renames. Codes can only contain letters, digits, spaces, `-`, `_` and `.`, with no space or `.` at either end and no `_vti_` (which SharePoint reserves), so trim padded codes (such as from a `CHAR` column) in the query that reads them. Codes that differ only in punctuation or case (`PE-KS3`, `peks3`) would share groups, so they're skipped with `group_name_collision`.
- **Who owns the groups.** Leave `app_owned` on for groups the automation manages. Turn it off for groups a person owns and curates, and give those to `grant_group_access` rather than the membership sync.
- **Scope.** If you roll out a few sites at a time, filter the input DataFrames and include the filter in the watermark's `process` key. Widening the scope then starts a fresh watermark, so the newly included rows are processed.

Outside Fabric, `KeyVaultCertificateCredential` isn't available (it needs `notebookutils`). Pass any credential with a `get_token` method that uses a certificate, such as `azure.identity.CertificateCredential`, and run the same flow with the input tables read however suits.

### Example notebook

[`examples/Manage Faculty Sharepoint Libraries.Notebook`](examples/Manage%20Faculty%20Sharepoint%20Libraries.Notebook/notebook-content.py) is a complete notebook, run from a pipeline, that drives every operation from curated lakehouse tables: incremental create, unlink and retire using a watermark, full membership and access passes, an audit table and a run outcome gate. [docs/faculty-sharepoint-notebook.md](docs/faculty-sharepoint-notebook.md) describes it cell by cell, and how it was moved onto this package.

## Troubleshooting

`SharePointLibraryManager`'s failure rows give the status code and the service's error code and message, such as `Status code: 403, Authorization_RequestDenied: Insufficient privileges to complete the operation.`

| Error | Likely cause |
| --- | --- |
| `Failed to get certificate ... Status code: 403, Forbidden` | The identity running the notebook can't read the certificate. Assign Key Vault Reader on the certificate, or the certificate **Get** access policy. |
| `Failed to get certificate ... Status code: 404` | The certificate name or Key Vault URL is wrong. |
| `ValueError` when creating `KeyVaultCertificateCredential` | `akv_url` isn't an `https://` Key Vault URL, such as `https://<key-vault-name>.vault.azure.net/`. |
| `Failed to sign client assertion ... Forbidden ... keys/sign/action` | The identity running the notebook can't sign with the key. Assign Key Vault Crypto User on the key, or the key **Sign** access policy. |
| `AADSTS700027` (certificate not registered on application) | The certificate version used to sign hasn't been uploaded to the app registration. This usually happens after a renewal, see [Renewing the certificate](#renewing-the-certificate). |
| `AADSTS700016` (application not found in the directory) | The client ID or tenant ID is wrong. |
| `AADSTS7000215` (invalid client secret) | Client secret only: the secret is wrong or has expired. |
| `Status code: 403, Error code: ErrorAccessDenied` from an `Emails` method | The app has no Exchange role for this mailbox. Check `Test-ServicePrincipalAuthorization`, or wait up to 2 hours after a change. |
| Graph `404` for the mailbox | The mailbox address is wrong, or the mailbox doesn't exist. |
| `ValueError` from an `Emails` method | No mailbox was given: pass one to `Emails(credential, mailbox)` or `shared_mailbox_email` to the call. |
| `MessageNotFoundError: No matching messages found.` | No email in the Inbox matches the exact subject and sender. Check both, or try `only_search_inbox=False`. |
| `get_attachment_ids_and_names` returns an empty list | The email has no file attachments. Inline images are left out unless you pass `include_inline=True`, and attached emails and cloud links are always left out. |
| `Attachment ... is a #microsoft.graph.itemAttachment` | `download_attachment` was given the ID of an attached email, not a file. |
| `ImportError: notebookutils is not available. KeyVaultCertificateCredential must be used in a Fabric notebook.` | The code is running outside Fabric. `KeyVaultCertificateCredential` only works in Fabric (or Synapse) notebooks. |
| SharePoint `401` with `Unsupported app only token` | The token came from a client secret. SharePoint needs a certificate, see [Giving the app access](#giving-the-app-access). |
| Graph or SharePoint `403` (`accessDenied`) for some sites but not others | The app has `Sites.Selected` but hasn't been granted those sites, see [Grant the app each site](#2-grant-the-app-each-site). |
| Graph or SharePoint `403` for every site | `Sites.Selected` (or another `Sites.*` permission) is missing from that API on the app registration, or lacks admin consent. Graph and SharePoint each need their own. Check the token's `roles` claim, see [Checking what a token contains](#checking-what-a-token-contains). |
| `403` when running `New-MgSitePermission` | The account running it needs `Sites.FullControl.All` consented for Microsoft Graph PowerShell, a SharePoint Administrator role, and access to the site. |
| Graph `403` after granting a permission | Tokens issued before the change don't include it. Create a new credential or start a new session to get a fresh token. |
| `ValueError: site_url '...' is not on https://...` | A `site_url` isn't on the `sharepoint_hostname` the manager was created with, uses `http://`, or has a user name or a backslash before its path. |
| `error` row: `Multiple Entra groups share displayName` | Two groups have the name the manager uses for a library. Delete or rename the duplicate. |
| `user_not_found` rows | The address matches no user's UPN, mail or proxy addresses. Fix it in the source data; the next run picks the user up. |
| `skip` row: `unknown permission '...'` | A membership row's `permission` isn't `read`, `contribute` or `owner`. Until it's fixed in the source data, no users (or, for a nested group, no nested groups) are removed from that library's groups. |
| `removal_guardrail_triggered` row with no library | The whole call would remove more members than `max_removals_per_run`, so none were removed. Check the input data, for example that the members table loaded in full, then re-run with `allow_mass_removal=True`. |
| `removal_guardrail_triggered` row | A sync would remove more members than `max_removals_per_group`. Check the input data, then re-run with `allow_mass_removal=True`. |
| `onedrive_shortcut_failed` row | The user was added to the group, but the shortcut to the library couldn't be made in their OneDrive. Usually their OneDrive doesn't exist yet: it's created the first time they open it, and an app can't create it for them. Later runs don't try again, as the user is already a member. Once their OneDrive exists, run `ensure_onedrive_shortcuts` with the members table to create the missing shortcuts, or the user can add one themselves: open the library and select **Add shortcut to My files**. To avoid it for new starters, have a SharePoint administrator create their OneDrive in advance with [`Request-SPOPersonalSite`](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/request-spopersonalsite). |
| `library_not_found` row | The library doesn't exist on the site, for example it was deleted by hand. Recreate it with `create_libraries` or remove it from the input. When unlinking, its Auto groups are still emptied. |
| `group_not_owned` row | A group with the name the manager uses is owned by someone, but not by the app registration, so it wasn't used. Check who created it: if it's legitimate, add the app's service principal as an owner; otherwise delete it. |
| `retire_deferred` row | The library isn't working on its new site yet (for example its create failed, or the copy there is the `_MOVED` leftover of an earlier move back), so the old one was left working. The status says what's missing. It's retired on a later run once the new one exists. |
| `library_name_conflict` row | Another list on the site already has the library's code, or its display name, as its title. With the code, the library can't be created; with the display name, the library keeps its code as its title. Rename the other list's title. |

## Upgrading to 1.0.0

1.0.0 removes the old names that 0.x versions kept working with a `FutureWarning`. Nothing else changes, so a notebook that runs on 0.7.x without a `FutureWarning` runs unchanged on 1.0.0.

| Removed | Use instead |
| --- | --- |
| `alt_fabric_m365.graph_api`, renamed in 0.7.0 | `alt_fabric_m365.exchange`, for example `from alt_fabric_m365.exchange import Emails, MessageNotFoundError` |
| The `fabric_python_helper` package, renamed in 0.6.0 | The `alt-fabric-m365` package, imported as `alt_fabric_m365` |

To upgrade:

1. In each notebook, change `fabric_python_helper` to `alt_fabric_m365` and `graph_api` to `exchange` in every import:

   ```python
   from alt_fabric_m365.auth import KeyVaultCertificateCredential
   from alt_fabric_m365.exchange import Emails
   ```

2. In each Fabric environment, remove `fabric_python_helper` under **Public libraries** if it's there. Its last release, 0.6.0, requires `alt-fabric-m365` 0.6, so it can't be installed alongside 1.0.0.
3. Set `alt-fabric-m365` to `1.0.0` and publish the environment.

To upgrade from 0.3.x or earlier, which used delegated (signed-in user) authentication and had a `pbi_admin` module, first follow [Upgrading from 0.3.x](https://github.com/ArchwayTrust/Insights-Fabric-M365/blob/v0.7.2/README.md#upgrading-from-03x) in the 0.7.2 README.

## Building and publishing

### Running the tests

The tests mock Microsoft Entra ID, Key Vault, Graph and SharePoint, and replace `notebookutils` and Spark, so they run anywhere without credentials. From the repository root:

```powershell
python -m pip install -e .[test]
python -m pytest
python -m pyflakes alt_fabric_m365 tests setup.py
```

The [Tests workflow](.github/workflows/tests.yml) runs the same checks on Python 3.10, 3.11 and 3.12, and builds the package, for every pull request into `main` and every push to `main`. It doesn't publish anything.

#### Requiring the tests to pass before merging

On its own, the Tests workflow only reports a result: GitHub still lets a pull request with failing tests be merged. To block that, add a ruleset on `main` once the workflow has run at least once (GitHub only lists a check after it has reported):

1. On GitHub, go to **Settings** > **Rules** > **Rulesets** > **New ruleset** > **New branch ruleset**.
2. Name it, for example `Protect main`, and set **Enforcement status** to **Active**.
3. Under **Target branches**, select **Add target** > **Include default branch**.
4. Turn on **Require a pull request before merging**. Under it, turn on **Require review from Code Owners**, so changes need a review from the teams in [`.github/CODEOWNERS`](https://github.com/ArchwayTrust/Insights-Fabric-M365/blob/main/.github/CODEOWNERS): Insights for the package, and IT Approvers for the SharePoint library manager, the example notebook that runs it and their tests, as they change the tenant's libraries, permissions and groups.
5. Turn on **Require status checks to pass**, select **Add checks**, and add `test (3.10)`, `test (3.11)` and `test (3.12)`.
6. Turn on **Block force pushes**, then select **Create**.

If the Python versions in `tests.yml` change, update the required checks to match, or pull requests will wait for a check that never runs.

What they cover:

| File | Covers |
| --- | --- |
| `test_auth.py` | `KeyVaultCertificateCredential`: token requests and caching, Key Vault signing, the client assertion, error descriptions |
| `test_exchange.py` | `Emails`, its retries and reading Graph errors |
| `test_sharepoint_http.py` | The retry wrapper: throttling, `Retry-After`, server and connection errors, redirects, a current token for each attempt |
| `test_sharepoint_helpers.py` | Group naming and collisions, input validation, user and group resolution, the property bag, navigation URLs |
| `test_sharepoint_manager.py` | Every `SharePointLibraryManager` operation end to end, including re-runs, dry runs, guardrails and failures, and authenticating with `KeyVaultCertificateCredential` |
| `test_readme_flow.py` | Runs the [recommended flow](#recommended-flow) example against the suggested input tables (a dry run, a live run, a re-run and a deletion), and again with the [setting-specific steps](#adding-setting-specific-steps) inserted |
| `test_faculty_notebook.py` | Runs the [example notebook](#example-notebook) cell by cell, with Spark SQL on SQLite and the real `KeyVaultCertificateCredential` |
| `test_packaging.py` | The version, public names, that removed names stay removed, and that the README's examples are valid Python |

The SharePoint tests run against `tests/sharepoint_fake.py`, an in-memory fake of the Graph and SharePoint endpoints the manager calls, which keeps state so a test can check what an operation changed. Add `--cov=alt_fabric_m365` to see coverage.

Code style, when to use a function rather than a method, and what a change's tests should check are in [CLAUDE.md](CLAUDE.md).

The tests can't show that the real services accept the requests, so test a new version in Fabric before releasing it (see [Testing a wheel in Fabric](#testing-a-wheel-in-fabric-before-publishing)). Get a token, find a known email and list its attachments. For the SharePoint manager, run the example notebook with `RUN_SMOKE_TEST = True`, first with `DRY_RUN = True`.

### Versioning

The version number is only set in `alt_fabric_m365/__init__.py`, as `__version__="1.2.0"`. `setup.py` reads it from there, so there's nothing else to keep in sync. Check it with `python setup.py --version`.

Versions follow [semantic versioning](https://semver.org/):

- **Major** (1.0.0 to 2.0.0) for breaking changes: removing or renaming a public name, changing the arguments a public function or method accepts or what it returns, or changing the columns of an action log.
- **Minor** (1.0.0 to 1.1.0) for new features that don't break existing code.
- **Patch** (1.0.0 to 1.0.1) for fixes and documentation.

Public names are those in each module's `__all__`; names starting with `_` can change in any release. To rename or remove a public name, keep the old name working with a `FutureWarning` in a minor release, and remove it in the next major release.

PyPI never accepts the same version twice, even if it's deleted.

### Building a wheel locally

From the repository root, in PowerShell:

```powershell
python -m pip install --upgrade build twine
Remove-Item dist -Recurse -Force -ErrorAction SilentlyContinue
python -m build
python -m twine check dist/*
```

This creates two files in `dist/`:

- `alt_fabric_m365-1.2.0-py3-none-any.whl`, the wheel
- `alt_fabric_m365-1.2.0.tar.gz`, the source distribution

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
%pip install /lakehouse/default/Files/alt_fabric_m365-1.2.0-py3-none-any.whl
```

Or add it to an environment under **Custom libraries** > **Upload**.

### Publishing to PyPI with a GitHub release

This is the normal route. Merging to `main` never publishes anything; only publishing a GitHub release does. The [Release workflow](.github/workflows/release.yml) then:

1. Runs the tests.
2. Checks the release tag matches `__version__` (tag `v1.2.0` for `__version__="1.2.0"`), and stops if it doesn't.
3. Builds the package and runs `twine check`.
4. Publishes it to PyPI.
5. Attaches the wheel and source distribution to the GitHub release.

To release:

1. Update `__version__` in `alt_fabric_m365/__init__.py`, merge it into `main` through a pull request.
2. On GitHub, go to **Releases** > **Draft a new release**.
3. Create a tag matching the version, for example `v1.2.0`, targeting `main`.
4. Add release notes, then select **Publish release**.
5. Check the workflow run under **Actions**, and that the new version appears on [PyPI](https://pypi.org/project/alt-fabric-m365/).

If a step fails, nothing after it runs. PyPI never accepts the same version twice, so a release whose version is already on PyPI fails at the publish step rather than silently uploading nothing.

How to recover depends on where the run failed. Check the **Publish to PyPI** step in the run, and whether the version is on [PyPI](https://pypi.org/project/alt-fabric-m365/#history):

- **Before publishing** (tests, tag check or build failed, nothing on PyPI): delete the GitHub release and its tag, fix the problem, and publish the release again.
- **After publishing** (the version is on PyPI, but attaching files to the release failed): don't delete the release, as publishing it again would fail at the PyPI step. Open the run under **Actions** and select **Re-run failed jobs**. Only the failed job runs again, using the files already built.
- **A published version is broken:** it can't be replaced. Fix it, increase the patch version (for example 1.0.1 to 1.0.2) and release that. Optionally [yank](https://pypi.org/help/#yanked) the broken version on PyPI so pip stops choosing it.

#### Trusted Publishing setup

The workflow uploads to PyPI with [Trusted Publishing](https://docs.pypi.org/trusted-publishers/), so no PyPI API token is stored in GitHub. PyPI only accepts uploads from this repository's `release.yml` workflow running in the `pypi` environment. Set it up once:

1. On PyPI, open the `alt-fabric-m365` project, then **Manage** > **Publishing** > **Add a new publisher** > **GitHub**. Enter owner `ArchwayTrust`, repository `Insights-Fabric-M365`, workflow `release.yml` and environment `pypi`. If the repository is renamed, add a publisher with the new name here before the next release, as PyPI rejects uploads from any other repository; remove the old one once a release has succeeded.
2. On GitHub, go to **Settings** > **Environments** > **New environment** and create `pypi`. Then, in the environment:
   - Under **Deployment branches and tags**, choose **Selected branches and tags**, select **Add deployment branch or tag rule**, set **Ref type** to **Tag** and enter the pattern `v*`. Only release tags can then publish, even if the workflow is started some other way.
   - Optionally, turn on **Required reviewers** and add the people who can approve a release, so each upload to PyPI waits for one of them.

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
%pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ alt-fabric-m365==1.2.0
```

### After publishing

Update the version in each Fabric environment that uses the package (**Public libraries**), and publish the environment.

### Release checklist

1. Update `__version__` in `alt_fabric_m365/__init__.py`.
2. Update this README, including the version history under [Installation](#installation), and an upgrading section for a new major version.
3. Build, run the tests and `twine check`, and test the wheel in Fabric.
4. Merge into `main` through a pull request, once the Tests workflow passes.
5. Publish a GitHub release tagged `v<version>`.
6. Confirm the Release workflow succeeded and the version is on PyPI.
7. Update and publish the Fabric environments.

## Licence

BSD 3-Clause, copyright Archway Learning Trust. See [LICENSE](LICENSE). The software is provided "as is", without warranty, and the licence excludes liability for any damage arising from its use.
