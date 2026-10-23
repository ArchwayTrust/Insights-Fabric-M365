"""
Authentication helpers for calling Microsoft APIs from Fabric notebooks as a service principal.

KeyVaultCertificateCredential gets app-only access tokens using a certificate held in Azure Key Vault.
Key Vault signs each client assertion, so the private key never leaves the key vault.

Scopes and audiences:
    Tokens are requested with the client credentials flow, so the scope is always the API's resource
    URI followed by /.default, for example https://graph.microsoft.com/.default. The token that comes
    back has that resource as its audience (aud claim). /.default means every permission the app has
    been granted on that API. An app can't request individual delegated scopes such as Mail.Read.

    Common scopes are in the Scopes class, with where to give the service principal access to each.
    One credential can get tokens for any number of APIs, as tokens are cached per scope.

    Some APIs use a scope specific to your tenant or environment:
        SharePoint REST API: https://{tenant}.sharepoint.com/.default
            Grant SharePoint application permissions (Sites.Selected where possible) on the app registration.
            Sites.Selected grants nothing until each site is also granted to the app, see the README.
            SharePoint only allows app-only access with a certificate, so client secrets are rejected.
        Dataverse: your environment URL followed by /.default, for example https://{org}.crm11.dynamics.com/.default
            Add the app as an application user in the Power Platform admin center.

Tokens for the identity running the notebook:
    notebookutils.credentials.getToken() returns tokens for the user or workspace identity running the
    notebook, but only for the storage, pbi, keyvault and kusto audiences. Use this module when you need
    another API, such as Graph or SharePoint, or need to act as a separate app.

Semantic link and Azure SDK clients:
    KeyVaultCertificateCredential implements azure.core.credentials.TokenCredential, so it can be passed as
    the credential to semantic link (sempy) functions and Azure SDK clients. They request the scopes they
    need, so no scope is needed when creating the credential:
        fabric.refresh_dataset("Sales", workspace="Finance", credential=credential)

Client secrets:
    This module only covers certificates. For a client secret use azure.identity.ClientSecretCredential,
    which is already in the Fabric runtime, with the secret read from Key Vault. It has the same get_token
    method, so it can be used anywhere this credential can:
        secret = notebookutils.credentials.getSecret(akv_url, secret_name)
        credential = ClientSecretCredential(tenant_id, client_id, secret)
    Certificates are preferred: the private key never leaves the key vault, whereas a secret is read into
    the notebook.
"""
from .keyvault import KeyVaultCertificateCredential
from .scopes import Scopes

__all__ = ["KeyVaultCertificateCredential", "Scopes"]
