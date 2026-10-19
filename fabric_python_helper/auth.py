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
    A credential gets tokens for one scope, so create one per API. They can share the same app
    registration and certificate.

    Some APIs use a scope specific to your tenant or environment:
        SharePoint REST API: https://{tenant}.sharepoint.com/.default
            Grant SharePoint application permissions (Sites.Selected where possible) on the app registration.
            SharePoint only allows app-only access with a certificate, so client secrets are rejected.
        Dataverse: your environment URL followed by /.default, for example https://{org}.crm11.dynamics.com/.default
            Add the app as an application user in the Power Platform admin center.

Tokens for the identity running the notebook:
    notebookutils.credentials.getToken() returns tokens for the user or workspace identity running the
    notebook, but only for the storage, pbi, keyvault and kusto audiences. Use this module when you need
    another API, such as Graph or SharePoint, or need to act as a separate app.
"""
import requests
import json
import base64
import hashlib
import time
import uuid
from notebookutils import mssparkutils

class Scopes:
    """
    Scopes for common Microsoft APIs, for use with KeyVaultCertificateCredential.

    The comment above each says where to give the service principal access. Entra ID issues a token
    whether or not that access exists, but the API rejects requests made without it.
    """

    # Microsoft Graph. Grant application permissions on the app registration, with admin consent.
    # For mailboxes use Exchange RBAC for Applications instead, which can limit the app to specific mailboxes.
    GRAPH = "https://graph.microsoft.com/.default"

    # Power BI REST API. Add the service principal to the workspace, and allow service principals to use
    # the APIs in the Fabric admin portal tenant settings (Developer settings). Access comes from workspace
    # roles, not from API permissions on the app registration.
    POWER_BI = "https://analysis.windows.net/powerbi/api/.default"

    # Fabric REST API. Access is granted the same way as for Power BI.
    FABRIC = "https://api.fabric.microsoft.com/.default"

    # Azure Storage and OneLake (OneLake only accepts this audience). For OneLake give the service principal
    # a workspace role or item permissions, for storage accounts an Azure RBAC role such as Storage Blob Data Reader.
    STORAGE = "https://storage.azure.com/.default"

    # Azure Key Vault. Azure RBAC role or access policy on the key vault.
    KEY_VAULT = "https://vault.azure.net/.default"

    # Azure SQL, and Fabric warehouses, SQL analytics endpoints and SQL databases over TDS. In Fabric give the
    # service principal a workspace role or item permissions, in Azure SQL create a user FROM EXTERNAL PROVIDER.
    SQL = "https://database.windows.net/.default"

    # Azure Data Explorer and Fabric Eventhouse / KQL databases. Give the service principal a workspace role, or add
    # it as a database principal, for example .add database <db> viewers ('aadapp=<client_id>;<tenant_id>').
    KUSTO = "https://api.kusto.windows.net/.default"

    # Azure Resource Manager, for managing Azure resources. Azure RBAC role on the subscription, resource group or resource.
    AZURE_MANAGEMENT = "https://management.azure.com/.default"

    # Azure Monitor Log Analytics query API. Log Analytics Reader on the Log Analytics workspace.
    LOG_ANALYTICS = "https://api.loganalytics.io/.default"

class KeyVaultCertificateCredential:
    """
    Authenticates as a service principal using a certificate held in Azure Key Vault.

    The certificate's private key never leaves the key vault. A short lived client assertion (JWT)
    is built in the notebook, Key Vault signs it, and the signed assertion is exchanged with
    Microsoft Entra ID for an access token using the client credentials flow.

    Requires:
        - An app registration with the certificate's public key (.cer) uploaded under Certificates & secrets.
        - A certificate in the key vault with an RSA key. Ideally make the private key non-exportable.
        - The identity running the notebook needs Get on the certificate and Sign on its key
          (for example Key Vault Reader and Key Vault Crypto User, scoped to the certificate and key).
        - The service principal to have access to the API being called, see Scopes.

    Example:
        credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, "Certificate Name", Scopes.GRAPH)
        headers = {"Authorization": f"Bearer {credential.get_access_token()}"}

    Attributes:
        tenant_id (str): Tenant ID for Azure authentication.
        client_id (str): Client ID of the app registration.
        akv_url (str): Azure Key Vault URL, in the form https://{key_vault_name}.vault.azure.net/.
        certificate_name (str): Name of the certificate in the Azure Key Vault.
        scope (str): Scope to request the access token for, see Scopes.
        certificate_version (str): Version of the certificate to sign with. None uses the latest version.
        access_token (str): The current access token, or None if one hasn't been requested yet.
        expires_at (float): Time (seconds since epoch) at which the current access token expires.
    """

    AKV_API_VERSION = "7.4"
    CLIENT_ASSERTION_TYPE = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    ASSERTION_LIFETIME_SECONDS = 600  # Entra ID recommends 5-10 minutes at most
    REFRESH_MARGIN_SECONDS = 300  # Request a new access token when the current one has less than this left

    def __init__(self, tenant_id, client_id, akv_url, certificate_name, scope, certificate_version=None):
        """
        Initializes the credential. No calls are made until an access token is requested.

        Parameters:
            tenant_id (str): Tenant ID for Azure authentication.
            client_id (str): Client ID of the app registration.
            akv_url (str): Azure Key Vault URL.
            certificate_name (str): Name of the certificate in the Azure Key Vault.
            scope (str): Scope to request the access token for, in the form {resource}/.default. See Scopes.
            certificate_version (str): Optional. Version of the certificate to sign with. Defaults to the latest version.
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.akv_url = akv_url.rstrip("/")
        self.certificate_name = certificate_name
        self.scope = scope
        self.certificate_version = certificate_version
        self.token_endpoint = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        self.access_token = None
        self.expires_at = 0

    def get_access_token(self):
        """
        Returns an access token for the configured scope.

        The token is cached and reused until it is within 5 minutes of expiring, so this can be
        called before every request.

        Returns:
            str: An access token.
        """
        if self.access_token is None or time.time() >= self.expires_at - self.REFRESH_MARGIN_SECONDS:
            self._request_access_token()

        return self.access_token

    def _request_access_token(self):
        """
        Exchanges a newly signed client assertion for an access token and caches it.
        """
        # Build a new assertion each time, as they're only valid for a few minutes
        client_assertion = self._build_client_assertion()

        response = requests.post(self.token_endpoint, data={
            "client_id": self.client_id,
            "scope": self.scope,
            "grant_type": "client_credentials",
            "client_assertion_type": self.CLIENT_ASSERTION_TYPE,
            "client_assertion": client_assertion
        })

        if response.status_code != 200:
            raise Exception(f"Failed to get access token for {self.client_id}. {self._describe_error(response)}")

        token_response = response.json()
        self.access_token = token_response["access_token"]
        self.expires_at = time.time() + int(token_response["expires_in"])

    def _build_client_assertion(self):
        """
        Builds a client assertion (JWT) and has Key Vault sign it with the certificate's private key.

        Follows https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials

        Returns:
            str: The signed client assertion.
        """
        key_id, certificate_der = self._get_certificate()
        now = int(time.time())

        header = {
            "alg": "PS256",
            "typ": "JWT",
            "x5t#S256": self._base64url(hashlib.sha256(certificate_der).digest())
        }

        claims = {
            "aud": self.token_endpoint,
            "iss": self.client_id,
            "sub": self.client_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "iat": now,
            "exp": now + self.ASSERTION_LIFETIME_SECONDS
        }

        # Key Vault signs a digest rather than the message, so hash the JWT's signing input here
        signing_input = f"{self._base64url(json.dumps(header).encode())}.{self._base64url(json.dumps(claims).encode())}"
        digest = hashlib.sha256(signing_input.encode("ascii")).digest()

        signature = self._sign_digest(key_id, digest)

        return f"{signing_input}.{signature}"

    def _get_certificate(self):
        """
        Retrieves the public part of the certificate from Azure Key Vault.

        Returns:
            tuple: The key identifier (URL including version) to sign with, and the certificate's DER encoded bytes.
        """
        url = f"{self.akv_url}/certificates/{self.certificate_name}"
        if self.certificate_version:
            url += f"/{self.certificate_version}"

        response = requests.get(url, headers=self._akv_headers(), params={"api-version": self.AKV_API_VERSION})

        if response.status_code != 200:
            raise Exception(f"Failed to get certificate {self.certificate_name} from key vault. {self._describe_error(response)}")

        certificate = response.json()

        return certificate["kid"], base64.b64decode(certificate["cer"])

    def _sign_digest(self, key_id, digest):
        """
        Signs a SHA-256 digest with the certificate's private key using the Azure Key Vault sign operation.

        Parameters:
            key_id (str): The key identifier returned with the certificate.
            digest (bytes): The SHA-256 digest to sign.

        Returns:
            str: The base64url encoded signature.
        """
        response = requests.post(
            f"{key_id}/sign",
            headers=self._akv_headers(),
            params={"api-version": self.AKV_API_VERSION},
            json={"alg": "PS256", "value": self._base64url(digest)}
        )

        if response.status_code != 200:
            raise Exception(f"Failed to sign client assertion with key vault key {self.certificate_name}. {self._describe_error(response)}")

        return response.json()["value"]

    def _akv_headers(self):
        """
        Returns headers for calling Azure Key Vault as the identity running the notebook.
        """
        akv_cred = mssparkutils.credentials.getToken('keyvault')

        return {
            "Authorization": f"Bearer {akv_cred}",
            "Content-Type": "application/json"
        }

    @staticmethod
    def _base64url(data):
        """
        Encodes bytes as unpadded base64url, as used in JWTs.
        """
        return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")

    @staticmethod
    def _describe_error(response):
        """
        Summarises an error response from Microsoft Entra ID or Azure Key Vault without including request data.
        """
        try:
            body = response.json()
        except ValueError:
            return f"Status code: {response.status_code}"

        error = body.get("error")

        # Key Vault nests the error as an object, Entra ID returns error and error_description at the top level
        if isinstance(error, dict):
            return f"Status code: {response.status_code}, {error.get('code')}: {error.get('message')}"

        return f"Status code: {response.status_code}, {error}: {body.get('error_description')}"
