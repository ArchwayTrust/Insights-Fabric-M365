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
import requests
import json
import base64
import hashlib
import threading
import time
import uuid
from urllib.parse import urlparse
from azure.core.credentials import AccessToken

class Scopes:
    """
    Scopes for common Microsoft APIs, for use with KeyVaultCertificateCredential or any other credential.

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

    Implements azure.core.credentials.TokenCredential, so it can also be passed as the credential to
    semantic link (sempy) functions and Azure SDK clients.

    Requires:
        - An app registration with the certificate's public key (.cer) uploaded under Certificates & secrets.
        - A certificate in the key vault with an RSA key. Ideally make the private key non-exportable.
        - The identity running the notebook needs Get on the certificate and Sign on its key
          (for example Key Vault Reader and Key Vault Crypto User, scoped to the certificate and key).
        - The service principal to have access to the API being called, see Scopes.

    Example:
        # Calling an API directly
        credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, "Certificate Name", Scopes.GRAPH)
        headers = {"Authorization": f"Bearer {credential.get_access_token()}"}

        # As the credential for semantic link, which requests the scopes it needs
        credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, "Certificate Name")
        fabric.refresh_dataset("Sales", workspace="Finance", credential=credential)

    Attributes:
        tenant_id (str): Tenant ID for Azure authentication.
        client_id (str): Client ID of the app registration.
        akv_url (str): Azure Key Vault URL, in the form https://{key_vault_name}.vault.azure.net/.
        certificate_name (str): Name of the certificate in the Azure Key Vault.
        scope (str): Default scope for get_access_token(), see Scopes. None if only used through get_token().
        certificate_version (str): Version of the certificate to sign with. None uses the latest version.
    """

    AKV_API_VERSION = "7.4"
    CLIENT_ASSERTION_TYPE = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    ASSERTION_LIFETIME_SECONDS = 600  # Entra ID recommends 5-10 minutes at most
    REFRESH_MARGIN_SECONDS = 300  # Request a new access token when the current one has less than this left
    REQUEST_TIMEOUT_SECONDS = 30  # Timeout for each call to Entra ID and Key Vault

    # Key Vault DNS suffixes in the public and sovereign clouds. The Key Vault token is only sent to these hosts.
    AKV_DNS_SUFFIXES = (".vault.azure.net", ".vault.azure.cn", ".vault.usgovcloudapi.net", ".vault.microsoftazure.de")

    def __init__(self, tenant_id, client_id, akv_url, certificate_name, scope=None, certificate_version=None):
        """
        Initializes the credential. No calls are made until an access token is requested.

        Parameters:
            tenant_id (str): Tenant ID for Azure authentication.
            client_id (str): Client ID of the app registration.
            akv_url (str): Azure Key Vault URL, in the form https://{key_vault_name}.vault.azure.net/.
                Must use https and a Key Vault host name, as the Key Vault token is sent to it.
            certificate_name (str): Name of the certificate in the Azure Key Vault.
            scope (str): Optional. Default scope for get_access_token(), in the form {resource}/.default. See Scopes.
                Not needed when the credential is passed to semantic link or an Azure SDK client.
            certificate_version (str): Optional. Version of the certificate to sign with. Defaults to the latest version.
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.akv_url = akv_url.rstrip("/")
        self._akv_host = self._validate_akv_url(self.akv_url)
        self.certificate_name = certificate_name
        self.scope = scope
        self.certificate_version = certificate_version
        self.token_endpoint = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        # Cached AccessTokens, keyed by the sorted, de-duplicated scopes. The lock stops concurrent callers
        # each requesting a new token when the cached one expires.
        self._tokens = {}
        self._tokens_lock = threading.Lock()

    @classmethod
    def _validate_akv_url(cls, akv_url):
        """
        Checks the Key Vault URL uses https and a Key Vault host name, as the Key Vault token is sent to it.

        Parameters:
            akv_url (str): Azure Key Vault URL.

        Returns:
            str: The Key Vault host name, in lower case.
        """
        parsed = urlparse(akv_url)
        host = (parsed.hostname or "").lower()

        if parsed.scheme != "https" or parsed.username or parsed.password or not host.endswith(cls.AKV_DNS_SUFFIXES):
            raise ValueError(f"akv_url must be a Key Vault URL such as https://{{key_vault_name}}.vault.azure.net/, not {akv_url}")

        return host

    def get_access_token(self, scope=None):
        """
        Returns an access token string, for calling an API directly.

        The token is cached and reused until it is within 5 minutes of expiring, so this can be
        called before every request.

        Parameters:
            scope (str): Optional. Scope to request the access token for. Defaults to the scope the credential was created with.

        Returns:
            str: An access token.
        """
        scope = scope or self.scope
        if scope is None:
            raise ValueError("No scope given. Pass one here or when creating the credential, see Scopes.")

        return self.get_token(scope).token

    def get_token(self, *scopes, claims=None, tenant_id=None, enable_cae=False, **kwargs):
        """
        Returns an access token for the given scopes. Implements azure.core.credentials.TokenCredential.

        Semantic link (sempy) functions and Azure SDK clients call this when the credential is passed to
        them. Tokens are cached per set of scopes, in any order, and reused until they are within 5 minutes
        of expiring.

        Parameters:
            scopes (str): Scopes to request the access token for, in the form {resource}/.default.
            claims (str): Optional. Additional claims required in the token, such as from an API's claims challenge.
                A new token is always requested when claims are given, and it replaces the cached one.
            tenant_id (str): Optional. Tenant to request the token from. Only the tenant the credential was created
                with is supported, so a ValueError is raised for any other tenant.
            enable_cae (bool): Optional. Accepted for compatibility but ignored, as continuous access evaluation
                isn't requested.
            kwargs: Other keyword arguments are accepted for compatibility but ignored.

        Returns:
            azure.core.credentials.AccessToken: The access token and its expiry time (seconds since epoch).
        """
        if not scopes:
            raise ValueError("At least one scope is required, see Scopes.")

        if tenant_id and tenant_id.lower() != self.tenant_id.lower():
            raise ValueError(f"Tokens can only be requested from tenant {self.tenant_id}, not {tenant_id}.")

        # Order doesn't matter to Entra ID, so get_token(a, b) and get_token(b, a) share a cached token
        cache_key = tuple(sorted(set(scopes)))

        # Return a valid cached token without waiting for the lock
        if not claims:
            token = self._tokens.get(cache_key)
            if self._is_valid(token):
                return token

        with self._tokens_lock:
            # Check again, as another caller may have got a new token while this one waited for the lock
            if not claims:
                token = self._tokens.get(cache_key)
                if self._is_valid(token):
                    return token

            token = self._request_access_token(" ".join(cache_key), claims)
            self._tokens[cache_key] = token

        return token

    def _is_valid(self, token):
        """
        Returns whether a cached token exists and has more than REFRESH_MARGIN_SECONDS left before it expires.
        """
        return token is not None and time.time() < token.expires_on - self.REFRESH_MARGIN_SECONDS

    def _request_access_token(self, scope, claims=None):
        """
        Exchanges a newly signed client assertion for an access token.

        Parameters:
            scope (str): Scope to request the access token for. Multiple scopes are separated by spaces.
            claims (str): Optional. Additional claims required in the token, as a JSON string.

        Returns:
            azure.core.credentials.AccessToken: The access token and its expiry time (seconds since epoch).
        """
        # Build a new assertion each time, as they're only valid for a few minutes
        client_assertion = self._build_client_assertion()

        data = {
            "client_id": self.client_id,
            "scope": scope,
            "grant_type": "client_credentials",
            "client_assertion_type": self.CLIENT_ASSERTION_TYPE,
            "client_assertion": client_assertion
        }
        if claims:
            data["claims"] = claims

        # Take the time before the request, so the expiry time errs on the early side
        now = int(time.time())

        response = requests.post(self.token_endpoint, data=data, timeout=self.REQUEST_TIMEOUT_SECONDS)

        if response.status_code != 200:
            raise Exception(f"Failed to get access token for {self.client_id} with scope {scope}. {self._describe_error(response)}")

        token_response = response.json()

        return AccessToken(token_response["access_token"], now + int(token_response["expires_in"]))

    def _build_client_assertion(self):
        """
        Builds a client assertion (JWT) and has Key Vault sign it with the certificate's private key.

        Follows https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials

        Returns:
            str: The signed client assertion.
        """
        # Get the Key Vault token once, for both the certificate and sign requests
        akv_headers = self._akv_headers()

        key_id, certificate_der = self._get_certificate(akv_headers)
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

        signature = self._sign_digest(key_id, digest, akv_headers)

        return f"{signing_input}.{signature}"

    def _get_certificate(self, akv_headers):
        """
        Retrieves the public part of the certificate from Azure Key Vault.

        Parameters:
            akv_headers (dict): Headers for calling Azure Key Vault, from _akv_headers().

        Returns:
            tuple: The key identifier (URL including version) to sign with, and the certificate's DER encoded bytes.
        """
        url = f"{self.akv_url}/certificates/{self.certificate_name}"
        if self.certificate_version:
            url += f"/{self.certificate_version}"

        response = requests.get(
            url,
            headers=akv_headers,
            params={"api-version": self.AKV_API_VERSION},
            timeout=self.REQUEST_TIMEOUT_SECONDS
        )

        if response.status_code != 200:
            raise Exception(f"Failed to get certificate {self.certificate_name} from key vault. {self._describe_error(response)}")

        certificate = response.json()

        return certificate["kid"], base64.b64decode(certificate["cer"])

    def _sign_digest(self, key_id, digest, akv_headers):
        """
        Signs a SHA-256 digest with the certificate's private key using the Azure Key Vault sign operation.

        Parameters:
            key_id (str): The key identifier returned with the certificate.
            digest (bytes): The SHA-256 digest to sign.
            akv_headers (dict): Headers for calling Azure Key Vault, from _akv_headers().

        Returns:
            str: The base64url encoded signature.
        """
        # The key identifier comes from the certificate response, so check it's in the same key vault
        # before sending the Key Vault token to it
        parsed_key_id = urlparse(key_id)
        if (parsed_key_id.scheme != "https" or parsed_key_id.username or parsed_key_id.password
                or (parsed_key_id.hostname or "").lower() != self._akv_host or parsed_key_id.port not in (None, 443)):
            raise Exception(f"Key identifier for certificate {self.certificate_name} is not in key vault {self.akv_url}.")

        response = requests.post(
            f"{key_id}/sign",
            headers=akv_headers,
            params={"api-version": self.AKV_API_VERSION},
            json={"alg": "PS256", "value": self._base64url(digest)},
            timeout=self.REQUEST_TIMEOUT_SECONDS
        )

        if response.status_code != 200:
            raise Exception(f"Failed to sign client assertion with key vault key {self.certificate_name}. {self._describe_error(response)}")

        return response.json()["value"]

    def _akv_headers(self):
        """
        Returns headers for calling Azure Key Vault as the identity running the notebook.

        notebookutils is imported here rather than at the top of the module, so the module (and Scopes)
        can be imported outside a Fabric notebook.
        """
        try:
            import notebookutils
        except ImportError as e:
            raise ImportError("notebookutils is not available. KeyVaultCertificateCredential must be used in a Fabric notebook.") from e

        akv_cred = notebookutils.credentials.getToken('keyvault')

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

        # Only summarise the documented error formats. Leave out anything else, such as a JSON list or string
        if not isinstance(body, dict) or body.get("error") is None:
            return f"Status code: {response.status_code}"

        error = body.get("error")

        # Key Vault nests the error as an object, Entra ID returns error and error_description at the top level
        if isinstance(error, dict):
            return f"Status code: {response.status_code}, {error.get('code')}: {error.get('message')}"

        return f"Status code: {response.status_code}, {error}: {body.get('error_description')}"
