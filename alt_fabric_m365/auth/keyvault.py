"""
KeyVaultCertificateCredential: app-only access tokens using a certificate held in Azure Key Vault.

See the alt_fabric_m365.auth package docstring for how scopes, notebookutils tokens and client secrets fit in.
"""

import requests
import json
import base64
import email.utils
import hashlib
import math
import random
import re
import threading
import time
import uuid
from collections.abc import Callable
from datetime import datetime, timezone
from urllib.parse import urlparse
from azure.core.credentials import AccessToken


def _base64url(data: bytes) -> str:
    """Encodes bytes as unpadded base64url, as used in JWTs."""
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _path(value: str) -> str:
    """Quotes a value for one segment of a URL path, so a / or ? in it can't select a different resource."""
    return requests.utils.quote(value, safe="")


def _retry_after_seconds(value: str | None) -> float | None:
    """Returns the wait a Retry-After header asks for, in seconds, or None if it's missing or can't be read.

    Args:
        value: The header's value: a number of seconds, or an HTTP date to wait until.

    Returns:
        The wait, at least 0, or None.
    """
    if not value:
        return None
    try:
        delay = float(value)
    except ValueError:
        try:
            retry_at = email.utils.parsedate_to_datetime(value)
        except (TypeError, ValueError):
            return None
        # A date without a zone, or with -0000, is read as naive; HTTP dates are GMT
        if retry_at.tzinfo is None:
            retry_at = retry_at.replace(tzinfo=timezone.utc)
        delay = (retry_at - datetime.now(timezone.utc)).total_seconds()
    # float() also reads "nan" and "inf", which time.sleep can't take
    return max(delay, 0.0) if math.isfinite(delay) else None


def _describe_error(response: requests.Response) -> str:
    """Summarises an error response from Microsoft Entra ID or Azure Key Vault without including request data."""
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


class KeyVaultCertificateCredential:
    """Authenticates as a service principal using a certificate held in Azure Key Vault.

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

    # Throttling, server errors and dropped connections are retried, as azure-identity's credentials and the rest of
    # the package do. Each request is safe to repeat. A token request is held up while it waits, so the waits are short.
    MAX_ATTEMPTS = 3
    RETRY_BASE_DELAY_SECONDS = 1
    RETRY_MAX_DELAY_SECONDS = 30  # Also caps a Retry-After header
    RETRY_STATUS_CODES = frozenset({429, 500, 502, 503, 504})

    # Key Vault DNS suffix in the Azure public cloud. The Key Vault token is only sent to these hosts. Sovereign clouds
    # aren't supported, as the token endpoint, Scopes and Microsoft Graph URL are all public cloud.
    AKV_DNS_SUFFIX = ".vault.azure.net"

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        akv_url: str,
        certificate_name: str,
        scope: str | None = None,
        certificate_version: str | None = None,
    ):
        """Initialises the credential. No calls are made until an access token is requested.

        Args:
            tenant_id: Tenant ID for Azure authentication.
            client_id: Client ID of the app registration.
            akv_url: Azure Key Vault URL, in the form https://{key_vault_name}.vault.azure.net/.
                Must use https and a Key Vault host name, as the Key Vault token is sent to it.
            certificate_name: Name of the certificate in the Azure Key Vault.
            scope: Optional. Default scope for get_access_token(), in the form {resource}/.default. See Scopes.
                Not needed when the credential is passed to semantic link or an Azure SDK client.
            certificate_version: Optional. Version of the certificate to sign with. Defaults to the latest version.

        Raises:
            ValueError: If akv_url isn't an https URL on a Key Vault host in the Azure public cloud.
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.akv_url = akv_url.rstrip("/")
        self._akv_host = self._validate_akv_url(self.akv_url)
        self.certificate_name = certificate_name
        self.scope = scope
        self.certificate_version = certificate_version
        self.token_endpoint = f"https://login.microsoftonline.com/{_path(self.tenant_id)}/oauth2/v2.0/token"

        # Cached AccessTokens, keyed by the sorted, de-duplicated scopes. The lock stops concurrent callers
        # each requesting a new token when the cached one expires.
        self._tokens = {}
        self._tokens_lock = threading.Lock()

    @classmethod
    def _validate_akv_url(cls, akv_url: str) -> str:
        """Checks the Key Vault URL uses https and a public cloud Key Vault host, as the Key Vault token is sent to it.

        Args:
            akv_url: Azure Key Vault URL.

        Returns:
            The Key Vault host name, in lower case.
        """
        parsed = urlparse(akv_url)
        host = (parsed.hostname or "").lower()

        # Match the whole host, not just its suffix: Key Vault names are letters, digits and hyphens. urlparse keeps
        # a backslash in the host but requests ends the host there, so https://evil.com\.vault.azure.net/ ends with the
        # suffix yet is requested from evil.com.
        is_key_vault_host = re.fullmatch(r"[a-z0-9-]+" + re.escape(cls.AKV_DNS_SUFFIX), host) is not None

        if parsed.scheme != "https" or parsed.username or parsed.password or not is_key_vault_host:
            raise ValueError(
                f"akv_url must be a Key Vault URL such as https://{{key_vault_name}}.vault.azure.net/, not {akv_url}. "
                "Only the Azure public cloud is supported."
            )

        return host

    @classmethod
    def _send_with_retry(cls, send: Callable[[], requests.Response]) -> requests.Response:
        """Sends a request, retrying throttling, server errors and dropped connections up to MAX_ATTEMPTS times.

        Args:
            send: Makes one attempt of the request.

        Returns:
            The response to the last attempt, which may be an error for the caller to report.

        Raises:
            requests.ConnectionError: If the last attempt couldn't connect, or a certificate didn't verify.
            requests.Timeout: If the last attempt timed out.
            requests.exceptions.ChunkedEncodingError: If the last attempt's response was cut off.
        """
        for attempt in range(1, cls.MAX_ATTEMPTS + 1):
            delay = min(cls.RETRY_BASE_DELAY_SECONDS * 2 ** (attempt - 1), cls.RETRY_MAX_DELAY_SECONDS)
            delay += delay * random.uniform(0, 0.25)
            try:
                response = send()
            # A connection dropped while the response is read raises ChunkedEncodingError
            except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError) as e:
                # A certificate that doesn't verify won't on a retry either
                if isinstance(e, requests.exceptions.SSLError) or attempt == cls.MAX_ATTEMPTS:
                    raise
            else:
                if response.status_code not in cls.RETRY_STATUS_CODES or attempt == cls.MAX_ATTEMPTS:
                    return response
                retry_after = _retry_after_seconds(response.headers.get("Retry-After"))
                if retry_after is not None:
                    delay = min(retry_after, cls.RETRY_MAX_DELAY_SECONDS)
            time.sleep(delay)

        raise AssertionError("unreachable: the last attempt always returns or raises")

    def get_access_token(self, scope: str | None = None) -> str:
        """Returns an access token string, for calling an API directly.

        The token is cached and reused until it is within 5 minutes of expiring, so this can be
        called before every request.

        Args:
            scope: Optional. Scope to request the access token for. Defaults to the scope the credential was
                created with.

        Returns:
            An access token.

        Raises:
            ValueError: If no scope is given here or when creating the credential.
            Exception: If Key Vault or Entra ID refuses a request for the token, with its status code and error,
                or Key Vault returns a key identifier outside akv_url's key vault.
            requests.RequestException: If Key Vault or Entra ID can't be reached, after retries.
            ImportError: Outside a Fabric notebook, as notebookutils is needed to call Key Vault.
        """
        scope = scope or self.scope
        if scope is None:
            raise ValueError("No scope given. Pass one here or when creating the credential, see Scopes.")

        return self.get_token(scope).token

    def get_token(
        self,
        *scopes: str,
        claims: str | None = None,
        tenant_id: str | None = None,
        enable_cae: bool = False,
        **kwargs,
    ) -> AccessToken:
        """Returns an access token for the given scopes. Implements azure.core.credentials.TokenCredential.

        Semantic link (sempy) functions and Azure SDK clients call this when the credential is passed to
        them. Tokens are cached per set of scopes, in any order, and reused until they are within 5 minutes
        of expiring.

        Args:
            *scopes: Scopes to request the access token for, in the form {resource}/.default.
            claims: Optional. Additional claims required in the token, such as from an API's claims challenge.
                A new token is always requested when claims are given, and it replaces the cached one.
            tenant_id: Optional. Tenant to request the token from. Only the tenant the credential was created
                with is supported, so a ValueError is raised for any other tenant.
            enable_cae: Optional. Accepted for compatibility but ignored, as continuous access evaluation
                isn't requested.
            **kwargs: Other keyword arguments are accepted for compatibility but ignored.

        Returns:
            The access token and its expiry time (seconds since epoch).

        Raises:
            ValueError: If no scopes are given, or tenant_id isn't the credential's tenant.
            Exception: If Key Vault or Entra ID refuses a request for the token, with its status code and error,
                or Key Vault returns a key identifier outside akv_url's key vault.
            requests.RequestException: If Key Vault or Entra ID can't be reached, after retries.
            ImportError: Outside a Fabric notebook, as notebookutils is needed to call Key Vault.
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

    def _is_valid(self, token: AccessToken | None) -> bool:
        """Returns whether a cached token exists and has more than REFRESH_MARGIN_SECONDS left before it expires."""
        return token is not None and time.time() < token.expires_on - self.REFRESH_MARGIN_SECONDS

    def _request_access_token(self, scope: str, claims: str | None = None) -> AccessToken:
        """Exchanges a newly signed client assertion for an access token.

        Args:
            scope: Scope to request the access token for. Multiple scopes are separated by spaces.
            claims: Optional. Additional claims required in the token, as a JSON string.

        Returns:
            The access token and its expiry time (seconds since epoch).
        """
        # Build a new assertion each time, as they're only valid for a few minutes. Retries resend it, as MSAL does:
        # they finish well within its lifetime.
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

        now = 0  # When the last attempt was sent, set by send()

        def send() -> requests.Response:
            nonlocal now
            # Take the time before the request, so the expiry time errs on the early side
            now = int(time.time())
            return requests.post(self.token_endpoint, data=data, timeout=self.REQUEST_TIMEOUT_SECONDS)

        response = self._send_with_retry(send)

        if response.status_code != 200:
            raise Exception(f"Failed to get access token for {self.client_id} with scope {scope}. {_describe_error(response)}")

        token_response = response.json()

        return AccessToken(token_response["access_token"], now + int(token_response["expires_in"]))

    def _build_client_assertion(self) -> str:
        """Builds a client assertion (JWT) and has Key Vault sign it with the certificate's private key.

        Follows https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials

        Returns:
            The signed client assertion.
        """
        # Get the Key Vault token once, for both the certificate and sign requests
        akv_headers = self._akv_headers()

        key_id, certificate_der = self._get_certificate(akv_headers)
        now = int(time.time())

        header = {
            "alg": "PS256",
            "typ": "JWT",
            "x5t#S256": _base64url(hashlib.sha256(certificate_der).digest())
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
        signing_input = f"{_base64url(json.dumps(header).encode())}.{_base64url(json.dumps(claims).encode())}"
        digest = hashlib.sha256(signing_input.encode("ascii")).digest()

        signature = self._sign_digest(key_id, digest, akv_headers)

        return f"{signing_input}.{signature}"

    def _get_certificate(self, akv_headers: dict) -> tuple[str, bytes]:
        """Retrieves the public part of the certificate from Azure Key Vault.

        Args:
            akv_headers: Headers for calling Azure Key Vault, from _akv_headers().

        Returns:
            The key identifier (URL including version) to sign with, and the certificate's DER encoded bytes.
        """
        url = f"{self.akv_url}/certificates/{_path(self.certificate_name)}"
        if self.certificate_version:
            url += f"/{_path(self.certificate_version)}"

        response = self._send_with_retry(lambda: requests.get(
            url,
            headers=akv_headers,
            params={"api-version": self.AKV_API_VERSION},
            timeout=self.REQUEST_TIMEOUT_SECONDS
        ))

        if response.status_code != 200:
            raise Exception(f"Failed to get certificate {self.certificate_name} from key vault. {_describe_error(response)}")

        certificate = response.json()

        return certificate["kid"], base64.b64decode(certificate["cer"])

    def _sign_digest(self, key_id: str, digest: bytes, akv_headers: dict) -> str:
        """Signs a SHA-256 digest with the certificate's private key using the Azure Key Vault sign operation.

        Args:
            key_id: The key identifier returned with the certificate.
            digest: The SHA-256 digest to sign.
            akv_headers: Headers for calling Azure Key Vault, from _akv_headers().

        Returns:
            The base64url encoded signature.
        """
        # The key identifier comes from the certificate response, so check it's in the same key vault
        # before sending the Key Vault token to it
        parsed_key_id = urlparse(key_id)
        if (parsed_key_id.scheme != "https" or parsed_key_id.username or parsed_key_id.password
                or (parsed_key_id.hostname or "").lower() != self._akv_host or parsed_key_id.port not in (None, 443)):
            raise Exception(f"Key identifier for certificate {self.certificate_name} is not in key vault {self.akv_url}.")

        response = self._send_with_retry(lambda: requests.post(
            f"{key_id}/sign",
            headers=akv_headers,
            params={"api-version": self.AKV_API_VERSION},
            json={"alg": "PS256", "value": _base64url(digest)},
            timeout=self.REQUEST_TIMEOUT_SECONDS
        ))

        if response.status_code != 200:
            raise Exception(f"Failed to sign client assertion with key vault key {self.certificate_name}. {_describe_error(response)}")

        return response.json()["value"]

    def _akv_headers(self) -> dict:
        """Returns headers for calling Azure Key Vault as the identity running the notebook.

        notebookutils is imported here rather than at the top of the module, so the package (and Scopes)
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
