"""
Shared fixtures. No test calls Microsoft Entra ID, Key Vault or Graph: requests is mocked and notebookutils,
which only exists in Fabric notebooks, is replaced with a fake.
"""
import base64
import json
import sys
import threading
import time
import types

import pytest
import requests


@pytest.fixture
def fake_notebookutils(monkeypatch):
    """
    Installs a fake notebookutils module. Returns the list of audiences getToken was called with.
    """
    audiences = []

    def get_token(audience):
        audiences.append(audience)
        return "kv-token"

    module = types.ModuleType("notebookutils")
    module.credentials = types.SimpleNamespace(getToken=get_token)
    monkeypatch.setitem(sys.modules, "notebookutils", module)

    return audiences


@pytest.fixture
def no_notebookutils(monkeypatch):
    """
    Makes sure notebookutils can't be imported, as outside Fabric.
    """
    monkeypatch.setitem(sys.modules, "notebookutils", None)


class FakeCredential:
    """
    A credential with a get_token method, standing in for KeyVaultCertificateCredential or ClientSecretCredential.
    """

    def __init__(self):
        self.scopes = []

    def get_token(self, *scopes, **kwargs):
        self.scopes.append(scopes)
        return types.SimpleNamespace(token="graph-token", expires_on=0)


@pytest.fixture
def credential():
    return FakeCredential()


def make_response(status_code, body=None, text=None):
    """
    Builds a real requests.Response, so raise_for_status, ok and json behave as they do against Graph.
    """
    response = requests.Response()
    response.status_code = status_code
    response._content = json.dumps(body).encode() if body is not None else (text or "").encode()
    response.url = "https://example.invalid/"
    return response


KID = "https://myvault.vault.azure.net/keys/cert/abc123"
CERT_DER = b"certificate-der-bytes"


def _key_vault_error(kwargs):
    """Returns the error Key Vault gives a request without its api-version or the notebook's token, or None."""
    if (kwargs.get("params") or {}).get("api-version") is None:
        return make_response(400, {"error": {"code": "BadParameter", "message": "api-version is required."}})
    scheme, _, token = (kwargs.get("headers") or {}).get("Authorization", "").partition(" ")
    if scheme != "Bearer" or not token:
        return make_response(401, {"error": {"code": "Unauthorized", "message": "AKV10000: Request is missing a Bearer"}})
    return None


class FakeAzure:
    """
    Stands in for Key Vault (certificate GET and sign POST) and the Entra ID token endpoint, recording each call.

    Like the real services, Key Vault refuses a request without an api-version or the notebook's Key Vault token, and
    Entra ID one without a client_id. certificate_status, sign_status and token_status (with their bodies) make a
    call fail.
    """

    def __init__(self, kid=KID, expires_in=3600, token_status=200, token_body=None, delay=0, object_id=None,
                 certificate_status=200, certificate_body=None, sign_status=200, sign_body=None):
        self.kid = kid
        # With object_id, access tokens are JWT-shaped with that oid claim, as Entra ID issues them
        self.object_id = object_id
        self.issued = []
        self.expires_in = expires_in
        self.token_status = token_status
        self.token_body = token_body
        self.certificate_status = certificate_status
        self.certificate_body = certificate_body
        self.sign_status = sign_status
        self.sign_body = sign_body
        self.delay = delay
        self.certificate_calls = []
        self.sign_calls = []
        self.token_calls = []
        self._lock = threading.Lock()

    def get(self, url, **kwargs):
        with self._lock:
            self.certificate_calls.append((url, kwargs))
        error = _key_vault_error(kwargs)
        if error is not None:
            return error
        if self.certificate_status != 200:
            return make_response(self.certificate_status, self.certificate_body)
        return make_response(200, {"kid": self.kid, "cer": base64.b64encode(CERT_DER).decode()})

    def post(self, url, **kwargs):
        if url.endswith("/sign"):
            with self._lock:
                self.sign_calls.append((url, kwargs))
            error = _key_vault_error(kwargs)
            if error is not None:
                return error
            if self.sign_status != 200:
                return make_response(self.sign_status, self.sign_body)
            return make_response(200, {"value": "c2lnbmF0dXJl"})

        if self.delay:
            time.sleep(self.delay)
        with self._lock:
            self.token_calls.append((url, kwargs))
            count = len(self.token_calls)
        if not (kwargs.get("data") or {}).get("client_id"):
            return make_response(400, {
                "error": "invalid_request",
                "error_description": "AADSTS900144: The request body must contain the following parameter: 'client_id'.",
            })
        if self.token_status != 200:
            return make_response(self.token_status, self.token_body)
        token = f"token-{count}"
        if self.object_id:
            claims = base64.urlsafe_b64encode(json.dumps({"oid": self.object_id, "n": count}).encode()).decode().rstrip("=")
            token = f"eyJhbGciOiJub25lIn0.{claims}.sig"
        with self._lock:
            self.issued.append(token)
        return make_response(200, {"access_token": token, "expires_in": self.expires_in})
