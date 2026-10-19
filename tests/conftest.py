"""
Shared fixtures. No test calls Microsoft Entra ID, Key Vault or Graph: requests is mocked and notebookutils,
which only exists in Fabric notebooks, is replaced with a fake.
"""
import json
import sys
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
