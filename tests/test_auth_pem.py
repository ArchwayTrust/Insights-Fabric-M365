"""
PemCertificateCredential: MSAL is replaced with a fake, and the private key normalisation is checked against a
real RSA key.
"""
import sys
import time
import types

import pytest
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa

from fabric_python_helper.auth import PemCertificateCredential, Scopes

TENANT = "tenant-a"
CLIENT = "client-id"


class FakeMsalApp:
    instances = []

    def __init__(self, client_id, authority, client_credential):
        self.client_id = client_id
        self.authority = authority
        self.client_credential = client_credential
        self.requests = []
        self.result = {"access_token": "token-1", "expires_in": 3600}
        FakeMsalApp.instances.append(self)

    def acquire_token_for_client(self, scopes, claims_challenge=None):
        self.requests.append((scopes, claims_challenge))
        return self.result


@pytest.fixture
def fake_msal(monkeypatch):
    FakeMsalApp.instances = []
    monkeypatch.setitem(sys.modules, "msal", types.SimpleNamespace(ConfidentialClientApplication=FakeMsalApp))
    return FakeMsalApp


@pytest.fixture(scope="module")
def pem_key():
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    return key.private_bytes(
        serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption(),
    ).decode()


def loads(pem):
    return serialization.load_pem_private_key(pem.encode(), password=None) is not None


def test_get_token_uses_the_certificate_and_returns_an_access_token(fake_msal, pem_key):
    credential = PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key)
    before = time.time()

    token = credential.get_token(Scopes.GRAPH)

    [app] = fake_msal.instances
    assert app.authority == f"https://login.microsoftonline.com/{TENANT}"
    assert app.client_credential["thumbprint"] == "ABC123"
    assert app.requests == [([Scopes.GRAPH], None)]
    assert token.token == "token-1"
    assert before + 3599 <= token.expires_on <= time.time() + 3600
    assert credential.get_access_token(Scopes.GRAPH) == "token-1"


def test_claims_are_passed_to_msal(fake_msal, pem_key):
    credential = PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key)

    credential.get_token("https://contoso.sharepoint.com/.default", claims='{"access_token":{}}')

    assert fake_msal.instances[0].requests == [(["https://contoso.sharepoint.com/.default"], '{"access_token":{}}')]


def test_failure_includes_the_entra_error(fake_msal, pem_key):
    credential = PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key)
    fake_msal.instances[0].result = {"error": "invalid_client", "error_description": "AADSTS700027: not registered"}

    with pytest.raises(Exception, match="invalid_client: AADSTS700027"):
        credential.get_token(Scopes.GRAPH)


def test_scope_and_tenant_are_checked(fake_msal, pem_key):
    credential = PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key)

    with pytest.raises(ValueError, match="At least one scope"):
        credential.get_token()
    with pytest.raises(ValueError, match="tenant-b"):
        credential.get_token(Scopes.GRAPH, tenant_id="tenant-b")
    credential.get_token(Scopes.GRAPH, tenant_id="TENANT-A")


@pytest.mark.parametrize("thumbprint, private_key", [("", "key"), ("ABC", ""), (None, "key")])
def test_thumbprint_and_key_are_required(fake_msal, thumbprint, private_key):
    with pytest.raises(ValueError, match="both required"):
        PemCertificateCredential(TENANT, CLIENT, thumbprint, private_key)


def test_missing_msal_is_explained(monkeypatch, pem_key):
    monkeypatch.setitem(sys.modules, "msal", None)

    with pytest.raises(ImportError, match="needs msal"):
        PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key)


def test_private_key_is_passed_normalised(fake_msal, pem_key):
    PemCertificateCredential(TENANT, CLIENT, "ABC123", pem_key.replace("\n", ""))

    assert loads(fake_msal.instances[0].client_credential["private_key"])


# Private key normalisation

def test_a_well_formed_key_is_unchanged(pem_key):
    assert PemCertificateCredential._normalise_private_key(pem_key) == pem_key


def test_a_key_with_its_newlines_stripped_is_rewrapped(pem_key):
    flattened = pem_key.replace("\n", "")

    normalised = PemCertificateCredential._normalise_private_key(flattened)

    assert loads(normalised)
    assert all(len(line) <= 64 for line in normalised.splitlines())


def test_a_key_with_escaped_newlines_is_restored(pem_key):
    assert loads(PemCertificateCredential._normalise_private_key(pem_key.replace("\n", "\\n")))


def test_text_before_the_pem_header_is_dropped(pem_key):
    exported = "Bag Attributes\n    localKeyID: 01 00 00 00\nKey Attributes: <No Attributes>\n" + pem_key

    assert PemCertificateCredential._normalise_private_key(exported) == pem_key
