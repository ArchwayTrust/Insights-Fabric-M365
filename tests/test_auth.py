import base64
import hashlib
import json
import threading
import time
from unittest import mock

import pytest

from alt_fabric_m365.auth import keyvault
from alt_fabric_m365.auth import KeyVaultCertificateCredential, Scopes

from conftest import CERT_DER, KID, FakeAzure, make_response

TENANT = "tenant-a"
CLIENT = "client-id"
AKV = "https://myvault.vault.azure.net/"
TIMEOUT = KeyVaultCertificateCredential.REQUEST_TIMEOUT_SECONDS
@pytest.fixture
def azure(fake_notebookutils):
    fake = FakeAzure()
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        yield fake


def make_credential(akv_url=AKV, **kwargs):
    return KeyVaultCertificateCredential(TENANT, CLIENT, akv_url, "cert", **kwargs)


def decode_segment(segment):
    return json.loads(base64.urlsafe_b64decode(segment + "=" * (-len(segment) % 4)))


# Importing and notebookutils

def test_imports_without_notebookutils(no_notebookutils):
    assert Scopes.GRAPH == "https://graph.microsoft.com/.default"
    make_credential()


def test_clear_error_when_used_outside_fabric(no_notebookutils):
    with pytest.raises(ImportError, match="Fabric notebook"):
        make_credential().get_token(Scopes.GRAPH)


# Caching

def test_token_is_cached(azure):
    credential = make_credential()

    first = credential.get_token("scope-a")
    second = credential.get_token("scope-a")

    assert first.token == "token-1"
    assert second is first
    assert len(azure.token_calls) == 1


def test_scope_order_and_duplicates_share_a_cached_token(azure):
    credential = make_credential()

    token = credential.get_token("b", "a")

    assert credential.get_token("a", "b") is token
    assert credential.get_token("a", "b", "a") is token
    assert len(azure.token_calls) == 1
    assert azure.token_calls[0][1]["data"]["scope"] == "a b"


def test_different_scopes_are_cached_separately(azure):
    credential = make_credential()

    assert credential.get_token("a").token == "token-1"
    assert credential.get_token("b").token == "token-2"
    assert credential.get_access_token("b") == "token-2"


def test_token_close_to_expiry_is_refreshed(fake_notebookutils):
    # Tokens with less than REFRESH_MARGIN_SECONDS left are treated as expired
    fake = FakeAzure(expires_in=KeyVaultCertificateCredential.REFRESH_MARGIN_SECONDS - 1)
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        credential = make_credential()
        credential.get_token("a")
        credential.get_token("a")

    assert len(fake.token_calls) == 2


def test_expiry_is_measured_from_before_the_request(fake_notebookutils):
    # A fake clock that only moves while the token request is in flight, so the test doesn't depend on timing
    clock = [1_000_000.0]
    fake = FakeAzure()

    def slow_token_post(url, **kwargs):
        response = fake.post(url, **kwargs)
        if not url.endswith("/sign"):
            clock[0] += 5
        return response

    with mock.patch.multiple(keyvault.requests, get=fake.get, post=slow_token_post), \
            mock.patch.object(keyvault.time, "time", lambda: clock[0]):
        token = make_credential().get_token("a")

    assert clock[0] == 1_000_005.0
    assert token.expires_on == 1_000_000 + 3600


def test_concurrent_callers_share_one_refresh(fake_notebookutils):
    fake = FakeAzure(delay=0.2)
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        credential = make_credential()
        barrier = threading.Barrier(10)
        results, errors = [], []

        def call(index):
            try:
                barrier.wait()
                scopes = ("a", "b") if index % 2 else ("b", "a")
                results.append(credential.get_token(*scopes))
            except Exception as e:
                errors.append(e)

        threads = [threading.Thread(target=call, args=(i,)) for i in range(10)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()

    assert not errors
    assert len(fake.token_calls) == 1
    assert len(results) == 10 and len({id(result) for result in results}) == 1


def test_no_scope_raises():
    with pytest.raises(ValueError):
        make_credential().get_token()
    with pytest.raises(ValueError):
        make_credential().get_access_token()


# get_token keyword arguments

def test_claims_bypass_the_cache_and_are_sent(azure):
    credential = make_credential()
    claims = '{"access_token":{"nbf":{"essential":true,"value":"1"}}}'

    credential.get_token("a")
    fresh = credential.get_token("a", claims=claims)

    assert fresh.token == "token-2"
    assert azure.token_calls[1][1]["data"]["claims"] == claims
    assert "claims" not in azure.token_calls[0][1]["data"]
    # The fresh token replaces the cached one
    assert credential.get_token("a") is fresh


def test_other_tenant_raises_without_a_request(azure):
    with pytest.raises(ValueError, match="tenant"):
        make_credential().get_token("a", tenant_id="tenant-b")

    assert azure.token_calls == []


def test_same_tenant_cae_and_unknown_kwargs_are_accepted(azure):
    credential = make_credential()

    assert credential.get_token("a", tenant_id=TENANT.upper()).token == "token-1"
    assert credential.get_token("a", enable_cae=True, something_new=1).token == "token-1"


# Requests to Key Vault and Entra ID

def test_every_request_has_a_timeout(azure):
    make_credential().get_token("a")

    for _, kwargs in azure.certificate_calls + azure.sign_calls + azure.token_calls:
        assert kwargs["timeout"] == TIMEOUT


def test_key_vault_token_is_fetched_once_per_refresh(azure, fake_notebookutils):
    credential = make_credential()

    credential.get_token("a")
    assert fake_notebookutils == ["keyvault"]

    credential.get_token("b")
    assert fake_notebookutils == ["keyvault", "keyvault"]


def test_requests_go_to_the_expected_urls(azure):
    make_credential(certificate_version="v2").get_token("a")

    assert azure.certificate_calls[0][0] == "https://myvault.vault.azure.net/certificates/cert/v2"
    assert azure.certificate_calls[0][1]["headers"]["Authorization"] == "Bearer kv-token"
    assert azure.sign_calls[0][0] == f"{KID}/sign"
    assert azure.token_calls[0][0] == f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"


def test_client_assertion_is_well_formed(azure):
    before = int(time.time())
    make_credential().get_token("a")

    data = azure.token_calls[0][1]["data"]
    assert data["grant_type"] == "client_credentials"
    assert data["client_assertion_type"] == "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"

    header_segment, claims_segment, signature = data["client_assertion"].split(".")
    header = decode_segment(header_segment)
    claims = decode_segment(claims_segment)

    assert header["alg"] == "PS256"
    expected_thumbprint = base64.urlsafe_b64encode(hashlib.sha256(CERT_DER).digest()).rstrip(b"=").decode()
    assert header["x5t#S256"] == expected_thumbprint

    assert claims["aud"] == f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    assert claims["iss"] == claims["sub"] == CLIENT
    assert claims["jti"]
    assert before <= claims["iat"] == claims["nbf"]
    assert claims["exp"] == claims["iat"] + KeyVaultCertificateCredential.ASSERTION_LIFETIME_SECONDS

    # Key Vault was asked to sign the SHA-256 digest of the header and claims, and its signature was used
    sign_body = azure.sign_calls[0][1]["json"]
    digest = hashlib.sha256(f"{header_segment}.{claims_segment}".encode("ascii")).digest()
    assert sign_body["alg"] == "PS256"
    assert sign_body["value"] == base64.urlsafe_b64encode(digest).rstrip(b"=").decode()
    assert signature == "c2lnbmF0dXJl"


def test_each_assertion_is_unique(azure):
    credential = make_credential()
    credential.get_token("a")
    credential.get_token("b")

    jtis = {decode_segment(call[1]["data"]["client_assertion"].split(".")[1])["jti"] for call in azure.token_calls}
    assert len(jtis) == 2


def test_token_error_includes_status_but_not_the_assertion(fake_notebookutils):
    fake = FakeAzure(token_status=401, token_body={"error": "invalid_client", "error_description": "AADSTS700027"})
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        with pytest.raises(Exception) as error:
            make_credential().get_token("a")

    message = str(error.value)
    assert "Status code: 401" in message and "AADSTS700027" in message
    assert fake.token_calls[0][1]["data"]["client_assertion"] not in message


# Key Vault URL checks

@pytest.mark.parametrize("akv_url", [
    "https://myvault.vault.azure.net",
    "https://myvault.vault.azure.net/",
    "https://MyVault.Vault.Azure.Net/",
])
def test_key_vault_urls_are_accepted(akv_url):
    make_credential(akv_url)


@pytest.mark.parametrize("akv_url", [
    "http://myvault.vault.azure.net/",
    "https://myvault.vault.azure.net.evil.com/",
    "https://evilvault.azure.net/",
    "https://vault.azure.net/",
    "https://evil.com/",
    "https://myvault.vault.azure.net@evil.com/",
    "https://user:password@myvault.vault.azure.net/",
    # requests ends the host at the backslash, so sends these to evil.com
    "https://evil.com\\.vault.azure.net/",
    "https://evil.com\\@myvault.vault.azure.net/",
    # Sovereign clouds aren't supported, as tokens are always requested from the public cloud
    "https://myvault.vault.azure.cn/",
    "https://myvault.vault.usgovcloudapi.net/",
    "myvault.vault.azure.net",
    "",
])
def test_other_urls_are_rejected(akv_url):
    with pytest.raises(ValueError):
        make_credential(akv_url)


@pytest.mark.parametrize("kid", [
    "https://evil.com/keys/cert/1",
    "http://myvault.vault.azure.net/keys/cert/1",
    "https://othervault.vault.azure.net/keys/cert/1",
    "https://myvault.vault.azure.net:8443/keys/cert/1",
    "https://myvault.vault.azure.net@evil.com/keys/cert/1",
    "not a url",
])
def test_key_vault_token_is_not_sent_to_another_host(fake_notebookutils, kid):
    fake = FakeAzure(kid=kid)
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        with pytest.raises(Exception, match="not in key vault"):
            make_credential().get_token("a")

    assert fake.sign_calls == [] and fake.token_calls == []


@pytest.mark.parametrize("kid", [
    "https://MYVAULT.vault.azure.net/keys/cert/1",
    "https://myvault.vault.azure.net:443/keys/cert/1",
])
def test_equivalent_key_ids_are_accepted(fake_notebookutils, kid):
    fake = FakeAzure(kid=kid)
    with mock.patch.multiple(keyvault.requests, get=fake.get, post=fake.post):
        make_credential().get_token("a")

    assert len(fake.sign_calls) == 1


# Error descriptions

@pytest.mark.parametrize("response, expected", [
    (make_response(400, {"error": "invalid_client", "error_description": "bad"}), "Status code: 400, invalid_client: bad"),
    (make_response(403, {"error": {"code": "Forbidden", "message": "no"}}), "Status code: 403, Forbidden: no"),
    (make_response(500, ["a", "b"]), "Status code: 500"),
    (make_response(500, "oops"), "Status code: 500"),
    (make_response(500, {"message": "no error key"}), "Status code: 500"),
    (make_response(502, text="<html>Bad Gateway</html>"), "Status code: 502"),
])
def test_describe_error(response, expected):
    assert keyvault._describe_error(response) == expected


# Encoding

@pytest.mark.parametrize("data, expected", [
    (b"", ""),
    (b"\xfb\xff", "-_8"),  # standard base64 would be "+/8=": base64url swaps + and /, and drops the padding
    (b"abcd", "YWJjZA"),
])
def test_base64url_is_unpadded_and_url_safe(data, expected):
    assert keyvault._base64url(data) == expected
