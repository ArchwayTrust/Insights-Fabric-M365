"""
The retry wrapper every Graph and SharePoint call goes through: throttling, server errors, connection errors
and redirects.
"""
import time
from datetime import datetime, timedelta, timezone
from email.utils import format_datetime
from unittest import mock

import pytest
import requests

from alt_fabric_m365.sharepoint import _http
from alt_fabric_m365.sharepoint._http import _parse_retry_after, _request_with_retry

from sharepoint_fake import response

URL = "https://contoso.sharepoint.com/sites/maths/_api/web/ensureuser"


@pytest.fixture
def sleeps(monkeypatch):
    calls = []
    monkeypatch.setattr(time, "sleep", calls.append)
    return calls


def patch_request(return_value=None, side_effect=None):
    """Patches requests.request to give return_value, or each of side_effect in turn (raising exceptions).

    Each response gets the URL it was requested from, with params encoded, as requests gives it.
    """
    results = iter(side_effect) if isinstance(side_effect, list) else None

    def request(method, url, **kwargs):
        result = next(results) if results is not None else side_effect if side_effect is not None else return_value
        if isinstance(result, BaseException):
            raise result
        if isinstance(result, requests.Response):
            result.url = requests.Request(method, url, params=kwargs.get("params")).prepare().url
        return result
    return mock.patch.object(_http.requests, "request", side_effect=request)


# Retry-After

@pytest.mark.parametrize("value, expected", [
    (None, None),
    ("120", 120.0),
    ("0.5", 0.5),
    ("-5", 0.0),
    ("100000", _http.RETRY_AFTER_MAX_DELAY),
    ("NaN", None),
    ("inf", None),
    ("soon", None),
])
def test_parse_retry_after_seconds(value, expected):
    assert _parse_retry_after(value) == expected


def test_parse_retry_after_http_date():
    retry_at = datetime.now(timezone.utc) + timedelta(seconds=90)

    delay = _parse_retry_after(format_datetime(retry_at, usegmt=True))

    assert 85 <= delay <= 90


def test_parse_retry_after_date_without_zone_is_treated_as_utc():
    retry_at = datetime.now(timezone.utc) + timedelta(seconds=60)

    delay = _parse_retry_after(retry_at.strftime("%a, %d %b %Y %H:%M:%S -0000"))

    assert 55 <= delay <= 60


def test_parse_retry_after_date_in_the_past_is_zero():
    assert _parse_retry_after("Wed, 21 Oct 2015 07:28:00 GMT") == 0.0


# Retries

def test_success_returns_immediately_with_default_timeout(sleeps):
    with patch_request(return_value=response(200, {})) as request:
        resp = _request_with_retry("GET", URL)

    assert resp.status_code == 200
    assert request.call_count == 1
    assert request.call_args.kwargs["timeout"] == _http.DEFAULT_TIMEOUT
    assert request.call_args.kwargs["allow_redirects"] is False
    assert sleeps == []


def test_client_errors_are_returned_without_retrying(sleeps):
    with patch_request(return_value=response(404)) as request:
        assert _request_with_retry("GET", URL).status_code == 404

    assert request.call_count == 1


# SharePoint and Graph throttle with 503 as well as 429, both with Retry-After
@pytest.mark.parametrize("status", [429, 503])
def test_throttling_honours_retry_after(sleeps, status):
    responses = [response(status, headers={"Retry-After": "7"}), response(200, {})]
    with patch_request(side_effect=responses):
        assert _request_with_retry("GET", URL).status_code == 200

    assert sleeps == [7.0]


def test_throttling_without_retry_after_backs_off_exponentially(sleeps):
    responses = [response(429), response(429), response(200, {})]
    with patch_request(side_effect=responses):
        _request_with_retry("GET", URL, base_delay=2)

    assert sleeps == [2.0, 4.0]


def test_server_errors_retry_with_jitter_then_return_the_last_response(sleeps):
    with patch_request(return_value=response(503)) as request:
        resp = _request_with_retry("GET", URL, max_retries=3, base_delay=1, max_delay=3)

    assert resp.status_code == 503
    assert request.call_count == 4
    # Backoff 1, 2, 3 (capped), each plus up to 25% jitter
    assert [1 <= sleeps[0] <= 1.25, 2 <= sleeps[1] <= 2.5, 3 <= sleeps[2] <= 3.75] == [True, True, True]


def test_connection_errors_retry_then_raise(sleeps):
    with patch_request(side_effect=requests.ConnectionError("reset")) as request:
        with pytest.raises(requests.ConnectionError):
            _request_with_retry("GET", URL, max_retries=2)

    assert request.call_count == 3
    assert len(sleeps) == 2


def test_timeout_then_success(sleeps):
    with patch_request(side_effect=[requests.Timeout("slow"), response(200, {})]):
        assert _request_with_retry("GET", URL).status_code == 200


# Redirects

def test_post_redirect_is_followed_with_the_same_method(sleeps):
    responses = [
        response(302, headers={"Location": "/sites/maths2/_api/web/ensureuser"}),
        response(200, {}),
    ]
    with patch_request(side_effect=responses) as request:
        _request_with_retry("POST", URL, json={"logonName": "x"})

    second = request.call_args_list[1]
    assert second.args == ("POST", "https://contoso.sharepoint.com/sites/maths2/_api/web/ensureuser")
    assert second.kwargs["json"] == {"logonName": "x"}
    assert second.kwargs["allow_redirects"] is False


def test_a_redirect_is_followed_without_sending_the_query_again(sleeps):
    location = "https://graph.microsoft.com/v1.0/groups/?%24filter=displayName+eq+%27x%27"
    responses = [response(301, headers={"Location": location}), response(503), response(200, {})]
    with patch_request(side_effect=responses) as request:
        _request_with_retry("GET", "https://graph.microsoft.com/v1.0/groups", params={"$filter": "displayName eq 'x'"})

    first, after_redirect, retry = request.call_args_list
    assert first.kwargs["params"] == {"$filter": "displayName eq 'x'"}
    # The redirect, and the retry after it, ask for Location as it is
    assert [(call.args[1], "params" in call.kwargs) for call in (after_redirect, retry)] == [(location, False)] * 2


def test_a_fragment_only_redirect_keeps_the_query(sleeps):
    responses = [response(302, headers={"Location": "#next"}), response(200, {})]
    with patch_request(side_effect=responses) as request:
        _request_with_retry("GET", "https://graph.microsoft.com/v1.0/groups", params={"$filter": "displayName eq 'x'"})

    after_redirect = request.call_args_list[1]
    # Resolved against the URL that was requested, query and all; requests doesn't send the fragment
    assert after_redirect.args[1] == "https://graph.microsoft.com/v1.0/groups?%24filter=displayName+eq+%27x%27#next"
    assert "params" not in after_redirect.kwargs


def test_cross_host_redirect_is_refused():
    with patch_request(return_value=response(307, headers={"Location": "https://evil.example.com/steal"})):
        with pytest.raises(requests.HTTPError, match="Refusing cross-host"):
            _request_with_retry("POST", URL)


def test_redirect_loop_stops_at_the_hop_limit():
    with patch_request(return_value=response(302, headers={"Location": URL})) as request:
        with pytest.raises(requests.HTTPError, match="Too many redirects"):
            _request_with_retry("PATCH", URL)

    assert request.call_count == _http.MAX_REDIRECT_HOPS + 1


@pytest.mark.parametrize("method", ["GET", "POST"])
def test_redirect_without_location_is_an_error(method):
    # raise_for_status() lets 3xx through, so returning it would look like success to a write helper
    with patch_request(return_value=response(302)):
        with pytest.raises(requests.HTTPError, match="Unfollowed redirect"):
            _request_with_retry(method, URL)


def test_redirect_from_https_to_http_is_refused():
    with patch_request(return_value=response(301, headers={"Location": "http://contoso.sharepoint.com/sites/x"})):
        with pytest.raises(requests.HTTPError, match="Refusing"):
            _request_with_retry("POST", URL)


@pytest.mark.parametrize("status", [301, 302, 303, 307, 308])
def test_get_redirect_on_the_same_host_is_followed(status):
    target = "https://contoso.sharepoint.com/sites/renamed/_api/web"
    with patch_request(side_effect=[response(status, headers={"Location": target}), response(200, {})]) as request:
        assert _request_with_retry("GET", URL).status_code == 200

    assert request.call_args_list[1].args == ("GET", target)


@pytest.mark.parametrize("location", [
    "https://evil.example.com/x",
    "http://contoso.sharepoint.com/sites/x",
    # urlparse reads this as on contoso.sharepoint.com, but requests sends it to evil.example.com
    "https://evil.example.com\\@contoso.sharepoint.com/sites/x",
])
def test_get_redirect_to_another_host_or_to_http_is_refused(location):
    with patch_request(return_value=response(302, headers={"Location": location})) as request:
        with pytest.raises(requests.HTTPError, match="Refusing"):
            _request_with_retry("GET", URL)

    assert request.call_count == 1


def test_303_is_not_followed_for_post_and_is_an_error():
    with patch_request(return_value=response(303, headers={"Location": "/sites/x"})) as request:
        with pytest.raises(requests.HTTPError, match="Unfollowed redirect"):
            _request_with_retry("POST", URL)

    assert request.call_count == 1


# Requests that must not run twice

def test_non_idempotent_request_is_not_repeated_after_a_server_error(sleeps):
    with patch_request(return_value=response(503)) as request:
        assert _request_with_retry("POST", URL, idempotent=False).status_code == 503

    assert request.call_count == 1


@pytest.mark.parametrize("error", [requests.ReadTimeout("slow"), requests.ConnectionError("reset")])
def test_non_idempotent_request_is_not_repeated_after_it_may_have_arrived(sleeps, error):
    with patch_request(side_effect=error) as request:
        with pytest.raises(type(error)):
            _request_with_retry("POST", URL, idempotent=False)

    assert request.call_count == 1


@pytest.mark.parametrize("first", [response(429, headers={"Retry-After": "1"}), requests.ConnectTimeout("no route")])
def test_non_idempotent_request_is_retried_when_it_certainly_did_not_arrive(sleeps, first):
    with patch_request(side_effect=[first, response(201, {})]) as request:
        assert _request_with_retry("POST", URL, idempotent=False).status_code == 201

    assert request.call_count == 2


# Tokens

def clock_token(sleeps):
    """A token that changes whenever time passes, as a credential's does when it renews one near expiry."""
    return _http._RenewableToken(lambda: f"token-{len(sleeps)}")


@pytest.mark.parametrize("first", [
    response(429, headers={"Retry-After": "300"}), response(503), requests.ConnectionError("reset"),
], ids=["throttled", "server error", "connection error"])
def test_each_attempt_sends_a_current_token(sleeps, first):
    headers = {"Authorization": _http._bearer(clock_token(sleeps)), "Accept": "application/json"}
    with patch_request(side_effect=[first, response(200, {})]) as request:
        _request_with_retry("GET", URL, headers=headers)

    sent = [call.kwargs["headers"] for call in request.call_args_list]
    assert [h["Authorization"] for h in sent] == ["Bearer token-0", "Bearer token-1"]
    assert all(h["Accept"] == "application/json" for h in sent)
    assert headers["Authorization"] == "Bearer token-0"  # the caller's headers are left alone


def test_a_plain_token_is_sent_as_given(sleeps):
    with patch_request(side_effect=[response(503), response(200, {})]) as request:
        _request_with_retry("GET", URL, headers={"Authorization": _http._bearer("abc")})

    assert [call.kwargs["headers"]["Authorization"] for call in request.call_args_list] == ["Bearer abc"] * 2
