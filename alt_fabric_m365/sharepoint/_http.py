"""
Resilient HTTP wrapper used for every Microsoft Graph and SharePoint call.
"""
import email.utils
import logging
import math
import random
import time
from collections.abc import Callable
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse

import requests

# Retry defaults
DEFAULT_MAX_RETRIES = 5
DEFAULT_BASE_DELAY = 1.0
DEFAULT_MAX_DELAY = 60.0

# Server-specified Retry-After values are honoured up to this higher cap:
# SharePoint/Graph throttle windows can legitimately exceed DEFAULT_MAX_DELAY
# (capping each wait at 60s would land every retry inside the same window),
# while a hostile/buggy server still cannot stall the run indefinitely.
RETRY_AFTER_MAX_DELAY = 300.0

# HTTP timeout (connect, read) in seconds — prevents indefinite hangs
DEFAULT_TIMEOUT = (10, 30)

# Maximum redirect hops followed manually for non-GET requests.
MAX_REDIRECT_HOPS = 5

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Resilient HTTP wrapper
# ---------------------------------------------------------------------------


class _RenewableToken(str):
    """An access token, or an Authorization header value, that can be replaced by a current one.

    Helpers take a token as a string and build their headers once, but one
    request can be retried for many minutes (Retry-After is honoured up to
    RETRY_AFTER_MAX_DELAY, for each of several retries), and a helper can
    wait out replication or follow pages with the same headers.  A
    credential only promises a token will last a few minutes when it hands
    it out, so _request_with_retry renews the header before every attempt.

    Attributes:
        renew (Callable[[], str]): Returns the current value, such as from a credential's get_token.
    """

    def __new__(cls, renew: Callable[[], str], value: str | None = None) -> "_RenewableToken":
        """Creates the token.

        Args:
            renew: Returns the current value. Credentials cache their tokens, so it's cheap to call.
            value: Optional. The current value, if the caller already has it. Defaults to calling renew.
        """
        token = super().__new__(cls, renew() if value is None else value)
        token.renew = renew
        return token


def _bearer(token: str) -> str:
    """Return the Authorization header value for *token*.

    Args:
        token: An access token. A _RenewableToken gives a header value _request_with_retry renews before each
            attempt.

    Returns:
        ``Bearer <token>``.
    """
    if isinstance(token, _RenewableToken):
        return _RenewableToken(lambda: f"Bearer {token.renew()}", f"Bearer {token}")
    return f"Bearer {token}"


def _with_current_token(kwargs: dict) -> None:
    """Replace a renewable Authorization header in a request's keyword arguments with a current one.

    Args:
        kwargs: The keyword arguments for requests.request. Its headers are copied rather than changed, as
            the caller may reuse them.
    """
    headers = kwargs.get("headers")
    authorization = headers.get("Authorization") if headers else None
    if isinstance(authorization, _RenewableToken):
        kwargs["headers"] = {**headers, "Authorization": _RenewableToken(authorization.renew)}


def _parse_retry_after(value: str | None) -> float | None:
    """Parse a Retry-After header into a delay in seconds.

    RFC 7231 allows either delta-seconds ("120") or an HTTP-date
    ("Wed, 21 Oct 2026 07:28:00 GMT").  Returns None when the header is
    absent or unparseable.  The result is clamped to [0, RETRY_AFTER_MAX_DELAY]
    so a hostile/buggy server cannot stall the run indefinitely.
    """
    if value is None:
        return None
    try:
        delay = float(value)
    except ValueError:
        try:
            retry_at = email.utils.parsedate_to_datetime(value)
        except (TypeError, ValueError):
            return None
        if retry_at.tzinfo is None:
            # parsedate_to_datetime yields naive datetimes for '-0000' or
            # missing zones — treat them as UTC rather than crash subtracting.
            retry_at = retry_at.replace(tzinfo=timezone.utc)
        delay = (retry_at - datetime.now(timezone.utc)).total_seconds()
    if not math.isfinite(delay):
        # float("NaN") parses but survives min/max clamping (NaN comparisons
        # are all False) and would crash time.sleep — treat as unparseable.
        return None
    return min(max(delay, 0.0), RETRY_AFTER_MAX_DELAY)


def _service_error(resp: requests.Response) -> tuple[str | None, str | None]:
    """Return the error code and message from a Graph or SharePoint error response.

    Graph sends ``{"error": {"code", "message"}}``.  SharePoint REST nests the
    message as ``{"lang", "value"}``, under ``error`` (odata=verbose) or
    ``odata.error``, and answers a token it won't accept with ``error`` and
    ``error_description`` at the top level.

    Args:
        resp: The response.

    Returns:
        The code and message. Either is None if the body isn't one of those shapes, or the value isn't a non-empty
        string: the message is written to the action log, so nothing else from the body is.
    """
    def text(value: object) -> str | None:
        return value if isinstance(value, str) and value.strip() else None

    try:
        body = resp.json()
    except ValueError:
        return None, None
    if not isinstance(body, dict):
        return None, None
    error = body.get("error", body.get("odata.error"))
    if isinstance(error, dict):
        message = error.get("message")
        if isinstance(message, dict):
            message = message.get("value")
        return text(error.get("code")), text(message)
    if isinstance(error, str):
        return text(error), text(body.get("error_description"))
    return None, None


def _raise_for_status(resp: requests.Response) -> None:
    """Raise requests.HTTPError for a 4xx or 5xx response, with the service's error code and message.

    requests' own message has the reason phrase and the whole URL, which can
    hold user principal names and filters, but not the service's error,
    which says what went wrong.  Callers write the message to the action log.

    Args:
        resp: The response.

    Raises:
        requests.HTTPError: If the status is 400 or above, with *resp* as its response, so callers can still check
            the status code.
    """
    if resp.status_code < 400:
        return
    code, message = _service_error(resp)
    detail = ": ".join(part for part in (code, message) if part) or resp.reason or "no error details"
    raise requests.HTTPError(f"Status code: {resp.status_code}, {detail}", response=resp)


def _has_ambiguous_host(url: str) -> bool:
    """Return whether requests might send a request for *url* to a host other than urlparse's hostname.

    requests (through urllib3) ends the host at a backslash, but urlparse
    doesn't, so urlparse reads 'https://evil.example\\@contoso.sharepoint.com'
    as contoso.sharepoint.com while requests connects to evil.example.  Host
    checks use urlparse, so a backslash before the path, or a user name,
    which no Graph or SharePoint URL needs, is refused rather than relying
    on the two parsers agreeing.

    Args:
        url: The URL a bearer token would be sent to.

    Returns:
        True if the URL has a backslash before its path or a user name or password.
    """
    parsed = urlparse(url)
    return "\\" in parsed.netloc or parsed.username is not None or parsed.password is not None


def _request_with_retry(
    method: str,
    url: str,
    max_retries: int = DEFAULT_MAX_RETRIES,
    base_delay: float = DEFAULT_BASE_DELAY,
    max_delay: float = DEFAULT_MAX_DELAY,
    idempotent: bool = True,
    **kwargs,
) -> requests.Response:
    """Make an HTTP request with automatic retry on transient failures.

    Retries on:
      - HTTP 429 (Too Many Requests) — respects the Retry-After header
      - HTTP 5xx — server errors, respecting Retry-After when it is sent, as
        SharePoint and Graph send it with 503 when throttling
      - Connection errors and timeouts

    Does NOT retry on other 4xx errors (returns the response for the caller).

    An Authorization header built with _bearer from a _RenewableToken is
    renewed before every attempt, so a long wait doesn't leave a retry with
    an expired token.

    With *idempotent* False (a request that must not run twice, such as
    creating a group), only failures where the server certainly did not act
    are retried: 429 and connect timeouts.  A 5xx, read timeout or dropped
    connection may follow a request that succeeded, so it is returned or
    raised rather than repeated — the next run finds what was created.

    Args:
        method:      HTTP method (GET, POST, DELETE, etc.).
        url:         Request URL.
        max_retries: Maximum number of retry attempts.
        base_delay:  Starting delay in seconds for exponential backoff.
        max_delay:   Cap on backoff delay in seconds.
        idempotent:  False for requests that must not be repeated, see above.
        **kwargs:    Passed through to requests.request().

    Returns:
        The requests.Response object.
    """
    last_exception = None
    kwargs.setdefault("timeout", DEFAULT_TIMEOUT)

    # Disable automatic redirects so we can follow them ourselves: every hop
    # is checked before the bearer token is sent to it, and the HTTP method
    # is preserved.  The requests library follows RFC 7231 and downgrades
    # POST → GET on 301/302, which breaks SharePoint POST endpoints like
    # ensureuser.  303 means "GET the new location", so it is only followed
    # for GET.
    kwargs.setdefault("allow_redirects", False)
    follow_manually = not kwargs["allow_redirects"]
    redirect_codes = (301, 302, 303, 307, 308) if method.upper() == "GET" else (301, 302, 307, 308)

    for attempt in range(max_retries + 1):
        _with_current_token(kwargs)
        try:
            resp = requests.request(method, url, **kwargs)

            # Handle redirects — re-issue with the same method to the new
            # Location.  Redirects to a different host, or from https to
            # http, are refused: silently following one would send the
            # Authorization bearer token there, and a response from it
            # could steer later requests (e.g. through a nextLink).
            hops = 0
            while follow_manually and resp.status_code in redirect_codes:
                redirect_url = resp.headers.get("Location")
                if not redirect_url:
                    break
                # Location may be relative (RFC 7231) — resolve it against the
                # URL requested, which has the query params were encoded
                # into, before the cross-host comparison.
                redirect_url = urljoin(resp.url or url, redirect_url)
                if hops >= MAX_REDIRECT_HOPS:
                    raise requests.HTTPError(
                        f"Too many redirects ({hops}) for {method} {url}",
                        response=resp,
                    )
                target, current = urlparse(redirect_url), urlparse(url)
                if _has_ambiguous_host(redirect_url) or target.hostname != current.hostname or (
                    current.scheme == "https" and target.scheme != "https"
                ):
                    raise requests.HTTPError(
                        f"Refusing cross-host or https-to-http redirect for {method}: "
                        f"{url} → {redirect_url}",
                        response=resp,
                    )
                logger.debug(
                    "Following %d redirect for %s: %s → %s",
                    resp.status_code, method, url, redirect_url,
                )
                url = redirect_url
                hops += 1
                # Location already carries the query, so params would add it
                # again, and OData refuses a repeated $filter or $select.
                # Retries after this go to Location too.
                kwargs.pop("params", None)
                # Don't count the redirect as a retry — just re-issue
                resp = requests.request(method, url, **kwargs)

            # A redirect that wasn't followed (no Location, or a 303 for a
            # write) is a failure: raise_for_status() lets 3xx through, so
            # returning it would let a write helper report success.
            if follow_manually and resp.status_code in (301, 302, 303, 307, 308):
                raise requests.HTTPError(
                    f"Unfollowed redirect ({resp.status_code}) for {method} {url}", response=resp,
                )

            # Success or non-retryable client error — return immediately;
            # callers treat non-2xx via raise_for_status as usual
            if resp.status_code < 500 and resp.status_code != 429:
                return resp
            if not idempotent and resp.status_code != 429:
                return resp

            # 429: respect Retry-After header
            if resp.status_code == 429:
                delay = _parse_retry_after(resp.headers.get("Retry-After"))
                if delay is None:
                    delay = min(base_delay * (2 ** attempt), max_delay)
                logger.warning(
                    "Rate limited (429) on %s %s — retrying in %.1fs (attempt %d/%d)",
                    method, url, delay, attempt + 1, max_retries,
                )
            else:
                # 5xx server error.  SharePoint and Graph also throttle with
                # 503 and a Retry-After header, so honour one when it's sent:
                # retrying sooner only extends the throttling.
                delay = _parse_retry_after(resp.headers.get("Retry-After"))
                if delay is None:
                    delay = min(base_delay * (2 ** attempt), max_delay)
                    # Add jitter (0-25%) to avoid thundering herd
                    delay += delay * random.uniform(0, 0.25)
                logger.warning(
                    "Server error %d on %s %s — retrying in %.1fs (attempt %d/%d)",
                    resp.status_code, method, url, delay, attempt + 1, max_retries,
                )

            if attempt < max_retries:
                time.sleep(delay)
                continue

            # Out of retries — return the last response so caller can raise_for_status
            return resp

        except (requests.ConnectionError, requests.Timeout) as exc:
            if not idempotent and not isinstance(exc, requests.ConnectTimeout):
                raise
            last_exception = exc
            delay = min(base_delay * (2 ** attempt), max_delay)
            delay += delay * random.uniform(0, 0.25)
            logger.warning(
                "Connection error on %s %s — retrying in %.1fs (attempt %d/%d): %s",
                method, url, delay, attempt + 1, max_retries, exc,
            )
            if attempt < max_retries:
                time.sleep(delay)
                continue
            raise

    # Should not reach here, but just in case
    if last_exception:
        raise last_exception
    return resp  # type: ignore[possibly-undefined]
