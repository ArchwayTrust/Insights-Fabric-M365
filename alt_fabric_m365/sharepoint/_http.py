"""
Resilient HTTP wrapper used for every Microsoft Graph and SharePoint call.
"""
import email.utils
import logging
import math
import random
import time
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
                # current URL before the cross-host comparison.
                redirect_url = urljoin(url, redirect_url)
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
