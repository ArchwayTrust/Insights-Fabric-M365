"""
Entra ID security group and user helpers: naming, lookup, creation, membership and resolution.
"""
import logging
import re
import time
from urllib.parse import urlparse

import requests

from ._common import GRAPH_BASE_URL, REPLICATION_BASE_DELAY, REPLICATION_MAX_ATTEMPTS, _append_log, _ThreadSafeLog
from ._graph import _graph_headers
from ._http import _request_with_retry

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Entra ID security group management
# ---------------------------------------------------------------------------


def sanitise_for_group_name(name: str) -> str:
    """Remove spaces and special chars to build a clean group name segment."""
    return "".join(c for c in name if c.isalnum())


def _extract_site_name(site_url: str) -> str:
    """
    Extract the site name from a SharePoint site URL.

    e.g. "https://contoso.sharepoint.com/sites/hr" -> "hr"
         "https://contoso.sharepoint.com/sites/finance-team" -> "financeteam"
    """

    path = urlparse(site_url.rstrip("/")).path  # e.g. "/sites/hr"
    # Take the last segment of the path as the site name
    site_name = path.rsplit("/", 1)[-1] if "/" in path else path
    return sanitise_for_group_name(site_name)


def build_group_name(
    group_prefix: str, site_url: str, library_name: str, role: str,
    *, include_site_name: bool = True,
) -> str:
    """
    Build a deterministic security group name.

    The role suffix is used as-is, so callers control capitalisation.

    Args:
        include_site_name: When True (default), the site name extracted from
            *site_url* is included in the group name.  When False it is
            omitted, producing shorter names.

    Examples:
        build_group_name("SP", "https://…/sites/hr", "Policy Documents", "Auto-Read")
        -> "SP-hr-PolicyDocuments-Auto-Read"

        build_group_name("SP", "https://…/sites/hr", "Policy Documents", "Auto-Read",
                          include_site_name=False)
        -> "SP-PolicyDocuments-Auto-Read"
    """
    sanitised_lib = sanitise_for_group_name(library_name)
    if include_site_name:
        site_name = _extract_site_name(site_url)
        return f"{group_prefix}-{site_name}-{sanitised_lib}-{role}"
    return f"{group_prefix}-{sanitised_lib}-{role}"


def _find_group_by_name(token: str, display_name: str) -> dict | None:
    """Find an Entra ID group by exact display name. Returns None if not found.

    Entra does not enforce display-name uniqueness.  Because the group found
    here is subsequently granted library permissions and has its membership
    synced, silently picking one of several matches would be dangerous —
    multiple matches raise instead so a human can resolve the duplicate.
    """
    # Use $filter with eq for exact match
    url = (
        f"{GRAPH_BASE_URL}/groups"
        f"?$filter=displayName eq '{_odata_escape(display_name)}'"
        f"&$select=id,displayName,description,mailNickname,securityEnabled,groupTypes"
    )
    resp = _request_with_retry("GET", url, headers=_graph_headers(token))
    resp.raise_for_status()
    groups = resp.json().get("value", [])
    if len(groups) > 1:
        ids = ", ".join(g["id"] for g in groups)
        raise ValueError(
            f"Multiple Entra groups share displayName '{display_name}' ({ids}) — "
            f"refusing to guess which one to manage. Delete/rename the duplicates."
        )
    if groups:
        return groups[0]
    return None


def _find_groups_by_name_prefix(token: str, name_prefix: str) -> dict[str, list[dict]]:
    """Find Entra ID groups whose displayName starts with *name_prefix*.

    Returns displayName -> list of matching group resources (Entra does not
    enforce display-name uniqueness).  One paged startswith query replaces a
    name-equality lookup per group, which matters in full-pass reconciles.
    Resolve individual names through _get_unique_group so an unrelated
    duplicate under the same prefix cannot fail the whole operation.
    """
    url = (
        f"{GRAPH_BASE_URL}/groups"
        f"?$filter=startswith(displayName, '{_odata_escape(name_prefix)}')"
        f"&$select=id,displayName,description,mailNickname"
    )
    groups_by_name: dict[str, list[dict]] = {}
    while url:
        resp = _request_with_retry("GET", url, headers=_graph_headers(token))
        resp.raise_for_status()
        data = resp.json()
        for group in data.get("value", []):
            groups_by_name.setdefault(group.get("displayName") or "", []).append(group)
        url = data.get("@odata.nextLink")
    return groups_by_name


def _get_unique_group(
    groups_by_name: dict[str, list[dict]], display_name: str,
) -> dict | None:
    """Resolve one name from a _find_groups_by_name_prefix result.

    Same contract as _find_group_by_name: None when absent, ValueError when
    duplicated — the caller is about to grant permissions or sync membership
    against the group, so guessing between duplicates is never safe.
    """
    matches = groups_by_name.get(display_name, [])
    if len(matches) > 1:
        ids = ", ".join(g["id"] for g in matches)
        raise ValueError(
            f"Multiple Entra groups share displayName '{display_name}' ({ids}) — "
            f"refusing to guess which one to manage. Delete/rename the duplicates."
        )
    return matches[0] if matches else None


def _build_mail_nickname(display_name: str) -> str:
    """Build a Graph-valid mailNickname from a display name.

    Graph rejects nicknames containing characters outside its allowed set or
    longer than 64 characters, and the group prefix comes from external input.
    """
    nickname = re.sub(r"[^a-z0-9]", "", display_name.lower())[:64]
    return nickname or "group"


def _create_security_group(
    token: str, display_name: str, description: str
) -> dict:
    """Create an Entra ID security group. Returns the group resource."""
    body = {
        "displayName": display_name,
        "description": description,
        "mailEnabled": False,
        "mailNickname": _build_mail_nickname(display_name),
        "securityEnabled": True,
    }
    resp = _request_with_retry(
        "POST", f"{GRAPH_BASE_URL}/groups", headers=_graph_headers(token), json=body
    )
    resp.raise_for_status()
    group = resp.json()
    logger.info("Created security group '%s' (id=%s)", display_name, group["id"])
    return group


def _ensure_security_group(
    token: str, display_name: str, description: str
) -> tuple[str, bool]:
    """Ensure a security group exists, creating it if needed.

    Returns:
        (group_id, newly_created) — *newly_created* is True when the group
        was just created and downstream callers should allow for Entra ID
        replication delay.
    """
    existing = _find_group_by_name(token, display_name)
    if existing:
        logger.info("Security group '%s' already exists (id=%s)", display_name, existing["id"])
        return existing["id"], False
    group = _create_security_group(token, display_name, description)
    return group["id"], True


def _rename_security_group(token: str, group_id: str, new_display_name: str) -> None:
    """Rename an Entra ID security group (displayName and mailNickname)."""
    body = {
        "displayName": new_display_name,
        "mailNickname": _build_mail_nickname(new_display_name),
    }
    url = f"{GRAPH_BASE_URL}/groups/{group_id}"
    resp = _request_with_retry("PATCH", url, headers=_graph_headers(token), json=body)
    resp.raise_for_status()
    logger.info("Renamed security group %s → '%s'", group_id, new_display_name)


def _get_group_members(token: str, group_id: str) -> dict[str, str]:
    """
    Get current (direct) members of a security group.

    Returns:
        Dict mapping member key -> member object ID.  Users are keyed by
        lowercase UPN, nested groups by ``group:{id}``, and anything else
        (service principals, devices, contacts) by ``objectid:{id}`` —
        dropping non-user members would let them silently survive unlink
        emptying and membership-sync removal.  A member whose type Graph
        does not report falls back to ``objectid:``, which membership sync
        preserves (the conservative failure mode).
    """
    members: dict[str, str] = {}
    url = f"{GRAPH_BASE_URL}/groups/{group_id}/members?$select=id,userPrincipalName"
    while url:
        resp = _request_with_retry("GET", url, headers=_graph_headers(token))
        resp.raise_for_status()
        data = resp.json()
        for member in data.get("value", []):
            upn = member.get("userPrincipalName", "")
            if upn:
                key = upn.lower()
            elif member.get("@odata.type") == "#microsoft.graph.group":
                key = f"group:{member['id']}"
            else:
                key = f"objectid:{member['id']}"
            members[key] = member["id"]
        url = data.get("@odata.nextLink")
    return members


def _odata_escape(value: str) -> str:
    """Escape a string value for use inside an OData single-quoted literal.

    Per the OData specification, a single quote inside a string literal is
    represented by doubling it (``''``).  Failing to escape allows malformed
    filter expressions and potential OData injection when the value comes from
    external input such as a DataFrame column.
    """
    return value.replace("'", "''")


class UserResolutionError(Exception):
    """A user principal name / email could not be resolved to an Entra user.

    Carries the original input and a human-readable reason so callers can
    log a clean, per-user warning instead of a raw HTTP error dump.
    """

    def __init__(self, user_principal_name: str, reason: str):
        self.user_principal_name = user_principal_name
        self.reason = reason
        super().__init__(f"Cannot resolve '{user_principal_name}': {reason}")


def _resolve_user_with_fallback(
    token: str, user_principal_name: str,
) -> tuple[str, str]:
    """Resolve a UPN or email alias to *(object_id, real_upn)*.

    Resolution order:

    1. Direct ``/users/{upn}`` lookup (fast-path for matching UPNs).
    2. Filter on the ``mail`` property (catches primary SMTP alias).
    3. Filter on ``proxyAddresses`` (catches any SMTP alias).

    Identifiers Graph rejects as malformed (HTTP 400 — stray spaces, bad
    characters from the source system) fall through each step rather than
    aborting: bad source data must surface as a clean per-user failure,
    not an opaque HTTP error that kills the whole library sync.

    Returns:
        Tuple of (Entra object-ID, canonical lowercase UPN).

    Raises:
        UserResolutionError: When the user cannot be found by any method or
            the identifier is malformed — carries a human-readable reason.
        requests.HTTPError: For genuine API failures (auth, exhausted
            throttling retries, 5xx) that should abort the sync.
    """
    if not user_principal_name.strip():
        raise UserResolutionError(user_principal_name, "empty user principal name")

    malformed = False

    # 1. Direct UPN lookup
    url = f"{GRAPH_BASE_URL}/users/{requests.utils.quote(user_principal_name)}"
    resp = _request_with_retry(
        "GET", url, headers=_graph_headers(token),
        params={"$select": "id,userPrincipalName"},
    )
    if resp.status_code == 200:
        data = resp.json()
        return data["id"], data.get("userPrincipalName", user_principal_name).lower()
    if resp.status_code == 400:
        # Graph rejects syntactically invalid UPNs with 400 — the search
        # fallbacks below may still match, so note it and continue.
        malformed = True
    elif resp.status_code != 404:
        resp.raise_for_status()

    # 2. Search by mail property
    search_url = f"{GRAPH_BASE_URL}/users"
    safe_upn = _odata_escape(user_principal_name)
    filter_mail = f"mail eq '{safe_upn}'"
    resp = _request_with_retry(
        "GET", search_url, headers=_graph_headers(token),
        params={"$filter": filter_mail, "$select": "id,userPrincipalName"},
    )
    if resp.status_code == 400:
        malformed = True
    else:
        resp.raise_for_status()
        results = resp.json().get("value", [])
        if results:
            user = results[0]
            real_upn = user.get("userPrincipalName", user_principal_name).lower()
            logger.info("Resolved '%s' via mail property → '%s'", user_principal_name, real_upn)
            return user["id"], real_upn

    # 3. Search by proxyAddresses.  Graph does not support tolower() in
    #    $filter (it 400s every call), and directory string comparisons are
    #    case-insensitive anyway, so the address is matched as-is.
    filter_proxy = f"proxyAddresses/any(p:p eq 'smtp:{safe_upn.lower()}')"
    resp = _request_with_retry(
        "GET", search_url, headers=_graph_headers(token),
        params={"$filter": filter_proxy, "$select": "id,userPrincipalName"},
    )
    if resp.status_code == 400:
        malformed = True
    else:
        resp.raise_for_status()
        results = resp.json().get("value", [])
        if results:
            user = results[0]
            real_upn = user.get("userPrincipalName", user_principal_name).lower()
            logger.info("Resolved '%s' via proxyAddresses → '%s'", user_principal_name, real_upn)
            return user["id"], real_upn

    # 4. Not found by any method
    raise UserResolutionError(
        user_principal_name,
        "rejected by Graph as malformed — check how the address is stored "
        "in the source system"
        if malformed
        else "not found by UPN, mail, or proxyAddresses",
    )


_GUID_PATTERN = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE,
)


class GroupResolutionError(Exception):
    """A group reference (display name or object ID) could not be resolved.

    The group counterpart of UserResolutionError: carries the original
    reference and a human-readable reason so callers can log a clean,
    per-reference warning.
    """

    def __init__(self, group_reference: str, reason: str):
        self.group_reference = group_reference
        self.reason = reason
        super().__init__(f"Cannot resolve group '{group_reference}': {reason}")


def _resolve_nestable_group(token: str, group_reference: str) -> dict:
    """Resolve a group object ID or exact display name to a group that can be nested.

    Accepting either form lets a hand-maintained table name a group
    readably, while an object ID stays available for display names that
    are not unique.  Entra only nests security groups inside security
    groups — a Microsoft 365 group or distribution list would be rejected
    by Graph on every run, so it is refused here as a resolution failure
    instead.

    Returns:
        The group resource (``id``, ``displayName``, …).

    Raises:
        GroupResolutionError: When the reference is empty, matches nothing,
            is a display name shared by several groups (the group is about
            to gain library access, so guessing is never safe), or is not a
            security group.
        requests.HTTPError: For genuine API failures.
    """
    reference = group_reference.strip()
    if not reference:
        raise GroupResolutionError(group_reference, "empty group reference")

    if _GUID_PATTERN.match(reference):
        resp = _request_with_retry(
            "GET", f"{GRAPH_BASE_URL}/groups/{reference}",
            headers=_graph_headers(token),
            params={"$select": "id,displayName,securityEnabled,groupTypes"},
        )
        if resp.status_code == 404:
            raise GroupResolutionError(group_reference, "no group has this object ID")
        resp.raise_for_status()
        group = resp.json()
    else:
        try:
            group = _find_group_by_name(token, reference)
        except ValueError:
            raise GroupResolutionError(
                group_reference,
                "display name is shared by more than one group — use the object ID instead",
            ) from None
        if group is None:
            raise GroupResolutionError(group_reference, "no group has this display name")

    if not group.get("securityEnabled") or "Unified" in (group.get("groupTypes") or []):
        raise GroupResolutionError(
            group_reference,
            "not a security group — Microsoft 365 groups and distribution lists "
            "cannot be nested in a security group",
        )
    return group


def _add_group_member(
    token: str, group_id: str, user_object_id: str, *, is_new_group: bool = False
) -> None:
    """Add a user to a security group.

    When *is_new_group* is True, retries on 404 (group not yet replicated).
    """
    url = f"{GRAPH_BASE_URL}/groups/{group_id}/members/$ref"
    body = {
        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_object_id}"
    }

    max_attempts = REPLICATION_MAX_ATTEMPTS if is_new_group else 1
    for attempt in range(max_attempts):
        resp = _request_with_retry("POST", url, headers=_graph_headers(token), json=body)
        # 400 with "already exist" means user is already a member — treat as success
        if resp.status_code == 400 and "already exist" in resp.text.lower():
            logger.info("User %s already a member of group %s", user_object_id, group_id)
            return
        if resp.status_code == 404 and is_new_group and attempt < max_attempts - 1:
            delay = REPLICATION_BASE_DELAY * (2 ** attempt)
            logger.info(
                "Group %s not yet replicated (404) — waiting %.0fs (attempt %d/%d)",
                group_id, delay, attempt + 1, max_attempts,
            )
            time.sleep(delay)
            continue
        resp.raise_for_status()
        return


def _remove_group_member(token: str, group_id: str, user_object_id: str) -> None:
    """Remove a user from a security group."""
    url = f"{GRAPH_BASE_URL}/groups/{group_id}/members/{user_object_id}/$ref"
    resp = _request_with_retry("DELETE", url, headers=_graph_headers(token))
    resp.raise_for_status()


def _empty_group_members(token: str, group_id: str) -> int:
    """Remove all members from a security group. Returns the count removed."""
    members = _get_group_members(token, group_id)
    for upn, member_oid in members.items():
        _remove_group_member(token, group_id, member_oid)
        logger.info("Removed '%s' from group %s (unlink)", upn, group_id)
    return len(members)


def _find_group_name_collisions(
    group_prefixes: dict[tuple, str], include_site_name: bool,
) -> dict[tuple, tuple]:
    """Find (site_url, library_code) pairs whose sanitised group names collide.

    ``sanitise_for_group_name`` strips punctuation, so distinct codes like
    'PE-KS3' and 'PEKS3' map to one group name — silently cross-wiring two
    libraries' permissions.  *group_prefixes* maps
    (site_url, library_code) -> group_prefix.

    Returns each colliding pair mapped to the pair it collides with (both
    directions), so callers can skip every involved library with a warning
    row while the rest of the run proceeds — a collision is a persistent
    data problem that a run-aborting exception could never retry away.
    """
    seen: dict[tuple, tuple] = {}
    collisions: dict[tuple, tuple] = {}
    for site_url, library_code in group_prefixes:
        prefix = group_prefixes[(site_url, library_code)]
        if include_site_name:
            key = (prefix, _extract_site_name(site_url), sanitise_for_group_name(library_code))
        else:
            key = (prefix, sanitise_for_group_name(library_code))
        other = seen.get(key)
        if other is not None and other != (site_url, library_code):
            collisions[(site_url, library_code)] = other
            collisions[other] = (site_url, library_code)
        else:
            seen[key] = (site_url, library_code)
    return collisions


def _log_group_name_collisions(
    collisions: dict[tuple, tuple], action_log: _ThreadSafeLog,
) -> None:
    """Write a warning row for every library skipped by a name collision."""
    for (site_url, library_code), other in collisions.items():
        logger.warning(
            "Group-name collision: (%s, %s) collides with %s — skipped",
            site_url, library_code, other,
        )
        _append_log(
            action_log, site_url, library_code, "group_name_collision",
            f"sanitised group names collide with {other} — library skipped; "
            f"rename one code (or set include_site_name=True) so permissions "
            f"cannot cross-wire",
        )
