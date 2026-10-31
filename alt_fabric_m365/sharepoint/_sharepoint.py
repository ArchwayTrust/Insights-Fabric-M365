"""
SharePoint REST API helpers: library permissions, property bag metadata and site sharing settings.
"""
import logging
import time
from xml.sax.saxutils import escape as xml_escape

import requests

from ._common import (
    METADATA_PROPERTY_KEYS, REPLICATION_BASE_DELAY, REPLICATION_MAX_ATTEMPTS, VALID_ROLES,
    _append_log, _as_text, _ThreadSafeLog,
)
from ._http import _bearer, _raise_for_status, _request_with_retry

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# SharePoint REST API helpers for library-level permission assignments
# ---------------------------------------------------------------------------


def _get_sharepoint_rest_url(site_url: str) -> str:
    """Build the base SharePoint REST API URL from the site URL."""
    return site_url.rstrip("/") + "/_api"


def _sharepoint_headers(token: str) -> dict:
    """Return headers for SharePoint REST API calls."""
    return {
        "Authorization": _bearer(token),
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
    }


def _sharepoint_post_headers(token: str) -> dict:
    """Return headers for SharePoint REST API POST (write) calls.

    With app-only OAuth bearer tokens (client-credentials flow), the form
    digest is not required for CSRF protection — that mechanism is only
    needed for cookie-based browser authentication.  Passing "0" satisfies
    the X-RequestDigest header requirement without a contextinfo round-trip.
    """
    headers = _sharepoint_headers(token)
    headers["X-RequestDigest"] = "0"
    return headers


def _has_unique_permissions(token: str, site_url: str, list_id: str) -> bool:
    """Check whether a library already has unique (non-inherited) permissions."""
    rest_url = _get_sharepoint_rest_url(site_url)
    url = (
        f"{rest_url}/web/lists(guid'{list_id}')"
        f"?$select=HasUniqueRoleAssignments"
    )
    resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
    _raise_for_status(resp)
    return resp.json()["d"]["HasUniqueRoleAssignments"]


def _break_role_inheritance(token: str, site_url: str, list_id: str) -> None:
    """Break role inheritance on a document library so it has its own permissions.

    copyRoleAssignments=false starts the library with no role assignments, so
    only the groups assigned to it give access.  clearSubscopes=false leaves
    folders and files that have their own permissions alone: a new library has
    none, but an existing library being taken over may have a folder restricted
    to a few people, which clearSubscopes=true would open to the library's groups.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    url = (
        f"{rest_url}/web/lists(guid'{list_id}')"
        f"/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=false)"
    )
    resp = _request_with_retry("POST", url, headers=_sharepoint_post_headers(token), data="")
    _raise_for_status(resp)
    logger.info("Broke role inheritance on library '%s'", list_id)


# SharePoint's built-in role definition for each permission level, by
# RoleTypeKind (SP.RoleType).  Names are localised and can be renamed, so
# they aren't used to find the roles.
SP_ROLE_TYPES = {
    "read": 2,        # Reader: "Read"
    "contribute": 3,  # Contributor: "Contribute"
    "owner": 5,       # Administrator: "Full Control"
}


def _group_claims(group_id: str) -> str:
    """Claims-based login name SharePoint uses for an Entra ID security group."""
    return f"c:0t.c|tenant|{group_id}"


def _resolve_role_def_ids(token: str, site_url: str) -> dict[str, int]:
    """Pre-resolve all role definition IDs for a site, in one request.

    Returns a dict mapping role name ('read', 'contribute', 'owner') to its
    numeric SharePoint role definition ID.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    resp = _request_with_retry(
        "GET", f"{rest_url}/web/roledefinitions?$select=Id,Name,RoleTypeKind",
        headers=_sharepoint_headers(token),
    )
    _raise_for_status(resp)
    ids_by_type = {rd["RoleTypeKind"]: rd["Id"] for rd in resp.json()["d"]["results"] if rd.get("RoleTypeKind")}
    missing = [role for role in VALID_ROLES if SP_ROLE_TYPES[role] not in ids_by_type]
    if missing:
        raise ValueError(f"Built-in role definition(s) for {missing} not found on {site_url}")
    return {role: ids_by_type[SP_ROLE_TYPES[role]] for role in VALID_ROLES}


def _resolve_group_principal_id(
    token: str, site_url: str, group_name: str, *, is_new_group: bool = False
) -> int:
    """Resolve an Entra ID security group to a SharePoint principal ID.

    Uses the ensureuser endpoint with the group's claims-based identity
    (c:0t.c|tenant|group-id).

    When *is_new_group* is True, retries on 400/404 to allow for Entra ID
    replication delay before SharePoint can recognise the group.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    url = f"{rest_url}/web/ensureuser"
    body = {"logonName": group_name}

    max_attempts = REPLICATION_MAX_ATTEMPTS if is_new_group else 1
    for attempt in range(max_attempts):
        resp = _request_with_retry("POST", url, headers=_sharepoint_post_headers(token), json=body)
        if resp.ok:
            return resp.json()["d"]["Id"]
        if is_new_group and resp.status_code in (400, 404) and attempt < max_attempts - 1:
            delay = REPLICATION_BASE_DELAY * (2 ** attempt)
            logger.info(
                "ensureuser failed for '%s' (%d) — group may not have replicated yet, "
                "waiting %.0fs (attempt %d/%d)",
                group_name, resp.status_code, delay, attempt + 1, max_attempts,
            )
            time.sleep(delay)
            continue
        _raise_for_status(resp)
    raise requests.HTTPError(
        f"ensureuser failed for '{group_name}' after {max_attempts} attempt(s)",
        response=resp,
    )


def _add_role_assignment(
    token: str,
    site_url: str,
    list_id: str,
    principal_id: int,
    role_def_id: int,
) -> None:
    """Grant a specific role to a principal (user/group) on a document library."""
    rest_url = _get_sharepoint_rest_url(site_url)
    url = (
        f"{rest_url}/web/lists(guid'{list_id}')"
        f"/roleassignments/addroleassignment(principalid={principal_id}, roledefid={role_def_id})"
    )
    resp = _request_with_retry("POST", url, headers=_sharepoint_post_headers(token), data="")
    _raise_for_status(resp)


def _remove_library_role_assignment(
    token: str, site_url: str, list_id: str, principal_id: int,
) -> bool:
    """Remove a principal's entire role assignment from a document library.

    Returns True when an assignment was removed, False when the principal had
    none (already gone — treated as success so retire runs stay idempotent).
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    headers = _sharepoint_post_headers(token)
    headers["X-HTTP-Method"] = "DELETE"
    headers["IF-MATCH"] = "*"
    resp = _request_with_retry(
        "POST",
        f"{rest_url}/web/lists(guid'{list_id}')"
        f"/roleassignments/getbyprincipalid({principal_id})",
        headers=headers, data="",
    )
    if resp.status_code == 404:
        return False
    _raise_for_status(resp)
    return True


def _get_library_group_roles(
    token: str, site_url: str, list_id: str,
) -> dict[str, set[int]]:
    """Return a library's role assignments as lowercase login name -> role definition IDs.

    Groups appear under their claims login name (see _group_claims), so a
    caller can tell whether a group already holds a role without resolving
    its principal ID through ensureuser first.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    url: str | None = (
        f"{rest_url}/web/lists(guid'{list_id}')/roleassignments"
        f"?$expand=Member,RoleDefinitionBindings"
        f"&$select=Member/LoginName,RoleDefinitionBindings/Id"
    )
    roles: dict[str, set[int]] = {}
    while url:
        resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
        _raise_for_status(resp)
        data = resp.json()["d"]
        for assignment in data["results"]:
            login_name = ((assignment.get("Member") or {}).get("LoginName") or "").lower()
            roles.setdefault(login_name, set()).update(
                role_def["Id"] for role_def in assignment["RoleDefinitionBindings"]["results"]
            )
        url = data.get("__next")
    return roles


def _ensure_group_assigned_to_library(
    token: str,
    site_url: str,
    list_id: str,
    group_id: str,
    group_display_name: str,
    role_def_id: int,
    *,
    is_new_group: bool = False,
    existing_roles: dict[str, set[int]] | None = None,
) -> bool:
    """Ensure a security group is assigned to a library with the given role.

    Uses the Entra ID group's claims-based identity to register it as a
    SharePoint principal, then grants the role on the library.

    When *is_new_group* is True, allows for Entra ID replication delay
    before SharePoint can recognise the group via ensureuser.

    *existing_roles* is the library's _get_library_group_roles result, when
    the caller has read it: a group that already holds the role is left
    alone, saving the ensureuser and role assignment writes.

    Returns True when the role was granted, False when it was already held.
    """
    claims_identity = _group_claims(group_id)
    if existing_roles is not None and role_def_id in existing_roles.get(claims_identity.lower(), set()):
        return False
    principal_id = _resolve_group_principal_id(
        token, site_url, claims_identity, is_new_group=is_new_group
    )
    _add_role_assignment(token, site_url, list_id, principal_id, role_def_id)
    logger.info(
        "Assigned group '%s' to library '%s' with role def %d",
        group_display_name, list_id, role_def_id,
    )
    return True


# ---------------------------------------------------------------------------
# SharePoint REST API helpers for library property bag metadata
# ---------------------------------------------------------------------------


def _set_library_properties(
    token: str, site_url: str, list_id: str,
    properties: dict[str, str],
) -> None:
    """Set property bag values on a library's root folder.

    The SharePoint REST API treats the ``Properties`` navigation property
    on ``SP.Folder`` as read-only — MERGE requests that include a nested
    ``Properties`` object succeed (HTTP 200) but silently discard the
    values.  This affects both the ``lists(guid'…')/rootfolder`` and
    ``GetFolderByServerRelativeUrl`` endpoints.

    To work around this we use the CSOM-over-REST ``ProcessQuery``
    endpoint, which calls ``PropertyValues.SetFieldValue`` /
    ``Folder.Update`` — the same code path that PnP PowerShell uses.

    Args:
        token:      SharePoint access token.
        site_url:   Full SharePoint site URL.
        list_id:    Graph list ID (GUID) of the document library.
        properties: Key-value pairs to write to the root folder property bag.
    """
    # Build the CSOM XML request.
    # Object path:  Site (current) → Web → Lists → GetById(list_id)
    #                → RootFolder (id 4) → Properties (id 5)
    # Actions:      SetFieldValue for each property, then Folder.Update.
    set_actions = []
    action_id = 1
    for key, value in properties.items():
        set_actions.append(
            f'<Method Name="SetFieldValue" Id="{action_id}" ObjectPathId="5">'
            f"<Parameters>"
            f'<Parameter Type="String">{xml_escape(key)}</Parameter>'
            f'<Parameter Type="String">{xml_escape(value)}</Parameter>'
            f"</Parameters>"
            f"</Method>"
        )
        action_id += 1
    set_actions.append(f'<Method Name="Update" Id="{action_id}" ObjectPathId="4" />')

    # Ensure list_id is wrapped in braces for CSOM Guid parameter.
    guid = list_id if list_id.startswith("{") else f"{{{list_id}}}"

    csom_xml = (
        '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"'
        ' SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0"'
        ' ApplicationName="Python">'
        "<Actions>" + "".join(set_actions) + "</Actions>"
        "<ObjectPaths>"
        '<StaticProperty Id="0"'
        ' TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}"'
        ' Name="Current" />'
        '<Property Id="1" ParentId="0" Name="Web" />'
        '<Property Id="2" ParentId="1" Name="Lists" />'
        '<Method Id="3" ParentId="2" Name="GetById">'
        f'<Parameters><Parameter Type="Guid">{guid}</Parameter></Parameters>'
        "</Method>"
        '<Property Id="4" ParentId="3" Name="RootFolder" />'
        '<Property Id="5" ParentId="4" Name="Properties" />'
        "</ObjectPaths>"
        "</Request>"
    )

    url = f"{site_url.rstrip('/')}/_vti_bin/client.svc/ProcessQuery"
    headers = {
        "Authorization": _bearer(token),
        "Content-Type": "text/xml",
        "X-RequestDigest": "0",
    }

    resp = _request_with_retry("POST", url, headers=headers, data=csom_xml.encode("utf-8"))
    _raise_for_status(resp)

    # ProcessQuery returns a JSON array.  Error entries contain an
    # ``ErrorInfo`` dict with ``ErrorMessage``.
    result = resp.json()
    for item in result:
        if isinstance(item, dict) and item.get("ErrorInfo"):
            raise ValueError(
                f"CSOM error setting property bag on list {list_id}: "
                f"{item['ErrorInfo']['ErrorMessage']}"
            )

    logger.info("Set property bag values on library '%s': %s", list_id, list(properties.keys()))


def _get_library_properties(
    token: str, site_url: str, list_id: str,
    keys: list[str],
) -> dict[str, str | None]:
    """Read property bag values from a library's root folder.

    Args:
        token:    SharePoint access token.
        site_url: Full SharePoint site URL.
        list_id:  Graph list ID (GUID) of the document library.
        keys:     Property bag keys to read.

    Returns:
        Dict mapping each requested key to its value, or ``None`` if the
        key does not exist in the property bag.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    # Fetch all property bag values — $select is unreliable for custom
    # (non-built-in) keys on the SP.PropertyValues endpoint.
    url = f"{rest_url}/web/lists(guid'{list_id}')/rootfolder/properties"
    resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
    _raise_for_status(resp)

    body = resp.json()
    props = body.get("d", {})
    return {key: props.get(key) for key in keys}


def _ensure_library_metadata(
    token: str,
    site_url: str,
    list_id: str | None,
    library_code: str,
    source_system: str | None,
    source_id: str | None,
    action_log: _ThreadSafeLog,
    *,
    dry_run: bool = False,
) -> None:
    """Set source_system / source_id on a library's root folder property bag.

    Skips any value that is ``None``.  The current values are read first, so
    only those that differ are written (``metadata_set``), and a re-run
    writes nothing (``metadata_unchanged``).  In dry-run mode the values are
    read but only logged as what would be set.

    Args:
        token: SharePoint access token.
        site_url: Full SharePoint site URL.
        list_id: Graph list ID of the library, or None for a library a dry run would create, which has no values.
        library_code: The library's code, for the action log.
        source_system: Value for SourceSystem, or None to leave it.
        source_id: Value for SourceId, or None to leave it.
        action_log: The operation's action log.
        dry_run: Optional. Log what would be set rather than set it. Defaults to False.
    """
    props = {}
    for df_col, prop_key in METADATA_PROPERTY_KEYS.items():
        value = {"source_system": source_system, "source_id": source_id}.get(df_col)
        if value is not None:
            props[prop_key] = _as_text(value)

    if not props:
        return

    current = _get_library_properties(token, site_url, list_id, list(props)) if list_id is not None else {}
    changed = {key: value for key, value in props.items() if current.get(key) != value}
    for prop_key, value in props.items():
        if prop_key not in changed:
            _append_log(
                action_log, site_url, library_code,
                "metadata_unchanged", f"{prop_key}={value}",
            )
    if not changed:
        return

    if dry_run:
        for prop_key, value in changed.items():
            logger.info("[DRY RUN] Would set metadata %s=%s", prop_key, value)
            _append_log(
                action_log, site_url, library_code,
                "metadata_set (dry run)", f"{prop_key}={value}",
            )
        return

    _set_library_properties(token, site_url, list_id, changed)

    for prop_key, value in changed.items():
        _append_log(
            action_log, site_url, library_code,
            "metadata_set", f"{prop_key}={value}",
        )


# ---------------------------------------------------------------------------
# SharePoint REST API helpers for site-level sharing settings
# ---------------------------------------------------------------------------


def _get_site_sharing_settings(token: str, site_url: str) -> dict:
    """Read the site-level sharing settings governed by the lockdown check.

    Returns:
        Dict with three booleans plus the member group id:
          - ``members_can_share``: "Allow members to share the site and
            individual files and folders" (``SP.Web.MembersCanShare``).
          - ``access_requests_enabled``: True when "Allow access requests"
            is on — either a request email is set, or requests go to the
            owners group (``UseAccessRequestDefault``).
          - ``members_can_invite``: "Allow members to invite others to the
            site members group" (``AllowMembersEditMembership`` on the
            associated member group).
          - ``member_group_id``: SharePoint id of the associated member
            group, or None when the web has none (``members_can_invite``
            is then False — there is nothing to lock down).
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    url = (
        f"{rest_url}/web"
        f"?$select=MembersCanShare,RequestAccessEmail,UseAccessRequestDefault"
    )
    resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
    _raise_for_status(resp)
    web = resp.json()["d"]

    url = (
        f"{rest_url}/web/associatedmembergroup"
        f"?$select=Id,AllowMembersEditMembership"
    )
    resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
    if resp.status_code == 404:
        member_group = {}
    else:
        _raise_for_status(resp)
        member_group = resp.json()["d"]

    return {
        "members_can_share": bool(web.get("MembersCanShare")),
        "access_requests_enabled": (
            bool(web.get("RequestAccessEmail")) or bool(web.get("UseAccessRequestDefault"))
        ),
        "members_can_invite": bool(member_group.get("AllowMembersEditMembership")),
        "member_group_id": member_group.get("Id"),
    }


def _disable_site_sharing(token: str, site_url: str, member_group_id: int | None) -> None:
    """Turn off member sharing, access requests, and member invites on a site.

    Sets ``MembersCanShare`` to False and clears ``RequestAccessEmail`` via
    MERGE, turns off ``UseAccessRequestDefault`` through its dedicated CSOM
    method — the property has no setter, so MERGE cannot change it (same
    sequence as PnP's ``DisableRequestAccess``) — and sets
    ``AllowMembersEditMembership`` to False on the associated member group
    when *member_group_id* is not None.
    """
    rest_url = _get_sharepoint_rest_url(site_url)

    headers = _sharepoint_post_headers(token)
    headers["X-HTTP-Method"] = "MERGE"
    body = {
        "__metadata": {"type": "SP.Web"},
        "MembersCanShare": False,
        "RequestAccessEmail": "",
    }
    resp = _request_with_retry("POST", f"{rest_url}/web", headers=headers, json=body)
    _raise_for_status(resp)

    resp = _request_with_retry(
        "POST",
        f"{rest_url}/web/setuseaccessrequestdefaultandupdate(useAccessRequestDefault=false)",
        headers=_sharepoint_post_headers(token),
        data="",
    )
    _raise_for_status(resp)

    if member_group_id is not None:
        headers = _sharepoint_post_headers(token)
        headers["X-HTTP-Method"] = "MERGE"
        body = {
            "__metadata": {"type": "SP.Group"},
            "AllowMembersEditMembership": False,
        }
        resp = _request_with_retry(
            "POST", f"{rest_url}/web/sitegroups/getbyid({member_group_id})",
            headers=headers, json=body,
        )
        _raise_for_status(resp)

    logger.info(
        "Disabled member sharing, access requests, and member invites on %s", site_url,
    )
