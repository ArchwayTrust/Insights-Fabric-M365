"""
Microsoft Graph helpers for sites, document libraries and OneDrive shortcuts.
"""
import logging
from urllib.parse import unquote, urlparse

from ._common import GRAPH_BASE_URL
from ._http import _bearer, _request_with_retry

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------


def _graph_headers(token: str) -> dict:
    """Return standard headers for Graph API calls."""
    return {
        "Authorization": _bearer(token),
        "Content-Type": "application/json",
    }


def _get_site_id(token: str, site_url: str) -> str:
    """Resolve a SharePoint site URL to its Graph site ID.

    The Graph API identifies sites as ``{hostname}:/{server-relative-path}``.
    """
    parsed = urlparse(site_url.rstrip("/"))
    hostname = parsed.hostname
    path = parsed.path

    url = f"{GRAPH_BASE_URL}/sites/{hostname}:{path}"
    resp = _request_with_retry("GET", url, headers=_graph_headers(token))
    resp.raise_for_status()
    return resp.json()["id"]


def _list_existing_lists(token: str, site_id: str) -> dict[str, dict]:
    """Fetch all lists/libraries on a site.

    Returns a dict keyed by the lowercase **last segment of the webUrl**
    (the URL slug, which is the library code).  Display names are not keys:
    a library is only ever identified by its URL, so code 'ART' can never
    pick up an unrelated library merely titled 'Art'.

    Only document libraries are keyed by slug.  Other lists (``/Lists/ENG``)
    are keyed ``list:{slug}``, so a code never matches one, but their titles
    still count for _find_title_conflict, as titles are unique across all of
    a site's lists.
    """
    url: str | None = f"{GRAPH_BASE_URL}/sites/{site_id}/lists?$select=id,displayName,webUrl,list"

    lists: list[dict] = []
    while url:
        resp = _request_with_retry("GET", url, headers=_graph_headers(token))
        resp.raise_for_status()
        data = resp.json()
        lists.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    lists_by_key: dict[str, dict] = {}
    for lst in lists:
        # webUrl is percent-encoded — decode so codes containing spaces or
        # '&' match the raw library_code used for lookups.
        web_url = lst.get("webUrl", "")
        if web_url:
            slug = unquote(web_url.rstrip("/").rsplit("/", 1)[-1]).lower()
            template = (lst.get("list") or {}).get("template")
            is_library = template is None or template == "documentLibrary"
            lists_by_key[slug if is_library else f"list:{slug}"] = lst
    return lists_by_key


def _find_title_conflict(existing: dict[str, dict], library_code: str, title: str | None = None) -> dict | None:
    """Return another list already titled *title*, if any.

    SharePoint list titles are unique per site, so creating a library with a
    title another list has, or renaming it to one, fails on every run.

    Args:
        existing: The site's lists, from _list_existing_lists.
        library_code: The library's code. The library with this URL slug never conflicts with itself.
        title: Optional. The title to look for. Defaults to library_code, the title a new library is
            created with (which sets its URL slug).

    Returns:
        The other list, or None if no other list has the title.
    """
    key = library_code.lower()
    wanted = (title or library_code).lower()
    for slug, lst in existing.items():
        if slug != key and (lst.get("displayName") or "").lower() == wanted:
            return lst
    return None


def _get_list_id(
    token: str, site_id: str, library_code: str,
    existing: dict[str, dict] | None = None,
) -> str:
    """Resolve a library code (URL slug) to its Graph list ID.

    Raises ValueError if the library does not exist.

    If *existing* is provided, uses it instead of fetching the list of lists.
    """
    if existing is None:
        existing = _list_existing_lists(token, site_id)
    key = library_code.lower()
    if key in existing:
        return existing[key]["id"]
    raise ValueError(
        f"Library '{library_code}' not found on site {site_id}"
    )


def _create_document_library(token: str, site_id: str, library_name: str) -> dict:
    """Create a new document library on the given site. Returns the list resource."""
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/lists"
    body = {
        "displayName": library_name,
        "list": {"template": "documentLibrary"},
    }
    resp = _request_with_retry("POST", url, headers=_graph_headers(token), json=body)
    resp.raise_for_status()
    new_list = resp.json()
    logger.info("Created document library '%s' (id=%s)", library_name, new_list["id"])
    return new_list


def _ensure_document_library(
    token: str, site_id: str, library_code: str,
    existing: dict[str, dict] | None = None,
) -> dict:
    """Ensure a document library exists, creating it if needed.

    Returns the Graph list resource (``id``, ``webUrl``, …).

    The library is created with *library_code* as its title, which sets the
    URL slug.  Its display name is left to the caller (_set_library_display_name),
    after its permissions are set: renaming can fail when another list has
    the name, and a new library still inherits the site's permissions.

    If *existing* is provided, uses it instead of fetching the list of lists.
    """
    if existing is None:
        existing = _list_existing_lists(token, site_id)
    key = library_code.lower()

    if key in existing:
        lib = existing[key]
        logger.info("Document library '%s' already exists (id=%s)", library_code, lib["id"])
        return lib

    # Create with the code — this becomes the URL slug
    return _create_document_library(token, site_id, library_code)


def _set_library_display_name(
    token: str, site_id: str, list_id: str, display_name: str,
) -> None:
    """Update the display name of a library (the URL slug is unchanged)."""
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}"
    body = {"displayName": display_name}
    resp = _request_with_retry("PATCH", url, headers=_graph_headers(token), json=body)
    resp.raise_for_status()
    logger.info("Set display name to '%s' (id=%s)", display_name, list_id)


# ---------------------------------------------------------------------------
# OneDrive shortcut helpers
# ---------------------------------------------------------------------------


def _get_library_display_name(existing_lists: dict[str, dict], library_code: str) -> str:
    """Return the human-readable display name for a library.

    Falls back to *library_code* when the list is not found in the cache.
    """
    entry = existing_lists.get(library_code.lower())
    if entry:
        return entry.get("displayName", library_code)
    return library_code


def _create_onedrive_shortcut(
    graph_token: str,
    user_object_id: str,
    shortcut_name: str,
    site_id: str,
    list_id: str,
    site_url: str,
) -> None:
    """Create a OneDrive shortcut in a user's drive root pointing to a library.

    Args:
        graph_token:    Microsoft Graph bearer token.
        user_object_id: Entra ID object ID of the target user.
        shortcut_name:  Display name for the shortcut folder in OneDrive.
        site_id:        Composite Graph site ID
                        (``{hostname},{siteCollectionId},{webId}``).
        list_id:        Graph list ID (GUID).
        site_url:       Full SharePoint site URL.
    """
    # Parse the composite Graph site ID to extract the GUIDs that
    # sharepointIds expects.
    parts = site_id.split(",")
    site_collection_id = parts[1] if len(parts) >= 3 else ""
    web_id = parts[2] if len(parts) >= 3 else ""

    sharepoint_ids = {
        "listId": list_id,
        "listItemUniqueId": "root",
        "siteId": site_collection_id,
        "siteUrl": site_url,
        "webId": web_id,
    }

    body = {
        "name": shortcut_name,
        "remoteItem": {
            "sharepointIds": sharepoint_ids,
        },
        "@microsoft.graph.conflictBehavior": "rename",
    }

    logger.debug(
        "OneDrive shortcut request — user: %s, sharepointIds: %s",
        user_object_id, sharepoint_ids,
    )

    url = f"{GRAPH_BASE_URL}/users/{user_object_id}/drive/root/children"
    # Not repeated after a failure that may have created the shortcut:
    # conflictBehavior 'rename' would add a second one.
    resp = _request_with_retry(
        "POST", url, headers=_graph_headers(graph_token), json=body, idempotent=False,
    )
    if not resp.ok:
        logger.warning(
            "OneDrive shortcut creation failed (HTTP %d): %s",
            resp.status_code, resp.text,
        )
    resp.raise_for_status()
