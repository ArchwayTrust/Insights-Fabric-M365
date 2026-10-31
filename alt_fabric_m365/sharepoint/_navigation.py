"""
SharePoint site navigation (quick launch) helpers.
"""
import logging
from urllib.parse import unquote, urlparse

from ._common import _append_log, _ThreadSafeLog
from ._http import _raise_for_status, _request_with_retry
from ._sharepoint import _get_sharepoint_rest_url, _sharepoint_headers, _sharepoint_post_headers

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Site navigation (quick launch) helpers
# ---------------------------------------------------------------------------

# Graph does not expose site navigation, so these use the SharePoint REST API.
# Modern SharePoint represents a non-clickable navigation label as a node
# whose URL is this sentinel — the same value the site UI writes when a
# "Label" is added to the navigation.
NAV_LABEL_SENTINEL_URL = "http://linkless.header/"


def _get_quicklaunch_nodes(token: str, site_url: str) -> list[dict]:
    """Return the site's top-level quick-launch nodes with their children."""
    rest_url = _get_sharepoint_rest_url(site_url)
    url = f"{rest_url}/web/navigation/quicklaunch?$expand=Children"
    resp = _request_with_retry("GET", url, headers=_sharepoint_headers(token))
    _raise_for_status(resp)
    return resp.json()["d"]["results"]


def _create_nav_node(
    token: str, site_url: str, title: str, node_url: str,
    *, parent_id: int | None = None, is_external: bool = False,
) -> dict:
    """Create a quick-launch navigation node; returns the created node.

    With *parent_id* the node is nested under that node, otherwise it is
    appended at the top level.
    """
    rest_url = _get_sharepoint_rest_url(site_url)
    if parent_id is None:
        url = f"{rest_url}/web/navigation/quicklaunch"
    else:
        url = f"{rest_url}/web/navigation/GetNodeById({parent_id})/Children"
    body = {
        "__metadata": {"type": "SP.NavigationNode"},
        "Title": title,
        "Url": node_url,
        "IsExternal": is_external,
    }
    # Not repeated after a failure that may have created the node, which
    # would duplicate the link.
    resp = _request_with_retry(
        "POST", url, headers=_sharepoint_post_headers(token), json=body, idempotent=False,
    )
    _raise_for_status(resp)
    return resp.json()["d"]


def _set_nav_node_title(token: str, site_url: str, node_id: int, title: str) -> None:
    """Update the display title of an existing navigation node."""
    rest_url = _get_sharepoint_rest_url(site_url)
    headers = _sharepoint_post_headers(token)
    headers["X-HTTP-Method"] = "MERGE"
    body = {"__metadata": {"type": "SP.NavigationNode"}, "Title": title}
    resp = _request_with_retry(
        "POST", f"{rest_url}/web/navigation/GetNodeById({node_id})",
        headers=headers, json=body,
    )
    _raise_for_status(resp)


def _normalise_nav_url(url_or_path: str) -> str:
    """Normalise a navigation URL for comparison (server-relative, decoded, lowercase).

    Decodes to a fixpoint: navigation nodes store the once-decoded path this
    module created them with, while Graph webUrl arrives percent-encoded, so
    the two sides can reach this function with different residual encoding.
    Repeated decoding makes every form converge on one key (at the
    theoretical cost of colliding two codes that differ only by encoding).
    """
    path = urlparse(url_or_path).path if "://" in url_or_path else url_or_path
    for _ in range(5):
        decoded = unquote(path)
        if decoded == path:
            break
        path = decoded
    return path.rstrip("/").lower()


def _find_nav_label(
    token: str, site_url: str, label: str,
) -> tuple[int, dict[str, dict]] | None:
    """Find a top-level quick-launch node titled *label*.

    Returns ``(label_node_id, children keyed by normalised URL)`` or None.
    Matching is by title only, so an existing ordinary link with the same
    title is treated as the label rather than duplicated.
    """
    for node in _get_quicklaunch_nodes(token, site_url):
        if (node.get("Title") or "").strip().lower() == label.strip().lower():
            children = (node.get("Children") or {}).get("results") or []
            return (
                node["Id"],
                {_normalise_nav_url(c["Url"]): c for c in children if c.get("Url")},
            )
    return None


def _ensure_nav_label(
    token: str, site_url: str, label: str,
) -> tuple[int, dict[str, dict], bool]:
    """Ensure a top-level quick-launch label node titled *label* exists.

    Returns ``(label_node_id, children keyed by normalised URL, created)``.
    """
    found = _find_nav_label(token, site_url, label)
    if found is not None:
        return found[0], found[1], False
    node = _create_nav_node(token, site_url, label, NAV_LABEL_SENTINEL_URL, is_external=True)
    logger.info("Created navigation label '%s' on %s", label, site_url)
    return node["Id"], {}, True


def _delete_nav_node(token: str, site_url: str, node_id: int) -> None:
    """Delete a navigation node by ID (children, if any, are removed with it)."""
    rest_url = _get_sharepoint_rest_url(site_url)
    headers = _sharepoint_post_headers(token)
    headers["IF-MATCH"] = "*"
    headers["X-HTTP-Method"] = "DELETE"
    resp = _request_with_retry(
        "POST", f"{rest_url}/web/navigation/GetNodeById({node_id})",
        headers=headers, data="",
    )
    _raise_for_status(resp)


def _ensure_library_nav_link(
    token: str,
    site_url: str,
    library_code: str,
    display_name: str,
    library_web_url: str | None,
    label_node_id: int,
    label_children: dict[str, dict],
    action_log: _ThreadSafeLog,
) -> None:
    """Ensure a library is linked under the navigation label node.

    Adds the link when missing; when present but titled differently (e.g.
    after a faculty rename) the node title is updated to *display_name*.
    *label_children* is the pre-fetched children map from _ensure_nav_label —
    workers only read it, so it is safe to share across threads.
    """
    if not library_web_url:
        raise ValueError(
            f"Library '{library_code}' has no webUrl — cannot add navigation link"
        )
    nav_url = unquote(urlparse(library_web_url).path)
    existing = label_children.get(_normalise_nav_url(nav_url))

    if existing is None:
        _create_nav_node(token, site_url, display_name, nav_url, parent_id=label_node_id)
        logger.info(
            "Added '%s' to site navigation on %s", display_name, site_url,
        )
        _append_log(
            action_log, site_url, library_code, "navigation_link_added",
            f"linked as '{display_name}' under navigation label",
        )
    else:
        _sync_nav_link_title(
            token, site_url, library_code, display_name, library_web_url,
            (label_node_id, label_children), action_log,
        )


def _sync_nav_link_title(
    token: str,
    site_url: str,
    library_code: str,
    display_name: str,
    library_web_url: str | None,
    label_found: tuple[int, dict[str, dict]] | None,
    action_log: _ThreadSafeLog,
) -> None:
    """Update a library's navigation link title to *display_name*, if linked.

    Title-sync only — a missing label or link is a no-op (adding links is the
    create path's job).  *label_found* is a per-site _find_nav_label result.
    """
    if not library_web_url or label_found is None:
        return
    _, children = label_found
    node = children.get(_normalise_nav_url(library_web_url))
    if node is None or node.get("Title") == display_name:
        return
    _set_nav_node_title(token, site_url, node["Id"], display_name)
    logger.info(
        "Renamed navigation link for '%s' to '%s' on %s",
        library_code, display_name, site_url,
    )
    _append_log(
        action_log, site_url, library_code, "navigation_link_renamed",
        f"navigation link title updated to '{display_name}'",
    )


def _remove_library_nav_link(
    token: str,
    site_url: str,
    library_code: str,
    library_web_url: str | None,
    label: str,
    label_found: tuple[int, dict[str, dict]] | None,
    action_log: _ThreadSafeLog,
) -> None:
    """Remove a library's link from under the navigation label, if present.

    *label_found* is a _find_nav_label result, fetched once per site by the
    caller rather than per library.  A no-op when the label or the link does
    not exist — the label itself is never created (or removed) here.  An
    emptied label is left in place for future libraries.
    """
    if not library_web_url or label_found is None:
        return
    _, children = label_found
    node = children.get(_normalise_nav_url(library_web_url))
    if node is None:
        return
    _delete_nav_node(token, site_url, node["Id"])
    logger.info(
        "Removed navigation link for '%s' from %s", library_code, site_url,
    )
    _append_log(
        action_log, site_url, library_code, "navigation_link_removed",
        f"removed link from under '{label}' navigation label",
    )
