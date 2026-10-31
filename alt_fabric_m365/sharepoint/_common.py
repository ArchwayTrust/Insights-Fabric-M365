"""
Constants, the action log and input validation shared by the SharePoint library operations.
"""
import logging
import threading
from collections.abc import Callable, Iterable
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import pandas as pd

from ._http import _has_ambiguous_host

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Permission levels accepted in input DataFrames (ordered for deterministic
# iteration when reconciling every role).
VALID_ROLES = ("read", "contribute", "owner")

# Entra ID replication — newly created groups may not be visible to
# downstream services (Graph member endpoints, SharePoint ensureuser)
# for several seconds.  These control the polling wait.
REPLICATION_MAX_ATTEMPTS = 6
REPLICATION_BASE_DELAY = 5.0  # seconds

# Group role suffixes.  Auto-* groups are owned by this automation (membership
# synced and emptied on unlink).  Manual-* groups are hand-curated: renamed
# alongside the library but NEVER emptied or synced by automation.
AUTO_GROUP_SUFFIXES = ["Auto-Read", "Auto-Contribute", "Auto-Owner"]
MANUAL_GROUP_SUFFIXES = ["Manual-Read", "Manual-Contribute"]
ALL_GROUP_SUFFIXES = AUTO_GROUP_SUFFIXES + MANUAL_GROUP_SUFFIXES

# Property bag keys used to store library-level provenance metadata on each
# library's root folder.  These are invisible to end users and only accessible
# via the SharePoint REST API.
# Maps DataFrame column name → property bag key name.  PascalCase avoids
# the _x005f_ encoding that SharePoint REST applies to underscores.
METADATA_PROPERTY_KEYS = {
    "source_system": "SourceSystem",
    "source_id": "SourceId",
}

# The type each SharePointLibraryManager option must have.  Notebook and
# pipeline parameters can arrive as strings, and 'False' is truthy, so an
# option's type is checked rather than its truthiness used.
OPTION_TYPES = {
    "dry_run": bool, "max_workers": int, "include_site_name": bool,
    "create_onedrive_shortcuts": bool, "max_removals_per_group": int,
    "allow_mass_removal": bool, "add_to_navigation": bool, "navigation_label": str,
}

# Columns of every action-log DataFrame the operations return.
ACTION_LOG_COLUMNS = [
    "site_url", "document_library", "group", "user",
    "action", "permission", "status",
]

# Actions that mean something failed, so a run should be treated as failed
# (and pipeline alerting fire) when any are present.
FATAL_ACTIONS = frozenset({"error", "add_failed", "remove_failed"})

# Actions that deserve attention but do not mean the run failed.
# library_not_found, library_name_conflict, group_name_collision,
# group_not_found and skip (an unknown permission) are persistent data
# problems: failing the run could not fix them, it would only wedge every
# future run behind them.
WARNING_ACTIONS = frozenset({
    "user_not_found", "removal_guardrail_triggered", "removals_skipped",
    "onedrive_shortcut_failed", "library_not_found", "library_name_conflict",
    "group_name_collision", "group_not_found", "group_not_owned", "retire_deferred",
    "skip",
})

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Thread-safe action log
# ---------------------------------------------------------------------------


class _ThreadSafeLog:
    """Append-only log safe for concurrent use from multiple threads."""

    def __init__(self):
        self._entries: list[dict] = []
        self._lock = threading.Lock()

    def append(self, entry: dict) -> None:
        with self._lock:
            self._entries.append(entry)

    def to_list(self) -> list[dict]:
        with self._lock:
            return list(self._entries)


def _append_log(
    action_log: _ThreadSafeLog,
    site_url: str,
    library: str,
    action: str,
    status: str,
    *,
    group: str = "",
    user: str = "",
    permission: str = "",
) -> None:
    """Append a standard action-log entry (single place the schema is defined)."""
    action_log.append({
        "site_url": site_url,
        "document_library": library,
        "group": group,
        "user": user,
        "action": action,
        "permission": permission,
        "status": status,
    })


def _validate_required_columns(df: pd.DataFrame, required_columns: set[str]) -> pd.DataFrame:
    """Validate that *df* contains *required_columns* with no nulls.

    Returns a copy of *df* with required columns cast to string dtype.
    """
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"DataFrame is missing required columns: {missing}")

    df = df.copy()
    for col in required_columns:
        if df[col].isna().any():
            nan_count = int(df[col].isna().sum())
            raise ValueError(
                f"Column '{col}' contains {nan_count} null value(s) — "
                f"all required columns must be non-null strings"
            )
        if not pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype(str)
    return df


def _check_option(name: str, value: object) -> bool | int | str:
    """Return a manager option's value, after checking its type.

    Args:
        name: The option, a key of OPTION_TYPES.
        value: The value given for it.

    Returns:
        The value as a bool, int or str. numpy's booleans and integers, as pandas returns them, are converted.

    Raises:
        TypeError: If the value isn't of the option's type, such as the string 'False' for a True/False option.
    """
    expected = OPTION_TYPES[name]
    if expected is bool and pd.api.types.is_bool(value):
        return bool(value)
    if expected is int and pd.api.types.is_integer(value) and not pd.api.types.is_bool(value):
        return int(value)
    if expected is str and isinstance(value, str):
        return value
    wanted = {bool: "True or False", int: "a whole number", str: "a string"}[expected]
    raise TypeError(f"{name} must be {wanted}, not {value!r}")


def _fill_flag_columns(df: pd.DataFrame, defaults: dict[str, bool]) -> pd.DataFrame:
    """Return *df* with each optional True/False column holding bools, and blanks set to its default.

    bool() of a blank is True for NaN, False for None and an error for pd.NA,
    and bool('False') is True, so values are checked rather than converted.

    Args:
        df: The input DataFrame. It isn't changed.
        defaults: Each optional column's default, used where the column is missing or a value is blank.

    Returns:
        A copy of *df* with every column in *defaults*, holding only True and False.

    Raises:
        ValueError: If a column holds anything other than True, False, 1, 0 or a blank.
    """
    df = df.copy()
    for column, default in defaults.items():
        if column not in df.columns:
            df[column] = default
            continue
        values = []
        for value in df[column]:
            if pd.isna(value):
                values.append(default)
            elif not isinstance(value, str) and value in (0, 1):
                values.append(bool(value))
            else:
                raise ValueError(f"Column '{column}' must hold True, False or blanks, not {value!r}")
        df[column] = pd.Series(values, index=df.index, dtype=bool)
    return df


def _validate_site_urls(site_urls: Iterable[str], sharepoint_hostname: str | None = None) -> None:
    """Raise ValueError if any URL is malformed.

    With *sharepoint_hostname*, each URL must also be https on that host:
    the SharePoint bearer token is sent to the site URL, so a URL on any
    other host (or over plain http) would hand the token to it.  So would a
    URL that requests reads as being on another host, see _has_ambiguous_host.
    """
    for site_url in site_urls:
        parsed = urlparse(site_url)
        if parsed.scheme not in ("http", "https") or not parsed.hostname:
            raise ValueError(
                f"Malformed site_url: '{site_url}' "
                f"— expected https://{{tenant}}.sharepoint.com/sites/{{name}}"
            )
        if sharepoint_hostname is not None and (
            parsed.scheme != "https"
            or _has_ambiguous_host(site_url)
            or parsed.hostname.lower() != sharepoint_hostname.lower()
        ):
            raise ValueError(
                f"site_url '{site_url}' is not on https://{sharepoint_hostname} — "
                f"the SharePoint token is only sent to that host"
            )


def _action_log_to_dataframe(action_log: _ThreadSafeLog) -> pd.DataFrame:
    """Convert a ThreadSafeLog to the standard action-log DataFrame."""
    return pd.DataFrame(
        action_log.to_list(),
        columns=ACTION_LOG_COLUMNS,
    )


def _run_parallel(
    func: Callable[..., None], items: Iterable[tuple], max_workers: int, action_log: _ThreadSafeLog,
) -> None:
    r"""Run *func(\*item, action_log=action_log)* over *items* sequentially or in parallel.

    Both paths have identical failure semantics: an unhandled exception in one
    item is logged as an 'error' row and the remaining items still run.  An
    interrupt, such as cancelling the notebook cell, stops both: the items
    already running finish, and the rest are never started.
    """
    if max_workers <= 1:
        for item in items:
            try:
                func(*item, action_log=action_log)
            except Exception as exc:
                logger.error("Unhandled error: %s — %s", item, exc)
                _append_log(action_log, "", "", "error", str(exc))
    else:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for item in items:
                future = executor.submit(func, *item, action_log=action_log)
                futures[future] = item
            try:
                for future in as_completed(futures):
                    exc = future.exception()
                    if exc:
                        item = futures[future]
                        logger.error("Unhandled error: %s — %s", item, exc)
                        _append_log(action_log, "", "", "error", str(exc))
            except BaseException:
                # Leaving the with block would wait for every queued item, so a
                # cancelled cell would carry on changing the tenant
                logger.warning("Interrupted — the items already running finish, the rest are skipped")
                executor.shutdown(wait=True, cancel_futures=True)
                raise
