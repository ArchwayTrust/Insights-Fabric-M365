"""
SharePointLibraryManager: manages SharePoint document libraries and their permissions through Entra ID security groups.
"""
import logging
import threading
import time
from collections.abc import Iterable
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests

from ..auth import Scopes
from ._common import (
    ALL_GROUP_SUFFIXES, AUTO_GROUP_SUFFIXES, METADATA_PROPERTY_KEYS, REPLICATION_BASE_DELAY, VALID_ROLES,
    _action_log_to_dataframe, _append_log, _blank_nulls, _check_option, _fill_flag_columns, _run_parallel, _ThreadSafeLog,
    _validate_library_codes, _validate_required_columns, _validate_site_urls,
)
from ._graph import (
    _create_onedrive_shortcut, _ensure_document_library, _find_library_shortcut, _find_title_conflict,
    _get_library_display_name, _get_list_id, _get_onedrive_root_items,
    _get_site_id, _list_existing_lists, _set_library_display_name,
)
from ._http import _RenewableToken
from ._groups import (
    GroupResolutionError, UserResolutionError, _add_group_member, _create_security_group, _empty_group_members,
    _ensure_app_owns_group, _find_existing_library_collisions, _find_group_by_name,
    _find_group_name_collisions, _find_groups_by_name_prefix, _get_group_members, _get_unique_group,
    _log_group_name_collisions, _remove_group_member, _resolve_nestable_group,
    _resolve_user_with_fallback, _set_group_description, _token_object_id, build_group_name,
)
from ._navigation import (
    _ensure_library_nav_link, _ensure_nav_label, _find_nav_label, _normalise_nav_url, _remove_library_nav_link,
    _sync_nav_link_title,
)
from ._sharepoint import (
    _add_role_assignment, _break_role_inheritance, _disable_site_sharing, _ensure_group_assigned_to_library,
    _ensure_library_metadata, _get_library_group_roles, _get_library_properties, _get_library_role_assignments,
    _get_site_sharing_settings, _group_claims, _has_unique_permissions, _remove_library_role_assignment,
    _resolve_group_principal_id, _resolve_role_def_ids,
)

logger = logging.getLogger(__name__)


def _site_and_code(item: tuple) -> tuple[str, str]:
    """Return a per-library work item's site URL and library code, for _run_parallel.

    Args:
        item: The positional arguments of a per-library _process_* method, which start (site_url, site_id,
            library_code).

    Returns:
        The site URL and library code.
    """
    return item[0], item[2]


class SharePointLibraryManager:
    """Manage SharePoint document library permissions via Entra ID security groups.

    Authenticate once in the constructor, then call any operation method.
    Instance-level defaults for ``dry_run``, ``max_workers``,
    ``include_site_name``, and ``create_onedrive_shortcuts`` can be
    overridden per call.

    The app registration needs these application permissions, with admin
    consent: Microsoft Graph ``Sites.FullControl.All`` (or
    ``Sites.Selected`` with full control on each site),
    ``Group.ReadWrite.All``, ``GroupMember.ReadWrite.All`` and
    ``User.Read.All``; SharePoint ``Sites.FullControl.All`` (or
    ``Sites.Selected``).  OneDrive shortcuts also need Graph
    ``Files.ReadWrite.All``.  SharePoint only accepts app-only tokens
    obtained with a certificate, not a client secret.

    Args:
        credential:           Credential with a ``get_token(*scopes)`` method,
                              normally ``alt_fabric_m365.auth.KeyVaultCertificateCredential``,
                              created without a scope: Graph and SharePoint
                              tokens are requested from it as needed.  It must
                              use a certificate, as SharePoint rejects app-only
                              tokens obtained with a client secret.
        sharepoint_hostname:  e.g. ``"contoso.sharepoint.com"``.
        dry_run:              Default dry-run mode for all operations.
        max_workers:          Default parallelism for all operations.
        include_site_name:    Include the site name in security group names.
        create_onedrive_shortcuts:
            Default for ``update_memberships``.  When True, a OneDrive
            shortcut is created for each newly added user.
        max_removals_per_group:
            Guardrail for ``update_memberships``: when a single group's sync
            would remove more members than this, the removals are skipped and
            logged as ``removal_guardrail_triggered`` (adds still proceed).
        max_removals_per_run:
            Guardrail for ``update_memberships``: when one call would remove
            more members than this across all its groups, as an empty or
            truncated members table would, no members are removed and one
            ``removal_guardrail_triggered`` row is logged (adds still proceed).
        allow_mass_removal:
            When True, disables both removal guardrails for the run.
        add_to_navigation:
            Default for ``create_libraries``, ``rename_libraries``,
            ``unlink_libraries`` and ``retire_moved_libraries``.  When True,
            each library is linked into the site's quick-launch navigation,
            nested under the ``navigation_label`` heading (created when
            missing); unlinking removes the link again.
        navigation_label:
            Title of the navigation label the library links are nested under.
        service_principal_id:
            Object ID of the app's service principal, which is made owner of
            the groups the manager creates.  Read from the Graph token's
            ``oid`` claim when not given.
    """

    def __init__(
        self,
        credential,
        sharepoint_hostname: str,
        *,
        dry_run: bool = False,
        max_workers: int = 1,
        include_site_name: bool = True,
        create_onedrive_shortcuts: bool = False,
        max_removals_per_group: int = 10,
        max_removals_per_run: int = 100,
        allow_mass_removal: bool = False,
        add_to_navigation: bool = False,
        navigation_label: str = "Document Libraries",
        service_principal_id: str | None = None,
    ):
        self.dry_run = _check_option("dry_run", dry_run)
        self.max_workers = _check_option("max_workers", max_workers)
        self.include_site_name = _check_option("include_site_name", include_site_name)
        self.create_onedrive_shortcuts = _check_option("create_onedrive_shortcuts", create_onedrive_shortcuts)
        self.max_removals_per_group = _check_option("max_removals_per_group", max_removals_per_group)
        self.max_removals_per_run = _check_option("max_removals_per_run", max_removals_per_run)
        self.allow_mass_removal = _check_option("allow_mass_removal", allow_mass_removal)
        self.add_to_navigation = _check_option("add_to_navigation", add_to_navigation)
        self.navigation_label = _check_option("navigation_label", navigation_label)
        self._service_principal_id = service_principal_id
        # Groups this manager has created, by lower-case name, with whether
        # the app was made their owner: Graph's name lookups can miss a group
        # for a while after it's created, and creating it again would make a
        # duplicate that blocks its library.
        self._created_groups: dict[str, tuple[dict, bool]] = {}
        self._created_groups_lock = threading.Lock()

        # Tokens are requested per use via the properties below, and again
        # before each attempt of a request (see _RenewableToken).  The
        # credential caches them and requests new ones near expiry, so runs
        # longer than a token lifetime (~60-90 min), and requests retried for
        # minutes while throttled, keep working.
        self.credential = credential
        self.sharepoint_hostname = sharepoint_hostname
        self._graph_scope = Scopes.GRAPH
        self._sp_scope = f"https://{sharepoint_hostname}/.default"

        logger.info("Acquiring access tokens …")
        _ = self._graph_token  # fail fast on bad credentials
        _ = self._sp_token
        logger.info("Tokens acquired successfully.")

    @property
    def service_principal_id(self) -> str:
        """Object ID of the app's service principal (see the constructor)."""
        if self._service_principal_id is None:
            self._service_principal_id = _token_object_id(self._graph_token)
        return self._service_principal_id

    def _remember_created_group(self, name: str, group: dict, *, app_owned: bool) -> None:
        """Record a group this manager created, so a lagging name lookup can't lead to a duplicate.

        Args:
            name: The group's display name.
            group: The group resource Graph returned when creating it.
            app_owned: Whether the group was created with the app as its owner.
        """
        with self._created_groups_lock:
            self._created_groups[name.lower()] = (group, app_owned)

    def _created_group(self, name: str) -> tuple[dict, bool] | None:
        """Return the group this manager created with *name*, if any (see _remember_created_group).

        Only for when a name lookup misses the group: once the lookup finds it,
        the group is checked like any other existing group.

        Args:
            name: The group's display name.

        Returns:
            The group resource and whether it was created with the app as its owner, or None.
        """
        with self._created_groups_lock:
            return self._created_groups.get(name.lower())

    def _check_group_owner(
        self, group_id: str, group_name: str, site_url: str, library_code: str,
        permission: str, action_log: _ThreadSafeLog, *, dry_run: bool = False, is_new_group: bool = False,
    ) -> bool:
        """Return whether an existing group may be granted access or synced.

        Group names are predictable, so a same-named group someone else
        created must not be adopted (see _ensure_app_owns_group).  A group
        with no owner gets the app added as owner and is logged.  A dry run
        only reads the owners, so it reports what a live run would do.  With
        *is_new_group*, a 404 while the group replicates is waited out.
        """
        outcome = _ensure_app_owns_group(
            self._graph_token, group_id, self.service_principal_id, dry_run=dry_run, is_new_group=is_new_group,
        )
        if outcome == "adopted":
            _append_log(
                action_log, site_url, library_code,
                "group_owner_added (dry run)" if dry_run else "group_owner_added",
                "the group has no owner — the app registration would be added as owner" if dry_run
                else "the group had no owner — the app registration was added as owner",
                group=group_name, permission=permission,
            )
        elif outcome == "not_owned":
            logger.warning(
                "Group '%s' is owned by someone other than the app — not used for '%s'",
                group_name, library_code,
            )
            _append_log(
                action_log, site_url, library_code, "group_not_owned",
                "the group's owners don't include the app registration, so it may not be the "
                "automation's group — not granted or synced; check who created it, then add the "
                "app as an owner or delete it",
                group=group_name, permission=permission,
            )
        return outcome != "not_owned"

    @property
    def _graph_token(self) -> str:
        return _RenewableToken(lambda: self.credential.get_token(self._graph_scope).token)

    @property
    def _sp_token(self) -> str:
        return _RenewableToken(lambda: self.credential.get_token(self._sp_scope).token)

    def _resolve_config(self, **overrides) -> tuple:
        """Resolve per-call overrides against instance defaults.

        Each keyword must name an instance attribute; None means "use the
        instance default".  Returns the effective values in keyword order,
        so callers unpack positionally:

            dry_run_, nav_ = self._resolve_config(
                dry_run=dry_run, add_to_navigation=add_to_navigation)

        Adding a per-call override to an operation is therefore one keyword
        here, not a copied resolution block.

        Each value's type is checked (see _check_option), including an
        instance default changed after construction, so a string such as
        'False' raises TypeError rather than counting as True.
        """
        return tuple(
            _check_option(name, value if value is not None else getattr(self, name))
            for name, value in overrides.items()
        )

    def _prefetch_sites(
        self,
        site_urls: Iterable[str],
        action_log: _ThreadSafeLog,
        *,
        role_definitions: bool = False,
        navigation_label: str | None = None,
        gone: tuple[str, str] | None = None,
    ) -> dict[str, dict]:
        """Read each site's state once, before the per-library work.

        Returns site_url -> dict with ``site_id``, ``lists`` (a
        _list_existing_lists result) and, when asked for, ``role_def_ids``
        and ``nav_label`` (a _find_nav_label result).  Reads only, so it
        runs in dry runs too.

        A site that can't be read gets one ``error`` row and is left out, so
        only its libraries are skipped and the rest of the operation
        proceeds; the watermark is held back, so they are retried next run.
        *gone* is an (action, status) row logged instead when a site doesn't
        exist (HTTP 404).  A 403 isn't one: the site exists, but the app
        can't read it, as when its Sites.Selected grant has been removed.
        """
        sites: dict[str, dict] = {}
        for site_url in dict.fromkeys(site_urls):
            try:
                try:
                    site_id = _get_site_id(self._graph_token, site_url)
                except requests.HTTPError as exc:
                    status = exc.response.status_code if exc.response is not None else None
                    if gone is not None and status == 404:
                        _append_log(action_log, site_url, "", *gone)
                        continue
                    raise
                site = {"site_id": site_id, "lists": _list_existing_lists(self._graph_token, site_id)}
                if role_definitions:
                    site["role_def_ids"] = _resolve_role_def_ids(self._sp_token, site_url)
                if navigation_label is not None:
                    site["nav_label"] = _find_nav_label(self._sp_token, site_url, navigation_label)
            except (requests.RequestException, ValueError, KeyError) as exc:
                logger.error("Error reading site state for '%s': %s", site_url, exc)
                _append_log(action_log, site_url, "", "error", str(exc))
                continue
            sites[site_url] = site
        return sites

    def _resolve_users(self, user_principal_names: Iterable[str], max_workers: int) -> dict[str, tuple]:
        """Resolve each distinct user once, in parallel unless max_workers is 1.

        Staff in several faculties would otherwise cost up to three Graph
        lookups per library.  Keys are the normalised (stripped, lowercase)
        input; values are (object_id, real_upn, failure_reason, fatal).
        A user that doesn't exist is not fatal (user_not_found); an API
        failure is, and fails the libraries that use the user.

        Args:
            user_principal_names: The users to resolve.
            max_workers: The operation's max_workers. 1 or less resolves one
                user at a time, as it runs one library at a time; otherwise up
                to 8 lookups run at once.

        Returns:
            Normalised user principal name -> (object_id, real_upn, failure_reason, fatal).
        """
        graph_token = self._graph_token

        def resolve(upn: str) -> tuple[str, tuple]:
            try:
                oid, real_upn = _resolve_user_with_fallback(graph_token, upn)
            except UserResolutionError as exc:
                return upn, (None, None, exc.reason, False)
            except requests.HTTPError as exc:
                if exc.response is not None and exc.response.status_code == 404:
                    return upn, (None, None, "not found in Entra ID (404)", False)
                return upn, (None, None, f"user lookup failed: {exc}", True)
            except requests.RequestException as exc:
                return upn, (None, None, f"user lookup failed: {exc}", True)
            if real_upn != upn:
                logger.info("Aliased domain: '%s' resolved to '%s'", upn, real_upn)
            return upn, (oid, real_upn, None, False)

        upns = sorted({upn.strip().lower() for upn in user_principal_names})
        if max_workers <= 1:
            return dict(map(resolve, upns))
        with ThreadPoolExecutor(max_workers=min(8, max(len(upns), 1))) as executor:
            return dict(executor.map(resolve, upns))

    # -------------------------------------------------------------------
    # create_libraries
    # -------------------------------------------------------------------

    def _process_create_library(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        library_display_name: str,
        group_prefix: str,
        auto_groups: bool,
        manual_read: bool,
        manual_contribute: bool,
        app_owned: bool,
        source_system: str | None,
        source_id: str | None,
        dry_run: bool,
        include_site_name: bool,
        role_def_ids: dict[str, int],
        existing_lists: dict[str, dict],
        nav: dict | None,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Create a single library with its optional Auto and Manual groups.

        With *app_owned*, new groups are created with the app as owner and an
        existing group is only used if the app owns it.  Without it (groups
        owned by people, such as the SLT groups), ownership is left alone.

        *nav* is None when navigation linking is off (or the site's label
        could not be ensured); otherwise a dict with ``label``, ``node_id``
        and ``children`` (the latter two None in dry-run).
        """
        logger.info("Creating library '%s' (%s) on site '%s' …", library_display_name, library_code, site_url)

        try:
            # Step 1: Ensure the library exists (created with code, display name set in step 7)
            existing_list = existing_lists.get(library_code.lower())
            if existing_list is None:
                conflict = _find_title_conflict(existing_lists, library_code)
                if conflict is not None:
                    # Creating it would fail on every run (titles are unique
                    # per site), and adopting the other library would hand
                    # it this faculty's groups.
                    logger.warning(
                        "Library '%s' not created on '%s': another library (%s) has that title",
                        library_code, site_url, conflict.get("webUrl"),
                    )
                    _append_log(
                        action_log, site_url, library_code, "library_name_conflict",
                        f"another library ({conflict.get('webUrl')}) is already titled "
                        f"'{library_code}' — rename it so this library can be created",
                    )
                    return
            if dry_run:
                list_id = None
                lib_web_url = None
                logger.info("[DRY RUN] Would ensure library '%s' (%s)", library_display_name, library_code)
                if existing_list is None:
                    _append_log(
                        action_log, site_url, library_code, "library_created (dry run)",
                        f"would create the library, with '{library_code}' as its URL",
                    )
            else:
                lib = _ensure_document_library(self._graph_token, site_id, library_code, existing=existing_lists)
                list_id = lib["id"]
                lib_web_url = lib.get("webUrl", "")
                if existing_list is None:
                    _append_log(action_log, site_url, library_code, "library_created", f"created at {lib_web_url}")

            # Step 2: Break role inheritance.  The library's existing group
            # bindings are read here so step 6 only writes missing ones.
            existing_roles: dict[str, set[int]] = {}
            if not dry_run:
                if _has_unique_permissions(self._sp_token, site_url, list_id):
                    logger.info(
                        "Library '%s' already has unique permissions — skipping inheritance break",
                        library_code,
                    )
                    existing_roles = _get_library_group_roles(self._sp_token, site_url, list_id)
                else:
                    # copyRoleAssignments=false: no group holds a role yet
                    _break_role_inheritance(self._sp_token, site_url, list_id)
                    # Taking over a library people already use takes away the
                    # access site members had through the site, so record it.
                    if existing_list is not None:
                        _append_log(
                            action_log, site_url, library_code, "inheritance_broken",
                            "stopped inheriting the site's permissions: only its groups, and any folders "
                            "or files with their own permissions, give access now",
                        )
            elif existing_list is not None and not _has_unique_permissions(
                self._sp_token, site_url, existing_list["id"],
            ):
                _append_log(
                    action_log, site_url, library_code, "inheritance_broken (dry run)",
                    "would stop it inheriting the site's permissions",
                )

            # Step 3: Build the full list of groups to create
            # Group names use the library code for stability
            all_groups: list[tuple[str, str, str, str]] = []
            if auto_groups:
                all_groups.append(("Auto-Read", "read", "Read", "group_assigned"))
                all_groups.append(("Auto-Contribute", "contribute", "Contribute", "group_assigned"))
            if manual_read:
                all_groups.append(("Manual-Read", "read", "Read", "manual_read_group_ensured"))
            if manual_contribute:
                all_groups.append(("Manual-Contribute", "contribute", "Contribute", "manual_contribute_group_ensured"))

            # Step 4: Create all Entra ID groups first (before any assignment)
            # This lets Entra replication proceed for ALL groups in parallel,
            # rather than waiting sequentially for each group.
            created_groups: list[tuple[str, str, str, str, str, bool]] = []
            for suffix, role_key, role_label, action_name in all_groups:
                group_display_name = build_group_name(
                    group_prefix, site_url, library_code, suffix,
                    include_site_name=include_site_name,
                )
                is_manual = suffix.startswith("Manual-")
                group_description = (
                    f"Manual {role_label} access to '{library_display_name}' ({library_code}) "
                    f"on {site_url} — membership managed manually, NOT synced by automation"
                ) if is_manual else (
                    f"{role_label} access to '{library_display_name}' ({library_code}) "
                    f"on {site_url} — managed by SharePoint Library Permissions Manager"
                )

                if dry_run:
                    # Ownership is only read, so the dry run shows groups a
                    # live run would refuse or adopt.  Groups people own are
                    # looked up too, as a duplicated name fails a live run.
                    existing_group = _find_group_by_name(self._graph_token, group_display_name)
                    if existing_group and app_owned and not self._check_group_owner(
                        existing_group["id"], group_display_name, site_url, library_code, role_key, action_log,
                        dry_run=True,
                    ):
                        continue
                    dry_action = (
                        f"ensure_{suffix.lower().replace('-', '_')}_group (dry run)"
                        if is_manual else "ensure_group (dry run)"
                    )
                    dry_status = (
                        f"would ensure {suffix} group exists and is assigned to library"
                        if is_manual else "would ensure group exists and is assigned to library"
                    )
                    logger.info("[DRY RUN] Would ensure group '%s'", group_display_name)
                    _append_log(
                        action_log, site_url, library_code, dry_action, dry_status,
                        group=group_display_name, permission=role_key,
                    )
                else:
                    group = _find_group_by_name(self._graph_token, group_display_name)
                    created = None if group else self._created_group(group_display_name)
                    if created is not None:
                        # Made earlier in the run and not yet found by name, so
                        # still replicating.  The app owns it if it was made
                        # its owner; otherwise it's checked, as a group found
                        # by name would be.  Its description is checked once
                        # the lookup finds it.
                        (group, created_app_owned), is_new = created, True
                        if app_owned and not created_app_owned and not self._check_group_owner(
                            group["id"], group_display_name, site_url, library_code, role_key, action_log,
                            is_new_group=True,
                        ):
                            continue
                    elif group is None:
                        group = _create_security_group(
                            self._graph_token, group_display_name, group_description,
                            self.service_principal_id if app_owned else None,
                        )
                        self._remember_created_group(group_display_name, group, app_owned=app_owned)
                        is_new = True
                    else:
                        logger.info("Security group '%s' already exists (id=%s)", group_display_name, group["id"])
                        is_new = False
                    group_id = group["id"]
                    if app_owned and not is_new:
                        if not self._check_group_owner(
                            group_id, group_display_name, site_url, library_code, role_key, action_log,
                        ):
                            continue
                        # The name follows the code, which never changes, but the
                        # description names the library, so it follows a rename.
                        # Groups people own (app_owned False) are left as they are.
                        if group.get("description") != group_description:
                            _set_group_description(self._graph_token, group_id, group_description)
                            _append_log(
                                action_log, site_url, library_code, "group_description_updated",
                                f"description set to: {group_description}",
                                group=group_display_name, permission=role_key,
                            )
                    created_groups.append((group_display_name, suffix, role_key, action_name, group_id, is_new))

            # Step 5: If any groups were newly created, wait once for Entra ID
            # replication before attempting SharePoint ensureuser for any of them.
            # This single wait replaces the per-group sequential waits.
            if not dry_run:
                any_new = any(is_new for _, _, _, _, _, is_new in created_groups)
                if any_new:
                    logger.info(
                        "Waiting %.0fs for Entra ID replication before assigning groups …",
                        REPLICATION_BASE_DELAY,
                    )
                    time.sleep(REPLICATION_BASE_DELAY)

            # Step 6: Assign all groups to the library (only missing bindings)
            for group_display_name, suffix, role_key, action_name, group_id, is_new in created_groups:
                role_def_id = role_def_ids[role_key]
                granted = _ensure_group_assigned_to_library(
                    self._sp_token, site_url, list_id,
                    group_id, group_display_name, role_def_id,
                    is_new_group=is_new, existing_roles=existing_roles,
                )
                is_manual = suffix.startswith("Manual-")
                if granted:
                    _append_log(
                        action_log, site_url, library_code, action_name,
                        "success (membership not synced)" if is_manual else "success",
                        group=group_display_name, permission=role_key,
                    )
                else:
                    _append_log(
                        action_log, site_url, library_code, "already_assigned", "no change needed",
                        group=group_display_name, permission=role_key,
                    )

            # Step 7: Set the display name.  After the permissions on purpose:
            # if renaming fails, the library already has its own permissions
            # and groups rather than still inheriting the site's.  A new
            # library is titled with its code until then.
            current_title = existing_list["displayName"] if existing_list is not None else library_code
            if library_display_name and library_display_name != current_title:
                title_conflict = _find_title_conflict(existing_lists, library_code, title=library_display_name)
                if title_conflict is not None and not dry_run:
                    # The lists were read before this run's renames, and the
                    # other list may have given up the title since
                    title_conflict = _find_title_conflict(
                        _list_existing_lists(self._graph_token, site_id), library_code, title=library_display_name,
                    )
                if title_conflict is not None:
                    # Titles are unique per site, so the rename would fail on every run
                    _append_log(
                        action_log, site_url, library_code, "library_name_conflict",
                        f"another list ({title_conflict.get('webUrl')}) is already titled "
                        f"'{library_display_name}' — the library keeps the title '{current_title}' "
                        f"until that list is renamed",
                    )
                elif dry_run:
                    if existing_list is not None:
                        _append_log(
                            action_log, site_url, library_code, "rename_library (dry run)",
                            f"would set display name to '{library_display_name}'",
                        )
                else:
                    _set_library_display_name(self._graph_token, site_id, list_id, library_display_name)
                    # A new library's display name is part of its creation
                    if existing_list is not None:
                        _append_log(
                            action_log, site_url, library_code, "library_renamed",
                            f"display name set to '{library_display_name}'",
                        )

            # Step 8: Record where the library came from in its root folder's
            # property bag.  After the permissions and display name on
            # purpose: users never see it, so a failure writing it is an
            # error row, but never leaves the library without its groups.
            if source_system is not None or source_id is not None:
                try:
                    # A dry run has no list_id, but reads an existing library's values
                    _ensure_library_metadata(
                        self._sp_token, site_url,
                        list_id if list_id is not None else (existing_list or {}).get("id"),
                        library_code, source_system, source_id, action_log, dry_run=dry_run,
                    )
                except (requests.RequestException, ValueError) as exc:
                    # ValueError: ProcessQuery can refuse with an ErrorInfo in a 200 response
                    logger.error("Metadata not set on '%s' on '%s': %s", library_code, site_url, exc)
                    _append_log(action_log, site_url, library_code, "error", f"metadata not set: {exc}")

            # Step 9: Ensure the library is linked in site navigation under
            # the label.  Last on purpose: a navigation failure is logged as
            # an error but never blocks the permissions work above.
            if nav is not None:
                if dry_run:
                    logger.info(
                        "[DRY RUN] Would ensure navigation link for '%s' under '%s'",
                        library_display_name, nav["label"],
                    )
                    _append_log(
                        action_log, site_url, library_code,
                        "ensure_navigation_link (dry run)",
                        f"would ensure link under '{nav['label']}' navigation label",
                    )
                else:
                    _ensure_library_nav_link(
                        self._sp_token, site_url, library_code,
                        library_display_name, lib_web_url,
                        nav["node_id"], nav["children"], action_log,
                    )

        except (requests.RequestException, ValueError) as exc:
            logger.error("Error creating library '%s' on '%s': %s", library_code, site_url, exc)
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def create_libraries(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
        include_site_name: bool | None = None,
        add_to_navigation: bool | None = None,
        navigation_label: str | None = None,
    ) -> pd.DataFrame:
        """Create SharePoint document libraries and their Entra ID security groups.

        For each unique (site_url, library_code) pair in the DataFrame:
          1. Creates the document library using ``library_code`` (which becomes the
             URL slug), then sets the display name to ``library_display_name``.
          2. Breaks role inheritance so the library has independent permissions.
          3. Unless ``auto_groups`` is False, creates and assigns Auto-Read (Read)
             and Auto-Contribute (Contribute) security groups.  Group names use
             the code, not the display name.
          4. If ``manual_read`` is True, creates a Manual-Read group (Read role)
             whose membership is NOT synced.
          5. If ``manual_contribute`` is True, creates a Manual-Contribute group
             (Contribute role) whose membership is NOT synced.
          6. Writes ``source_system`` and ``source_id``, when given, to the
             library's property bag.  After the groups, so a failure here is an
             ``error`` row but never leaves the library without them.
          7. If ``add_to_navigation`` is on, ensures the library is linked in the
             site's quick-launch navigation, nested under the ``navigation_label``
             heading (the label is created once per site when missing).  Libraries
             created via Graph do not appear in navigation on their own; this is
             ensure-semantics, so an existing link is not duplicated and its title
             follows the display name on rename.

        Required DataFrame columns:
          - site_url:              Full SharePoint site URL
          - library_code:          Short code used as the URL slug (e.g. 'ENG')
          - library_display_name:  Friendly display name shown in the UI (e.g. 'English')
          - group_prefix:          Prefix for security group names (e.g. 'SP')

        Optional DataFrame columns:
          - auto_groups:       Bool — create the automation-synced Auto-Read and
                               Auto-Contribute groups (default True).  False
                               gives a library governed by Manual groups only.
          - manual_read:       Bool — create a Manual-Read group (default False)
          - manual_contribute: Bool — create a Manual-Contribute group (default False)
          - app_owned:         Bool — the app registration owns the groups (default
                               True): new groups are created with it as owner, and
                               an existing group is only granted access if it owns
                               it (an ownerless group has it added).  False for
                               groups people own, such as an SLT library's.
          - source_system:     Str — provenance tag stored in the library's root
                               folder property bag (default None)
          - source_id:         Str — source record ID stored in the library's root
                               folder property bag (default None)

        Args:
            df:               DataFrame with the required columns.
            dry_run:          Override instance default.
            max_workers:      Override instance default.
            include_site_name: Override instance default.
            add_to_navigation: Override instance default.
            navigation_label:  Override instance default.

        Returns:
            Action-log DataFrame with columns: site_url, document_library,
            group, user, action, permission, status.
        """
        (dry_run_, max_workers_, include_site_name_,
         add_to_navigation_, navigation_label_) = self._resolve_config(
            dry_run=dry_run, max_workers=max_workers,
            include_site_name=include_site_name,
            add_to_navigation=add_to_navigation,
            navigation_label=navigation_label,
        )

        required = {"site_url", "library_code", "library_display_name", "group_prefix"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        df = _fill_flag_columns(df, {
            "auto_groups": True, "manual_read": False, "manual_contribute": False, "app_owned": True,
        })
        if "source_system" not in df.columns:
            df["source_system"] = None
        if "source_id" not in df.columns:
            df["source_id"] = None

        action_log = _ThreadSafeLog()

        # Read each site once: its lists and role definition IDs.
        sites = self._prefetch_sites(df["site_url"].unique(), action_log, role_definitions=True)

        # Pre-ensure the navigation label once per site — doing it here rather
        # than in the per-library workers stops concurrent workers racing to
        # create the same label twice.  A site whose label cannot be ensured
        # gets an error row and its libraries skip the navigation step; the
        # permissions work still proceeds.
        nav_by_site: dict[str, dict | None] = {}
        if add_to_navigation_:
            for site_url_key in sites:
                try:
                    if dry_run_:
                        # Read only, so the dry run reports a label a live run would create
                        found = _find_nav_label(self._sp_token, site_url_key, navigation_label_)
                        node_id, children = None, None
                        if found is None:
                            _append_log(
                                action_log, site_url_key, "", "navigation_label_created (dry run)",
                                f"would create '{navigation_label_}' navigation label",
                            )
                    else:
                        node_id, children, created = _ensure_nav_label(
                            self._sp_token, site_url_key, navigation_label_,
                        )
                        if created:
                            _append_log(
                                action_log, site_url_key, "", "navigation_label_created",
                                f"created '{navigation_label_}' navigation label",
                            )
                    nav_by_site[site_url_key] = {
                        "label": navigation_label_, "node_id": node_id, "children": children,
                    }
                except (requests.RequestException, ValueError, KeyError) as exc:
                    logger.error(
                        "Error ensuring navigation label on '%s': %s", site_url_key, exc,
                    )
                    _append_log(action_log, site_url_key, "", "error", str(exc))
                    nav_by_site[site_url_key] = None

        # Deduplicate by (site_url, library_code) — whole rows, so values from
        # different duplicate rows are never mixed
        deduped = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")
        deduped = deduped[deduped["site_url"].isin(list(sites))]

        prefixes = {(r["site_url"], r["library_code"]): r["group_prefix"] for _, r in deduped.iterrows()}
        collisions = _find_group_name_collisions(prefixes, include_site_name_)
        collisions.update(_find_existing_library_collisions(
            prefixes, {site_url: site["lists"] for site_url, site in sites.items()},
        ))
        _log_group_name_collisions(collisions, action_log)

        items = []
        for _, row in deduped.iterrows():
            site_url = row["site_url"]
            if (site_url, row["library_code"]) in collisions:
                continue
            items.append((
                site_url, sites[site_url]["site_id"],
                row["library_code"],
                row["library_display_name"],
                row["group_prefix"],
                bool(row["auto_groups"]),
                bool(row["manual_read"]),
                bool(row["manual_contribute"]),
                bool(row["app_owned"]),
                row["source_system"] if pd.notna(row["source_system"]) else None,
                row["source_id"] if pd.notna(row["source_id"]) else None,
                dry_run_,
                include_site_name_,
                sites[site_url]["role_def_ids"],
                sites[site_url]["lists"],
                nav_by_site.get(site_url),
            ))

        _run_parallel(self._process_create_library, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # set_library_metadata / get_library_metadata
    # -------------------------------------------------------------------

    def _process_set_metadata(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        source_system: str | None,
        source_id: str | None,
        dry_run: bool,
        existing_lists: dict[str, dict],
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Set metadata properties on a single library."""
        logger.info("Setting metadata on library '%s' on site '%s' …", library_code, site_url)

        if source_system is None and source_id is None:
            logger.info(
                "Skipping library '%s' on '%s': both source_system and source_id are None",
                library_code, site_url,
            )
            _append_log(
                action_log, site_url, library_code,
                "skipped", "no metadata values provided",
            )
            return

        try:
            # From the prefetched lists, so a dry run reports a missing
            # library as a live run does
            list_id = _get_list_id(self._graph_token, site_id, library_code, existing=existing_lists)
            _ensure_library_metadata(
                self._sp_token, site_url, list_id, library_code,
                source_system, source_id, action_log, dry_run=dry_run,
            )

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error setting metadata on library '%s' on '%s': %s",
                library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def set_library_metadata(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
    ) -> pd.DataFrame:
        """Write source_system / source_id metadata to existing libraries.

        Values are stored in the library's root folder property bag, which is
        invisible to end users and only accessible via the SharePoint REST API.

        Required DataFrame columns:
          - site_url:      Full SharePoint site URL
          - library_code:  Library code (URL slug)

        At least one of the following columns must be present:
          - source_system: Str — provenance tag (e.g. 'SAP', 'MIS')
          - source_id:     Str — source record ID

        Args:
            df:          DataFrame with the required columns.
            dry_run:     Override instance default.
            max_workers: Override instance default.

        Returns:
            Action-log DataFrame.
        """
        dry_run_, max_workers_ = self._resolve_config(dry_run=dry_run, max_workers=max_workers)

        required = {"site_url", "library_code"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        has_source_system = "source_system" in df.columns
        has_source_id = "source_id" in df.columns
        if not has_source_system and not has_source_id:
            raise ValueError(
                "DataFrame must contain at least one of 'source_system' or 'source_id' columns"
            )

        if not has_source_system:
            df["source_system"] = None
        if not has_source_id:
            df["source_id"] = None

        action_log = _ThreadSafeLog()
        sites = self._prefetch_sites(df["site_url"].unique(), action_log)

        deduped = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")

        items = []
        for _, row in deduped.iterrows():
            site_url = row["site_url"]
            if site_url not in sites:
                continue  # site couldn't be read — already logged
            items.append((
                site_url, sites[site_url]["site_id"],
                row["library_code"],
                row["source_system"] if pd.notna(row["source_system"]) else None,
                row["source_id"] if pd.notna(row["source_id"]) else None,
                dry_run_,
                sites[site_url]["lists"],
            ))

        _run_parallel(self._process_set_metadata, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    def get_library_metadata(
        self,
        df: pd.DataFrame,
    ) -> pd.DataFrame:
        """Read source_system / source_id metadata from existing libraries.

        Values are read from each library's root folder property bag.

        Required DataFrame columns:
          - site_url:      Full SharePoint site URL
          - library_code:  Library code (URL slug)

        Args:
            df:     DataFrame with the required columns.

        Returns:
            DataFrame with columns: site_url, library_code, source_system,
            source_id, status.  ``status`` is ``ok`` where the library was
            read, ``library_not_found`` where it doesn't exist on the site,
            and ``error: <reason>`` where the site or library couldn't be
            read.  Values are ``None`` where a property has not been set or
            the status isn't ``ok``.
        """
        required = {"site_url", "library_code"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        site_log = _ThreadSafeLog()
        sites = self._prefetch_sites(df["site_url"].unique(), site_log)
        site_errors = {row["site_url"]: row["status"] for row in site_log.to_list() if row["action"] == "error"}

        deduped = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")

        results = []
        for _, row in deduped.iterrows():
            site_url = row["site_url"]
            library_code = row["library_code"]
            site = sites.get(site_url)
            library = site["lists"].get(library_code.lower()) if site else None
            values = {"source_system": None, "source_id": None}
            # Unset values are None too, so the status is what says whether they were read
            if site is None:
                status = f"error: site could not be read: {site_errors.get(site_url, '')}"
            elif library is None:
                logger.warning("Library '%s' not found on '%s' — metadata not read", library_code, site_url)
                status = "library_not_found"
            else:
                try:
                    props = _get_library_properties(
                        self._sp_token, site_url, library["id"],
                        list(METADATA_PROPERTY_KEYS.values()),
                    )
                    values = {"source_system": props.get("SourceSystem"), "source_id": props.get("SourceId")}
                    status = "ok"
                except (requests.RequestException, ValueError) as exc:
                    logger.error(
                        "Error reading metadata for library '%s' on '%s': %s",
                        library_code, site_url, exc,
                    )
                    status = f"error: {exc}"
            results.append({"site_url": site_url, "library_code": library_code, **values, "status": status})

        # object dtype keeps unset values as None — pandas 3 would otherwise
        # infer a string dtype and turn them into NaN.
        return pd.DataFrame(
            results, columns=["site_url", "library_code", "source_system", "source_id", "status"], dtype=object,
        )

    # -------------------------------------------------------------------
    # rename_libraries
    # -------------------------------------------------------------------

    def _process_rename_library(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        new_display_name: str,
        dry_run: bool,
        add_to_navigation: bool,
        navigation_label: str,
        existing_lists: dict[str, dict],
        nav_label_found: tuple[int, dict[str, dict]] | None,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Rename a library's display name, and its navigation link's title.

        *existing_lists* and *nav_label_found* are per-site prefetches from
        rename_libraries.
        """
        logger.info(
            "Renaming library code='%s' display='%s' on '%s' …",
            library_code, new_display_name, site_url,
        )

        try:
            # Step 1: Update the SharePoint library display name.  The library
            # is found in the prefetched lists, so a dry run reports a missing
            # library as a live run does.
            list_id = _get_list_id(self._graph_token, site_id, library_code, existing=existing_lists)
            library = existing_lists[library_code.lower()]
            lib_web_url = library.get("webUrl", "")
            if library.get("displayName") == new_display_name:
                # A re-run writes nothing, as the README promises
                _append_log(
                    action_log, site_url, library_code,
                    "already_named", f"display name is already '{new_display_name}'",
                )
            elif dry_run:
                logger.info(
                    "[DRY RUN] Would set display name to '%s' on library '%s'",
                    new_display_name, library_code,
                )
                _append_log(
                    action_log, site_url, library_code,
                    "rename_library (dry run)",
                    f"would set display name to '{new_display_name}'",
                )
            else:
                _set_library_display_name(self._graph_token, site_id, list_id, new_display_name)
                _append_log(
                    action_log, site_url, library_code,
                    "library_renamed",
                    f"display name set to '{new_display_name}'",
                )

            # Step 2: Keep the navigation link's title in step with the new
            # display name.  Last on purpose — a navigation failure must not
            # block the rename itself.  Title-sync only: a missing link is
            # not added here (that is create_libraries' job).
            if add_to_navigation:
                if dry_run:
                    _append_log(
                        action_log, site_url, library_code,
                        "sync_navigation_link (dry run)",
                        f"would update the '{navigation_label}' navigation link "
                        f"title to '{new_display_name}' if the link exists",
                    )
                else:
                    _sync_nav_link_title(
                        self._sp_token, site_url, library_code, new_display_name,
                        lib_web_url, nav_label_found, action_log,
                    )

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error renaming library '%s' on '%s': %s", library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def rename_libraries(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
        include_site_name: bool | None = None,
        add_to_navigation: bool | None = None,
        navigation_label: str | None = None,
    ) -> pd.DataFrame:
        """Rename SharePoint libraries' display names.

        For each row in the DataFrame:
          1. Updates the library's display name (the URL slug is unchanged),
             unless it already has that name (``already_named``).
          2. If ``add_to_navigation`` is on, updates the library's quick-launch
             navigation link title to the new display name (a missing label or
             link is a no-op — links are only added by ``create_libraries``).

        A library's code can't change: it is the library's URL slug, which
        SharePoint never changes, and part of its group names.  To recode a
        library, unlink it and create a new one under the new code.

        Required DataFrame columns:
          - site_url:              Full SharePoint site URL
          - library_code:          Library code (URL slug)
          - library_display_name:  New display name for the library

        Args:
            df:               DataFrame with the required columns.
            dry_run:          Override instance default.
            max_workers:      Override instance default.
            include_site_name: Ignored, as group names don't change.  Kept so
                               existing calls keep working.
            add_to_navigation: Override instance default.
            navigation_label:  Override instance default.

        Returns:
            Action-log DataFrame.

        Raises:
            ValueError: If *df* has values in an ``old_library_code`` column.
                Earlier versions accepted it to rename the groups after a code
                change, but looked the library up by the new code, so it never
                worked.
        """
        dry_run_, max_workers_, add_to_navigation_, navigation_label_ = self._resolve_config(
            dry_run=dry_run, max_workers=max_workers,
            add_to_navigation=add_to_navigation,
            navigation_label=navigation_label,
        )

        required = {"site_url", "library_code", "library_display_name"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        # Refused rather than ignored, so a caller expecting its groups to be
        # renamed finds out
        if "old_library_code" in df.columns and df["old_library_code"].notna().any():
            raise ValueError(
                "old_library_code is no longer supported: a library's code is its URL, which can't change. "
                "To recode a library, unlink it and create a new one under the new code."
            )
        df = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")

        action_log = _ThreadSafeLog()

        # One list enumeration (rename + webUrl lookups) and, with
        # navigation, one quick-launch read per site.
        sites = self._prefetch_sites(
            df["site_url"].unique(), action_log,
            navigation_label=navigation_label_ if add_to_navigation_ else None,
        )

        items = []
        for _, row in df.iterrows():
            site_url = row["site_url"]
            if site_url not in sites:
                continue  # site couldn't be read — already logged
            items.append((
                site_url, sites[site_url]["site_id"],
                row["library_code"],
                row["library_display_name"],
                dry_run_,
                add_to_navigation_,
                navigation_label_,
                sites[site_url]["lists"],
                sites[site_url].get("nav_label"),
            ))

        _run_parallel(self._process_rename_library, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # unlink_libraries
    # -------------------------------------------------------------------

    def _process_unlink_library(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        group_prefix: str,
        dry_run: bool,
        include_site_name: bool,
        add_to_navigation: bool,
        navigation_label: str,
        existing_lists: dict[str, dict],
        nav_label_found: tuple[int, dict[str, dict]] | None,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Unlink a single library: set display name to _UNLINKED and empty Auto groups.

        *existing_lists* and *nav_label_found* are per-site prefetches from
        unlink_libraries.

        Manual-* groups are deliberately left untouched — their descriptions
        promise membership is never managed by automation, and a transient
        source deletion must not destroy hand-curated memberships.
        """
        unlinked_display = f"{library_code}_UNLINKED"
        logger.info("Unlinking library '%s' on '%s' …", library_code, site_url)

        try:
            # Step 1: Set the display name to {code}_UNLINKED
            lib_web_url = None
            library = existing_lists.get(library_code.lower())
            if library is None:
                # Never created (created and deleted between runs) or deleted
                # by hand: nothing to rename.  A data problem, not a fault —
                # an error here would fail every run on the same event.  The
                # Auto groups are still emptied below.
                logger.warning("Library '%s' not found on '%s' — not renamed", library_code, site_url)
                _append_log(
                    action_log, site_url, library_code, "library_not_found",
                    "library does not exist on the site — nothing to rename; "
                    "its Auto groups are still emptied",
                )
            elif library.get("displayName") == unlinked_display:
                # Unlinked by an earlier run: a re-run writes nothing
                lib_web_url = library.get("webUrl", "")
                _append_log(
                    action_log, site_url, library_code,
                    "already_unlinked", f"display name is already '{unlinked_display}'",
                )
            elif dry_run:
                logger.info("[DRY RUN] Would set display name to '%s'", unlinked_display)
                _append_log(
                    action_log, site_url, library_code,
                    "unlink_library (dry run)",
                    f"would set display name to '{unlinked_display}'",
                )
            else:
                # The prefetched lists resolve both the list ID for the rename
                # and the webUrl the navigation link is keyed on.
                list_id = _get_list_id(
                    self._graph_token, site_id, library_code, existing=existing_lists,
                )
                lib_web_url = existing_lists[library_code.lower()].get("webUrl", "")
                _set_library_display_name(self._graph_token, site_id, list_id, unlinked_display)
                _append_log(
                    action_log, site_url, library_code,
                    "library_unlinked",
                    f"display name set to '{unlinked_display}'",
                )

            # Step 2: Empty the automation-owned security groups only.
            # One startswith query resolves all of this library's groups
            # instead of a name-equality lookup per suffix.
            library_group_prefix = build_group_name(
                group_prefix, site_url, library_code, "",
                include_site_name=include_site_name,
            )
            groups_by_name = _find_groups_by_name_prefix(
                self._graph_token, library_group_prefix,
            )
            for suffix in AUTO_GROUP_SUFFIXES:
                group_name = build_group_name(
                    group_prefix, site_url, library_code, suffix,
                    include_site_name=include_site_name,
                )
                existing_group = _get_unique_group(groups_by_name, group_name)
                if not existing_group:
                    logger.info("Group '%s' not found — skipping", group_name)
                    continue
                # Emptying is destructive, so only for the app's own groups
                if not self._check_group_owner(
                    existing_group["id"], group_name, site_url, library_code, "", action_log,
                    dry_run=dry_run,
                ):
                    continue

                if dry_run:
                    logger.info("[DRY RUN] Would empty group '%s'", group_name)
                    _append_log(
                        action_log, site_url, library_code,
                        "empty_group (dry run)", "would remove all members",
                        group=group_name,
                    )
                else:
                    removed, failures = _empty_group_members(self._graph_token, existing_group["id"])
                    _append_log(
                        action_log, site_url, library_code,
                        "group_emptied",
                        f"removed {removed} member(s)" + (
                            f"; {len(failures)} couldn't be removed" if failures else ""
                        ),
                        group=group_name,
                    )
                    for member_key, exc in failures.items():
                        _append_log(
                            action_log, site_url, library_code, "remove_failed", str(exc),
                            group=group_name, user=member_key,
                        )

            # Manual groups are deliberately preserved (never emptied or
            # synced).  Recorded unconditionally — a Graph existence check is
            # not worth a log row that only says nothing was done, and its
            # duplicate-name error could fail an otherwise-successful unlink.
            _append_log(
                action_log, site_url, library_code,
                "manual_groups_preserved",
                "Manual-* groups (if any) left intact — not managed by automation",
            )

            # Step 3: Remove the library's navigation link.  Last on purpose —
            # a navigation failure must not block the unlink itself.  The
            # label is never created here, and an emptied label is kept.
            if add_to_navigation:
                if dry_run:
                    logger.info(
                        "[DRY RUN] Would remove navigation link for '%s' from under '%s'",
                        library_code, navigation_label,
                    )
                    _append_log(
                        action_log, site_url, library_code,
                        "remove_navigation_link (dry run)",
                        f"would remove link from under '{navigation_label}' "
                        f"navigation label if present",
                    )
                else:
                    _remove_library_nav_link(
                        self._sp_token, site_url, library_code, lib_web_url,
                        navigation_label, nav_label_found, action_log,
                    )

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error unlinking library '%s' on '%s': %s", library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def unlink_libraries(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
        include_site_name: bool | None = None,
        add_to_navigation: bool | None = None,
        navigation_label: str | None = None,
    ) -> pd.DataFrame:
        """Unlink removed SharePoint libraries.

        For each row in the DataFrame:
          1. Sets the library display name to ``{library_code}_UNLINKED``
             (the URL slug is unchanged).
          2. Empties the membership of the automation-owned groups only
             (Auto-Read, Auto-Contribute, Auto-Owner).  Manual-Read and
             Manual-Contribute are logged as preserved, never emptied —
             their membership is hand-curated.  Groups that do not exist
             are silently skipped.
          3. If ``add_to_navigation`` is on, removes the library's link from
             under the ``navigation_label`` heading in the site's quick-launch
             navigation (a missing label or link is a no-op; the label itself
             is never created or removed here).

        A library whose group names would be another library's (its code
        sanitises like that of another row, or of another library on the
        site, such as 'PEKS3' and 'PE-KS3') is skipped with
        ``group_name_collision``, as create_libraries skips it.

        Required DataFrame columns:
          - site_url:       Full SharePoint site URL
          - library_code:   Library code (URL slug) to identify the library
          - group_prefix:   Prefix for security group names

        Args:
            df:               DataFrame with the required columns.
            dry_run:          Override instance default.
            max_workers:      Override instance default.
            include_site_name: Override instance default.
            add_to_navigation: Override instance default.
            navigation_label:  Override instance default.

        Returns:
            Action-log DataFrame.
        """
        (dry_run_, max_workers_, include_site_name_,
         add_to_navigation_, navigation_label_) = self._resolve_config(
            dry_run=dry_run, max_workers=max_workers,
            include_site_name=include_site_name,
            add_to_navigation=add_to_navigation,
            navigation_label=navigation_label,
        )

        required = {"site_url", "library_code", "group_prefix"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        # Deduplicate like create_libraries — duplicate rows would race two
        # workers over the same library, groups, and navigation node.
        df = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")

        action_log = _ThreadSafeLog()

        # One list enumeration (rename + webUrl lookups) and, with
        # navigation, one quick-launch read per site.
        sites = self._prefetch_sites(
            df["site_url"].unique(), action_log,
            navigation_label=navigation_label_ if add_to_navigation_ else None,
        )

        # Groups are found by name, so a code whose group names are another
        # library's ('PEKS3' after 'PE-KS3', which create skipped) would empty
        # that library's groups.  Skipped the same way create skips them.
        df = df[df["site_url"].isin(list(sites))]
        prefixes = {(r["site_url"], r["library_code"]): r["group_prefix"] for _, r in df.iterrows()}
        collisions = _find_group_name_collisions(prefixes, include_site_name_)
        collisions.update(_find_existing_library_collisions(
            prefixes, {site_url: site["lists"] for site_url, site in sites.items()},
        ))
        _log_group_name_collisions(collisions, action_log)

        items = []
        for _, row in df.iterrows():
            site_url = row["site_url"]
            if (site_url, row["library_code"]) in collisions:
                continue
            items.append((
                site_url, sites[site_url]["site_id"],
                row["library_code"],
                row["group_prefix"],
                dry_run_,
                include_site_name_,
                add_to_navigation_,
                navigation_label_,
                sites[site_url]["lists"],
                sites[site_url].get("nav_label"),
            ))

        _run_parallel(self._process_unlink_library, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # retire_moved_libraries
    # -------------------------------------------------------------------

    def _process_retire_library(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        group_prefix: str,
        dry_run: bool,
        include_site_name: bool,
        add_to_navigation: bool,
        navigation_label: str,
        existing_lists: dict[str, dict],
        nav_label_found: tuple[int, dict[str, dict]] | None,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Retire the library a site move left behind on the old site.

        Renames it to {code}_MOVED, strips the automation groups' role
        assignments (ending the stale access they would otherwise keep
        granting), and removes its navigation link.  Group membership is
        never touched — the same groups now serve the new site's library.
        """
        retired_display = f"{library_code}_MOVED"
        logger.info("Retiring moved library '%s' on '%s' …", library_code, site_url)

        try:
            if library_code.lower() not in existing_lists:
                # Nothing left to retire (deleted by hand, or already cleaned
                # up) — benign, so no warning.
                _append_log(
                    action_log, site_url, library_code, "retire_skipped",
                    "no library on the old site — nothing to retire",
                )
                return

            # Step 1: Rename so the leftover is clearly not the live library
            list_id = _get_list_id(self._graph_token, site_id, library_code, existing=existing_lists)
            library = existing_lists[library_code.lower()]
            lib_web_url = library.get("webUrl", "")
            if library.get("displayName") == retired_display:
                # Retired by an earlier run: a re-run writes nothing
                _append_log(
                    action_log, site_url, library_code,
                    "already_retired", f"display name is already '{retired_display}'",
                )
            elif dry_run:
                logger.info("[DRY RUN] Would set display name to '%s'", retired_display)
                _append_log(
                    action_log, site_url, library_code,
                    "retire_library (dry run)",
                    f"would set display name to '{retired_display}' and remove "
                    f"automation group access",
                )
            else:
                _set_library_display_name(self._graph_token, site_id, list_id, retired_display)
                _append_log(
                    action_log, site_url, library_code,
                    "library_retired",
                    f"display name set to '{retired_display}' (site moved)",
                )

            # Step 2: Remove the automation groups' role assignments from the
            # old library.  Without this, the groups — reused for the new
            # site's library — would keep granting live access to the leftover.
            # Only groups that still hold a role are touched, so a re-run
            # doesn't register each group on the old site again to find it
            # has nothing to remove.
            library_roles = _get_library_group_roles(self._sp_token, site_url, list_id)
            groups_by_name = _find_groups_by_name_prefix(
                self._graph_token,
                build_group_name(
                    group_prefix, site_url, library_code, "",
                    include_site_name=include_site_name,
                ),
            )
            for suffix in ALL_GROUP_SUFFIXES:
                group_name = build_group_name(
                    group_prefix, site_url, library_code, suffix,
                    include_site_name=include_site_name,
                )
                existing_group = _get_unique_group(groups_by_name, group_name)
                if not existing_group or not library_roles.get(_group_claims(existing_group["id"]).lower()):
                    continue
                if dry_run:
                    _append_log(
                        action_log, site_url, library_code,
                        "remove_group_access (dry run)",
                        "would remove the group's role assignment from the old library",
                        group=group_name,
                    )
                    continue
                principal_id = _resolve_group_principal_id(
                    self._sp_token, site_url, _group_claims(existing_group["id"]),
                )
                if _remove_library_role_assignment(
                    self._sp_token, site_url, list_id, principal_id,
                ):
                    _append_log(
                        action_log, site_url, library_code,
                        "group_access_removed",
                        "role assignment removed from the old library",
                        group=group_name,
                    )

            # Step 3: Remove the old site's navigation link.  Last on purpose —
            # a navigation failure must not block the retire itself.
            if add_to_navigation:
                if dry_run:
                    _append_log(
                        action_log, site_url, library_code,
                        "remove_navigation_link (dry run)",
                        f"would remove link from under '{navigation_label}' "
                        f"navigation label if present",
                    )
                else:
                    _remove_library_nav_link(
                        self._sp_token, site_url, library_code, lib_web_url,
                        navigation_label, nav_label_found, action_log,
                    )

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error retiring library '%s' on '%s': %s", library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def _replacement_problem(
        self, site_url: str, site: dict, library_code: str, group_prefix: str, include_site_name: bool,
    ) -> str | None:
        """Return why a moved library's copy on its new site isn't working yet, or None if it is.

        Retiring removes the old library's access, so the new one must exist,
        have its own permissions, and give at least one of its groups Read,
        Contribute or Full Control.  A library whose create failed part-way,
        or the {code}_MOVED leftover of an earlier move back, can exist
        without them.  Limited Access, which a group gets from permissions on
        something inside the library, doesn't count.

        Args:
            site_url: The new site's URL.
            site: The new site, from _prefetch_sites with role definitions, or
                an empty dict if it couldn't be read.
            library_code: The library's code.
            group_prefix: Prefix of the library's group names.
            include_site_name: Whether group names include the site name.

        Returns:
            The reason, to follow "the library", or None.
        """
        library = site.get("lists", {}).get(library_code.lower())
        if library is None:
            return f"isn't on its new site ({site_url}) yet"
        try:
            if not _has_unique_permissions(self._sp_token, site_url, library["id"]):
                return f"on its new site ({site_url}) still inherits the site's permissions"
            roles = _get_library_group_roles(self._sp_token, site_url, library["id"])
            library_roles = set(site["role_def_ids"].values())
            groups_by_name = _find_groups_by_name_prefix(
                self._graph_token,
                build_group_name(group_prefix, site_url, library_code, "", include_site_name=include_site_name),
            )
            for suffix in ALL_GROUP_SUFFIXES:
                name = build_group_name(
                    group_prefix, site_url, library_code, suffix, include_site_name=include_site_name,
                )
                # A duplicated name raises ValueError: its library can't be synced, so it isn't working
                group = _get_unique_group(groups_by_name, name)
                if group is not None and roles.get(_group_claims(group["id"]).lower(), set()) & library_roles:
                    return None
        except (requests.RequestException, ValueError, KeyError) as exc:
            return f"on its new site ({site_url}) couldn't be checked: {exc}"
        return f"on its new site ({site_url}) gives none of its groups access yet"

    def retire_moved_libraries(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
        include_site_name: bool | None = None,
        add_to_navigation: bool | None = None,
        navigation_label: str | None = None,
    ) -> pd.DataFrame:
        """Retire libraries left behind on their faculty's previous site.

        Run this with ``site_url`` set to the OLD hub (the SCD's
        ``previous_hub_url``) after a site migration.  For each row:
          1. Renames the old-site library to ``{library_code}_MOVED``.
          2. Removes the automation security groups' role assignments from it,
             ending the stale access those groups would otherwise keep
             granting.  Their membership — shared with the new site's library
             — is never touched.
          3. If ``add_to_navigation`` is on, removes the old site's
             navigation link.

        An old site that no longer exists, or a library already gone, is
        logged as ``retire_skipped`` — the leftover has evidently been
        cleaned up already.  An old or new site the app can't read (such as
        a 403 when its grant has been removed) is an ``error`` row, and the
        library is left as it is.

        Required DataFrame columns:
          - site_url:       OLD SharePoint site URL (the previous hub)
          - library_code:   Library code (URL slug)
          - group_prefix:   Prefix for security group names

        Optional DataFrame columns:
          - new_site_url:   The site the library moved to.  When given, the
                            old library is only retired once the library is
                            working there (it exists, has its own permissions
                            and gives one of its groups a role), so a failed
                            create on the new site never leaves the faculty
                            without a working library; until then the row is
                            logged as
                            ``retire_deferred`` (a warning).  Leave it null
                            for a library deleted at source, whose leftover
                            should be retired anyway.  Not checked in a dry
                            run, where the create pass makes nothing.  A row
                            whose new site is the same site as site_url,
                            spelt differently (in case, or with a trailing
                            slash), is logged as ``retire_skipped``, dry run
                            or not, as the library hasn't moved.

        Args:
            df:               DataFrame with the required columns.
            dry_run:          Override instance default.
            max_workers:      Override instance default.
            include_site_name: Override instance default.
            add_to_navigation: Override instance default.
            navigation_label:  Override instance default.

        Returns:
            Action-log DataFrame.
        """
        (dry_run_, max_workers_, include_site_name_,
         add_to_navigation_, navigation_label_) = self._resolve_config(
            dry_run=dry_run, max_workers=max_workers,
            include_site_name=include_site_name,
            add_to_navigation=add_to_navigation,
            navigation_label=navigation_label,
        )

        required = {"site_url", "library_code", "group_prefix"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)
        df = df.drop_duplicates(subset=["site_url", "library_code"], keep="first")
        if "new_site_url" not in df.columns:
            df["new_site_url"] = None
        new_site_urls = list(df["new_site_url"].dropna().unique())
        _validate_site_urls(new_site_urls, self.sharepoint_hostname)

        action_log = _ThreadSafeLog()

        # The new sites: to check the move's replacement exists, and that the
        # new site isn't the old one spelt differently.  Reads only, so dry
        # runs read them too.
        # A new site that doesn't exist yet only defers the retire, but one
        # that can't be read is an error for each library moving to it, so
        # the run isn't clean while the reason is hidden.
        new_site_log = _ThreadSafeLog()
        new_sites = (
            self._prefetch_sites(new_site_urls, new_site_log, role_definitions=True, gone=("missing", ""))
            if new_site_urls else {}
        )
        unreadable_new_sites = {
            row["site_url"]: row["status"] for row in new_site_log.to_list() if row["action"] == "error"
        }

        # Old hubs may have been deleted since the move — those are skipped
        # (nothing to retire), not errors.
        sites = self._prefetch_sites(
            df["site_url"].unique(), action_log,
            navigation_label=navigation_label_ if add_to_navigation_ else None,
            gone=("retire_skipped", "old site no longer exists — nothing to retire"),
        )

        items = []
        for _, row in df.iterrows():
            site_url = row["site_url"]
            if site_url not in sites:
                continue  # site gone or couldn't be read — already logged
            new_site_url = row["new_site_url"]
            if pd.notna(new_site_url) and new_site_url in unreadable_new_sites:
                _append_log(
                    action_log, site_url, row["library_code"], "error",
                    f"its new site ({new_site_url}) couldn't be read, so it wasn't retired: "
                    f"{unreadable_new_sites[new_site_url]}",
                )
                continue
            new_site = new_sites.get(new_site_url) if pd.notna(new_site_url) else None
            if new_site is not None and new_site["site_id"] == sites[site_url]["site_id"]:
                # Graph finds a site ignoring case and a trailing slash, so
                # both URLs are this site: retiring would retire the live library.
                logger.warning(
                    "Not retiring '%s' on '%s': its new site %s is the same site",
                    row["library_code"], site_url, new_site_url,
                )
                _append_log(
                    action_log, site_url, row["library_code"], "retire_skipped",
                    f"its new site ({new_site_url}) is the same site, so it hasn't moved — left as it is",
                )
                continue
            problem = None
            if not dry_run_ and pd.notna(new_site_url):
                problem = self._replacement_problem(
                    new_site_url, new_sites.get(new_site_url, {}),
                    row["library_code"], row["group_prefix"], include_site_name_,
                )
            if problem is not None:
                logger.warning("Not retiring '%s' on '%s': the library %s", row["library_code"], site_url, problem)
                _append_log(
                    action_log, site_url, row["library_code"], "retire_deferred",
                    f"the library {problem}, so the old one keeps its access — "
                    f"retire it once the new library has been set up",
                )
                continue
            items.append((
                site_url, sites[site_url]["site_id"],
                row["library_code"],
                row["group_prefix"],
                dry_run_,
                include_site_name_,
                add_to_navigation_,
                navigation_label_,
                sites[site_url]["lists"],
                sites[site_url].get("nav_label"),
            ))

        _run_parallel(self._process_retire_library, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # update_memberships
    # -------------------------------------------------------------------

    def _process_update_memberships(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        group_prefix: str,
        group_df: pd.DataFrame,
        nested_df: pd.DataFrame,
        group_refs: dict[str, tuple[str | None, str | None, str | None, str | None]],
        user_refs: dict[str, tuple],
        dry_run: bool,
        include_site_name: bool,
        role_def_ids: dict[str, int],
        existing_lists: dict[str, dict],
        create_onedrive_shortcuts: bool,
        max_removals_per_group: int,
        allow_mass_removal: bool,
        planned_removals: _ThreadSafeLog,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Sync membership for a single library's security groups.

        Every role in VALID_ROLES is reconciled: roles absent from the input
        are treated as desired-empty so stale members are removed from any
        existing Auto group — but groups are never *created* just to be
        emptied.  Desired members are users (*group_df*) and nested groups
        (*nested_df*, whose references were resolved once per run into
        *group_refs*: reference -> (object_id, display_name, failure_action,
        failure_reason)).
        Removals of a member kind are skipped for a role when any of that
        role's lookups of that kind failed (incomplete data must not drive
        removals), and all removals are skipped when they exceed
        *max_removals_per_group* without *allow_mass_removal*.  The removals
        that remain aren't made here: they're added to *planned_removals*, so
        update_memberships can check them against its run-wide limit first.
        """
        logger.info("Updating memberships for '%s' on '%s' …", library_code, site_url)

        try:
            # A library can disappear outside the automation (deleted in the
            # SharePoint UI) with no fresh source event to recreate it.  That
            # is a data problem to surface, not a transient fault: a fatal
            # error row here would wedge every future run behind the same
            # missing library, so warn and skip the sync instead.
            if library_code.lower() not in existing_lists:
                logger.warning(
                    "Library '%s' not found on '%s' — membership sync skipped",
                    library_code, site_url,
                )
                _append_log(
                    action_log, site_url, library_code, "library_not_found",
                    "library does not exist on the site — membership sync skipped; "
                    "recreate the library or exclude the faculty",
                )
                return

            # Resolve the SP list ID for role assignment calls (from the
            # prefetched lists, so dry runs can read the assignments too)
            list_id = _get_list_id(self._graph_token, site_id, library_code, existing=existing_lists)

            # Build desired state grouped by permission level.  Seeding every
            # valid role keeps absent roles in the reconcile loop (desired-empty).
            desired_by_role: dict[str, set[str]] = {role: set() for role in VALID_ROLES}
            # Roles containing a user that couldn't be resolved, or a row whose
            # permission is unknown (it could be meant for any role): their
            # removals are skipped, as an incomplete desired set must not drive
            # removals.
            failed_roles: set[str] = set()
            for _, row in group_df.iterrows():
                upn = row["user_principal_name"].strip().lower()
                perm = row["permission"].strip().lower()
                if perm not in VALID_ROLES:
                    logger.warning("Skipping unknown permission '%s' for user '%s'", perm, upn)
                    _append_log(
                        action_log, site_url, library_code,
                        "skip", f"unknown permission '{perm}' — not synced, and no users are "
                        f"removed from this library's groups until it's fixed",
                        user=upn, permission=perm,
                    )
                    failed_roles.update(VALID_ROLES)
                    continue
                desired_by_role[perm].add(upn)

            # UPNs were resolved once per call (see _resolve_users) to
            # handle aliased domains.  Maps input_upn -> (object_id,
            # real_upn) so membership comparison uses the canonical UPN that
            # Graph returns for existing members.
            all_upns = sorted({upn for upns in desired_by_role.values() for upn in upns})
            resolved_users: dict[str, tuple[str, str]] = {}  # input_upn -> (oid, real_upn)
            for upn in all_upns:
                oid, real_upn, fail_reason, fatal = user_refs[upn]
                if fatal:
                    raise requests.RequestException(f"user '{upn}': {fail_reason}")
                if oid is not None:
                    resolved_users[upn] = (oid, real_upn)
                    continue
                user_perms = [r for r, u in desired_by_role.items() if upn in u]
                failed_roles.update(user_perms)
                _append_log(
                    action_log, site_url, library_code,
                    "user_not_found", fail_reason,
                    user=upn, permission=", ".join(user_perms),
                )
                logger.warning("User '%s' could not be resolved: %s", upn, fail_reason)

            # Rebuild desired sets using real (canonical) UPNs and collect
            # the pre-resolved object IDs so the sync loop doesn't need to
            # call Graph again per user.
            real_desired_by_role: dict[str, set[str]] = {}
            upn_to_oid: dict[str, str] = {}  # real_upn -> object_id
            for role, upns in desired_by_role.items():
                real_upns: set[str] = set()
                for upn in upns:
                    if upn in resolved_users:
                        oid, real_upn = resolved_users[upn]
                        real_upns.add(real_upn)
                        upn_to_oid[real_upn] = oid
                real_desired_by_role[role] = real_upns

            # Desired nested groups per role, keyed as _get_group_members keys
            # them ("group:{object_id}") so the reconcile below treats users
            # and groups alike.  An unresolvable reference leaves that role's
            # desired groups incomplete, and an unknown permission every
            # role's: their nested-group removals are skipped (user removals
            # are unaffected — their desired set is complete).
            desired_groups_by_role: dict[str, dict[str, str]] = {role: {} for role in VALID_ROLES}
            failed_group_roles: set[str] = set()
            for _, row in nested_df.iterrows():
                reference = row["member_group"]
                perm = row["permission"].strip().lower()
                if perm not in VALID_ROLES:
                    logger.warning("Skipping unknown permission '%s' for group '%s'", perm, reference)
                    _append_log(
                        action_log, site_url, library_code,
                        "skip", f"unknown permission '{perm}' — not synced, and no nested groups are "
                        f"removed from this library's groups until it's fixed",
                        user=reference, permission=perm,
                    )
                    failed_group_roles.update(VALID_ROLES)
                    continue
                object_id, display_name, fail_action, fail_reason = group_refs[reference]
                if object_id is None:
                    failed_group_roles.add(perm)
                    _append_log(
                        action_log, site_url, library_code,
                        fail_action, fail_reason,
                        user=reference or "(empty)", permission=perm,
                    )
                    logger.warning("Group '%s' could not be resolved: %s", reference, fail_reason)
                    continue
                desired_groups_by_role[perm][f"group:{object_id}"] = display_name

            # One paged startswith query resolves every existing group for
            # this library upfront — replacing a name-equality Graph lookup
            # per role, which the full-pass reconcile would otherwise pay on
            # every run (mostly for 'read'/'owner' groups that do not exist).
            library_group_prefix = build_group_name(
                group_prefix, site_url, library_code, "",
                include_site_name=include_site_name,
            )
            existing_groups_by_name = _find_groups_by_name_prefix(
                self._graph_token, library_group_prefix,
            )

            # For each role, ensure group exists, assign to library, sync members.
            # Phase 1: Create all groups upfront so Entra ID replication can
            # proceed for all groups in parallel instead of waiting sequentially.
            # Roles with an empty desired set only reconcile groups that already
            # exist — a group is never created just to be emptied.
            created_role_groups: list[tuple[str, str, str | None, bool, set[str], dict[str, str], bool]] = []
            # Each entry: (role, group_display_name, group_id, is_new_group,
            #              desired_upns, desired_groups, assign_needed).
            # group_id is None for a group a dry run would create, and
            # is_new_group is True for a group that may still be replicating.
            # Groups created by this call have no members yet.
            created_here: set[str] = set()

            for role in VALID_ROLES:
                desired_upns = real_desired_by_role.get(role, set())
                desired_groups = desired_groups_by_role[role]
                group_display_name = build_group_name(
                    group_prefix, site_url, library_code,
                    f"Auto-{role.capitalize()}",
                    include_site_name=include_site_name,
                )
                group_description = (
                    f"{role.capitalize()} access to '{library_code}' "
                    f"on {site_url} — managed by SharePoint Library Permissions Manager"
                )

                existing_group = _get_unique_group(existing_groups_by_name, group_display_name)
                replicating = known_owned = False
                created = None if existing_group else self._created_group(group_display_name)
                if created is not None:
                    # Made earlier in the run and missing from the lookup while
                    # it replicates.  It's synced like a group found by name,
                    # as an earlier operation may have added members.
                    (existing_group, known_owned), replicating = created, True
                # Auto groups are the automation's own: never add members to,
                # grant, or empty a same-named group it doesn't own.  A dry
                # run only reads the owners, so it shows groups a live run
                # would refuse or adopt.  A group this manager created with
                # the app as owner is known to be the app's.
                if existing_group and not known_owned and not self._check_group_owner(
                    existing_group["id"], group_display_name, site_url, library_code,
                    role, action_log, dry_run=dry_run, is_new_group=replicating,
                ):
                    continue
                if desired_upns or desired_groups:
                    if existing_group:
                        group_id, is_new_group = existing_group["id"], replicating
                    elif dry_run:
                        logger.info("[DRY RUN] Would create group '%s'", group_display_name)
                        _append_log(
                            action_log, site_url, library_code,
                            "ensure_group (dry run)",
                            "would create the group and assign it to the library",
                            group=group_display_name, permission=role,
                        )
                        group_id, is_new_group = None, True
                    else:
                        group = _create_security_group(
                            self._graph_token, group_display_name, group_description,
                            self.service_principal_id,
                        )
                        self._remember_created_group(group_display_name, group, app_owned=True)
                        group_id, is_new_group = group["id"], True
                        created_here.add(group_id)
                    created_role_groups.append(
                        (role, group_display_name, group_id, is_new_group,
                         desired_upns, desired_groups, True)
                    )
                elif existing_group:
                    created_role_groups.append(
                        (role, group_display_name, existing_group["id"], replicating,
                         desired_upns, desired_groups, False)
                    )

            # Phase 2: If any groups were newly created, wait once for Entra ID
            # replication before attempting SharePoint ensureuser for any of them.
            if not dry_run and created_role_groups:
                any_new = any(is_new for _, _, _, is_new, _, _, _ in created_role_groups)
                if any_new:
                    logger.info(
                        "Waiting %.0fs for Entra ID replication before assigning groups …",
                        REPLICATION_BASE_DELAY,
                    )
                    time.sleep(REPLICATION_BASE_DELAY)

            # Phase 3: Assign groups to library and sync membership.  The
            # library's group bindings are read once so only missing ones are
            # written — this runs as a full pass, so most already exist.
            existing_roles: dict[str, set[int]] = {}
            if any(
                assign and group_id is not None and group_id not in created_here
                for _, _, group_id, _, _, _, assign in created_role_groups
            ):
                existing_roles = _get_library_group_roles(self._sp_token, site_url, list_id)

            # Track users who already received a OneDrive shortcut for this
            # library (a user with multiple roles should only get one shortcut).
            shortcuts_created: set[str] = set()

            for (role, group_display_name, group_id, is_new_group,
                 desired_upns, desired_groups, assign_needed) in created_role_groups:
                if assign_needed and dry_run:
                    # A group the dry run would create was logged as ensure_group above
                    if group_id is not None:
                        held = role_def_ids[role] in existing_roles.get(_group_claims(group_id).lower(), set())
                        _append_log(
                            action_log, site_url, library_code,
                            *(("already_assigned", "no change needed") if held
                              else ("group_assigned (dry run)", "would assign the group to the library")),
                            group=group_display_name, permission=role,
                        )
                elif assign_needed:
                    role_def_id = role_def_ids[role]
                    granted = _ensure_group_assigned_to_library(
                        self._sp_token, site_url, list_id,
                        group_id, group_display_name, role_def_id,
                        is_new_group=is_new_group, existing_roles=existing_roles,
                    )
                    _append_log(
                        action_log, site_url, library_code,
                        *(("group_assigned", "success") if granted else ("already_assigned", "no change needed")),
                        group=group_display_name, permission=role,
                    )

                # Sync group membership.  A group created by this call has no
                # members, and reading them straight away can 404 while it
                # replicates.
                current_members = (
                    {} if group_id is None or group_id in created_here
                    else _get_group_members(self._graph_token, group_id, is_new_group=is_new_group)
                )

                # Add users who should be in the group (UPNs are already
                # resolved to canonical form; OIDs are in upn_to_oid).
                for upn in desired_upns:
                    try:
                        if upn in current_members:
                            logger.info("'%s' already in group '%s'", upn, group_display_name)
                            _append_log(
                                action_log, site_url, library_code,
                                "already_member", "no change needed",
                                group=group_display_name, user=upn, permission=role,
                            )
                            continue

                        if dry_run:
                            _append_log(
                                action_log, site_url, library_code,
                                "add_to_group (dry run)", "would add to group",
                                group=group_display_name, user=upn, permission=role,
                            )
                            if create_onedrive_shortcuts and upn not in shortcuts_created:
                                shortcuts_created.add(upn)
                                _append_log(
                                    action_log, site_url, library_code,
                                    "onedrive_shortcut (dry run)", "would create OneDrive shortcut",
                                    group=group_display_name, user=upn, permission=role,
                                )
                            continue

                        _add_group_member(
                            self._graph_token, group_id, upn_to_oid[upn],
                            is_new_group=is_new_group,
                        )
                        _append_log(
                            action_log, site_url, library_code,
                            "added_to_group", "success",
                            group=group_display_name, user=upn, permission=role,
                        )
                        logger.info("Added '%s' to group '%s'", upn, group_display_name)

                        # Create a OneDrive shortcut so the library appears in
                        # the user's drive immediately after being granted access.
                        if create_onedrive_shortcuts and upn not in shortcuts_created:
                            try:
                                shortcut_name = _get_library_display_name(
                                    existing_lists, library_code,
                                )
                                _create_onedrive_shortcut(
                                    graph_token=self._graph_token,
                                    user_object_id=upn_to_oid[upn],
                                    shortcut_name=shortcut_name,
                                    site_id=site_id,
                                    list_id=list_id,
                                    site_url=site_url,
                                )
                                shortcuts_created.add(upn)
                                _append_log(
                                    action_log, site_url, library_code,
                                    "onedrive_shortcut_created", "success",
                                    group=group_display_name, user=upn, permission=role,
                                )
                                logger.info(
                                    "Created OneDrive shortcut for '%s' -> '%s'",
                                    upn, shortcut_name,
                                )
                            except Exception as exc:
                                _append_log(
                                    action_log, site_url, library_code,
                                    "onedrive_shortcut_failed", str(exc),
                                    group=group_display_name, user=upn, permission=role,
                                )
                                logger.warning(
                                    "Failed to create OneDrive shortcut for '%s': %s",
                                    upn, exc,
                                )
                    except requests.HTTPError as exc:
                        _append_log(
                            action_log, site_url, library_code,
                            "add_failed", str(exc),
                            group=group_display_name, user=upn, permission=role,
                        )
                        logger.error(
                            "Failed to add '%s' to group '%s': %s", upn, group_display_name, exc,
                        )

                # Nest desired groups (keys are "group:{object_id}").
                for member_key, member_name in desired_groups.items():
                    try:
                        if member_key in current_members:
                            _append_log(
                                action_log, site_url, library_code,
                                "already_member",
                                f"no change needed (nested group '{member_name}')",
                                group=group_display_name, user=member_key, permission=role,
                            )
                            continue

                        if dry_run:
                            _append_log(
                                action_log, site_url, library_code,
                                "add_to_group (dry run)", f"would nest group '{member_name}' in group",
                                group=group_display_name, user=member_key, permission=role,
                            )
                            continue

                        _add_group_member(
                            self._graph_token, group_id, member_key.split(":", 1)[1],
                            is_new_group=is_new_group,
                        )
                        _append_log(
                            action_log, site_url, library_code,
                            "added_to_group", f"success (nested group '{member_name}')",
                            group=group_display_name, user=member_key, permission=role,
                        )
                        logger.info(
                            "Nested group '%s' in group '%s'", member_name, group_display_name,
                        )
                    except requests.HTTPError as exc:
                        _append_log(
                            action_log, site_url, library_code,
                            "add_failed", f"nested group '{member_name}': {exc}",
                            group=group_display_name, user=member_key, permission=role,
                        )
                        logger.error(
                            "Failed to nest group '%s' in group '%s': %s",
                            member_name, group_display_name, exc,
                        )

                # Remove members who are in the group but NOT desired.  Auto
                # groups are owned by this automation, so users AND nested
                # groups are reconciled — hand-curated access belongs in the
                # Manual groups.  Other non-user members (service principals,
                # devices — the objectid: keys) can never be expressed as
                # desired state, so they are preserved and surfaced in the
                # audit log; unlink's group emptying still removes them.
                removal_candidates = {}
                for key, oid in current_members.items():
                    if key in desired_upns or key in desired_groups:
                        continue
                    if key.startswith("objectid:"):
                        _append_log(
                            action_log, site_url, library_code,
                            "non_user_member_preserved",
                            "service principal/device left in place — "
                            "not managed by membership sync",
                            group=group_display_name, user=key, permission=role,
                        )
                        continue
                    removal_candidates[key] = oid

                # A failed lookup leaves the desired set of that member kind
                # incomplete — the "stale" member may be the unresolved one —
                # so removals of that kind are skipped for the role.
                blocked = {
                    key for key in removal_candidates
                    if role in (failed_group_roles if key.startswith("group:") else failed_roles)
                }
                if blocked:
                    _append_log(
                        action_log, site_url, library_code,
                        "removals_skipped",
                        f"{len(blocked)} removal(s) skipped: one or more '{role}' "
                        f"user or nested-group lookups failed, refusing to remove "
                        f"members based on incomplete data",
                        group=group_display_name, permission=role,
                    )
                    logger.warning(
                        "Skipping %d removal(s) from '%s' — unresolved lookups in role '%s'",
                        len(blocked), group_display_name, role,
                    )
                    removal_candidates = {
                        key: oid for key, oid in removal_candidates.items() if key not in blocked
                    }

                if len(removal_candidates) > max_removals_per_group and not allow_mass_removal:
                    _append_log(
                        action_log, site_url, library_code,
                        "removal_guardrail_triggered",
                        f"would remove {len(removal_candidates)} member(s), above the cap "
                        f"of {max_removals_per_group} — removals skipped; verify the input "
                        f"data and re-run with allow_mass_removal=True to proceed",
                        group=group_display_name, permission=role,
                    )
                    logger.warning(
                        "Removal guardrail: %d planned removals from '%s' exceed cap %d — skipped",
                        len(removal_candidates), group_display_name, max_removals_per_group,
                    )
                elif removal_candidates:
                    # Made once every library is planned, so the run-wide
                    # limit can see them all (see update_memberships)
                    planned_removals.append({
                        "site_url": site_url, "library_code": library_code, "group": group_display_name,
                        "role": role, "group_id": group_id, "members": removal_candidates,
                    })

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error updating memberships for '%s' on '%s': %s", library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def _process_removals(
        self, removal: dict, dry_run: bool, *, action_log: _ThreadSafeLog,
    ) -> None:
        """Remove the members _process_update_memberships planned to remove from one group.

        Args:
            removal: A planned removal: ``site_url``, ``library_code``,
                ``group``, ``role``, ``group_id`` and ``members`` (member key ->
                object ID).
            dry_run: Log the removals rather than make them.
            action_log: The operation's action log.
        """
        site_url, library_code = removal["site_url"], removal["library_code"]
        group_name, role = removal["group"], removal["role"]
        for member_key, member_oid in removal["members"].items():
            if dry_run:
                _append_log(
                    action_log, site_url, library_code,
                    "remove_from_group (dry run)", "would remove from group",
                    group=group_name, user=member_key, permission=role,
                )
                continue
            try:
                removed = _remove_group_member(self._graph_token, removal["group_id"], member_oid)
                _append_log(
                    action_log, site_url, library_code,
                    "removed_from_group", "success" if removed else "already removed",
                    group=group_name, user=member_key, permission=role,
                )
                logger.info("Removed '%s' from group '%s'", member_key, group_name)
            except requests.RequestException as exc:
                # Includes timeouts and dropped connections, so the row names
                # the group and member rather than being a bare error row
                _append_log(
                    action_log, site_url, library_code,
                    "remove_failed", str(exc),
                    group=group_name, user=member_key, permission=role,
                )
                logger.error("Failed to remove '%s' from group '%s': %s", member_key, group_name, exc)

    def update_memberships(
        self,
        df: pd.DataFrame,
        *,
        libraries_df: pd.DataFrame | None = None,
        nested_groups_df: pd.DataFrame | None = None,
        dry_run: bool | None = None,
        max_workers: int | None = None,
        include_site_name: bool | None = None,
        create_onedrive_shortcuts: bool | None = None,
        max_removals_per_group: int | None = None,
        max_removals_per_run: int | None = None,
        allow_mass_removal: bool | None = None,
    ) -> pd.DataFrame:
        """Sync Entra ID security group memberships for SharePoint libraries.

        For each unique (site_url, library_code) pair in the DataFrames:
          1. Groups users (and any nested groups) by permission level (read,
             contribute, owner).
          2. Resolves each user principal name to an Entra ID object.  If the
             UPN does not match directly, the ``mail`` and ``proxyAddresses``
             properties are searched so that **aliased domains** are handled
             transparently.  Nested-group references (display name or object
             ID) are resolved once per call.
          3. For each role with desired members, ensures the corresponding Auto
             security group exists and is assigned to the library.  Roles with
             no desired members are still reconciled against any *existing*
             Auto group (stale members removed) but no group is created.
          4. Adds the desired users and nests the desired groups.
          5. Removes users and nested groups present in the group but NOT
             desired — Auto groups are automation-owned, so a hand-nested
             group is removed too (hand-curated access belongs in the Manual
             groups).  Service principals and devices are preserved.  A
             guardrail can intervene:
               - removals of a member kind are skipped for a role when any of
                 that role's user / nested-group lookups failed
                 (``removals_skipped``),
               - a group's removals are skipped when they exceed
                 *max_removals_per_group* without *allow_mass_removal*
                 (``removal_guardrail_triggered``), and
               - every group's removals are skipped when, across the call,
                 they would exceed *max_removals_per_run* without
                 *allow_mass_removal* (one ``removal_guardrail_triggered``
                 row).  Removals are made after every library's adds, once
                 the total is known.
          6. Optionally creates a OneDrive shortcut in each newly added user's
             OneDrive so the library appears in their drive root immediately.
             A shortcut that fails is logged as ``onedrive_shortcut_failed``
             and not tried again, as later runs find the user already a member:
             ``ensure_onedrive_shortcuts`` makes the missing ones.

        This method does NOT create libraries or break role inheritance — call
        ``create_libraries`` first for new libraries.

        A dry run reads each group's owners, current members and library
        assignment, so it logs the same rows a live run would (adds and
        removals as ``add_to_group (dry run)`` and ``remove_from_group (dry
        run)``, including ``removals_skipped`` and
        ``removal_guardrail_triggered``), without changing anything.

        Required DataFrame columns:
          - site_url:             Full SharePoint site URL
          - library_code:         Library code (URL slug)
          - user_principal_name:  User's UPN (e.g. jsmith@contoso.com)
          - permission:           Permission level — 'read', 'contribute', or 'owner'
          - group_prefix:         Prefix for security group names (e.g. 'SP')

        Args:
            df:               DataFrame with the required columns.
            libraries_df:     Optional DataFrame (site_url, library_code,
                              group_prefix) of libraries to reconcile even when
                              they have no rows in *df* — e.g. a faculty whose
                              last member left still needs its groups emptied.
            nested_groups_df: Optional DataFrame (site_url, library_code,
                              member_group, permission, group_prefix) of
                              existing Entra security groups to nest in the
                              library's Auto group for *permission*.
                              ``member_group`` is a group object ID or exact
                              display name; one that cannot be resolved (or is
                              a Microsoft 365 group / distribution list, which
                              Entra cannot nest) is logged as
                              ``group_not_found``.  Its libraries are reconciled
                              even when they have no rows in *df*.
            dry_run:          Override instance default.
            max_workers:      Override instance default.
            include_site_name: Override instance default.
            create_onedrive_shortcuts:
                Override instance default.  When True, create a OneDrive
                shortcut in each newly added user's drive root pointing to
                the SharePoint library.  Shortcut failures are logged but
                never block membership sync.
            max_removals_per_group: Override instance default.
            max_removals_per_run:   Override instance default.
            allow_mass_removal:     Override instance default.

        Returns:
            Action-log DataFrame.  Rows with ``action == "user_not_found"``
            identify users that could not be resolved by UPN, mail, or
            proxyAddresses.  To extract them::

                not_found = results[results["action"] == "user_not_found"]
        """
        (dry_run_, max_workers_, include_site_name_, create_onedrive_shortcuts_,
         max_removals_, max_removals_per_run_, allow_mass_removal_) = self._resolve_config(
            dry_run=dry_run, max_workers=max_workers,
            include_site_name=include_site_name,
            create_onedrive_shortcuts=create_onedrive_shortcuts,
            max_removals_per_group=max_removals_per_group,
            max_removals_per_run=max_removals_per_run,
            allow_mass_removal=allow_mass_removal,
        )

        required = {
            "site_url", "library_code", "user_principal_name",
            "permission", "group_prefix",
        }
        df = _validate_required_columns(_blank_nulls(df, ("user_principal_name", "permission")), required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        if libraries_df is not None:
            libraries_df = _validate_required_columns(
                libraries_df, {"site_url", "library_code", "group_prefix"},
            )
            _validate_site_urls(libraries_df["site_url"].unique(), self.sharepoint_hostname)
            libraries_df = libraries_df.drop_duplicates(
                subset=["site_url", "library_code"], keep="first",
            )

        nested_required = {"site_url", "library_code", "member_group", "permission", "group_prefix"}
        if nested_groups_df is None:
            nested_groups_df = pd.DataFrame(columns=sorted(nested_required))
        nested_groups_df = _validate_required_columns(
            _blank_nulls(nested_groups_df, ("member_group", "permission")), nested_required,
        )
        _validate_site_urls(nested_groups_df["site_url"].unique(), self.sharepoint_hostname)
        nested_groups_df["member_group"] = nested_groups_df["member_group"].str.strip()

        # Effective prefix per (site_url, library_code): first from df, then
        # from nested_groups_df and libraries_df for libraries without users.
        prefix_by_library: dict[tuple[str, str], str] = {}
        for _, row in df.iterrows():
            key = (row["site_url"], row["library_code"])
            if key in prefix_by_library and prefix_by_library[key] != row["group_prefix"]:
                logger.warning(
                    "Multiple group_prefix values for library '%s' on '%s' — using first: '%s'.",
                    key[1], key[0], prefix_by_library[key],
                )
            prefix_by_library.setdefault(key, row["group_prefix"])
        for extra_df in (nested_groups_df, libraries_df):
            if extra_df is None:
                continue
            for _, row in extra_df.iterrows():
                prefix_by_library.setdefault(
                    (row["site_url"], row["library_code"]), row["group_prefix"],
                )

        # Resolve each nested-group reference once — an academy-wide group is
        # typically nested in many libraries.  Values are (object_id,
        # display_name, failure_action, failure_reason).  An unresolvable
        # reference is a persistent data problem (group_not_found, a
        # warning); an API failure is transient (error, fatal).  Both are
        # recorded per reference, so one bad lookup cannot abort the sync
        # of libraries that do not use it.
        group_refs: dict[str, tuple[str | None, str | None, str | None, str | None]] = {}
        for reference in nested_groups_df["member_group"].unique():
            try:
                group = _resolve_nestable_group(self._graph_token, reference)
                group_refs[reference] = (
                    group["id"], group.get("displayName") or reference, None, None,
                )
            except GroupResolutionError as exc:
                group_refs[reference] = (None, None, "group_not_found", exc.reason)
            except requests.RequestException as exc:
                group_refs[reference] = (None, None, "error", f"group lookup failed: {exc}")

        action_log = _ThreadSafeLog()

        # Read each site once: its lists and role definition IDs.  A site
        # that can't be read skips only its own libraries.
        unique_urls = sorted({site_url for site_url, _ in prefix_by_library})
        sites = self._prefetch_sites(unique_urls, action_log, role_definitions=True)
        prefix_by_library = {pair: prefix for pair, prefix in prefix_by_library.items() if pair[0] in sites}

        # Collisions within this call, and with libraries already on the
        # site (e.g. an unlinked library still holding the shared groups).
        collisions = _find_group_name_collisions(prefix_by_library, include_site_name_)
        collisions.update(_find_existing_library_collisions(
            prefix_by_library, {site_url: site["lists"] for site_url, site in sites.items()},
        ))
        for pair in collisions:
            prefix_by_library.pop(pair, None)
        _log_group_name_collisions(collisions, action_log)

        user_refs = self._resolve_users(
            df.loc[df.set_index(["site_url", "library_code"]).index.isin(list(prefix_by_library)),
                   "user_principal_name"],
            max_workers_,
        )
        grouped = dict(tuple(df.groupby(["site_url", "library_code"])))
        nested_grouped = dict(tuple(nested_groups_df.groupby(["site_url", "library_code"])))

        # Libraries with no membership rows reconcile against an empty frame.
        empty_group_df = df.iloc[0:0]
        empty_nested_df = nested_groups_df.iloc[0:0]

        planned_removals = _ThreadSafeLog()
        items = []
        for (site_url, library_code), group_prefix in prefix_by_library.items():
            group_df = grouped.get((site_url, library_code), empty_group_df)
            nested_df = nested_grouped.get((site_url, library_code), empty_nested_df)
            items.append((
                site_url, sites[site_url]["site_id"],
                library_code,
                group_prefix,
                group_df,
                nested_df,
                group_refs,
                user_refs,
                dry_run_,
                include_site_name_,
                sites[site_url]["role_def_ids"],
                sites[site_url]["lists"],
                create_onedrive_shortcuts_,
                max_removals_,
                allow_mass_removal_,
                planned_removals,
            ))

        _run_parallel(self._process_update_memberships, items, max_workers_, action_log, locate=_site_and_code)

        # Each group's removals were within max_removals_per_group, but an
        # empty or truncated members table would empty every small group, so
        # the call's total is checked before any member is removed.
        removals = planned_removals.to_list()
        total = sum(len(removal["members"]) for removal in removals)
        if total > max_removals_per_run_ and not allow_mass_removal_:
            logger.warning(
                "Removal guardrail: %d planned removals from %d groups exceed the run limit %d — none made",
                total, len(removals), max_removals_per_run_,
            )
            _append_log(
                action_log, "", "", "removal_guardrail_triggered",
                f"would remove {total} member(s) from {len(removals)} group(s), above the run limit of "
                f"{max_removals_per_run_} — no members removed; verify the input data and re-run with "
                f"allow_mass_removal=True to proceed",
            )
        else:
            _run_parallel(
                self._process_removals, [(removal, dry_run_) for removal in removals], max_workers_, action_log,
                locate=lambda item: (item[0]["site_url"], item[0]["library_code"]),
            )
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # ensure_onedrive_shortcuts
    # -------------------------------------------------------------------

    def _process_ensure_shortcuts(
        self,
        site_url: str,
        site_id: str,
        library_code: str,
        existing_lists: dict[str, dict],
        upns: list[str],
        user_refs: dict[str, tuple],
        dry_run: bool,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Ensure each user has a OneDrive shortcut to one library.

        Args:
            site_url: The library's site.
            site_id: The site's Graph ID.
            library_code: The library's code.
            existing_lists: The site's lists, from _list_existing_lists.
            upns: Normalised user principal names, keys of *user_refs*.
            user_refs: _resolve_users result.
            dry_run: Log the shortcuts that would be created rather than create them.
            action_log: The operation's action log.
        """
        library = existing_lists.get(library_code.lower())
        if library is None:
            _append_log(
                action_log, site_url, library_code, "library_not_found",
                "library does not exist on the site — no shortcuts made",
            )
            return
        shortcut_name = _get_library_display_name(existing_lists, library_code)
        for upn in upns:
            oid, _, fail_reason, fatal = user_refs[upn]
            if oid is None:
                _append_log(
                    action_log, site_url, library_code, "error" if fatal else "user_not_found", fail_reason, user=upn,
                )
                continue
            try:
                if _find_library_shortcut(_get_onedrive_root_items(self._graph_token, oid), library) is not None:
                    _append_log(
                        action_log, site_url, library_code, "onedrive_shortcut_exists", "no change needed", user=upn,
                    )
                    continue
                if dry_run:
                    _append_log(
                        action_log, site_url, library_code,
                        "onedrive_shortcut (dry run)", "would create OneDrive shortcut", user=upn,
                    )
                    continue
                _create_onedrive_shortcut(
                    graph_token=self._graph_token, user_object_id=oid, shortcut_name=shortcut_name,
                    site_id=site_id, list_id=library["id"], site_url=site_url,
                )
                _append_log(action_log, site_url, library_code, "onedrive_shortcut_created", "success", user=upn)
            except requests.RequestException as exc:
                status = getattr(getattr(exc, "response", None), "status_code", None)
                reason = (
                    "the user's OneDrive doesn't exist yet: it's created the first time they open it"
                    if status == 404 else str(exc)
                )
                logger.warning("OneDrive shortcut for '%s' to '%s' failed: %s", upn, library_code, reason)
                _append_log(action_log, site_url, library_code, "onedrive_shortcut_failed", reason, user=upn)

    def ensure_onedrive_shortcuts(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
    ) -> pd.DataFrame:
        """Create any missing OneDrive shortcuts to libraries, for example ones that failed before.

        With ``create_onedrive_shortcuts``, update_memberships only creates a
        shortcut when it adds a user to a group, and doesn't try again if that
        fails, most often because the user's OneDrive didn't exist yet.  This
        isn't part of the usual flow: run it to make those shortcuts later.
        For each user and library it reads the root of the user's OneDrive,
        and creates a shortcut to the library unless one is already there,
        even if the user has renamed it.  Safe to re-run.

        Needs ``Files.ReadWrite.All`` on Microsoft Graph, as shortcuts do in
        update_memberships.

        Required DataFrame columns (other columns are ignored, so the
        DataFrame given to update_memberships can be reused):
          - site_url:             Full SharePoint site URL
          - library_code:         Library code (URL slug)
          - user_principal_name:  User's UPN

        Args:
            df:          DataFrame with the required columns.
            dry_run:     Override instance default.
            max_workers: Override instance default.

        Returns:
            Action-log DataFrame: ``onedrive_shortcut_created``,
            ``onedrive_shortcut_exists``, or ``onedrive_shortcut (dry run)``
            for each user and library, and ``onedrive_shortcut_failed`` where
            one couldn't be made, for example because the user's OneDrive
            doesn't exist yet.
        """
        dry_run_, max_workers_ = self._resolve_config(dry_run=dry_run, max_workers=max_workers)

        df = _validate_required_columns(df, {"site_url", "library_code", "user_principal_name"})
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)

        action_log = _ThreadSafeLog()
        sites = self._prefetch_sites(df["site_url"].unique(), action_log)
        df = df[df["site_url"].isin(list(sites))]
        user_refs = self._resolve_users(df["user_principal_name"], max_workers_)

        # Rows can name one library differently, as MATH and math, or its site
        # with and without a trailing slash.  Each library is one work item,
        # so two workers can't both find a user has no shortcut and both make
        # one.
        libraries: dict[tuple[str, str], tuple[str, str, str, dict]] = {}
        library_upns: dict[tuple[str, str], set[str]] = {}
        for (site_url, library_code), library_df in df.groupby(["site_url", "library_code"]):
            site = sites[site_url]
            library = site["lists"].get(library_code.lower())
            key = (site["site_id"], library["id"] if library is not None else library_code.lower())
            libraries.setdefault(key, (site_url, site["site_id"], library_code, site["lists"]))
            library_upns.setdefault(key, set()).update(
                upn.strip().lower() for upn in library_df["user_principal_name"]
            )
        items = [(*libraries[key], sorted(library_upns[key]), user_refs, dry_run_) for key in libraries]

        _run_parallel(self._process_ensure_shortcuts, items, max_workers_, action_log, locate=_site_and_code)
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # grant_group_access
    # -------------------------------------------------------------------

    def _process_grant_group_access(
        self,
        site_url: str,
        library_code: str,
        grants: list[tuple[str, str, str, int | None]],
        dry_run: bool,
        role_def_ids: dict[str, int],
        existing_lists: dict[str, dict],
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Ensure each (group_name, role, group_id, principal_id) grant on one library.

        Principal IDs are pre-resolved per site by grant_group_access (None
        in a dry run, as resolving one is a write).  The library's role
        assignments are read once so only missing bindings are written, and
        the audit log distinguishes real grants from no-ops.
        """
        try:
            # Same reasoning as membership sync: a missing library is a
            # persistent data problem, so warn rather than fail every run.
            if library_code.lower() not in existing_lists:
                logger.warning(
                    "Library '%s' not found on '%s' — group access not granted",
                    library_code, site_url,
                )
                _append_log(
                    action_log, site_url, library_code, "library_not_found",
                    "library does not exist on the site — group access not granted",
                )
                return

            list_id = existing_lists[library_code.lower()]["id"]
            if dry_run:
                # Without principal IDs, match groups by their claims login name
                held_by_login = _get_library_group_roles(self._sp_token, site_url, list_id)
            else:
                bindings = _get_library_role_assignments(self._sp_token, site_url, list_id)
            for group_name, role, group_id, principal_id in grants:
                role_def_id = role_def_ids[role]
                held = (
                    held_by_login.get(_group_claims(group_id).lower(), set()) if dry_run
                    else bindings.get(principal_id, set())
                )
                if role_def_id in held:
                    _append_log(
                        action_log, site_url, library_code,
                        "already_assigned", "no change needed",
                        group=group_name, permission=role,
                    )
                    continue
                if dry_run:
                    _append_log(
                        action_log, site_url, library_code,
                        "grant_group_access (dry run)",
                        f"would grant the group '{role}' on the library",
                        group=group_name, permission=role,
                    )
                    continue
                _add_role_assignment(self._sp_token, site_url, list_id, principal_id, role_def_id)
                logger.info(
                    "Granted '%s' %s on library '%s' (%s)",
                    group_name, role, library_code, site_url,
                )
                _append_log(
                    action_log, site_url, library_code,
                    "group_access_granted", "success",
                    group=group_name, permission=role,
                )

        except (requests.RequestException, ValueError) as exc:
            logger.error(
                "Error granting group access on '%s' on '%s': %s", library_code, site_url, exc,
            )
            _append_log(action_log, site_url, library_code, "error", str(exc))

    def grant_group_access(
        self,
        df: pd.DataFrame,
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
    ) -> pd.DataFrame:
        """Ensure existing Entra ID groups hold a permission level on existing libraries.

        For each row, the named group is granted the permission level on the
        library unless it already holds it.  Additive and idempotent: no
        role assignment is ever removed (including other levels the same
        group holds), so it is safe to run as a full pass every run — which
        also restores an assignment removed by hand.

        Unlike ``create_libraries`` nothing is created.  A group that does not
        exist is logged once as ``group_not_found`` and a library that does
        not exist as ``library_not_found`` (both warnings); their grants are
        skipped.  A display name shared by several groups is an ``error`` —
        never guess which group gets access.  Libraries must already have
        unique permissions (as ``create_libraries`` leaves them); SharePoint
        rejects role assignments on a library that still inherits.

        A dry run reads the groups, sites and role assignments, so it logs
        the same rows a live run would, with ``grant_group_access (dry run)``
        in place of ``group_access_granted``.

        Required DataFrame columns:
          - site_url:      Full SharePoint site URL
          - library_code:  Library code (URL slug)
          - group_name:    Exact Entra ID display name of the group
          - permission:    Permission level — 'read', 'contribute', or 'owner'

        Args:
            df:          DataFrame with the required columns.
            dry_run:     Override instance default.
            max_workers: Override instance default.

        Returns:
            Action-log DataFrame.
        """
        dry_run_, max_workers_ = self._resolve_config(dry_run=dry_run, max_workers=max_workers)

        required = {"site_url", "library_code", "group_name", "permission"}
        df = _validate_required_columns(df, required)
        _validate_site_urls(df["site_url"].unique(), self.sharepoint_hostname)
        df["permission"] = df["permission"].str.strip().str.lower()
        unknown = set(df["permission"]) - set(VALID_ROLES)
        if unknown:
            raise ValueError(
                f"Unknown permission(s) {sorted(unknown)} — must be one of {list(VALID_ROLES)}"
            )
        df = df.drop_duplicates(subset=["site_url", "library_code", "group_name", "permission"])

        action_log = _ThreadSafeLog()

        # Per-site prefetch: lists, role definitions, and each group's
        # SharePoint principal ID — the principal is per site, so ensureuser
        # runs once per (site, group) rather than once per library.  A site
        # whose prefetch fails gets an error row and its grants are skipped.
        # Everything but ensureuser is a read, so dry runs do the rest too.
        group_ids: dict[str, str] = {}
        for group_name in df["group_name"].unique():
            # Exact-name lookup, as create_libraries uses when ensuring
            # these groups — so a duplicated name (ValueError) is fatal
            # here too: never guess which group gets access.
            try:
                group = _find_group_by_name(self._graph_token, group_name)
            except (requests.RequestException, ValueError) as exc:
                logger.error("Group '%s': %s", group_name, exc)
                _append_log(action_log, "", "", "error", str(exc), group=group_name)
                continue
            if group is None:
                logger.warning("Group '%s' not found — its grants are skipped", group_name)
                _append_log(
                    action_log, "", "", "group_not_found",
                    "group does not exist — its grants were skipped",
                    group=group_name,
                )
                continue
            group_ids[group_name] = group["id"]

        sites = self._prefetch_sites(df["site_url"].unique(), action_log, role_definitions=True)
        principal_ids: dict[tuple[str, str], int] = {}
        if not dry_run_:
            for site_url_key in sites:
                site_groups = df.loc[df["site_url"] == site_url_key, "group_name"].unique()
                for group_name in site_groups:
                    if group_name not in group_ids:
                        continue
                    # Per group, so one group that SharePoint can't resolve
                    # doesn't skip the other groups' grants on the site.
                    try:
                        principal_ids[(site_url_key, group_name)] = _resolve_group_principal_id(
                            self._sp_token, site_url_key, _group_claims(group_ids[group_name]),
                        )
                    except (requests.RequestException, ValueError, KeyError) as exc:
                        logger.error("Group '%s' on '%s': %s", group_name, site_url_key, exc)
                        _append_log(
                            action_log, site_url_key, "", "error", str(exc), group=group_name,
                        )

        grants_by_library: dict[tuple[str, str], list[tuple[str, str, str, int | None]]] = {}
        for row in df.itertuples(index=False):
            if row.site_url not in sites or row.group_name not in group_ids:
                continue  # site prefetch failed, or group missing or ambiguous — already logged
            principal_id = None
            if not dry_run_:
                principal_id = principal_ids.get((row.site_url, row.group_name))
                if principal_id is None:
                    continue  # SharePoint couldn't resolve the group — already logged
            grants_by_library.setdefault((row.site_url, row.library_code), []).append(
                (row.group_name, row.permission, group_ids[row.group_name], principal_id)
            )

        items = [
            (
                site_url, library_code, grants, dry_run_,
                sites[site_url]["role_def_ids"],
                sites[site_url]["lists"],
            )
            for (site_url, library_code), grants in grants_by_library.items()
        ]
        _run_parallel(
            self._process_grant_group_access, items, max_workers_, action_log,
            locate=lambda item: (item[0], item[1]),
        )
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # enforce_site_sharing_lockdown
    # -------------------------------------------------------------------

    def _process_enforce_site_sharing(
        self,
        site_url: str,
        dry_run: bool,
        *,
        action_log: _ThreadSafeLog,
    ) -> None:
        """Check one site's sharing settings; turn off anything non-compliant."""
        logger.info("Checking site sharing settings on '%s' …", site_url)

        try:
            settings = _get_site_sharing_settings(self._sp_token, site_url)
            violations = [
                name for name in (
                    "members_can_share", "access_requests_enabled", "members_can_invite",
                )
                if settings[name]
            ]

            if not violations:
                logger.info("Site '%s' sharing settings already compliant", site_url)
                _append_log(
                    action_log, site_url, "", "site_sharing_check",
                    "compliant — member sharing, access requests, and member "
                    "invites already off",
                )
                return

            if dry_run:
                logger.info(
                    "[DRY RUN] Would turn off on '%s': %s", site_url, ", ".join(violations),
                )
                _append_log(
                    action_log, site_url, "", "site_sharing_lockdown (dry run)",
                    f"would turn off: {', '.join(violations)}",
                )
                return

            _disable_site_sharing(self._sp_token, site_url, settings["member_group_id"])

            # SharePoint has form for accepting writes it then discards (see
            # the property bag helpers) — read back and fail loudly if the
            # settings didn't stick.
            after = _get_site_sharing_settings(self._sp_token, site_url)
            still_on = [name for name in violations if after[name]]
            if still_on:
                raise ValueError(
                    f"site sharing lockdown did not stick on {site_url}: "
                    f"{', '.join(still_on)} still enabled after write"
                )

            _append_log(
                action_log, site_url, "", "site_sharing_disabled",
                f"turned off: {', '.join(violations)}",
            )
            logger.info("Site '%s': turned off %s", site_url, ", ".join(violations))

        except (requests.RequestException, ValueError) as exc:
            logger.error("Error enforcing site sharing lockdown on '%s': %s", site_url, exc)
            _append_log(action_log, site_url, "", "error", str(exc))

    def enforce_site_sharing_lockdown(
        self,
        site_urls: Iterable[str],
        *,
        dry_run: bool | None = None,
        max_workers: int | None = None,
    ) -> pd.DataFrame:
        """Turn off site-level member sharing, access requests, and invites.

        For each unique site URL, checks and (when needed) turns off:
          - "Allow members to share the site and individual files and folders"
            (``SP.Web.MembersCanShare``),
          - "Allow access requests" (``RequestAccessEmail`` cleared and
            ``UseAccessRequestDefault`` turned off), and
          - "Allow members to invite others to the site members group"
            (``AllowMembersEditMembership`` on the associated member group).

        The check runs once per site, however many libraries the site hosts.
        Already-compliant sites are logged as ``site_sharing_check`` and left
        untouched, so the call is idempotent and cheap to make every run.
        Writes are verified with a read-back; a write that does not stick is
        logged as an ``error`` row.

        Args:
            site_urls:   Iterable of full SharePoint site URLs.  A site listed
                         more than once, including spelt differently (in
                         case, or with a trailing slash), is checked once,
                         under its first spelling.
            dry_run:     Override instance default.
            max_workers: Override instance default.

        Returns:
            Action-log DataFrame (``document_library`` is empty for these
            site-level rows).

        Raises:
            TypeError: If site_urls is a single URL rather than a list of them.
        """
        dry_run_, max_workers_ = self._resolve_config(dry_run=dry_run, max_workers=max_workers)

        # A string is an iterable of characters, each of which would fail as a URL
        if isinstance(site_urls, str):
            raise TypeError("site_urls is a list of site URLs: for one site, pass [site_url]")
        site_urls = list(site_urls)
        _validate_site_urls(site_urls, self.sharepoint_hostname)
        # Graph finds a site ignoring case and a trailing slash, so each
        # spelling would be checked and written separately, at the same time
        # with max_workers above 1
        spellings: dict[str, str] = {}
        for site_url in site_urls:
            spellings.setdefault(site_url.rstrip("/").lower(), site_url)
        unique_urls = sorted(spellings.values())

        action_log = _ThreadSafeLog()
        items = [(site_url, dry_run_) for site_url in unique_urls]
        _run_parallel(
            self._process_enforce_site_sharing, items, max_workers_, action_log,
            locate=lambda item: (item[0], ""),
        )
        return _action_log_to_dataframe(action_log)

    # -------------------------------------------------------------------
    # Read-only helpers
    # -------------------------------------------------------------------

    def group_name(
        self,
        group_prefix: str,
        site_url: str,
        library_code: str,
        suffix: str,
        *,
        include_site_name: bool | None = None,
    ) -> str:
        """Return the security group name the operations use for a library.

        Built exactly as create_libraries builds it, so callers that target
        these groups by name (e.g. grant_group_access rows) cannot drift from
        the groups the automation ensures.

        Args:
            group_prefix:      Prefix for security group names (e.g. 'SP').
            site_url:          Full SharePoint site URL.
            library_code:      Library code (URL slug).
            suffix:            Role suffix, e.g. 'Auto-Read' or 'Manual-Contribute'.
            include_site_name: Override instance default.
        """
        (include_site_name_,) = self._resolve_config(include_site_name=include_site_name)
        return build_group_name(
            group_prefix, site_url, library_code, suffix,
            include_site_name=include_site_name_,
        )

    def get_site_sharing_settings(self, site_url: str) -> dict:
        """Read the site sharing settings enforce_site_sharing_lockdown governs.

        Returns:
            Dict with ``members_can_share``, ``access_requests_enabled`` and
            ``members_can_invite`` (True means the setting is on) and
            ``member_group_id``.
        """
        _validate_site_urls([site_url], self.sharepoint_hostname)
        return _get_site_sharing_settings(self._sp_token, site_url)

    def get_navigation_link(
        self,
        site_url: str,
        library_code: str,
        *,
        navigation_label: str | None = None,
    ) -> dict | None:
        """Return a library's quick-launch link under the navigation label.

        Reads the site's navigation as it is, so it shows what
        create_libraries wrote rather than what it reported.

        Args:
            site_url:         Full SharePoint site URL.
            library_code:     Library code (URL slug).
            navigation_label: Override instance default.

        Returns:
            The navigation node (``Id``, ``Title``, ``Url``, …), or None when
            the label or the link does not exist.

        Raises:
            ValueError: If the site URL isn't on the SharePoint host, or the
                code isn't one SharePoint keeps as a library's URL (see
                _validate_library_codes): the link is found by the library's
                URL, which is the site URL and the code.
        """
        (navigation_label_,) = self._resolve_config(navigation_label=navigation_label)
        _validate_site_urls([site_url], self.sharepoint_hostname)
        # A '#' or '?' in the code would cut the URL short and look up another path
        _validate_library_codes([library_code])
        found = _find_nav_label(self._sp_token, site_url, navigation_label_)
        if found is None:
            return None
        _, children = found
        return children.get(_normalise_nav_url(f"{site_url.rstrip('/')}/{library_code}"))
