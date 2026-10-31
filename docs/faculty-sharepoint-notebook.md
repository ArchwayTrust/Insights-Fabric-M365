# Manage Faculty Sharepoint Libraries notebook

The **Manage Faculty Sharepoint Libraries** notebook gives each faculty its own SharePoint document library on its academy's hub site, and keeps the library's permissions in line with the curated faculty tables. It used to carry about 4,100 lines of core code in one cell. That code now lives in this package, as [`alt_fabric_m365.sharepoint`](../alt_fabric_m365/sharepoint), and the notebook imports it.

The notebook's final state is [`examples/Manage Faculty Sharepoint Libraries.Notebook/notebook-content.py`](../examples/Manage%20Faculty%20Sharepoint%20Libraries.Notebook/notebook-content.py), in the format Fabric's Git integration uses. It is 1,160 lines, down from 5,341.

- [Moving the notebook over](#moving-the-notebook-over)
- [Final state of the notebook](#final-state-of-the-notebook)
- [What changed in the notebook](#what-changed-in-the-notebook)
- [What changed in behaviour](#what-changed-in-behaviour)
- [Where the core code went](#where-the-core-code-went)
- [Test coverage](#test-coverage)
- [Possible next steps](#possible-next-steps)

## Moving the notebook over

1. **Release version 0.6.0** of the package (see [Publishing to PyPI with a GitHub release](../README.md#publishing-to-pypi-with-a-github-release)), or upload the wheel to test it first (see [Testing a wheel in Fabric](../README.md#testing-a-wheel-in-fabric-before-publishing)).
2. **Add the package to the notebook's environment.** In the Fabric environment the notebook and its pipeline use, add `alt-fabric-m365` version `0.6.0` under **Public libraries** and publish the environment. The current notebook has no environment attached (`"environment": {}` in its metadata), so attach one under **Environment** in the notebook toolbar. Its dependencies, including `pandas`, are already in the Fabric runtime.
3. **Put the app's certificate in Key Vault.** The notebook now authenticates with the package's `KeyVaultCertificateCredential`: Key Vault signs each request with the certificate's private key, which never leaves the key vault, instead of the notebook reading the key from two secrets. For the existing app registration (`CLIENT_ID`), follow steps 2 to 4 of [Setting up a certificate](../README.md#setting-up-a-certificate): create a certificate in the `KEY_VAULT` key vault, upload it to the app registration with `--append` (so the current certificate keeps working until you remove it), and give the identity that runs the notebook, interactively and in the pipeline, **Get** on the certificate and **Sign** on its key. Then set the `CERTIFICATE_NAME` parameter to the certificate's name.
4. **Replace the notebook's content** with the example file. With Git integration, copy it over the notebook's `notebook-content.py` and commit; otherwise make the edits listed in [What changed in the notebook](#what-changed-in-the-notebook) by hand. The example has placeholders where the tenant's values go, so fill them in:
   - In the parameters cell, set `CLIENT_ID`, `SHAREPOINT_HOSTNAME` and `CERTIFICATE_NAME`, or have the pipeline pass them as parameters, as it passes `DRY_RUN`.
   - In the smoke test cells, set `SMOKE_SITE_URL` to the site with the demo library, and replace the four `<test-user-…-upn>` values with the UPNs of the people to add to it.

   Nothing else (tables, the app registration's API permissions) needs to change.
5. **Dry run it.** Run the notebook with `DRY_RUN = True` and `RUN_SMOKE_TEST = True`. The imports cell prints the package version. Compare the action log rows for this run (in `it.sharepoint_action_log`, by `run_id`) with those of a recent dry run from the old notebook: they should match, apart from the [behaviour changes](#what-changed-in-behaviour). Membership sync's dry run reads each Auto group's owners: `group_owner_added (dry run)` rows are ownerless groups (how the notebook created them), which the live run adopts automatically, and `group_not_owned` rows are groups with an owner who isn't the app. Add the app's service principal as a co-owner of those, or the live run will leave them alone. Manual groups are checked the same way when their faculty next goes through the create pass, or all at once with [Adopting the existing Manual groups](#adopting-the-existing-manual-groups).
6. **Let the pipeline run it** as before. The pipeline passes `DRY_RUN = False`.
7. **Retire the old certificate.** Once the pipeline has run successfully, remove the old certificate from the app registration and delete the two Key Vault secrets that held its thumbprint and private key.

To roll back, restore the previous notebook content; it works for as long as the old certificate is on the app registration and its two secrets exist, so keep them until step 7. The package doesn't change any state the old notebook relies on: group names, library URLs, property bag keys and the watermark are all the same.

## Adopting the existing Manual groups

The notebook created its groups without an owner. The app is added as each group's owner the first time a live run uses it. Membership sync is a full pass, so from 1.0.2 it adopts every Auto group (1.0.1 skipped Auto groups with no members to add). The create pass only covers faculties changed since the watermark, so a faculty that hasn't changed never has its Manual groups checked.

To adopt them all, run the create pass once over every live faculty. Add this as a cell after **Create/Modify Libraries**, run it once with `DRY_RUN = True`, then once with `DRY_RUN = False`, and remove it again:

```python
# One-off: the create pass over every live faculty, not only those changed since
# the watermark, so the app is made owner of every Manual group. Safe to re-run.
pdf_all_libraries = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.current_faculty_name AS library_display_name
        ,f.academy_code || '-Sharepoint' AS group_prefix
        ,TRUE AS manual_read
        ,TRUE AS manual_contribute
        ,'Arbor' AS source_system
        ,f.faculty_unique_id AS source_id
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f')}
""").toPandas()

results = mgr.create_libraries(pdf_all_libraries)
persist_action_log(results, "adopt_manual_groups")
display(results[results["action"].isin(["group_owner_added", "group_owner_added (dry run)", "group_not_owned"])])
```

`group_owner_added` rows (ending `(dry run)` in a dry run) are the groups adopted. `group_not_owned` rows are groups owned by someone other than the app: check who created each one, then add the app's service principal as a co-owner or delete it. Apart from ownership, the cell only changes what had drifted from the faculty list, such as a display name or description, as the create pass is safe to re-run. The SLT groups are `app_owned=False`, so they're left with the headteacher.

## Final state of the notebook

Cells are listed in order. **Changed** marks the cells that differ from the previous version; every other cell is unchanged.

| # | Cell | What it does | Changed |
| --- | --- | --- | --- |
| 1-2 | *Manage Faculty Sharepoint Libraries*, *Setup* | Headings | |
| 3 | `from notebookutils import variableLibrary` | | |
| 4 | **Parameters** | `DRY_RUN` (default True; the pipeline passes False), `MAX_WORKERS`, `INCLUDE_SITE_NAME`, `CREATE_ONEDRIVE_SHORTCUTS`, the removal guardrail, `ENFORCE_SITE_SHARING_LOCKDOWN`, navigation, the SLT library, the rollout scope (`ACADEMY_FILTER`, `FACULTY_CODE_FILTER`), `RUN_SMOKE_TEST`, the app's client ID, the SharePoint host and `CERTIFICATE_NAME`, the app's certificate in Key Vault | **Changed** |
| 5 | Variable library | Reads the workspace, lakehouse, Key Vault and tenant from `VL_CORE` | **Changed** |
| 6 | Lakehouse tables | Names the curated faculty list and membership tables, and the action log, sync state and unresolved users tables | |
| 7-9 | *Load Libraries* | Imports, including the package, and prints the package version. Turns on INFO logging | **Changed** |
| 10-11 | *Core Setup* | Creates a `KeyVaultCertificateCredential` for the app's certificate, and the `SharePointLibraryManager` from the parameters | **Changed** |
| 12 | Run context | Run ID, `persist_action_log` (appends each operation's log to the audit table), the rollout scope predicate `scoped()`, the reserved SLT code and the shared eligibility predicate `eligible()` | **Changed** |
| 13-18 | *Smoke Test* | Only with `RUN_SMOKE_TEST`: against the IT site's demo library, reads metadata, creates the library, checks its navigation link, locks down sharing and checks it, and syncs four test members | **Changed** (cells 16 and 17) |
| 19-20 | *Real Data*, watermark | Creates the sync state table if needed, reads this scope's watermark, and reads the newest change as the upper bound for this run's incremental passes | **Changed** |
| 21-23 | *Site Sharing Lockdown* | Every in-scope hub, every run | |
| 24-26 | *SLT Libraries* | One SLT library per hub, with Manual groups only, owned by the headteacher rather than the app (`app_owned=False`). Reports faculties whose code collides with the SLT code, and hubs shared by two academies | **Changed** (cell 25) |
| 27-29 | *Create/Modify Libraries* | Faculties new or changed since the watermark | **Changed** (bounded by the run's upper watermark) |
| 30-32 | *Unlink Deleted Faculties* | Faculties deleted at source since the watermark | **Changed** (bounded by the run's upper watermark) |
| 33-35 | *Retire Moved Libraries* | Faculties whose hub changed since the watermark: retires the copy on the old hub, once the library exists on the new one | **Changed** (bounded by the run's upper watermark; waits for the new library) |
| 36-38 | *SLT Access to Faculty Libraries* | Grants each academy's SLT groups Contribute and Read on every faculty library, every run | **Changed** (cell 37) |
| 39-41 | *Update Memberships* | Full membership sync for every in-scope faculty, including the all-staff nested groups | |
| 42-43 | *Unresolved Users* | Appends users that couldn't be resolved to the unresolved users table, with academy and subject | |
| 44-45 | *Run Outcome* | Shows warnings, and fails the run if any action failed | |
| 46 | Watermark | Advances the watermark, only after a clean run that wasn't a dry run | **Changed** (escaped process key) |

The cell that held the core code, and its *Core Code* heading, have been removed.

## What changed in the notebook

**Parameters and variable library.** `CERT_THUMBPRINT_SECRET_NAME` and `CERT_PRIVATE_KEY_SECRET_NAME` are replaced by one parameter, `CERTIFICATE_NAME`, the name of the app's certificate in the `KEY_VAULT` key vault. The variable library cell no longer reads the two secrets.

**Load Libraries.** The imports the core code needed (`msal`, `requests`, `threading` and others) are gone. The notebook imports the package instead:

```python
import logging
import uuid
from datetime import datetime, timezone

import pandas as pd

# The library manager lives in alt_fabric_m365 (pinned in the attached
# environment): https://github.com/ArchwayTrust/Insights-Fabric-M365
import alt_fabric_m365
from alt_fabric_m365.auth import KeyVaultCertificateCredential
from alt_fabric_m365.sharepoint import (
    ACTION_LOG_COLUMNS,
    FATAL_ACTIONS,
    WARNING_ACTIONS,
    SharePointLibraryManager,
    sanitise_for_group_name,
)

print(f"alt_fabric_m365 {alt_fabric_m365.__version__}")
```

**Core Setup.** The manager takes a credential rather than the tenant, client ID and certificate. It's the package's `KeyVaultCertificateCredential`, created without a scope: the manager asks it for a Graph token and a SharePoint token, and it caches each until shortly before it expires. This replaces the notebook's own MSAL code, which read the private key from Key Vault secrets into the notebook. The other arguments are unchanged.

```python
credential = KeyVaultCertificateCredential(
    tenant_id=TENANT_ID,
    client_id=CLIENT_ID,
    akv_url=KEY_VAULT_PATH,
    certificate_name=CERTIFICATE_NAME,
)

mgr = SharePointLibraryManager(
    credential,
    sharepoint_hostname=SHAREPOINT_HOSTNAME,
    dry_run=DRY_RUN,
    # ... the other arguments as before
)
```

**Run context.** `FATAL_ACTIONS` and `WARNING_ACTIONS` are imported rather than defined here, as the package defines the action names. They hold the same values. `SLT_CODE_KEY` uses the now public `sanitise_for_group_name`. `FACULTY_CODE_FILTER` is now quote-escaped before it goes into the SQL, as `ACADEMY_FILTER` already was.

**Watermark.** Two fixes from review, each covered by a notebook test:

- The create, unlink and retire passes each read the faculty list separately. A row landing between two reads could be missed by an earlier pass, yet still be covered by the watermark advanced from a later pass. The watermark cell now reads the newest `last_changed_ts` once, as `HIGH_WATERMARK`, and all three passes only take rows up to it. Rows newer than that are left for the next run.
- The watermark key contains the rollout filters, so a `'` in a filter broke the watermark's lookup and its `MERGE`. Both now use `SYNC_PROCESS_LITERAL`, the quote-escaped key.

**Smoke test.** The navigation and sharing checks read the site through the manager instead of calling private functions with its token:

| Before | After |
| --- | --- |
| `_find_nav_label(mgr._sp_token, SMOKE_SITE_URL, NAVIGATION_LABEL)`, then a lookup of `_normalise_nav_url(...)` in its children | `mgr.get_navigation_link(SMOKE_SITE_URL, SMOKE_LIBRARY_CODE)` |
| `_get_site_sharing_settings(mgr._sp_token, SMOKE_SITE_URL)` | `mgr.get_site_sharing_settings(SMOKE_SITE_URL)` |

**SLT checks.** The warning rows' columns come from `ACTION_LOG_COLUMNS` rather than a hand-written list.

**SLT access.** Group names are built with `mgr.group_name(row.group_prefix, row.site_url, SLT_LIBRARY_CODE, suffix)`, which uses the manager's own `include_site_name`, instead of `_build_group_name(..., include_site_name=INCLUDE_SITE_NAME)`. The names are the same.

## What changed in behaviour

The core code moved over unchanged. A script compared every function and constant in the package with the notebook's version, as parsed Python: 105 are identical, and the rest differ only by the authentication change above (the MSAL functions are gone) and by using the shared `ACTION_LOG_COLUMNS` list. These fixes were then made, each with a test that failed before it:

| Fix | Effect on the notebook |
| --- | --- |
| Group lookups put their `$filter` straight into the URL, so a group name containing `&`, `#` or `+` broke the query. | An `all_staff_security_group` such as `Maths & Science Staff` failed every run with an `error` row. It now resolves. |
| One group lookup matched names case-sensitively, the other didn't (as Graph doesn't). The group-name collision check now also ignores case, so codes such as `PE-KS3` and `peks3` are reported instead of sharing groups. | A group whose name differed only in case was missed by membership sync, which created a duplicate. Now it's reused. |
| `site_url` was only checked to be http or https, and the SharePoint token is sent to it. | A `site_url` that isn't `https://` on `SHAREPOINT_HOSTNAME` now raises a `ValueError` for that operation, before any request. All hub URLs should already be on that host. |
| Redirects from https to http were followed, with the token, and GET redirects weren't checked at all. | Every redirect hop, for every method, must stay on the same host and on https. Requests to SharePoint and Graph aren't redirected elsewhere in normal use. |
| A user principal name containing `/` could change the path of the Graph user lookup. | The `/` is now encoded. Valid UPNs are sent exactly as before. |
| Groups were found by name alone, and names are predictable, so a same-named group created by someone else would have been granted the library's access. | The app is made owner of the groups it creates, and an existing group is only granted or synced if the app owns it. Groups the notebook created before have no owner, so the app is added as their owner on first use (`group_owner_added`). A group owned only by others is skipped with a `group_not_owned` warning. The SLT groups are exempt (`app_owned=False`), as the headteacher owns them. No new permission is needed: `Group.ReadWrite.All` covers owners. |
| A faculty rename changed the library's display name but left each group's description naming the old one. | The create pass now updates the descriptions of the groups the app owns. Group names never change, as they use the faculty code. |
| The retire pass ran even when creating the faculty's library on its new hub had failed, removing the old library's access with no replacement. | The retire query passes `new_site_url`, and a library is only retired once it exists on the new hub (`retire_deferred` until then). A faculty deleted since its move is still retired. |
| An ordinary SharePoint list whose URL ended in a faculty code (`/Lists/ENG`) was treated as that faculty's library. | Only document libraries are matched. |
| Permission levels were looked up by their English names. | They're found by type, so renamed or translated levels still work. |
| Unlink emptied Auto groups without checking who owns them. | Like the other passes, it only empties groups the app owns. |
| A redirect that couldn't be followed, such as a POST 303, was returned as if it had worked. | It's an error. OneDrive shortcut creation, like group creation, is no longer repeated after a failure that may have succeeded. |
| Unlinking a library that doesn't exist (a faculty created and deleted between runs, or a library deleted by hand) was a fatal `error`. | Now a `library_not_found` warning, and the Auto groups are still emptied. Before, the run failed and the watermark never advanced past the event. |
| Creating a group or navigation link was retried after a 5xx or timeout, even though it may have succeeded. | Only retried when the request certainly didn't arrive (429, connect timeout), so no duplicate group can block a library. |
| The collision check only compared libraries in the same call, but create runs are incremental. | A new code that would share groups with a library already on the site (e.g. `PEKS3` after `PE-KS3`, or an unlinked library still holding the groups) is now skipped with `group_name_collision`. |
| A library code also matched other libraries' titles, so code `ART` could take over a library titled "Art". | Libraries are only found by URL. A code that is another library's title is reported as `library_name_conflict`. |
| Membership sync read a just-created group's members, which can 404 while it replicates. | No longer read: a new group has none. |
| In create and membership sync, one site that couldn't be read (e.g. a mistyped hub URL) stopped the whole operation. | Only that site's libraries are skipped, with an `error` row. |
| `rename_libraries` could rename groups onto names already in use. | It no longer renames groups at all: a faculty code is the library's URL, so it never changes. The notebook doesn't call it. |
| In grants, one group SharePoint couldn't resolve skipped every grant on its site. | Only that group's grants are skipped. |
| `get_library_metadata` returned `NaN` instead of `None` for unset values under pandas 3. | No effect under the current runtime's pandas 2. |

### Fewer requests

Membership sync and the SLT passes run over every faculty on every run, so work that changes nothing adds up. On a run where nothing has changed:

- **Role assignments** are read once per library, and only missing ones are written. Before, every group was re-registered (`ensureuser`) and re-assigned on every run: two writes per group per library. Unchanged groups are now logged as `already_assigned`.
- **Users** are resolved once per membership sync, not once per library they belong to. Each resolution can take up to three Graph lookups.
- **Role definitions** are read in one request per site, not three.
- **Libraries** are listed once per site for metadata reads and writes, instead of once per library.

## Where the core code went

| Section of the old core cell | Now in |
| --- | --- |
| Constants, thread-safe action log, shared validation helpers | [`sharepoint/_common.py`](../alt_fabric_m365/sharepoint/_common.py) |
| Resilient HTTP wrapper | [`sharepoint/_http.py`](../alt_fabric_m365/sharepoint/_http.py) |
| Authentication (`_build_msal_app`, `_token_from_app`, `_validate_and_normalise_certificate`) | Replaced by the package's existing [`auth.KeyVaultCertificateCredential`](../alt_fabric_m365/auth/keyvault.py) |
| Graph API helpers, OneDrive shortcut helpers | [`sharepoint/_graph.py`](../alt_fabric_m365/sharepoint/_graph.py) |
| Entra ID security group management, group-name collisions | [`sharepoint/_groups.py`](../alt_fabric_m365/sharepoint/_groups.py) |
| SharePoint REST: permissions, property bag, site sharing settings | [`sharepoint/_sharepoint.py`](../alt_fabric_m365/sharepoint/_sharepoint.py) |
| Site navigation (quick launch) | [`sharepoint/_navigation.py`](../alt_fabric_m365/sharepoint/_navigation.py) |
| `SharePointLibraryManager` | [`sharepoint/manager.py`](../alt_fabric_m365/sharepoint/manager.py) |

Function names are unchanged, so a fix made to another copy of this code can be found and applied here, except that `_build_group_name` and `_sanitise_for_group_name` are now public, without the underscore.

The old core cell said it was an inlined copy of `SharePointLibraryPermissions` in the Systems-CoreScripts repository, to be kept in sync. This package is now the copy the notebook runs, so either retire that one or note there that this package supersedes it.

## Test coverage

All tests run without credentials or a Fabric workspace: [`tests/sharepoint_fake.py`](../tests/sharepoint_fake.py) is an in-memory fake of the Graph and SharePoint REST endpoints the manager calls. It replaces `requests.request`, so the real request building, retries and response handling all run, and it keeps state (groups, members, libraries, role assignments, navigation, sharing settings), so tests check what an operation actually changed and that running it again changes nothing. Like the real services, it parses query strings out of the URL, which is how the `&` bug above was caught.

The suite went from 102 to 323 tests. It passes on Python 3.10, 3.11, 3.12 and 3.13, and with pandas 2 and pandas 3.

| File | Tests | Covers |
| --- | --- | --- |
| [`test_sharepoint_http.py`](../tests/test_sharepoint_http.py) | 36 | `Retry-After` in seconds and as a date (with clamping, NaN and past dates); retries on 429, 5xx, connection errors and timeouts, with backoff and jitter; redirects for POST keep the method; cross-host, https-to-http and looping redirects are refused |
| [`test_sharepoint_helpers.py`](../tests/test_sharepoint_helpers.py) | 74 | Group naming, mail nicknames and collisions; DataFrame and URL validation; parallel runs logging failures without stopping; library lookup by slug and display name, with paging; duplicate and case-differing group names; member keys for users, groups and service principals; replication waits; user resolution by UPN, mail and proxy address, malformed and unknown users; nestable group resolution; property bag XML escaping and errors; navigation URL normalisation |
| [`test_sharepoint_manager.py`](../tests/test_sharepoint_manager.py) | 79 | Every operation end to end: creating libraries with groups, metadata and navigation, and again changing nothing; renaming; dry runs making no writes; manual-only (SLT) libraries; colliding codes; one failing library not stopping others; parallel runs. Membership sync adding users and aliases, removing stale members, keeping service principals, emptying roles without creating groups, the removal guardrail, unresolved users blocking removals, nested groups, OneDrive shortcuts, `add_failed` and `remove_failed`. Unlink keeping Manual groups; retire removing access but not membership; grants being additive and idempotent; site sharing lockdown, including a write that doesn't stick. Authenticating with the real `KeyVaultCertificateCredential` (Entra ID and Key Vault faked): one Graph and one SharePoint token, each sent only to its own API |
| [`test_faculty_notebook.py`](../tests/test_faculty_notebook.py) | 10 | Runs the example notebook cell by cell, as Fabric does, with Spark SQL run on SQLite over small faculty tables and the real `KeyVaultCertificateCredential`, and checks a live run, a second run, a dry run, a failed run and a widened scope (below). Also checks the notebook has no undefined names and uses only the package's public API |

Line and branch coverage of `alt_fabric_m365.sharepoint` is 97%. Most of what's left is handling for rarer failures, such as a site's prefetch failing part way through a rename.

The notebook test's scenarios, against a pilot scope of `BAA` and `MATH%`:

- **Live run.** Creates the SLT library and the in-scope faculty library, with metadata, groups and navigation; unlinks the deleted faculty's library; retires the moved faculty's copy on its old hub; grants the SLT groups access; syncs members and the all-staff nested group; locks down sharing; persists the action log and the unresolved user; advances the watermark. The out-of-scope academy and faculty are left alone.
- **Second run.** Finds no new watermark events and makes no membership changes.
- **Dry run.** Makes no writes to SharePoint or Entra ID and leaves the watermark alone.
- **Authentication.** The certificate named by `CERTIFICATE_NAME` is read from the key vault and signs the assertions, and the run uses one Graph and one SharePoint token.
- **Rows landing mid-run.** A new faculty and a newer move added after the create read are both left for the next run, which picks them up.
- **Quoted scope values.** A `'` in `ACADEMY_FILTER` reaches the watermark lookup and `MERGE` safely.
- **Failed run.** A failed member addition makes the run outcome cell raise, and the watermark isn't advanced.
- **Widened scope.** Changing the filters starts a new watermark and replays events for the newly in-scope faculties, and the faculty coded `SLT` is reported and skipped.

**Not covered:** whether Microsoft Graph and SharePoint accept these requests. The fake follows the documented APIs, but only a run against the tenant can show it. The notebook's smoke test does that against the IT site's demo library, so run it (with `DRY_RUN = True` first) before switching the pipeline over, and after upgrading the package.

## Possible next steps

- **Use a non-exportable key.** Creating the certificate in Key Vault as in step 3, rather than importing one, means nobody can download its private key.
- **Pin the package version** in the environment, and upgrade deliberately: run the smoke test with `DRY_RUN = True` after each upgrade.
