# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "synapse_pyspark"
# META   },
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse_name": "",
# META       "default_lakehouse_workspace_id": ""
# META     },
# META     "environment": {}
# META   }
# META }

# MARKDOWN ********************

# # Manage Faculty Sharepoint Libraries

# MARKDOWN ********************

# ## Setup

# CELL ********************

from notebookutils import variableLibrary

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# PARAMETERS CELL ********************

# Safety: True = log intended actions only, mutate nothing. The pipeline passes
# False explicitly; an interactive "Run all" therefore never touches the tenant.
# Notebook should be left as defaulting to True while the pipeline overrides it.
DRY_RUN = True

# How many libraries are processed concurrently (1 = sequential).
MAX_WORKERS = 4

# True inserts the site name into security group names
# ({prefix}-{site}-{code}-{suffix} rather than {prefix}-{code}-{suffix}).
# Redundant here — the academy prefix (e.g. BAA-Sharepoint) already scopes
# the name — so kept False. Do NOT flip after rollout: names are how the
# automation finds existing groups, so changing the format orphans every
# group already created and provisions duplicates under the new names.
INCLUDE_SITE_NAME = False

# True also drops a shortcut to the library into the OneDrive root of each
# member newly added by membership sync (existing members are untouched).
CREATE_ONEDRIVE_SHORTCUTS = False

# Guardrails for membership sync: a group's removals are skipped when they
# exceed its cap, and every removal is skipped when the run's total would
# exceed the run cap, unless mass removal is explicitly allowed for the run.
MAX_REMOVALS_PER_GROUP = 10
MAX_REMOVALS_PER_RUN = 100
ALLOW_MASS_REMOVAL = False

# Site-level hardening: once per site per run, turn off "Allow members to
# share the site and individual files and folders", "Allow access requests",
# and "Allow members to invite others to the site members group". False
# leaves site sharing settings untouched.
ENFORCE_SITE_SHARING_LOCKDOWN = True

# True links each library into the site's quick-launch navigation, nested
# under the NAVIGATION_LABEL heading (the label is created on first use).
# Ensure semantics: pre-existing libraries are backfilled the next time
# their faculty row changes, and a manually removed link is re-added then.
ADD_TO_SITE_NAVIGATION = True
NAVIGATION_LABEL = "Document Libraries"

# One SLT library per academy hub (created when missing), governed only by
# the hand-curated {academy}-Sharepoint-SLT-Manual-Read / -Manual-Contribute
# groups — the headteacher is assigned as their owner by hand and the
# automation never syncs their membership.  The same two groups are also
# granted Read / Contribute on every faculty library of the academy. The
# code is reserved: a faculty whose code collides with it is skipped with a
# warning. Do NOT change the code after rollout — it is the library's URL
# slug and the group-name segment, so changing it orphans both.
MANAGE_SLT_LIBRARY = True
SLT_LIBRARY_CODE = "SLT"
SLT_LIBRARY_DISPLAY_NAME = "SLT"

# Rollout scope: ACADEMY_FILTER is a comma-separated list of academy codes
# matched with IN (e.g. "BAA" or "BAA,NCA,CSJ"); FACULTY_CODE_FILTER is a
# SQL LIKE pattern on static_faculty_code. Empty string = no filter. Kept
# at the pilot scope until sign-off; widening the rollout is a parameter
# change, not a code change (the scope-keyed watermark replays newly
# in-scope events automatically).
ACADEMY_FILTER = "BAA"
FACULTY_CODE_FILTER = "MATH%"

# Set True to run the smoke-test cells against the demo library.
RUN_SMOKE_TEST = False

# Service principal used for Graph/SharePoint calls, and the name of its
# certificate in the KEY_VAULT key vault. Key Vault signs with the
# certificate's private key, which never leaves the key vault; the identity
# running the notebook needs Get on the certificate and Sign on its key.
# Replace these placeholders with your tenant's values, or pass them as
# pipeline parameters.
CLIENT_ID = "<app-registration-client-id>"
SHAREPOINT_HOSTNAME = "<tenant>.sharepoint.com"
CERTIFICATE_NAME = "<certificate-name>"

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

vl = variableLibrary.getLibrary("VL_CORE")

STORE_WORKSPACE_NAME = vl.getVariable("STORE_WORKSPACE_NAME")
CURATED_LH_NAME      = vl.getVariable("CURATED_LH_NAME")
BASE_LH_NAME         = vl.getVariable("BASE_LH_NAME")
KEY_VAULT_PATH       = vl.getVariable("KEY_VAULT")
TENANT_ID            = vl.getVariable("TENANT_ID")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Explicitly name all lakehouse tables
curated_faculty_list       = f"{STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.it.sharepoint_faculty_list"
curated_faculty_membership = f"{STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.it.sharepoint_faculty_membership"
action_log_table           = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.it.sharepoint_action_log"
sync_state_table           = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.it.sharepoint_sync_state"
unresolved_users_table     = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.it.sharepoint_unresolved_users"


print("Lakehouse tables:")
print(curated_faculty_list)
print(curated_faculty_membership)
print(action_log_table)
print(sync_state_table)
print(unresolved_users_table)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Load Libraries

# CELL ********************

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

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

logging.basicConfig(level=logging.INFO)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Core Setup

# CELL ********************

# No scope: the manager requests Graph and SharePoint tokens as it needs them.
credential = KeyVaultCertificateCredential(
    tenant_id=TENANT_ID,
    client_id=CLIENT_ID,
    akv_url=KEY_VAULT_PATH,
    certificate_name=CERTIFICATE_NAME,
)

mgr = SharePointLibraryManager(
    credential,
    sharepoint_hostname=SHAREPOINT_HOSTNAME,
    dry_run=DRY_RUN,                                                # default for all operations — override per call
    max_workers=MAX_WORKERS,                                        # default parallelism — override per call
    include_site_name=INCLUDE_SITE_NAME,                            # controls whether security groups include the site name
    create_onedrive_shortcuts=CREATE_ONEDRIVE_SHORTCUTS,            # controls whether links are created in user one drive
    max_removals_per_group=MAX_REMOVALS_PER_GROUP,                  # membership-sync removal guardrails
    max_removals_per_run=MAX_REMOVALS_PER_RUN,
    allow_mass_removal=ALLOW_MASS_REMOVAL,
    add_to_navigation=ADD_TO_SITE_NAVIGATION,                       # link libraries into quick-launch navigation
    navigation_label=NAVIGATION_LABEL,                              # heading the links are nested under
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Run context shared by every operation below: one run id for the audit trail,
# a per-operation persist helper, and the collected results for the error gate.

RUN_ID = str(uuid.uuid4())
run_results: list[tuple[str, pd.DataFrame]] = []

# FATAL_ACTIONS (fail the run and pipeline alerting) and WARNING_ACTIONS
# (surfaced, but do not fail the run) are imported from the package, which
# defines the action names.


def persist_action_log(results_pdf: pd.DataFrame, operation: str) -> None:
    """Append an operation's action log to the Delta audit table."""
    run_results.append((operation, results_pdf))
    if results_pdf.empty:
        print(f"{operation}: no actions logged.")
        return
    log_pdf = results_pdf.copy()
    log_pdf["run_id"] = RUN_ID
    log_pdf["operation"] = operation
    log_pdf["dry_run"] = DRY_RUN
    log_pdf["logged_at"] = datetime.now(timezone.utc)
    (spark.createDataFrame(log_pdf)
        .write.format("delta").mode("append").saveAsTable(action_log_table))
    print(f"{operation}: {len(log_pdf)} action(s) persisted to {action_log_table} (run {RUN_ID}).")


# ACADEMY_FILTER parsed once: comma-separated codes → sorted, deduplicated
# list. Sorted so the watermark scope key below is order-independent
# ("BAA,NCA" and "NCA,BAA" are the same scope).
ACADEMY_CODES = sorted({
    code.strip() for code in ACADEMY_FILTER.split(",") if code.strip()
})


def scoped(alias: str) -> str:
    """Rollout-scope predicate fragment for the faculty list aliased as *alias*."""
    clause = ""
    if ACADEMY_CODES:
        quoted = ", ".join("'" + code.replace("'", "''") + "'" for code in ACADEMY_CODES)
        clause += f" AND {alias}.academy_code IN ({quoted})"
    if FACULTY_CODE_FILTER:
        pattern = FACULTY_CODE_FILTER.replace("'", "''")
        clause += f" AND {alias}.static_faculty_code LIKE '{pattern}'"
    return clause


# The SLT library's code, as the group-name segment it becomes. A faculty
# whose code sanitises to the same segment would share the SLT library's URL
# slug and groups, so such faculties are never eligible.
SLT_CODE_KEY = sanitise_for_group_name(SLT_LIBRARY_CODE).upper()


def eligible(alias: str, *, reserved_code: bool = False) -> str:
    """Shared eligibility predicate for the faculty list aliased as *alias*.

    The single definition of "a faculty this automation manages" — every
    driver query below composes it (adding its own deleted_in_source /
    watermark conditions), so create, unlink, retire, membership sync, the
    SLT passes, and the sharing lockdown cannot drift out of scope with one
    another.

    Faculties whose code collides with the reserved SLT library code are
    excluded.  ``reserved_code=True`` selects exactly those otherwise-eligible
    faculties instead, so they can be surfaced as warnings rather than
    skipped silently.
    """
    # SQL mirror of sanitise_for_group_name.  Deliberately ASCII: a Unicode
    # class (\p{L}) would need backslashes whose meaning depends on Spark's
    # string-literal escaping config.  The only divergence — a non-ASCII code
    # such as 'SLTé' — errs towards a loud reserved-code warning, never a
    # silent collision.
    code_key = f"UPPER(regexp_replace({alias}.static_faculty_code, '[^A-Za-z0-9]', ''))"
    return f"""COALESCE({alias}.exclude_from_sharepoint, FALSE) = FALSE
        AND {alias}.sharepoint_hub_url IS NOT NULL
        AND {alias}.static_faculty_code IS NOT NULL
        AND {alias}.academy_code IS NOT NULL
        AND {code_key} {'=' if reserved_code else '<>'} '{SLT_CODE_KEY}'{scoped(alias)}"""

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Smoke Test (RUN_SMOKE_TEST only)
# Exercises the demo library end-to-end: metadata read, create (including the
# quick-launch navigation link, verified by read-back), site sharing lockdown
# (verified by read-back), and membership sync. Skipped unless the
# RUN_SMOKE_TEST parameter is True — and even then, DRY_RUN still governs
# whether anything is mutated; the read-back assertions only fire on live
# runs (a dry run creates nothing to find). The metadata read doubles as the
# re-link example should the SCD table ever be lost (libraries carry
# SourceSystem/SourceId in their property bags).

# CELL ********************

SMOKE_SITE_URL = "https://<tenant>.sharepoint.com/sites/<smoke-test-site>"
SMOKE_LIBRARY_CODE = "DEMO_LIB"
SMOKE_DISPLAY_NAME = "Demo Library & Link Test"   # & exercises URL normalisation

if RUN_SMOKE_TEST:
    query = pd.DataFrame({
        "site_url": [SMOKE_SITE_URL],
        "library_code": [SMOKE_LIBRARY_CODE],
    })

    metadata_df = mgr.get_library_metadata(query)
    display(metadata_df)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if RUN_SMOKE_TEST:
    pdf_test_libs = pd.DataFrame({
        "site_url": [SMOKE_SITE_URL],
        "library_code": [SMOKE_LIBRARY_CODE],        # becomes the URL slug
        "library_display_name": [SMOKE_DISPLAY_NAME],  # friendly name shown in UI
        "group_prefix": ["IT-Sharepoint"],
        "manual_read": [True],
        "manual_contribute": [True],
        "source_system": ["Manual"],            # optional: provenance tag
        "source_id": ["Manual01"],              # optional: source record ID
    })

    results = mgr.create_libraries(pdf_test_libs)
    persist_action_log(results, "smoke_test_create_libraries")
    display(results)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if RUN_SMOKE_TEST and ADD_TO_SITE_NAVIGATION:
    # Verify the navigation link by reading the quick launch back (read-only,
    # so it runs in dry-run too — but only asserts on live runs). This reads
    # the site's navigation, so it checks what create_libraries wrote, not
    # what it reported.
    link_node = mgr.get_navigation_link(SMOKE_SITE_URL, SMOKE_LIBRARY_CODE)
    print(f"Navigation link node: {link_node}")

    if not DRY_RUN:
        if link_node is None:
            raise RuntimeError(
                f"smoke test: no '{NAVIGATION_LABEL}' link found for "
                f"{SMOKE_LIBRARY_CODE} on {SMOKE_SITE_URL}"
            )
        if link_node.get("Title") != SMOKE_DISPLAY_NAME:
            raise RuntimeError(
                f"smoke test: navigation link title is '{link_node.get('Title')}', "
                f"expected '{SMOKE_DISPLAY_NAME}'"
            )
        print("Navigation link smoke check passed.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if RUN_SMOKE_TEST:
    # Site sharing hardening against the demo site, then a read-back
    # assertion. NOTE: a live (non-dry) run genuinely turns off member
    # sharing, access requests, and member invites on the IT team site —
    # the same change the real pass applies to every faculty hub.
    results = mgr.enforce_site_sharing_lockdown([SMOKE_SITE_URL])
    persist_action_log(results, "smoke_test_site_sharing_lockdown")
    display(results)

    settings = mgr.get_site_sharing_settings(SMOKE_SITE_URL)
    print(f"Effective sharing settings on {SMOKE_SITE_URL}: {settings}")

    if not DRY_RUN:
        still_on = [
            name for name in (
                "members_can_share", "access_requests_enabled", "members_can_invite",
            )
            if settings[name]
        ]
        if still_on:
            raise RuntimeError(
                f"smoke test: site sharing lockdown left enabled: {', '.join(still_on)}"
            )
        print("Site sharing lockdown smoke check passed.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if RUN_SMOKE_TEST:
    pdf_test_memberships = pd.DataFrame({
        "user_principal_name": [
            # Placeholders: the UPNs of a few people to add to the demo library
            "<test-user-1-upn>",
            "<test-user-2-upn>",
            "<test-user-3-upn>",
            "<test-user-4-upn>",
        ],
    }).assign(
        site_url=SMOKE_SITE_URL,
        library_code=SMOKE_LIBRARY_CODE,
        permission="contribute",
        group_prefix="IT-Sharepoint",
    )

    results = mgr.update_memberships(pdf_test_memberships, create_onedrive_shortcuts=True)
    persist_action_log(results, "smoke_test_update_memberships")
    display(results)

    # Users whose email didn't match any UPN, mail, or proxyAddresses in Entra ID:
    not_found = results[results["action"] == "user_not_found"]
    if not not_found.empty:
        display(not_found)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Real Data
# Driven by the curated faculty tables. Create/unlink/retire are incremental:
# they process rows whose `last_changed_ts` is newer than this notebook's
# stored watermark (advanced only after a clean, non-dry run). Membership
# sync, the SLT library and SLT access passes, and the site sharing lockdown
# are full passes over every live, in-scope faculty/site each run.

# CELL ********************

# Watermark: process faculty-list rows changed since the last successful run.

# CREATE TABLE does not create its schema — a fresh schema-enabled
# lakehouse fails with SCHEMA_NOT_FOUND without this.
spark.sql(f"CREATE SCHEMA IF NOT EXISTS {STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.it")

spark.sql(f"""
    CREATE TABLE IF NOT EXISTS {sync_state_table} (
        process            STRING,
        last_processed_ts  TIMESTAMP,
        updated_at         TIMESTAMP
    )
    USING DELTA
""")

# The watermark is keyed by rollout scope: changing ACADEMY_FILTER /
# FACULTY_CODE_FILTER starts a fresh watermark row at 1900 for the new
# scope, so events the old scope filtered out are replayed instead of being
# permanently consumed — widening the rollout really is just a parameter
# change.  Replays are harmless: create/unlink/retire are ensure-semantics.
# The academy part uses the parsed, sorted codes so reordering the list is
# not a scope change (adding/removing a code is, and triggers the replay).
SYNC_PROCESS_NAME = (
    f"faculty_sharepoint_sync|{','.join(ACADEMY_CODES) or '*'}|{FACULTY_CODE_FILTER or '*'}"
)

# The filters are free text, so quote-escape the key wherever it is a SQL literal.
SYNC_PROCESS_LITERAL = SYNC_PROCESS_NAME.replace("'", "''")

wm_rows = spark.sql(f"""
    SELECT last_processed_ts FROM {sync_state_table}
    WHERE process = '{SYNC_PROCESS_LITERAL}'
""").collect()

WATERMARK = (
    wm_rows[0]["last_processed_ts"]
    if wm_rows and wm_rows[0]["last_processed_ts"] is not None
    else datetime(1900, 1, 1)
)
WATERMARK_LITERAL = WATERMARK.strftime("%Y-%m-%d %H:%M:%S.%f")

# Upper bound for the create, unlink and retire passes, read once before any
# of them. Each pass reads the faculty list separately, so without a shared
# bound a row landing between two reads could be missed by the earlier pass
# yet covered by the watermark advanced from a later one. Rows newer than
# this are left for the next run.
hw_rows = spark.sql(f"SELECT MAX(last_changed_ts) AS max_ts FROM {curated_faculty_list}").collect()
HIGH_WATERMARK = (
    hw_rows[0]["max_ts"]
    if hw_rows and hw_rows[0]["max_ts"] is not None
    else WATERMARK
)
HIGH_WATERMARK_LITERAL = HIGH_WATERMARK.strftime("%Y-%m-%d %H:%M:%S.%f")

print(
    f"Run {RUN_ID}: processing changes newer than {WATERMARK_LITERAL}, "
    f"up to {HIGH_WATERMARK_LITERAL} (DRY_RUN={DRY_RUN})."
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Site Sharing Lockdown
# Full check each run (not watermark-gated): sharing settings can be turned
# back on by hand in the site UI, so drift is corrected every run. One check
# per site, not per library — already-compliant sites are only read.

# CELL ********************

df_sites = spark.sql(f"""
    SELECT DISTINCT
        f.sharepoint_hub_url AS site_url
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f')}
""")

pdf_sites = df_sites.toPandas()
display(pdf_sites)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if ENFORCE_SITE_SHARING_LOCKDOWN:
    results = mgr.enforce_site_sharing_lockdown(pdf_sites["site_url"])
    persist_action_log(results, "site_sharing_lockdown")
    display(results)
else:
    print("ENFORCE_SITE_SHARING_LOCKDOWN is False — site sharing settings not checked.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### SLT Libraries
# Full pass each run (not watermark-gated): every academy hub with at least one
# live, in-scope faculty gets an SLT library, created when missing. It has no
# Auto groups — access comes only from the hand-curated Manual groups, e.g. for
# NES:
# - NES-Sharepoint-SLT-Manual-Contribute (Contribute, membership NOT synced)
# - NES-Sharepoint-SLT-Manual-Read (Read, membership NOT synced)
# The headteacher is assigned as owner of both groups by hand. With
# ADD_TO_SITE_NAVIGATION on, the library is also linked under the
# NAVIGATION_LABEL heading. Faculties whose code collides with the reserved SLT
# code are reported here and skipped by every pass.

# CELL ********************

df_slt_sites = spark.sql(f"""
    SELECT DISTINCT
        f.sharepoint_hub_url AS site_url
        ,f.academy_code || '-Sharepoint' AS group_prefix
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f')}
""")

pdf_slt_libraries = df_slt_sites.toPandas().assign(
    library_code=SLT_LIBRARY_CODE,
    library_display_name=SLT_LIBRARY_DISPLAY_NAME,
    auto_groups=False,          # outside the dynamic flow — Manual groups only
    manual_read=True,
    manual_contribute=True,
    app_owned=False,            # the headteacher owns the SLT groups, not the app
)

# Faculties the reserved-code exclusion in eligible() skips.
pdf_reserved_code_faculties = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.current_faculty_name
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f', reserved_code=True)}
""").toPandas()

# A hub shared by two academies would give its one /SLT library two sets of
# SLT groups, and create_libraries would silently keep only the first row.
# dim_sites maps each academy to its own hub today, so this only guards the
# assumption: such hubs are reported and skipped rather than decided by row
# order.
shared_hub = pdf_slt_libraries["site_url"].duplicated(keep=False)

# Both conditions are surfaced as warnings rather than skipped silently.
slt_warnings = [
    pd.DataFrame({
        "site_url": pdf_reserved_code_faculties["site_url"],
        "document_library": pdf_reserved_code_faculties["library_code"],
        "status": (
            "faculty '" + pdf_reserved_code_faculties["current_faculty_name"].fillna("")
            + f"' has a code reserved for the academy SLT library ('{SLT_LIBRARY_CODE}') — "
            "skipped by every pass. If it was provisioned before the SLT library "
            "existed, empty its Auto groups by hand; set exclude_from_sharepoint "
            "to acknowledge"
        ),
    }),
    pd.DataFrame({
        "site_url": pdf_slt_libraries.loc[shared_hub, "site_url"],
        "document_library": SLT_LIBRARY_CODE,
        "status": (
            "hub is shared with another academy — SLT library for "
            + pdf_slt_libraries.loc[shared_hub, "group_prefix"]
            + " skipped; one hub per academy is assumed"
        ),
    }),
]
pdf_slt_warnings = pd.concat(slt_warnings, ignore_index=True).assign(
    group="", user="", action="group_name_collision", permission="",
)[ACTION_LOG_COLUMNS]
if not pdf_slt_warnings.empty:
    persist_action_log(pdf_slt_warnings, "slt_library_checks")
    display(pdf_slt_warnings)

pdf_slt_libraries = pdf_slt_libraries[~shared_hub]
display(pdf_slt_libraries)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if MANAGE_SLT_LIBRARY:
    results = mgr.create_libraries(pdf_slt_libraries)
    persist_action_log(results, "create_slt_libraries")
    display(results)
else:
    print("MANAGE_SLT_LIBRARY is False — SLT libraries not managed.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Create/Modify Libraries
# New, renamed, and resurrected faculties since the watermark. Rename is
# handled by create_libraries (ensure semantics update the display name).

# CELL ********************

df_library_changes = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.current_faculty_name AS library_display_name
        ,f.academy_code || '-Sharepoint' AS group_prefix
        ,TRUE AS manual_read
        ,TRUE AS manual_contribute
        ,'Arbor' AS source_system
        ,f.faculty_unique_id AS source_id
        ,f.last_changed_ts
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND f.last_changed_ts > TIMESTAMP'{WATERMARK_LITERAL}'
        AND f.last_changed_ts <= TIMESTAMP'{HIGH_WATERMARK_LITERAL}'
        AND {eligible('f')}
""")

pdf_library_changes = df_library_changes.toPandas()

display(pdf_library_changes)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Per faculty this ensures the library (URL slug = static code, display name =
# current name) plus its groups, e.g. for code MATH under academy BAA:
#   BAA-Sharepoint-MATH-Auto-Read         (Read, membership synced)
#   BAA-Sharepoint-MATH-Auto-Contribute   (Contribute, membership synced)
#   BAA-Sharepoint-MATH-Manual-Read       (Read, membership NOT synced)
#   BAA-Sharepoint-MATH-Manual-Contribute (Contribute, membership NOT synced)
# With ADD_TO_SITE_NAVIGATION on, the library is also linked in the site's
# quick-launch navigation under the NAVIGATION_LABEL heading.

results = mgr.create_libraries(pdf_library_changes)
persist_action_log(results, "create_libraries")
display(results)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Unlink Deleted Faculties
# - Renames the display name to {library_code}_UNLINKED
# - Empties the Auto security groups (Manual groups are preserved)
# - Removes the quick-launch navigation link (when ADD_TO_SITE_NAVIGATION is on)

# CELL ********************

df_deleted_in_source = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.academy_code || '-Sharepoint' AS group_prefix
        ,f.last_changed_ts
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = TRUE
        AND f.last_changed_ts > TIMESTAMP'{WATERMARK_LITERAL}'
        AND f.last_changed_ts <= TIMESTAMP'{HIGH_WATERMARK_LITERAL}'
        AND {eligible('f')}
""")

pdf_deleted_in_source = df_deleted_in_source.toPandas()

display(pdf_deleted_in_source)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

results = mgr.unlink_libraries(pdf_deleted_in_source)
persist_action_log(results, "unlink_libraries")
display(results)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Retire Moved Libraries
# Faculties whose hub URL changed since the watermark (site migrations). The
# create pass above provisions the library on the new site; this pass cleans
# up the old one — renames it to {code}_MOVED, removes the automation groups'
# role assignments (the groups are reused on the new site, so without this
# they would keep granting live access to the leftover copy), and removes the
# old site's navigation link. Group membership is untouched.

# CELL ********************

# deleted_in_source is deliberately not filtered: a faculty that moved and
# was later deleted still needs its old-site leftover retired.

df_hub_moves = spark.sql(f"""
    SELECT
        f.previous_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.academy_code || '-Sharepoint' AS group_prefix
        -- The old library is only retired once its replacement exists on the
        -- new hub; not checked for a faculty since deleted, whose leftover is
        -- retired anyway.
        ,CASE WHEN COALESCE(f.deleted_in_source, FALSE) THEN NULL
              ELSE f.sharepoint_hub_url END AS new_site_url
        ,f.last_changed_ts
    FROM
        {curated_faculty_list} AS f
    WHERE
        f.previous_hub_url IS NOT NULL
        AND f.previous_hub_url <> f.sharepoint_hub_url
        AND f.last_changed_ts > TIMESTAMP'{WATERMARK_LITERAL}'
        AND f.last_changed_ts <= TIMESTAMP'{HIGH_WATERMARK_LITERAL}'
        AND {eligible('f')}
""")

pdf_hub_moves = df_hub_moves.toPandas()

display(pdf_hub_moves)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

results = mgr.retire_moved_libraries(pdf_hub_moves)
persist_action_log(results, "retire_moved_libraries")
display(results)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### SLT Access to Faculty Libraries
# Full pass each run: ensures each academy's SLT-Manual-Contribute group holds
# Contribute, and its SLT-Manual-Read group holds Read, on every live, in-scope
# faculty library of that academy. Runs after the create pass so libraries
# created this run are covered; being a full pass it also backfills libraries
# created before the SLT groups existed and restores an assignment removed by
# hand. Additive only — unlinked libraries keep their SLT access.

# CELL ********************

df_slt_grant_libraries = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.academy_code || '-Sharepoint' AS group_prefix
    FROM
        {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f')}
""")

# Group names are built exactly as create_libraries built them for the SLT
# library, so the grants always target the groups the SLT pass ensures.
pdf_slt_grants = pd.DataFrame(
    [
        {
            "site_url": row.site_url,
            "library_code": row.library_code,
            "group_name": mgr.group_name(
                row.group_prefix, row.site_url, SLT_LIBRARY_CODE, suffix,
            ),
            "permission": permission,
        }
        for row in df_slt_grant_libraries.toPandas().itertuples(index=False)
        for suffix, permission in (("Manual-Contribute", "contribute"), ("Manual-Read", "read"))
    ],
    columns=["site_url", "library_code", "group_name", "permission"],
)
display(pdf_slt_grants)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

if MANAGE_SLT_LIBRARY:
    results = mgr.grant_group_access(pdf_slt_grants)
    persist_action_log(results, "grant_slt_access")
    display(results)
else:
    print("MANAGE_SLT_LIBRARY is False — SLT access to faculty libraries not managed.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Update Memberships
# Full reconcile for every live, in-scope faculty. The library list comes from
# the faculty SCD (current hub URL — site moves are handled by the retire pass
# above — and frozen static code, a single lineage with create/unlink), so a
# faculty whose last member left still appears and has its Auto groups
# emptied, subject to the removal guardrail.
# The hand-maintained all_staff_* columns on the faculty list add nested
# groups: all_staff_read = 1 nests all_staff_security_group (display name or
# object ID) in the faculty's Auto-Read group, all_staff_contribute = 1 in its
# Auto-Contribute group. Auto groups are automation-owned, so setting a flag
# back to 0 (or changing the group) removes the nesting on the next run — no
# last_changed_ts bump needed.

# CELL ********************

df_sync_libraries = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,f.academy_code || '-Sharepoint' AS group_prefix
    FROM {curated_faculty_list} AS f
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND {eligible('f')}
""")

df_memberships = spark.sql(f"""
    SELECT
        f.sharepoint_hub_url AS site_url
        ,f.static_faculty_code AS library_code
        ,m.member_email AS user_principal_name
        ,'contribute' AS permission  -- read, contribute or owner
        ,f.academy_code || '-Sharepoint' AS group_prefix
    FROM {curated_faculty_list} AS f
    INNER JOIN {curated_faculty_membership} AS m
        ON m.faculty_unique_id = f.faculty_unique_id
    WHERE
        COALESCE(f.deleted_in_source, FALSE) = FALSE
        AND m.member_email IS NOT NULL
        AND {eligible('f')}
""")

# One row per (faculty, flag set to 1). A flag set without a group passes
# through as an empty reference, which the manager reports as
# group_not_found rather than silently ignoring.
df_nested_groups = spark.sql(f"""
    WITH live AS (
        SELECT
            f.sharepoint_hub_url AS site_url
            ,f.static_faculty_code AS library_code
            ,COALESCE(TRIM(f.all_staff_security_group), '') AS member_group
            ,f.academy_code || '-Sharepoint' AS group_prefix
            ,COALESCE(f.all_staff_read, 0) AS all_staff_read
            ,COALESCE(f.all_staff_contribute, 0) AS all_staff_contribute
        FROM {curated_faculty_list} AS f
        WHERE
            COALESCE(f.deleted_in_source, FALSE) = FALSE
            AND {eligible('f')}
    )
    SELECT site_url, library_code, member_group, 'read' AS permission, group_prefix
    FROM live
    WHERE all_staff_read = 1

    UNION ALL

    SELECT site_url, library_code, member_group, 'contribute' AS permission, group_prefix
    FROM live
    WHERE all_staff_contribute = 1
""")

pdf_sync_libraries = df_sync_libraries.toPandas()
pdf_memberships = df_memberships.toPandas()
pdf_nested_groups = df_nested_groups.toPandas()
display(pdf_memberships)
display(pdf_nested_groups)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

results = mgr.update_memberships(
    pdf_memberships,
    libraries_df=pdf_sync_libraries,
    nested_groups_df=pdf_nested_groups,
)
persist_action_log(results, "update_memberships")
display(results)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ### Unresolved Users
# Memberships whose user principal name could not be resolved in Entra ID
# (typically a bad or stale address in the source system). Joined back to the
# faculty list so each row carries Academy and Subject, then appended to the
# unresolved-users log table. These are warnings, not failures — fix the
# address in the source system and the next run picks the user up.

# CELL ********************

unresolved = (
    results[results["action"] == "user_not_found"]
    if not results.empty else results
)

if unresolved.empty:
    print("All user principal names resolved.")
else:
    pdf_faculty_lookup = spark.sql(f"""
        SELECT DISTINCT
            f.sharepoint_hub_url AS site_url
            ,f.static_faculty_code AS library_code
            ,f.academy_code AS academy
            ,f.current_faculty_name AS subject
        FROM {curated_faculty_list} AS f
        WHERE
            COALESCE(f.deleted_in_source, FALSE) = FALSE
            AND {eligible('f')}
    """).toPandas()

    pdf_unresolved = (
        unresolved.rename(columns={"document_library": "library_code"})
        .merge(pdf_faculty_lookup, on=["site_url", "library_code"], how="left")
        [["academy", "subject", "user", "status"]]
        .rename(columns={"user": "user_principal_name", "status": "reason"})
        .sort_values(["academy", "subject", "user_principal_name"])
        .reset_index(drop=True)
    )
    # A miss on the lookup join should stay visible, not blank (and NaN in a
    # string column breaks spark.createDataFrame's schema inference).
    pdf_unresolved[["academy", "subject"]] = (
        pdf_unresolved[["academy", "subject"]].fillna("(unknown)")
    )

    print(f"{len(pdf_unresolved)} user principal name(s) could not be resolved:")
    display(pdf_unresolved)

    log_pdf = pdf_unresolved.copy()
    log_pdf["run_id"] = RUN_ID
    log_pdf["dry_run"] = DRY_RUN
    log_pdf["logged_at"] = datetime.now(timezone.utc)
    (spark.createDataFrame(log_pdf)
        .write.format("delta").mode("append").saveAsTable(unresolved_users_table))
    print(f"Persisted to {unresolved_users_table} (run {RUN_ID}).")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Run Outcome
# Surfaces warnings, fails the run (and pipeline alerting) on any failed
# action, then advances the watermark only after a clean non-dry run.

# CELL ********************

all_results = (
    pd.concat([r for _, r in run_results], ignore_index=True)
    if run_results else pd.DataFrame(columns=["action"])
)

warning_rows = all_results[all_results["action"].isin(WARNING_ACTIONS)]
if not warning_rows.empty:
    print(f"{len(warning_rows)} warning(s) this run:")
    display(warning_rows)

fatal_rows = all_results[all_results["action"].isin(FATAL_ACTIONS)]
if not fatal_rows.empty:
    display(fatal_rows)
    raise RuntimeError(
        f"{len(fatal_rows)} failed action(s) this run — "
        f"see {action_log_table} where run_id = '{RUN_ID}'"
    )

print("No failed actions this run.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Advance the watermark to the newest change processed this run. Skipped in
# dry-run (events must remain unconsumed) and unreachable after failures
# (the error gate above raises first), so failed events are retried next run.

if DRY_RUN:
    print("DRY_RUN — watermark not advanced.")
else:
    processed_ts = [
        ts for frame in (pdf_library_changes, pdf_deleted_in_source, pdf_hub_moves)
        if not frame.empty
        for ts in [frame["last_changed_ts"].max()]
        if pd.notna(ts)
    ]
    if not processed_ts:
        print("No watermark events processed — watermark unchanged.")
    else:
        new_watermark = pd.Timestamp(max(processed_ts)).strftime("%Y-%m-%d %H:%M:%S.%f")
        spark.sql(f"""
            MERGE INTO {sync_state_table} AS t
            USING (
                SELECT
                    '{SYNC_PROCESS_LITERAL}' AS process,
                    TIMESTAMP'{new_watermark}' AS last_processed_ts
            ) AS s
            ON t.process = s.process
            WHEN MATCHED THEN UPDATE SET
                t.last_processed_ts = s.last_processed_ts,
                t.updated_at        = current_timestamp()
            WHEN NOT MATCHED THEN INSERT (process, last_processed_ts, updated_at)
            VALUES (s.process, s.last_processed_ts, current_timestamp())
        """)
        print(f"Watermark advanced to {new_watermark}.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }
