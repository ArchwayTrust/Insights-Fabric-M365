"""
SharePointLibraryManager operations, run end to end against the in-memory FakeTenant.

Each test checks the action log the operation returns and what it actually changed, and where it matters that
running it again changes nothing and that a dry run makes no writes.
"""
import time
from decimal import Decimal

import numpy as np
import pandas as pd
import pytest

import alt_fabric_m365
from alt_fabric_m365.auth import Scopes
from alt_fabric_m365.sharepoint import (
    ACTION_LOG_COLUMNS, FATAL_ACTIONS, WARNING_ACTIONS, SharePointLibraryManager,
)

from conftest import fake_token
from sharepoint_fake import APP_ID, HOST, ROLE_DEFINITIONS, FakeTenant

READ = ROLE_DEFINITIONS["Read"]
CONTRIBUTE = ROLE_DEFINITIONS["Contribute"]
FULL_CONTROL = ROLE_DEFINITIONS["Full Control"]
PREFIX = "BAA-Sharepoint"
LABEL = "Document Libraries"


@pytest.fixture
def tenant(monkeypatch):
    tenant = FakeTenant().install(monkeypatch)
    tenant.sleeps = []
    monkeypatch.setattr(time, "sleep", tenant.sleeps.append)
    return tenant


@pytest.fixture
def site(tenant):
    return tenant.add_site("/sites/BAA")


@pytest.fixture
def make_manager(credential):
    def make(**options):
        options.setdefault("include_site_name", False)
        options.setdefault("service_principal_id", APP_ID)
        return SharePointLibraryManager(credential, HOST, **options)
    return make


def libraries_df(site, *codes, **columns):
    """One row per library code, with display name '<code> Faculty' unless library_display_name is given."""
    df = pd.DataFrame({
        "site_url": site.url,
        "library_code": list(codes),
        "library_display_name": [f"{code} Faculty" for code in codes],
        "group_prefix": PREFIX,
    })
    return df.assign(**columns)


def memberships_df(site, rows):
    """rows: (library_code, user_principal_name, permission)."""
    return pd.DataFrame(
        [(site.url, code, upn, permission, PREFIX) for code, upn, permission in rows],
        columns=["site_url", "library_code", "user_principal_name", "permission", "group_prefix"],
    )


def actions(results):
    return sorted(results["action"])


def rows(results, action):
    return results[results["action"] == action]


def principal_id(site, group):
    return site.principals[f"c:0t.c|tenant|{group['id']}"]


def created(tenant, path_fragment, method="POST"):
    return [c for c in tenant.calls if c[0] == method and path_fragment in c[1]]


# Construction

def test_constructor_gets_graph_and_sharepoint_tokens_up_front(credential, tenant):
    SharePointLibraryManager(credential, HOST)

    assert credential.scopes == [(Scopes.GRAPH,), (f"https://{HOST}/.default",)]
    assert tenant.calls == []


def test_works_with_key_vault_certificate_credential(fake_notebookutils, tenant, site, monkeypatch):
    from alt_fabric_m365.auth import keyvault
    from alt_fabric_m365.auth import KeyVaultCertificateCredential
    from conftest import FakeAzure

    azure = FakeAzure(object_id=APP_ID)
    monkeypatch.setattr(keyvault.requests, "get", azure.get)
    monkeypatch.setattr(keyvault.requests, "post", azure.post)
    credential = KeyVaultCertificateCredential("tenant", "client", "https://myvault.vault.azure.net/", "sp-cert")

    manager = SharePointLibraryManager(credential, HOST, include_site_name=False)
    manager.create_libraries(libraries_df(site, "MATH", "ENG"))
    manager.enforce_site_sharing_lockdown([site.url])

    # Graph and SharePoint tokens are requested once each and reused for every call
    assert [kwargs["data"]["scope"] for _, kwargs in azure.token_calls] == [Scopes.GRAPH, f"https://{HOST}/.default"]
    graph_auth = {c[2]["headers"]["Authorization"] for c in tenant.calls if c[1].startswith("https://graph.microsoft.com/")}
    sharepoint_auth = {c[2]["headers"]["Authorization"] for c in tenant.calls if c[1].startswith(f"https://{HOST}/")}
    assert (graph_auth, sharepoint_auth) == ({f"Bearer {azure.issued[0]}"}, {f"Bearer {azure.issued[1]}"})
    assert fake_notebookutils == ["keyvault", "keyvault"]
    assert len(site.lists) == 2
    # The app's object ID comes from the Graph token, and owns the groups it created
    assert manager.service_principal_id == APP_ID
    assert all(group["owners"] == {APP_ID} for group in tenant.groups.values())


def test_a_request_retried_after_a_long_throttle_sends_a_current_token(tenant, site):
    from types import SimpleNamespace
    from sharepoint_fake import response

    class RenewingCredential:
        """Hands out a new token whenever time passes, as a credential does near a token's expiry."""

        def get_token(self, *scopes):
            return SimpleNamespace(token=fake_token(scopes[0], n=len(tenant.sleeps)), expires_on=0)

    tenant.fail("POST", r"/ensureuser", response(429, headers={"Retry-After": "300"}))
    manager = SharePointLibraryManager(RenewingCredential(), HOST, include_site_name=False, service_principal_id=APP_ID)

    results = manager.create_libraries(libraries_df(site, "MATH"))

    throttled_at = tenant.sleeps.index(300.0)
    sent = [call[2]["headers"]["Authorization"] for call in tenant.calls if call[1].endswith("/ensureuser")]
    scope = f"https://{HOST}/.default"
    assert sent[:2] == [f"Bearer {fake_token(scope, n=n)}" for n in (throttled_at, throttled_at + 1)]
    assert rows(results, "error").empty


def test_constructor_fails_fast_on_bad_credentials():
    class Broken:
        def get_token(self, *scopes, **kwargs):
            raise Exception("Failed to get access token")

    with pytest.raises(Exception, match="Failed to get access token"):
        SharePointLibraryManager(Broken(), HOST)


def test_manager_is_exported_from_the_package():
    assert alt_fabric_m365.SharePointLibraryManager is SharePointLibraryManager


def test_action_classes_do_not_overlap():
    assert FATAL_ACTIONS == {"error", "add_failed", "remove_failed"}
    assert not FATAL_ACTIONS & WARNING_ACTIONS


def test_group_name_follows_the_instance_setting(make_manager, site):
    assert make_manager().group_name(PREFIX, site.url, "MATH", "Auto-Read") == "BAA-Sharepoint-MATH-Auto-Read"
    assert make_manager(include_site_name=True).group_name(PREFIX, site.url, "MATH", "Auto-Read") == (
        "BAA-Sharepoint-BAA-MATH-Auto-Read"
    )


def test_missing_columns_raise_before_any_request(make_manager, tenant, site):
    with pytest.raises(ValueError, match="missing required columns"):
        make_manager().create_libraries(pd.DataFrame({"site_url": [site.url]}))

    assert tenant.calls == []


# A pipeline parameter typed as String arrives as a str, and 'False' is truthy
@pytest.mark.parametrize("option, value", [
    ("dry_run", "False"), ("allow_mass_removal", "False"), ("include_site_name", 1), ("add_to_navigation", None),
    ("create_onedrive_shortcuts", "True"), ("max_workers", "4"), ("max_workers", True),
    ("max_removals_per_group", 10.5), ("navigation_label", 5),
])
def test_options_of_the_wrong_type_are_refused(credential, option, value):
    with pytest.raises(TypeError, match=option):
        SharePointLibraryManager(credential, HOST, **{option: value})


def test_per_call_options_of_the_wrong_type_are_refused(make_manager, tenant, site):
    with pytest.raises(TypeError, match="allow_mass_removal"):
        make_manager().update_memberships(memberships_df(site, []), allow_mass_removal="False")

    assert tenant.calls == []


def test_options_changed_after_construction_are_checked(make_manager, tenant, site):
    manager = make_manager()
    manager.allow_mass_removal = "False"

    with pytest.raises(TypeError, match="allow_mass_removal"):
        manager.update_memberships(memberships_df(site, []))

    assert tenant.calls == []


def test_numpy_booleans_and_integers_are_accepted(credential):
    manager = SharePointLibraryManager(credential, HOST, dry_run=np.bool_(True), max_workers=np.int64(2))

    assert manager.dry_run is True
    assert manager.max_workers == 2 and type(manager.max_workers) is int


# create_libraries

def test_create_library_with_groups_metadata_and_navigation(make_manager, tenant, site):
    df = libraries_df(site, "MATH", library_display_name="Maths", manual_read=True, manual_contribute=True,
                      source_system="Arbor", source_id="F1")

    results = make_manager(add_to_navigation=True).create_libraries(df)

    assert list(results.columns) == ACTION_LOG_COLUMNS
    assert actions(results) == sorted([
        "navigation_label_created", "library_created", "metadata_set", "metadata_set", "group_assigned",
        "group_assigned", "manual_read_group_ensured", "manual_contribute_group_ensured", "navigation_link_added",
    ])
    assert rows(results, "library_created")["status"].tolist() == [f"created at {site.url}/MATH"]

    lib = site.list_by_slug("MATH")
    assert (lib["displayName"], lib["unique"]) == ("Maths", True)
    assert lib["properties"] == {"SourceSystem": "Arbor", "SourceId": "F1"}

    expected_roles = {
        "BAA-Sharepoint-MATH-Auto-Read": READ,
        "BAA-Sharepoint-MATH-Auto-Contribute": CONTRIBUTE,
        "BAA-Sharepoint-MATH-Manual-Read": READ,
        "BAA-Sharepoint-MATH-Manual-Contribute": CONTRIBUTE,
    }
    for name, role in expected_roles.items():
        group = tenant.group_named(name)
        assert lib["role_assignments"][principal_id(site, group)] == {role}
    assert "NOT synced" in tenant.group_named("BAA-Sharepoint-MATH-Manual-Read")["description"]

    [label] = site.nav
    assert (label["Title"], label["Url"]) == (LABEL, "http://linkless.header/")
    assert [(c["Title"], c["Url"]) for c in label["Children"]] == [("Maths", "/sites/BAA/MATH")]

    # One wait for replication covers all four new groups
    assert tenant.sleeps == [5.0]


def test_create_library_again_changes_nothing(make_manager, tenant, site):
    manager = make_manager(add_to_navigation=True)
    df = libraries_df(site, "MATH", manual_read=True, source_system="Arbor", source_id="F1")
    manager.create_libraries(df)
    tenant.calls.clear()

    results = manager.create_libraries(df)

    assert sorted(rows(results, "metadata_unchanged")["status"]) == ["SourceId=F1", "SourceSystem=Arbor"]
    assert "metadata_set" not in actions(results)
    assert created(tenant, "ProcessQuery") == []

    assert "navigation_label_created" not in actions(results)
    assert "navigation_link_added" not in actions(results)
    assert not {"library_created", "library_renamed"} & set(actions(results))
    assert actions(results).count("already_assigned") == 3
    assert created(tenant, "/web/ensureuser") == [] and created(tenant, "addroleassignment") == []
    assert created(tenant, "/groups") == []
    assert created(tenant, f"/sites/{site.id}/lists") == []
    assert created(tenant, "breakroleinheritance") == []
    assert len(site.lists) == 1 and len(site.nav[0]["Children"]) == 1
    # And nothing else: the checks above name the writes a re-run could make, this catches any other
    assert tenant.writes() == []


@pytest.mark.parametrize("operation", ["create_libraries", "set_library_metadata"])
@pytest.mark.parametrize("dry_run", [False, True])
def test_only_metadata_that_changed_is_written(make_manager, tenant, site, operation, dry_run):
    make_manager().create_libraries(libraries_df(site, "MATH", source_system="Arbor", source_id="F1"))
    tenant.calls.clear()

    results = getattr(make_manager(dry_run=dry_run), operation)(
        libraries_df(site, "MATH", source_system="Arbor", source_id="F2"),
    )

    set_action = "metadata_set (dry run)" if dry_run else "metadata_set"
    assert rows(results, "metadata_unchanged")["status"].tolist() == ["SourceSystem=Arbor"]
    assert rows(results, set_action)["status"].tolist() == ["SourceId=F2"]
    assert site.list_by_slug("MATH")["properties"] == {"SourceSystem": "Arbor", "SourceId": "F1" if dry_run else "F2"}
    assert len(created(tenant, "ProcessQuery")) == (0 if dry_run else 1)


def test_every_request_says_which_app_it_comes_from(make_manager, tenant, site):
    tenant.add_user("jo@contoso.com")
    manager = make_manager(add_to_navigation=True)

    manager.create_libraries(libraries_df(site, "MATH", source_system="Arbor"))
    manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "read")]))

    # SharePoint serves traffic that says who it's from ahead of traffic that doesn't when it throttles
    assert {kwargs["headers"]["User-Agent"].split(" ")[0] for _, _, kwargs in tenant.calls} == {
        f"NONISV|ArchwayTrust|alt-fabric-m365/{alt_fabric_m365.__version__}"
    }
    assert created(tenant, "ProcessQuery")  # the CSOM call that sets metadata is covered too


# CSOM refuses with an HTTP error, or with an ErrorInfo in a 200 response
@pytest.mark.parametrize("failure, reason", [
    ((403, {"error": {"code": "-2147024891", "message": {"value": "Denied."}}}), "Status code: 403"),
    ((200, [{"SchemaVersion": "15.0.0.0", "ErrorInfo": {"ErrorMessage": "Access denied."}}]), "CSOM error"),
])
def test_a_failed_metadata_write_still_leaves_the_library_its_groups_and_name(make_manager, tenant, site, failure, reason):
    from sharepoint_fake import response
    tenant.fail("POST", r"/ProcessQuery", response(*failure))

    results = make_manager(add_to_navigation=True).create_libraries(
        libraries_df(site, "MATH", library_display_name="Maths", source_system="Arbor", source_id="F1"),
    )

    [error] = rows(results, "error")["status"]
    assert error.startswith(f"metadata not set: {reason}")
    assert actions(results).count("group_assigned") == 2
    assert {"library_created", "navigation_link_added"} <= set(actions(results))
    library = site.list_by_slug("MATH")
    assert (library["displayName"], library["unique"], len(library["role_assignments"])) == ("Maths", True, 2)
    assert library["properties"] == {}


def test_create_library_renames_display_name_and_navigation_link(make_manager, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH", library_display_name="Maths"))

    results = manager.create_libraries(libraries_df(site, "MATH", library_display_name="Mathematics"))

    assert site.list_by_slug("MATH")["displayName"] == "Mathematics"
    assert site.nav[0]["Children"][0]["Title"] == "Mathematics"
    assert "navigation_link_renamed" in actions(results)
    assert rows(results, "library_renamed")["status"].tolist() == ["display name set to 'Mathematics'"]


def test_create_dry_run_reports_a_new_display_name(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH", library_display_name="Maths"))
    tenant.calls.clear()

    results = make_manager(dry_run=True).create_libraries(
        libraries_df(site, "MATH", library_display_name="Mathematics"),
    )

    assert tenant.writes() == []
    assert "rename_library (dry run)" in actions(results) and "library_created (dry run)" not in actions(results)


@pytest.mark.parametrize("dry_run", [False, True])
def test_a_display_name_another_list_has_is_a_warning(make_manager, tenant, site, dry_run):
    site.add_list("English", slug="EnglishOld")  # made by hand, titled with the new library's display name

    results = make_manager(dry_run=dry_run).create_libraries(libraries_df(site, "ENG", library_display_name="English"))

    assert list(rows(results, "library_name_conflict")["status"].str.contains("'English'")) == [True]
    assert rows(results, "error").empty
    if not dry_run:
        # The library still has its own permissions and groups, under its code as its title
        library = site.list_by_slug("ENG")
        assert (library["displayName"], library["unique"], len(library["role_assignments"])) == ("ENG", True, 2)


def test_a_failed_rename_leaves_the_library_with_its_own_permissions(make_manager, site):
    # Both want the same title, so the second rename fails: the first took it during this run
    results = make_manager().create_libraries(libraries_df(site, "SCI1", "SCI2", library_display_name="Science"))

    second = site.list_by_slug("SCI2")
    assert (second["displayName"], second["unique"], len(second["role_assignments"])) == ("SCI2", True, 2)
    assert list(rows(results, "error")["document_library"]) == ["SCI2"]


def test_a_title_another_library_gives_up_in_the_same_run_is_used(make_manager, site):
    make_manager().create_libraries(libraries_df(site, "ENG", "OTH", library_display_name=["English", "Other"]))

    # ENG gives up "English" before OTH takes it, after the site's lists were read
    results = make_manager().create_libraries(
        libraries_df(site, "ENG", "OTH", library_display_name=["Maths", "English"]),
    )

    assert (site.list_by_slug("ENG")["displayName"], site.list_by_slug("OTH")["displayName"]) == ("Maths", "English")
    assert "library_name_conflict" not in actions(results)


def test_renaming_onto_another_lists_title_is_a_warning(make_manager, site):
    make_manager().create_libraries(libraries_df(site, "ENG"))
    site.add_list("English", slug="EnglishOld")

    results = make_manager().create_libraries(libraries_df(site, "ENG", library_display_name="English"))

    assert site.list_by_slug("ENG")["displayName"] == "ENG Faculty"
    assert "library_name_conflict" in actions(results) and rows(results, "error").empty


def test_taking_over_a_library_keeps_permissions_set_inside_it(make_manager, tenant, site):
    library = site.add_list("MATH")  # made by hand, still inheriting the site's permissions
    library["unique_items"]["Exam papers"] = {7: {FULL_CONTROL}}

    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert library["unique"]
    assert library["unique_items"] == {"Exam papers": {7: {FULL_CONTROL}}}
    assert actions(results).count("inheritance_broken") == 1

    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert "inheritance_broken" not in actions(results)


def test_new_library_logs_no_inheritance_break(make_manager, site):
    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert site.list_by_slug("MATH")["unique"]
    assert "inheritance_broken" not in actions(results)


def test_create_dry_run_reports_taking_over_an_inheriting_library(make_manager, tenant, site):
    site.add_list("MATH")
    site.add_list("ENG", unique=True)

    results = make_manager(dry_run=True).create_libraries(libraries_df(site, "MATH", "ENG"))

    assert tenant.writes() == []
    assert list(rows(results, "inheritance_broken (dry run)")["document_library"]) == ["MATH"]


def test_create_dry_run_makes_no_writes(make_manager, tenant, site):
    df = libraries_df(site, "MATH", manual_read=True, manual_contribute=True, source_system="Arbor", source_id="F1")

    results = make_manager(dry_run=True, add_to_navigation=True).create_libraries(df)

    assert tenant.writes() == []
    assert site.lists == {}
    assert actions(results) == sorted([
        "navigation_label_created (dry run)", "library_created (dry run)",
        "metadata_set (dry run)", "metadata_set (dry run)", "ensure_group (dry run)", "ensure_group (dry run)",
        "ensure_manual_read_group (dry run)", "ensure_manual_contribute_group (dry run)",
        "ensure_navigation_link (dry run)",
    ])


def test_per_call_dry_run_overrides_the_instance_default(make_manager, tenant, site):
    make_manager(dry_run=False).create_libraries(libraries_df(site, "MATH"), dry_run=True)

    assert tenant.writes() == []


def test_create_manual_only_library(make_manager, tenant, site):
    df = libraries_df(site, "SLT", library_display_name="SLT", auto_groups=False, manual_read=True,
                      manual_contribute=True)

    make_manager().create_libraries(df)

    assert sorted(g["displayName"] for g in tenant.groups.values()) == [
        "BAA-Sharepoint-SLT-Manual-Contribute", "BAA-Sharepoint-SLT-Manual-Read",
    ]


@pytest.mark.parametrize("blank", [None, float("nan"), pd.NA])
def test_blank_flags_take_their_defaults(make_manager, tenant, site, blank):
    flags = {flag: pd.Series([blank], dtype=object) for flag in
             ("auto_groups", "manual_read", "manual_contribute", "app_owned")}

    results = make_manager().create_libraries(libraries_df(site, "MATH", **flags))

    assert rows(results, "error").empty
    assert sorted(g["displayName"] for g in tenant.groups.values()) == [
        "BAA-Sharepoint-MATH-Auto-Contribute", "BAA-Sharepoint-MATH-Auto-Read",
    ]
    assert all(g["owners"] == {APP_ID} for g in tenant.groups.values())


def test_flags_given_as_numbers_are_accepted(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "SLT", auto_groups=0, manual_read=1.0))

    assert [g["displayName"] for g in tenant.groups.values()] == ["BAA-Sharepoint-SLT-Manual-Read"]


@pytest.mark.parametrize("value", ["False", "yes", 2])
def test_flags_that_arent_booleans_are_refused(make_manager, tenant, site, value):
    with pytest.raises(ValueError, match="manual_read"):
        make_manager().create_libraries(libraries_df(site, "MATH", manual_read=value))

    assert tenant.calls == []


def test_colliding_codes_are_skipped_and_others_proceed(make_manager, tenant, site):
    results = make_manager().create_libraries(libraries_df(site, "PE-KS3", "PEKS3", "ENG"))

    assert len(rows(results, "group_name_collision")) == 2
    assert [lst["slug"] for lst in site.lists.values()] == ["ENG"]


def test_duplicate_rows_create_one_library(make_manager, site):
    df = pd.concat([libraries_df(site, "MATH"), libraries_df(site, "MATH", library_display_name="Other")])

    make_manager().create_libraries(df)

    [lib] = site.lists.values()
    assert lib["displayName"] == "MATH Faculty"


def test_a_failing_library_is_logged_and_the_rest_proceed(make_manager, tenant, site):
    from sharepoint_fake import graph_error
    tenant.fail("POST", r"/lists$", graph_error(403, "accessDenied"))

    results = make_manager(max_workers=1).create_libraries(libraries_df(site, "ART", "ENG"))

    [error] = rows(results, "error").itertuples()
    assert error.document_library == "ART" and "403" in error.status
    assert [lst["slug"] for lst in site.lists.values()] == ["ENG"]


def test_create_in_parallel(make_manager, site):
    codes = [f"F{i}" for i in range(8)]

    results = make_manager(max_workers=4).create_libraries(libraries_df(site, *codes))

    assert sorted(lst["slug"] for lst in site.lists.values()) == codes
    assert len(rows(results, "group_assigned")) == 16
    assert rows(results, "error").empty


# pandas makes an integer column with a null a float column, and Spark gives a DecimalType column as Decimals
@pytest.mark.parametrize("code", [101.0, np.float32(101), Decimal("101.00")])
def test_a_whole_number_code_from_a_float_or_decimal_column_is_used_without_decimals(make_manager, tenant, site, code):
    df = libraries_df(site, "101").assign(library_code=[code])

    results = make_manager().create_libraries(df)

    assert set(results["document_library"]) == {"101"}
    assert site.list_by_slug("101") is not None
    assert tenant.group_named("BAA-Sharepoint-101-Auto-Read") is not None


@pytest.mark.parametrize("operation", ["create_libraries", "set_library_metadata"])
def test_a_whole_number_source_id_from_a_column_with_nulls_is_stored_without_decimals(
    make_manager, site, library, operation,
):
    df = libraries_df(site, "MATH", "ENG").assign(source_id=[42, None])
    assert df["source_id"].dtype == float

    results = getattr(make_manager(), operation)(df)

    assert site.list_by_slug("MATH")["properties"] == {"SourceId": "42"}
    assert list(rows(results, "metadata_set")["status"]) == ["SourceId=42"]


# Metadata

def test_set_and_get_library_metadata(make_manager, site):
    site.add_list("MATH")
    manager = make_manager()

    results = manager.set_library_metadata(pd.DataFrame({
        "site_url": [site.url, site.url], "library_code": ["MATH", "NONE"], "source_id": ["F1", None],
    }))

    assert actions(results) == ["metadata_set", "skipped"]
    metadata = manager.get_library_metadata(pd.DataFrame({"site_url": [site.url] * 2, "library_code": ["MATH", "ENG"]}))
    assert metadata.to_dict("records") == [
        {"site_url": site.url, "library_code": "MATH", "source_system": None, "source_id": "F1", "status": "ok"},
        {"site_url": site.url, "library_code": "ENG", "source_system": None, "source_id": None,
         "status": "library_not_found"},
    ]


def test_get_metadata_of_a_library_with_none_set_is_ok(make_manager, site):
    site.add_list("MATH")

    metadata = make_manager().get_library_metadata(libraries_df(site, "MATH"))

    assert metadata[["source_system", "source_id", "status"]].values.tolist() == [[None, None, "ok"]]


def _fail_site_read(tenant, monkeypatch):
    from sharepoint_fake import graph_error
    tenant.fail("GET", r":/sites/BAA$", graph_error(403, "accessDenied", "Access denied"))


def _fail_properties_read(tenant, monkeypatch):
    from sharepoint_fake import response
    tenant.fail("GET", r"/rootfolder/properties$", response(403, {"error": {
        "code": "-2147024891, System.UnauthorizedAccessException", "message": {"value": "Access denied."},
    }}))


def _drop_properties_connection(tenant, monkeypatch):
    import requests
    from alt_fabric_m365.sharepoint import _http

    def request(method, url, **kwargs):
        if "/sites/BAA/" in url and url.endswith("/rootfolder/properties"):
            raise requests.ConnectionError("Connection reset by peer")
        return tenant.request(method, url, **kwargs)
    monkeypatch.setattr(_http.requests, "request", request)


# Unset values are None too, so a row that couldn't be read must say so rather than look unset
@pytest.mark.parametrize("fail, reason", [
    (_fail_site_read, "error: site could not be read: Status code: 403, accessDenied: Access denied"),
    (_fail_properties_read, "error: Status code: 403"),
    (_drop_properties_connection, "error: Connection reset by peer"),
])
def test_get_metadata_reports_a_library_it_could_not_read(make_manager, tenant, site, monkeypatch, fail, reason):
    site.add_list("MATH")["properties"] = {"SourceSystem": "MIS", "SourceId": "F1"}
    other = tenant.add_site("/sites/BAA2")
    other.add_list("ENG")["properties"] = {"SourceSystem": "MIS", "SourceId": "F2"}
    fail(tenant, monkeypatch)

    metadata = make_manager().get_library_metadata(
        pd.concat([libraries_df(site, "MATH"), libraries_df(other, "ENG")]),
    )

    [math, eng] = metadata.to_dict("records")
    assert (math["source_system"], math["source_id"]) == (None, None)
    assert math["status"].startswith(reason)
    assert (eng["source_id"], eng["status"]) == ("F2", "ok")


def test_set_metadata_needs_a_metadata_column(make_manager, site):
    with pytest.raises(ValueError, match="source_system' or 'source_id"):
        make_manager().set_library_metadata(pd.DataFrame({"site_url": [site.url], "library_code": ["MATH"]}))


@pytest.mark.parametrize("dry_run", [False, True])
def test_set_metadata_on_a_missing_library_is_an_error_row(make_manager, tenant, site, dry_run):
    results = make_manager(dry_run=dry_run).set_library_metadata(pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "source_system": ["Arbor"],
    }))

    assert actions(results) == ["error"]
    assert tenant.writes() == []


def test_set_metadata_dry_run_makes_no_writes(make_manager, tenant, site):
    library = site.add_list("MATH")

    results = make_manager(dry_run=True).set_library_metadata(pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "source_system": ["Arbor"], "source_id": ["F1"],
    }))

    assert actions(results) == ["metadata_set (dry run)", "metadata_set (dry run)"]
    assert tenant.writes() == [] and library["properties"] == {}


# rename_libraries

def test_rename_sets_the_display_name_and_leaves_groups(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH", manual_read=True))
    groups = sorted(g["displayName"] for g in tenant.groups.values())

    results = make_manager().rename_libraries(
        libraries_df(site, "MATH", library_display_name="Maths").drop(columns="group_prefix"),
    )

    assert actions(results) == ["library_renamed"]
    assert site.list_by_slug("MATH")["displayName"] == "Maths"
    assert sorted(g["displayName"] for g in tenant.groups.values()) == groups


@pytest.mark.parametrize("dry_run", [False, True])
def test_rename_of_a_missing_library_is_an_error(make_manager, tenant, site, dry_run):
    results = make_manager(dry_run=dry_run).rename_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["error"]
    assert tenant.writes() == []


def test_old_library_code_is_refused(make_manager, tenant, site):
    # A code is the library's URL, which can't change, so renaming groups to a new code never worked
    with pytest.raises(ValueError, match="old_library_code"):
        make_manager().rename_libraries(libraries_df(site, "MATH", old_library_code="MATHS"))

    assert tenant.calls == []


def test_an_empty_old_library_code_column_is_ignored(make_manager, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))

    results = make_manager().rename_libraries(
        libraries_df(site, "MATH", old_library_code=None, library_display_name="Maths"),
    )

    assert actions(results) == ["library_renamed"]


def test_rename_syncs_the_navigation_title(make_manager, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH"))

    results = manager.rename_libraries(libraries_df(site, "MATH", library_display_name="Mathematics"))

    assert actions(results) == ["library_renamed", "navigation_link_renamed"]
    assert site.nav[0]["Children"][0]["Title"] == "Mathematics"


@pytest.mark.parametrize("dry_run", [False, True])
def test_rename_to_the_name_a_library_already_has_writes_nothing(make_manager, tenant, site, dry_run):
    make_manager().create_libraries(libraries_df(site, "MATH", library_display_name="Maths"))
    tenant.calls.clear()

    results = make_manager(dry_run=dry_run).rename_libraries(
        libraries_df(site, "MATH", library_display_name="Maths"),
    )

    assert results[["action", "status"]].values.tolist() == [["already_named", "display name is already 'Maths'"]]
    assert tenant.writes() == []


def test_rename_dry_run_makes_no_writes(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))
    tenant.calls.clear()

    results = make_manager(dry_run=True).rename_libraries(libraries_df(site, "MATH", library_display_name="X"))

    assert tenant.writes() == []
    assert actions(results) == ["rename_library (dry run)"]


# unlink_libraries

def test_unlink_empties_auto_groups_keeps_manual_and_removes_the_link(make_manager, tenant, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH", manual_read=True))
    jo = tenant.add_user("jo@contoso.com")
    sam = tenant.add_user("sam@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][jo["id"]] = jo
    tenant.group_named("BAA-Sharepoint-MATH-Manual-Read")["members"][sam["id"]] = sam

    results = manager.unlink_libraries(libraries_df(site, "MATH").drop(columns="library_display_name"))

    assert actions(results) == [
        "group_emptied", "group_emptied", "library_unlinked", "manual_groups_preserved", "navigation_link_removed",
    ]
    assert site.list_by_slug("MATH")["displayName"] == "MATH_UNLINKED"
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == []
    assert tenant.member_upns("BAA-Sharepoint-MATH-Manual-Read") == ["sam@contoso.com"]
    assert site.nav[0]["Title"] == LABEL and site.nav[0]["Children"] == []


def test_unlink_empties_the_rest_of_a_group_when_one_member_cant_be_removed(make_manager, tenant, site):
    from sharepoint_fake import graph_error
    make_manager().create_libraries(libraries_df(site, "MATH"))
    group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    jo, sam, ali = (tenant.add_user(f"{name}@contoso.com") for name in ("jo", "sam", "ali"))
    group["members"] = {user["id"]: user for user in (jo, sam, ali)}
    tenant.fail("DELETE", rf"/members/{sam['id']}/\$ref$", graph_error(403, "Authorization_RequestDenied", "Denied."))

    results = make_manager().unlink_libraries(libraries_df(site, "MATH"))

    emptied = rows(results, "group_emptied").set_index("group")["status"]
    assert emptied["BAA-Sharepoint-MATH-Auto-Contribute"] == "removed 2 member(s); 1 couldn't be removed"
    [failed] = rows(results, "remove_failed").itertuples()
    assert (failed.group, failed.user) == ("BAA-Sharepoint-MATH-Auto-Contribute", "sam@contoso.com")
    assert "remove_failed" in FATAL_ACTIONS
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["sam@contoso.com"]


def test_unlinking_a_library_deleted_by_hand_removes_its_navigation_link(make_manager, tenant, site):
    make_manager(add_to_navigation=True).create_libraries(libraries_df(site, "MATH"))
    del site.lists[site.list_by_slug("MATH")["id"]]

    results = make_manager(add_to_navigation=True).unlink_libraries(libraries_df(site, "MATH"))

    assert {"library_not_found", "navigation_link_removed"} <= set(actions(results))
    assert site.nav[0]["Children"] == []
    # The removal is reported by its own row, once it's done
    assert "navigation" not in rows(results, "library_not_found")["status"].iloc[0]


def test_unlink_dry_run_makes_no_writes(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))
    tenant.calls.clear()

    results = make_manager(dry_run=True, add_to_navigation=True).unlink_libraries(libraries_df(site, "MATH"))

    assert tenant.writes() == []
    assert "empty_group (dry run)" in actions(results)


@pytest.mark.parametrize("dry_run", [False, True])
def test_unlinking_a_code_that_shares_a_live_librarys_groups_leaves_them(make_manager, tenant, site, dry_run):
    # PEKS3 was skipped by the create pass, as its group names are PE-KS3's
    make_manager().create_libraries(libraries_df(site, "PE-KS3"))
    jo = tenant.add_user("jo@contoso.com")
    tenant.group_named("BAA-Sharepoint-PEKS3-Auto-Contribute")["members"][jo["id"]] = jo
    tenant.calls.clear()

    results = make_manager(dry_run=dry_run, add_to_navigation=True).unlink_libraries(libraries_df(site, "PEKS3"))

    [row] = results.itertuples()
    assert (row.action, row.document_library) == ("group_name_collision", "PEKS3")
    assert tenant.writes() == []
    assert tenant.member_upns("BAA-Sharepoint-PEKS3-Auto-Contribute") == ["jo@contoso.com"]


def test_unlinking_codes_whose_groups_collide_leaves_both(make_manager, tenant, site):
    results = make_manager().unlink_libraries(libraries_df(site, "PE-KS3", "PEKS3"))

    assert actions(results) == ["group_name_collision", "group_name_collision"]
    assert tenant.writes() == []


def test_unlink_skips_a_site_whose_prefetch_fails(make_manager, tenant, site):
    from sharepoint_fake import graph_error
    tenant.fail("GET", r"/lists\?", graph_error(403, "accessDenied"))

    results = make_manager().unlink_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["error"]
    assert tenant.writes() == []


# retire_moved_libraries

def test_retire_renames_and_removes_group_access_but_keeps_membership(make_manager, tenant, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH", manual_read=True))
    jo = tenant.add_user("jo@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][jo["id"]] = jo

    results = manager.retire_moved_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["group_access_removed"] * 3 + ["library_retired", "navigation_link_removed"]
    lib = site.list_by_slug("MATH")
    assert lib["displayName"] == "MATH_MOVED"
    assert lib["role_assignments"] == {}
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]


@pytest.mark.parametrize("operation, done", [
    ("unlink_libraries", "already_unlinked"),
    ("retire_moved_libraries", "already_retired"),
])
@pytest.mark.parametrize("dry_run", [False, True])
def test_unlinking_or_retiring_again_writes_nothing(make_manager, tenant, site, operation, done, dry_run):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH", manual_read=True))
    getattr(manager, operation)(libraries_df(site, "MATH"))
    tenant.calls.clear()

    results = getattr(make_manager(add_to_navigation=True, dry_run=dry_run), operation)(libraries_df(site, "MATH"))

    assert done in actions(results)
    assert not {
        "library_unlinked", "library_retired", "group_access_removed",
        "unlink_library (dry run)", "retire_library (dry run)", "remove_group_access (dry run)",
    } & set(actions(results))
    assert tenant.writes() == []
    # Retire no longer registers each group on the old site to find it has no access to remove
    assert created(tenant, "/ensureuser") == []


def test_retire_skips_a_site_or_library_that_is_gone(make_manager, tenant, site):
    gone_site = "https://contoso.sharepoint.com/sites/Deleted"
    df = pd.DataFrame({
        "site_url": [gone_site, site.url], "library_code": ["MATH", "ENG"], "group_prefix": PREFIX,
    })

    results = make_manager().retire_moved_libraries(df)

    assert results[["site_url", "action"]].values.tolist() == [
        [gone_site, "retire_skipped"], [site.url, "retire_skipped"],
    ]
    assert tenant.writes() == []


# Under Sites.Selected, a site whose grant has been removed gives 403; only a deleted site gives 404
@pytest.mark.parametrize("dry_run", [False, True])
def test_retire_reports_an_old_site_it_cant_read_as_an_error(make_manager, tenant, site, dry_run):
    from sharepoint_fake import graph_error
    make_manager().create_libraries(libraries_df(site, "MATH"))
    tenant.calls.clear()
    tenant.fail("GET", r":/sites/BAA$", graph_error(403, "accessDenied", "Access denied"))

    results = make_manager(dry_run=dry_run).retire_moved_libraries(libraries_df(site, "MATH"))

    [row] = results.itertuples()
    assert (row.site_url, row.action) == (site.url, "error")
    assert "Status code: 403, accessDenied" in row.status
    assert tenant.writes() == []


@pytest.mark.parametrize("failure, action", [
    ("403", "error"),            # exists, but the app can't read it: the reason is reported
    ("404", "retire_deferred"),  # doesn't exist yet: retired once the library works there
])
def test_retire_keeps_a_library_whose_new_site_it_cant_read(make_manager, tenant, site, failure, action):
    from sharepoint_fake import graph_error
    make_manager().create_libraries(libraries_df(site, "MATH"))
    new_site = tenant.add_site("/sites/BAA2")
    tenant.calls.clear()
    if failure == "403":
        new_site_url = new_site.url
        tenant.fail("GET", r":/sites/BAA2$", graph_error(403, "accessDenied", "Access denied"))
    else:
        new_site_url = "https://contoso.sharepoint.com/sites/NotMadeYet"

    results = make_manager().retire_moved_libraries(libraries_df(site, "MATH").assign(new_site_url=new_site_url))

    [row] = results.itertuples()
    assert (row.site_url, row.document_library, row.action) == (site.url, "MATH", action)
    if action == "error":
        assert row.status.startswith(f"its new site ({new_site_url}) couldn't be read, so it wasn't retired: ")
        assert "accessDenied" in row.status
    assert tenant.writes() == []
    assert site.list_by_slug("MATH")["displayName"] == "MATH Faculty"


@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("spelling", ["lower case", "trailing slash"])
def test_retire_leaves_a_library_whose_new_site_is_the_same_site(make_manager, tenant, site, spelling, dry_run):
    make_manager(add_to_navigation=True).create_libraries(libraries_df(site, "MATH"))
    tenant.calls.clear()
    # Graph resolves both to the same site, so the library hasn't moved
    same_site = site.url.lower() if spelling == "lower case" else f"{site.url}/"

    results = make_manager(add_to_navigation=True, dry_run=dry_run).retire_moved_libraries(
        libraries_df(site, "MATH").assign(new_site_url=same_site))

    assert actions(results) == ["retire_skipped"]
    assert tenant.writes() == []
    library = site.list_by_slug("MATH")
    assert (library["displayName"], len(library["role_assignments"])) == ("MATH Faculty", 2)


# Bringing a library back: restored after an unlink, moved back to its first site, or deleted by hand

def test_a_restored_library_gets_its_name_link_access_and_members_back(make_manager, tenant, site):
    tenant.add_user("jo@contoso.com")
    manager = make_manager(add_to_navigation=True)
    df = libraries_df(site, "MATH", manual_read=True)
    members = memberships_df(site, [("MATH", "jo@contoso.com", "contribute")])
    manager.create_libraries(df)
    manager.update_memberships(members)
    access = dict(site.list_by_slug("MATH")["role_assignments"])
    manager.unlink_libraries(df)

    manager.create_libraries(df)
    manager.update_memberships(members)

    library = site.list_by_slug("MATH")
    assert (library["displayName"], library["role_assignments"]) == ("MATH Faculty", access)
    assert [(c["Title"], c["Url"]) for c in site.nav[0]["Children"]] == [("MATH Faculty", "/sites/BAA/MATH")]
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]
    tenant.calls.clear()
    manager.create_libraries(df)
    manager.update_memberships(members)
    assert tenant.writes() == []


def test_a_library_moved_away_and_back_works_on_its_first_site_again(make_manager, tenant, site):
    other = tenant.add_site("/sites/BAA2")
    manager = make_manager(add_to_navigation=True)
    on_first, on_other = libraries_df(site, "MATH"), libraries_df(other, "MATH")
    manager.create_libraries(on_first)
    access = dict(site.list_by_slug("MATH")["role_assignments"])
    manager.create_libraries(on_other)
    manager.retire_moved_libraries(on_first.assign(new_site_url=other.url))
    assert (site.list_by_slug("MATH")["displayName"], site.list_by_slug("MATH")["role_assignments"]) == (
        "MATH_MOVED", {})

    # Retire removed the groups' access to the first site's copy, so moving back gives it back
    manager.create_libraries(on_first)
    results = manager.retire_moved_libraries(on_other.assign(new_site_url=site.url))

    assert "library_retired" in actions(results)
    first, second = site.list_by_slug("MATH"), other.list_by_slug("MATH")
    assert (first["displayName"], first["role_assignments"]) == ("MATH Faculty", access)
    assert (second["displayName"], second["role_assignments"]) == ("MATH_MOVED", {})
    assert [c["Title"] for c in site.nav[0]["Children"]] == ["MATH Faculty"]
    assert other.nav[0]["Children"] == []


def test_a_library_deleted_by_hand_is_created_again_with_one_link(make_manager, tenant, site):
    manager = make_manager(add_to_navigation=True)
    df = libraries_df(site, "MATH")
    manager.create_libraries(df)
    del site.lists[site.list_by_slug("MATH")["id"]]

    results = manager.create_libraries(df)

    assert "library_created" in actions(results)
    library = site.list_by_slug("MATH")
    assert (library["displayName"], library["unique"], len(library["role_assignments"])) == ("MATH Faculty", True, 2)
    # Its link was left behind, so the new library's link is the same one rather than a second
    assert [(c["Title"], c["Url"]) for c in site.nav[0]["Children"]] == [("MATH Faculty", "/sites/BAA/MATH")]


@pytest.fixture
def library(make_manager, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))
    return site.list_by_slug("MATH")


def test_memberships_add_users_including_aliases(make_manager, tenant, site, library):
    jo = tenant.add_user("jo@contoso.com")
    tenant.add_user("sam.jones@contoso.onmicrosoft.com", mail="sam@contoso.com")

    results = make_manager().update_memberships(memberships_df(site, [
        ("MATH", "Jo@Contoso.com ", "contribute"),
        ("MATH", "sam@contoso.com", "contribute"),
    ]))

    # create_libraries already gave the group Contribute, so nothing is re-written
    assert actions(results) == ["added_to_group", "added_to_group", "already_assigned"]
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == [
        "jo@contoso.com", "sam.jones@contoso.onmicrosoft.com",
    ]
    assert jo["id"] in tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"]


def test_memberships_again_change_nothing(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    df = memberships_df(site, [("MATH", "jo@contoso.com", "contribute")])
    make_manager().update_memberships(df)
    tenant.calls.clear()

    results = make_manager().update_memberships(df)

    assert actions(results) == ["already_assigned", "already_member"]
    assert tenant.writes() == []


def test_memberships_remove_stale_members_but_keep_service_principals(make_manager, tenant, site, library):
    jo = tenant.add_user("jo@contoso.com")
    stale = tenant.add_user("left@contoso.com")
    service_principal = tenant.add_service_principal()
    group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    group["members"] = {stale["id"]: stale, service_principal["id"]: service_principal}

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    assert actions(results) == ["added_to_group", "already_assigned", "non_user_member_preserved", "removed_from_group"]
    assert set(group["members"]) == {jo["id"], service_principal["id"]}


def test_roles_with_no_desired_members_are_emptied_but_not_created(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    reader = tenant.add_user("reader@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"][reader["id"]] = reader

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Read") == []
    assert tenant.group_named("BAA-Sharepoint-MATH-Auto-Owner") is None
    assert rows(results, "removed_from_group")["group"].tolist() == ["BAA-Sharepoint-MATH-Auto-Read"]


def test_owner_role_creates_and_assigns_its_group(make_manager, tenant, site, library):
    tenant.add_user("head@contoso.com")
    tenant.sleeps.clear()

    make_manager().update_memberships(memberships_df(site, [("MATH", "head@contoso.com", "owner")]))

    owner = tenant.group_named("BAA-Sharepoint-MATH-Auto-Owner")
    assert library["role_assignments"][principal_id(site, owner)] == {FULL_CONTROL}
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Owner") == ["head@contoso.com"]
    assert tenant.sleeps == [5.0]


def test_unresolved_users_block_that_roles_removals(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    stale = tenant.add_user("maybe@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][stale["id"]] = stale

    results = make_manager().update_memberships(memberships_df(site, [
        ("MATH", "jo@contoso.com", "contribute"),
        ("MATH", "typo@contoso.com", "contribute"),
    ]))

    [not_found] = rows(results, "user_not_found").itertuples()
    assert (not_found.user, not_found.permission) == ("typo@contoso.com", "contribute")
    assert "removals_skipped" in actions(results)
    assert "maybe@contoso.com" in tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute")


def test_removal_guardrail(make_manager, tenant, site, library):
    group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    for i in range(3):
        user = tenant.add_user(f"old{i}@contoso.com")
        group["members"][user["id"]] = user
    tenant.add_user("jo@contoso.com")
    df = memberships_df(site, [("MATH", "jo@contoso.com", "contribute")])

    results = make_manager(max_removals_per_group=2).update_memberships(df)

    assert "removal_guardrail_triggered" in actions(results)
    assert len(group["members"]) == 4

    results = make_manager(max_removals_per_group=2).update_memberships(df, allow_mass_removal=True)

    assert len(rows(results, "removed_from_group")) == 3
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]


def test_removals_up_to_the_guardrail_are_made(make_manager, tenant, site, library):
    group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    for i in range(2):
        user = tenant.add_user(f"old{i}@contoso.com")
        group["members"][user["id"]] = user
    tenant.add_user("jo@contoso.com")

    # The limit is how many may be removed, so removing exactly that many isn't blocked
    results = make_manager(max_removals_per_group=2).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]),
    )

    assert "removal_guardrail_triggered" not in actions(results)
    assert len(rows(results, "removed_from_group")) == 2
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]


def add_stale_members(tenant, group_name, count):
    group = tenant.group_named(group_name)
    for i in range(count):
        user = tenant.add_user(f"{group_name}-{i}@contoso.com")
        group["members"][user["id"]] = user
    return group


@pytest.fixture
def two_libraries_with_stale_members(make_manager, tenant, site):
    """MATH and ENG, each with three members their source data no longer lists."""
    make_manager().create_libraries(libraries_df(site, "MATH", "ENG"))
    groups = [add_stale_members(tenant, f"BAA-Sharepoint-{code}-Auto-Contribute", 3) for code in ("MATH", "ENG")]
    tenant.calls.clear()
    return groups


@pytest.mark.parametrize("dry_run", [False, True])
def test_removals_across_the_run_above_the_limit_are_all_skipped(
    make_manager, tenant, site, two_libraries_with_stale_members, dry_run,
):
    # Each group is under max_removals_per_group, as when a members table comes back empty
    results = make_manager(dry_run=dry_run, max_removals_per_run=5).update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH", "ENG"),
    )

    [row] = rows(results, "removal_guardrail_triggered").itertuples()
    assert "6 member(s) from 2 group(s)" in row.status
    assert not {"removed_from_group", "remove_from_group (dry run)"} & set(actions(results))
    assert tenant.writes() == []
    assert [len(group["members"]) for group in two_libraries_with_stale_members] == [3, 3]


@pytest.mark.parametrize("options", [
    {"max_removals_per_run": 6},
    {"max_removals_per_run": 5, "allow_mass_removal": True},
])
def test_removals_within_the_run_limit_or_allowed_go_ahead(
    make_manager, tenant, site, two_libraries_with_stale_members, options,
):
    results = make_manager(**options).update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH", "ENG"),
    )

    assert actions(results) == ["removed_from_group"] * 6
    assert [len(group["members"]) for group in two_libraries_with_stale_members] == [0, 0]


def test_libraries_df_reconciles_libraries_with_no_members(make_manager, tenant, site, library):
    left = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][left["id"]] = left

    results = make_manager().update_memberships(
        memberships_df(site, []),
        libraries_df=libraries_df(site, "MATH").drop(columns="library_display_name"),
    )

    assert actions(results) == ["removed_from_group"]
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == []


def test_missing_library_is_a_warning_not_an_error(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")

    results = make_manager().update_memberships(memberships_df(site, [("ART", "jo@contoso.com", "contribute")]))

    assert actions(results) == ["library_not_found"]
    assert "library_not_found" in WARNING_ACTIONS


def test_unknown_permission_is_skipped(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "edit")]))

    assert rows(results, "skip")["status"].str.startswith("unknown permission 'edit'").tolist() == [True]
    assert "skip" in WARNING_ACTIONS


@pytest.mark.parametrize("role", ["read", "contribute", "owner"])
def test_unknown_permission_never_removes_anyone(make_manager, tenant, site, library, role):
    # The row could be meant for any role, so every role's desired members are incomplete
    name = f"BAA-Sharepoint-MATH-Auto-{role.capitalize()}"
    group = tenant.group_named(name) or tenant.add_group(name, owners={APP_ID})
    stale = tenant.add_user("old@contoso.com")  # not in the input at all
    group["members"][stale["id"]] = stale
    tenant.add_user("jo@contoso.com")

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "Edit")]))

    assert tenant.member_upns(name) == ["old@contoso.com"]
    assert actions(results) == ["removals_skipped", "skip"]


def test_unknown_permission_for_a_nested_group_never_removes_nested_groups(make_manager, tenant, site, library):
    staff = tenant.add_group("BAA All Staff")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"][staff["id"]] = staff

    nested = pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "group_prefix": [PREFIX],
        "member_group": ["BAA All Staff"], "permission": ["edit"],
    })

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested)

    assert staff["id"] in tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"]
    assert actions(results) == ["removals_skipped", "skip"]


# HR and MIS extracts often have staff with no email
@pytest.mark.parametrize("missing, action", [
    (("MATH", None, "read"), "user_not_found"),
    (("MATH", "sam@contoso.com", None), "skip"),
    (("MATH", np.nan, "read"), "user_not_found"),
])
def test_a_member_row_missing_its_user_or_permission_is_a_row_not_a_failed_sync(
    make_manager, tenant, site, library, missing, action,
):
    tenant.add_user("jo@contoso.com")
    left = tenant.add_user("left@contoso.com")
    read_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")
    read_group["members"][left["id"]] = left
    df = memberships_df(site, [("MATH", "jo@contoso.com", "read"), missing])

    results = make_manager().update_memberships(df)

    assert actions(results) == sorted(["added_to_group", "already_assigned", action, "removals_skipped"])
    # The others are synced, and, as for any user who can't be found, nobody is removed
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Read") == ["jo@contoso.com", "left@contoso.com"]


@pytest.mark.parametrize("pandas_1_5", [False, True])
def test_a_numeric_permission_beside_a_missing_one_is_skipped_on_every_pandas(
    make_manager, tenant, site, library, monkeypatch, pandas_1_5,
):
    from alt_fabric_m365.sharepoint import _common
    if pandas_1_5:
        # pandas 1.5, which older Fabric runtimes have, counts any object column as strings
        monkeypatch.setattr(
            _common.pd.api.types, "is_string_dtype", lambda column: getattr(column, "dtype", None) == object,
        )
    tenant.add_user("jo@contoso.com")
    df = memberships_df(site, [("MATH", "jo@contoso.com", 7), ("MATH", "jo@contoso.com", None)])

    results = make_manager().update_memberships(df)

    assert rows(results, "error").empty
    assert sorted(rows(results, "skip")["permission"]) == ["", "7.0"]


def test_a_nested_group_row_missing_its_group_is_a_row_not_a_failed_sync(make_manager, tenant, site, library):
    staff = tenant.add_group("BAA All Staff")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"][staff["id"]] = staff

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested_df(site.url, None))

    assert "removals_skipped" in actions(results)
    assert rows(results, "error").empty
    assert staff["id"] in tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"]


def test_nested_groups_are_added_and_unresolvable_ones_reported(make_manager, tenant, site, library):
    all_staff = tenant.add_group("BAA All Staff")
    old_nested = tenant.add_group("Old Nested")
    reader = tenant.add_user("reader@contoso.com")
    read_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")
    contribute_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    stale_contributor = tenant.add_user("left@contoso.com")
    read_group["members"] = {reader["id"]: reader}
    contribute_group["members"] = {old_nested["id"]: old_nested, stale_contributor["id"]: stale_contributor}
    nested = pd.DataFrame({
        "site_url": site.url, "library_code": "MATH", "group_prefix": PREFIX,
        "member_group": ["BAA All Staff", "Missing Group"], "permission": ["read", "contribute"],
    })

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested)

    assert set(read_group["members"]) == {all_staff["id"]}  # stale user removed, group nested
    # Contribute's nested-group lookup failed, so its stale nested group is kept. Its users all
    # resolved, so the stale user is still removed: only removals of the kind that failed are skipped.
    assert set(contribute_group["members"]) == {old_nested["id"]}
    assert rows(results, "group_not_found")["user"].tolist() == ["Missing Group"]
    assert "removals_skipped" in actions(results)


def test_onedrive_shortcut_once_per_new_user(make_manager, tenant, site, library):
    jo = tenant.add_user("jo@contoso.com")

    results = make_manager(create_onedrive_shortcuts=True).update_memberships(memberships_df(site, [
        ("MATH", "jo@contoso.com", "read"),
        ("MATH", "jo@contoso.com", "contribute"),
    ]))

    [(user_id, body)] = tenant.shortcuts
    assert user_id == jo["id"]
    assert body["name"] == "MATH Faculty"
    assert body["remoteItem"]["sharepointIds"]["listId"] == library["id"]
    assert len(rows(results, "onedrive_shortcut_created")) == 1


def test_onedrive_shortcut_failure_does_not_block_the_sync(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    tenant.add_user("jo@contoso.com")
    tenant.fail("POST", r"/drive/root/children$", graph_error(403, "accessDenied"))

    results = make_manager(create_onedrive_shortcuts=True).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]),
    )

    assert "onedrive_shortcut_failed" in actions(results)
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]


def test_memberships_dry_run_shows_the_adds_and_removals_a_live_run_makes(make_manager, tenant, site, library):
    jo, sam, _ = (tenant.add_user(f"{name}@contoso.com") for name in ("jo", "sam", "ali"))
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"].update({jo["id"]: jo, sam["id"]: sam})
    df = memberships_df(site, [("MATH", "jo@contoso.com", "contribute"), ("MATH", "ali@contoso.com", "contribute")])
    tenant.calls.clear()

    dry = make_manager(dry_run=True).update_memberships(df)

    assert tenant.writes() == []
    assert sorted(dry[["action", "user"]].values.tolist()) == [
        ["add_to_group (dry run)", "ali@contoso.com"], ["already_assigned", ""],
        ["already_member", "jo@contoso.com"], ["remove_from_group (dry run)", "sam@contoso.com"],
    ]

    live = make_manager().update_memberships(df)

    assert sorted(live[["action", "user"]].values.tolist()) == [
        ["added_to_group", "ali@contoso.com"], ["already_assigned", ""],
        ["already_member", "jo@contoso.com"], ["removed_from_group", "sam@contoso.com"],
    ]


@pytest.mark.parametrize("allow_mass_removal", [False, True])
def test_memberships_dry_run_shows_the_removal_guardrail(make_manager, tenant, site, library, allow_mass_removal):
    contribute = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    for i in range(11):
        user = tenant.add_user(f"user{i}@contoso.com")
        contribute["members"][user["id"]] = user
    tenant.calls.clear()

    results = make_manager(dry_run=True, allow_mass_removal=allow_mass_removal).update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH"),
    )

    assert tenant.writes() == []
    assert actions(results) == (
        ["remove_from_group (dry run)"] * 11 if allow_mass_removal else ["removal_guardrail_triggered"]
    )


def test_memberships_dry_run_shows_removals_skipped_for_unresolved_users(make_manager, tenant, site, library):
    sam = tenant.add_user("sam@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][sam["id"]] = sam
    tenant.calls.clear()

    results = make_manager(dry_run=True).update_memberships(
        memberships_df(site, [("MATH", "nobody@contoso.com", "contribute")]),
    )

    assert tenant.writes() == []
    assert actions(results) == ["removals_skipped", "user_not_found"]


def test_memberships_dry_run_makes_no_writes(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    tenant.calls.clear()

    results = make_manager(dry_run=True).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]),
    )

    assert tenant.writes() == []
    assert "add_to_group (dry run)" in actions(results)


def test_memberships_in_parallel(make_manager, tenant, site):
    codes = [f"F{i}" for i in range(6)]
    make_manager().create_libraries(libraries_df(site, *codes))
    for i in range(4):
        tenant.add_user(f"user{i}@contoso.com")
    df = memberships_df(site, [(code, f"user{i}@contoso.com", "contribute") for code in codes for i in range(4)])

    results = make_manager(max_workers=4).update_memberships(df)

    assert len(rows(results, "added_to_group")) == 24
    assert rows(results, "error").empty


# ensure_onedrive_shortcuts

def test_ensure_onedrive_shortcuts_makes_the_missing_ones_once(make_manager, tenant, site, library):
    jo, sam, ali = (tenant.add_user(f"{name}@contoso.com") for name in ("jo", "sam", "ali"))
    make_manager(create_onedrive_shortcuts=True).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "read")]),
    )
    tenant.shortcuts[0][1]["name"] = "My maths"  # renamed by jo, and still a shortcut to the library
    tenant.without_onedrive.add(ali["id"])  # a new starter who hasn't opened OneDrive yet
    df = memberships_df(site, [
        ("MATH", "jo@contoso.com", "read"), ("MATH", "jo@contoso.com", "contribute"),
        ("MATH", "sam@contoso.com", "read"), ("MATH", "ali@contoso.com", "read"),
    ])

    results = make_manager().ensure_onedrive_shortcuts(df)

    assert results[["user", "action"]].values.tolist() == [
        ["ali@contoso.com", "onedrive_shortcut_failed"],
        ["jo@contoso.com", "onedrive_shortcut_exists"],
        ["sam@contoso.com", "onedrive_shortcut_created"],
    ]
    assert "doesn't exist yet" in rows(results, "onedrive_shortcut_failed")["status"].iloc[0]
    assert [(user_id, body["name"]) for user_id, body in tenant.shortcuts] == [
        (jo["id"], "My maths"), (sam["id"], "MATH Faculty"),
    ]

    tenant.without_onedrive.clear()  # ali has now opened OneDrive

    results = make_manager().ensure_onedrive_shortcuts(df)

    assert actions(results) == ["onedrive_shortcut_created", "onedrive_shortcut_exists", "onedrive_shortcut_exists"]
    assert len(tenant.shortcuts) == 3


def test_ensure_onedrive_shortcuts_dry_run_makes_no_writes(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    tenant.calls.clear()

    results = make_manager(dry_run=True).ensure_onedrive_shortcuts(
        memberships_df(site, [("MATH", "jo@contoso.com", "read")]),
    )

    assert actions(results) == ["onedrive_shortcut (dry run)"]
    assert tenant.writes() == []


def test_ensure_onedrive_shortcuts_reports_missing_libraries_and_users(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")

    results = make_manager().ensure_onedrive_shortcuts(memberships_df(site, [
        ("ART", "jo@contoso.com", "read"), ("MATH", "nobody@contoso.com", "read"),
    ]))

    assert actions(results) == ["library_not_found", "user_not_found"]
    assert tenant.shortcuts == []


def test_ensure_onedrive_shortcuts_does_not_count_a_shortcut_to_a_folder_in_the_library(
    make_manager, tenant, site, library,
):
    jo = tenant.add_user("jo@contoso.com")
    tenant.shortcuts.append((jo["id"], {"name": "Homework", "remoteItem": {
        "sharepointIds": {"listId": library["id"], "listItemId": "3"}, "webUrl": f"{site.web_url(library)}/Homework",
    }}))
    df = memberships_df(site, [("MATH", "jo@contoso.com", "read")])

    results = make_manager().ensure_onedrive_shortcuts(df)
    again = make_manager().ensure_onedrive_shortcuts(df)

    assert (actions(results), actions(again)) == (["onedrive_shortcut_created"], ["onedrive_shortcut_exists"])
    assert [body["name"] for _, body in tenant.shortcuts] == ["Homework", "MATH Faculty"]


def test_ensure_onedrive_shortcuts_handles_each_library_once_however_the_rows_name_it(
    make_manager, tenant, site, library,
):
    tenant.add_user("jo@contoso.com")
    df = pd.DataFrame(
        [(site.url, "MATH", "jo@contoso.com"), (f"{site.url}/", "math", "Jo@contoso.com")],
        columns=["site_url", "library_code", "user_principal_name"],
    )

    results = make_manager(max_workers=4).ensure_onedrive_shortcuts(df)

    assert actions(results) == ["onedrive_shortcut_created"]
    assert len(tenant.shortcuts) == 1


# grant_group_access

@pytest.fixture
def slt_groups(make_manager, site, library):
    make_manager().create_libraries(libraries_df(site, "SLT", auto_groups=False, manual_read=True,
                                                 manual_contribute=True))


def grants_df(site, rows_):
    return pd.DataFrame(
        [(site.url, code, group, permission) for code, group, permission in rows_],
        columns=["site_url", "library_code", "group_name", "permission"],
    )


def test_grant_group_access_is_additive_and_idempotent(make_manager, tenant, site, library, slt_groups):
    df = grants_df(site, [
        ("MATH", "BAA-Sharepoint-SLT-Manual-Contribute", "contribute"),
        ("MATH", "BAA-Sharepoint-SLT-Manual-Read", " Read "),
    ])

    results = make_manager().grant_group_access(df)

    assert actions(results) == ["group_access_granted", "group_access_granted"]
    contribute = tenant.group_named("BAA-Sharepoint-SLT-Manual-Contribute")
    assert library["role_assignments"][principal_id(site, contribute)] == {CONTRIBUTE}

    results = make_manager().grant_group_access(df)

    assert actions(results) == ["already_assigned", "already_assigned"]


# SharePoint pages a library's role assignments, so a group's may only be on a later page
@pytest.mark.parametrize("operation", ["create_libraries", "grant_group_access"])
def test_role_assignments_on_a_later_page_are_seen(make_manager, tenant, site, library, slt_groups, operation):
    df = grants_df(site, [
        ("MATH", "BAA-Sharepoint-SLT-Manual-Contribute", "contribute"),
        ("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read"),
    ])
    make_manager().grant_group_access(df)
    tenant.role_assignment_page_size = 1
    tenant.calls.clear()

    if operation == "create_libraries":
        results = make_manager().create_libraries(libraries_df(site, "MATH"))
        assert "already_assigned" in actions(results) and "group_assigned" not in actions(results)
    else:
        results = make_manager().grant_group_access(df)
        assert actions(results) == ["already_assigned", "already_assigned"]

    assert created(tenant, "addroleassignment") == []
    assert any("$skiptoken=" in url for method, url, _ in tenant.calls if url.split("?")[0].endswith("/roleassignments"))


def test_grant_reports_missing_groups_libraries_and_duplicate_names(make_manager, tenant, site, library, slt_groups):
    tenant.add_group("Dup")
    tenant.add_group("Dup")

    results = make_manager().grant_group_access(grants_df(site, [
        ("MATH", "Nobody", "read"),
        ("ART", "BAA-Sharepoint-SLT-Manual-Read", "read"),
        ("MATH", "Dup", "read"),
    ]))

    assert actions(results) == ["error", "group_not_found", "library_not_found"]


def test_grant_rejects_unknown_permissions(make_manager, site):
    with pytest.raises(ValueError, match="Unknown permission"):
        make_manager().grant_group_access(grants_df(site, [("MATH", "G", "edit")]))


def test_grant_dry_run_makes_no_writes(make_manager, tenant, site):
    results = make_manager(dry_run=True).grant_group_access(grants_df(site, [("MATH", "G", "read")]))

    assert tenant.writes() == []
    assert actions(results) == ["group_not_found"]


def test_grant_dry_run_reports_what_a_live_run_does(make_manager, tenant, site, library, slt_groups):
    tenant.add_group("Dup")
    tenant.add_group("Dup")
    make_manager().grant_group_access(grants_df(site, [("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read")]))
    df = grants_df(site, [
        ("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read"),  # already held
        ("MATH", "BAA-Sharepoint-SLT-Manual-Contribute", "contribute"),
        ("MATH", "Nobody", "read"),
        ("ART", "BAA-Sharepoint-SLT-Manual-Read", "read"),
        ("MATH", "Dup", "read"),
    ])
    tenant.calls.clear()

    dry = make_manager(dry_run=True).grant_group_access(df)

    assert tenant.writes() == []
    assert actions(dry) == sorted([
        "already_assigned", "grant_group_access (dry run)", "group_not_found", "library_not_found", "error",
    ])

    live = make_manager().grant_group_access(df)

    assert actions(live) == sorted([
        "already_assigned", "group_access_granted", "group_not_found", "library_not_found", "error",
    ])


def test_a_grant_run_again_makes_no_writes(make_manager, tenant, site, library, slt_groups):
    df = grants_df(site, [("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read")])
    make_manager().grant_group_access(df)
    tenant.calls.clear()

    results = make_manager().grant_group_access(df)

    assert actions(results) == ["already_assigned"]
    # A grant the library already holds doesn't register the group on the site (ensureuser) either
    assert tenant.writes() == []


def test_a_group_is_registered_once_per_site_for_all_its_libraries(make_manager, tenant, site, library, slt_groups):
    make_manager().create_libraries(libraries_df(site, "ENG", "ART"))
    tenant.calls.clear()

    results = make_manager(max_workers=3).grant_group_access(grants_df(site, [
        (code, "BAA-Sharepoint-SLT-Manual-Read", "read") for code in ("MATH", "ENG", "ART")
    ]))

    assert actions(results) == ["group_access_granted"] * 3
    assert len(created(tenant, "/web/ensureuser")) == 1
    assert len(created(tenant, "addroleassignment")) == 3


# enforce_site_sharing_lockdown

def test_lockdown_turns_sharing_off_once_per_site(make_manager, tenant, site):
    manager = make_manager()

    results = manager.enforce_site_sharing_lockdown([site.url, site.url])

    [row] = results.itertuples()
    assert row.action == "site_sharing_disabled"
    assert row.status == "turned off: members_can_share, access_requests_enabled, members_can_invite"
    assert manager.get_site_sharing_settings(site.url) == {
        "members_can_share": False, "access_requests_enabled": False, "members_can_invite": False,
        "member_group_id": 5,
    }

    assert actions(manager.enforce_site_sharing_lockdown([site.url])) == ["site_sharing_check"]


@pytest.mark.parametrize("max_workers", [1, 4])
def test_lockdown_checks_a_site_spelt_differently_once(make_manager, tenant, site, max_workers):
    spellings = [site.url, f"{site.url}/", site.url.upper()]

    results = make_manager(max_workers=max_workers).enforce_site_sharing_lockdown(spellings)

    assert results[["site_url", "action"]].values.tolist() == [[site.url, "site_sharing_disabled"]]


def test_lockdown_of_one_url_given_as_a_string_raises_before_any_request(make_manager, tenant, site):
    with pytest.raises(TypeError, match=r"pass \[site_url\]"):
        make_manager().enforce_site_sharing_lockdown(site.url)

    assert tenant.calls == []


def test_lockdown_dry_run_makes_no_writes(make_manager, tenant, site):
    results = make_manager(dry_run=True).enforce_site_sharing_lockdown([site.url])

    assert tenant.writes() == []
    assert actions(results) == ["site_sharing_lockdown (dry run)"]


def test_lockdown_that_does_not_stick_is_an_error(make_manager, site):
    site.ignore_sharing_writes = True

    results = make_manager().enforce_site_sharing_lockdown([site.url])

    [row] = results.itertuples()
    assert row.action == "error" and "did not stick" in row.status


def test_lockdown_without_a_member_group(make_manager, tenant, site):
    site.member_group = None
    site.web.update(MembersCanShare=False, RequestAccessEmail="", UseAccessRequestDefault=True)

    results = make_manager().enforce_site_sharing_lockdown([site.url])

    assert results["status"].tolist() == ["turned off: access_requests_enabled"]
    assert not any("sitegroups" in c[1] for c in tenant.calls)


# Read-only helpers

@pytest.mark.parametrize("code", ["MATH#notes", "MATH?x=1", "D&T"])
def test_get_navigation_link_refuses_a_code_sharepoint_cant_use_as_a_url(make_manager, tenant, site, code):
    with pytest.raises(ValueError, match="library_code"):
        make_manager().get_navigation_link(site.url, code)

    assert tenant.calls == []


def test_get_navigation_link(make_manager, site):
    manager = make_manager(add_to_navigation=True)
    assert manager.get_navigation_link(site.url, "MATH") is None

    manager.create_libraries(libraries_df(site, "MATH", library_display_name="Maths & Stats"))

    assert manager.get_navigation_link(site.url, "math")["Title"] == "Maths & Stats"
    assert manager.get_navigation_link(site.url, "ENG") is None
    assert manager.get_navigation_link(site.url, "MATH", navigation_label="Other") is None


# Failure paths and remaining dry runs

def test_navigation_label_failure_does_not_block_permissions(make_manager, tenant, site):
    from sharepoint_fake import response
    tenant.fail("GET", r"/navigation/quicklaunch", response(403))

    results = make_manager(add_to_navigation=True).create_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["error", "group_assigned", "group_assigned", "library_created"]
    assert site.nav == []


def test_a_navigation_link_is_not_posted_again_after_a_server_error(make_manager, tenant, site):
    from sharepoint_fake import response
    # A 5xx may come after SharePoint created the link, so posting it again could duplicate it
    tenant.fail("POST", r"/GetNodeById\(\d+\)/Children$", response(503))

    results = make_manager(add_to_navigation=True).create_libraries(libraries_df(site, "MATH"))

    posts = [url for method, url, _ in tenant.calls if method == "POST" and url.endswith("/Children")]
    assert len(posts) == 1
    [error] = rows(results, "error").itertuples()
    assert error.document_library == "MATH" and "Status code: 503" in error.status
    [label] = site.nav
    assert label["Children"] == []

    make_manager(add_to_navigation=True).create_libraries(libraries_df(site, "MATH"))

    assert [child["Title"] for child in label["Children"]] == ["MATH Faculty"]


def test_failure_rows_give_graphs_error_code_and_message(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    tenant.add_user("jo@contoso.com")
    tenant.fail("POST", r"/members/\$ref$", graph_error(
        403, "Authorization_RequestDenied", "Insufficient privileges to complete the operation."))

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    assert list(rows(results, "add_failed")["status"]) == [
        "Status code: 403, Authorization_RequestDenied: Insufficient privileges to complete the operation."
    ]


def test_error_rows_give_sharepoints_error_message(make_manager, tenant, site):
    from sharepoint_fake import response
    tenant.fail("POST", r"/addroleassignment", response(403, {"error": {
        "code": "-2147024891, System.UnauthorizedAccessException", "message": {"lang": "en-US", "value": "Access denied."},
    }}))

    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert list(rows(results, "error")["status"]) == [
        "Status code: 403, -2147024891, System.UnauthorizedAccessException: Access denied."
    ]


def test_a_member_already_removed_by_a_retried_request_is_not_a_failure(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    left = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][left["id"]] = left
    # The first DELETE removed the member, but its response was lost, so the retry finds them gone
    tenant.fail("DELETE", r"/members/.+/\$ref$", graph_error(404, "Request_ResourceNotFound", "Does not exist."))

    results = make_manager().update_memberships(memberships_df(site, []), libraries_df=libraries_df(site, "MATH"))

    assert results[["action", "user", "status"]].values.tolist() == [
        ["removed_from_group", "left@contoso.com", "already removed"],
    ]


@pytest.mark.parametrize("operation", ["update_memberships", "ensure_onedrive_shortcuts"])
def test_one_worker_makes_every_request_in_turn_including_user_lookups(
    make_manager, tenant, site, library, monkeypatch, operation,
):
    import threading
    from alt_fabric_m365.sharepoint import _http
    for name in ("jo", "sam", "ali"):
        tenant.add_user(f"{name}@contoso.com")
    threads = set()

    def request(method, url, **kwargs):
        threads.add(threading.get_ident())
        return tenant.request(method, url, **kwargs)
    monkeypatch.setattr(_http.requests, "request", request)

    getattr(make_manager(max_workers=1), operation)(memberships_df(site, [
        ("MATH", f"{name}@contoso.com", "read") for name in ("jo", "sam", "ali")
    ]))

    assert threads == {threading.get_ident()}


def test_add_and_remove_failures_are_fatal_rows(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    tenant.add_user("jo@contoso.com")
    stale = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][stale["id"]] = stale
    tenant.fail("POST", r"/members/\$ref$", graph_error(403, "Authorization_RequestDenied"))
    tenant.fail("DELETE", r"/members/.+/\$ref$", graph_error(403, "Authorization_RequestDenied"))

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    failed = results[results["action"].isin(FATAL_ACTIONS)]
    assert sorted(failed["action"]) == ["add_failed", "remove_failed"]


def test_a_removal_that_cannot_reach_graph_is_logged_against_its_group_and_member(
    make_manager, tenant, site, library, monkeypatch,
):
    import requests
    from alt_fabric_m365.sharepoint import _http
    stale = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][stale["id"]] = stale

    def request(method, url, **kwargs):
        if method == "DELETE":
            raise requests.ConnectionError("Connection reset by peer")
        return tenant.request(method, url, **kwargs)
    monkeypatch.setattr(_http.requests, "request", request)

    results = make_manager().update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH"),
    )

    [row] = results[results["action"].isin(FATAL_ACTIONS)].itertuples()
    assert (row.action, row.document_library, row.group, row.user) == (
        "remove_failed", "MATH", "BAA-Sharepoint-MATH-Auto-Contribute", "left@contoso.com",
    )
    assert "Connection reset by peer" in row.status


# An error a worker doesn't handle, such as a KeyError on an unexpected response, for each kind of work item
@pytest.mark.parametrize("worker, run, library_code", [
    ("_process_create_library", lambda m, site: m.create_libraries(libraries_df(site, "MATH")), "MATH"),
    ("_process_update_memberships", lambda m, site: m.update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "read")])), "MATH"),
    ("_process_removals", lambda m, site: m.update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH")), "MATH"),
    ("_process_grant_group_access", lambda m, site: m.grant_group_access(
        grants_df(site, [("MATH", "BAA-Sharepoint-MATH-Auto-Read", "read")])), "MATH"),
    ("_process_enforce_site_sharing", lambda m, site: m.enforce_site_sharing_lockdown([site.url]), ""),
])
def test_an_unhandled_error_is_logged_against_its_site_and_library(
    make_manager, tenant, site, library, monkeypatch, worker, run, library_code,
):
    tenant.add_user("jo@contoso.com")
    left = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][left["id"]] = left

    def fail(*args, **kwargs):
        raise KeyError("Id")
    monkeypatch.setattr(SharePointLibraryManager, worker, fail)

    results = run(make_manager(), site)

    [row] = rows(results, "error").itertuples()
    assert (row.site_url, row.document_library, row.status) == (site.url, library_code, "KeyError: 'Id'")


def test_a_user_lookup_graph_fails_fails_the_library_without_removing_anyone(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    tenant.add_user("jo@contoso.com")
    stale = tenant.add_user("left@contoso.com")
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["members"][stale["id"]] = stale
    tenant.calls.clear()
    # Not a 404: Graph couldn't say whether the user exists, so the "stale" member may be them
    tenant.fail("GET", r"/users/jo%40contoso\.com$", graph_error(503, "serviceNotAvailable"), times=99)

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    [error] = rows(results, "error").itertuples()
    assert error.document_library == "MATH"
    assert "user 'jo@contoso.com': user lookup failed: Status code: 503, serviceNotAvailable" in error.status
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["left@contoso.com"]
    assert tenant.writes() == []


def test_a_nested_group_lookup_graph_fails_is_an_error_and_keeps_nested_groups(
    make_manager, tenant, site, library,
):
    from sharepoint_fake import graph_error
    tenant.add_group("BAA All Staff")
    old_nested = tenant.add_group("Old Nested")
    read_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")
    read_group["members"] = {old_nested["id"]: old_nested}
    tenant.calls.clear()
    # The nested group is looked up before the library's own groups
    tenant.fail("GET", r"/groups$", graph_error(403, "Authorization_RequestDenied", "Insufficient privileges"))

    results = make_manager().update_memberships(
        memberships_df(site, []), nested_groups_df=nested_df(site.url, "BAA All Staff"),
    )

    [error] = rows(results, "error").itertuples()
    assert error.user == "BAA All Staff" and "group lookup failed" in error.status
    assert set(read_group["members"]) == {old_nested["id"]}
    assert tenant.writes() == []


def test_nested_group_add_failure_is_fatal(make_manager, tenant, site, library):
    from sharepoint_fake import graph_error
    tenant.add_group("BAA All Staff")
    tenant.fail("POST", r"/members/\$ref$", graph_error(400, "Request_BadRequest", "cannot nest"))
    nested = pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "group_prefix": [PREFIX],
        "member_group": ["BAA All Staff"], "permission": ["read"],
    })

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested)

    [failed] = rows(results, "add_failed").itertuples()
    assert "nested group 'BAA All Staff'" in failed.status


def test_memberships_dry_run_lists_new_groups_shortcuts_and_nested_groups(make_manager, tenant, site, library):
    tenant.add_user("jo@contoso.com")
    tenant.add_group("BAA All Staff")
    nested = pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "group_prefix": [PREFIX],
        "member_group": ["BAA All Staff"], "permission": ["read"],
    })
    tenant.calls.clear()

    results = make_manager(dry_run=True, create_onedrive_shortcuts=True).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "owner")]), nested_groups_df=nested,
    )

    assert tenant.writes() == []
    assert actions(results) == sorted([
        "ensure_group (dry run)", "add_to_group (dry run)", "onedrive_shortcut (dry run)",  # owner, a new group
        "already_assigned", "add_to_group (dry run)",  # read, nested group
    ])  # contribute exists, but has no members to remove


def test_retire_dry_run_makes_no_writes(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))
    tenant.calls.clear()

    results = make_manager(dry_run=True, add_to_navigation=True).retire_moved_libraries(libraries_df(site, "MATH"))

    assert tenant.writes() == []
    assert actions(results) == sorted([
        "retire_library (dry run)", "remove_group_access (dry run)", "remove_group_access (dry run)",
        "remove_navigation_link (dry run)",
    ])


def test_grant_skips_a_site_whose_prefetch_fails(make_manager, tenant, site, library, slt_groups):
    from sharepoint_fake import response
    tenant.fail("GET", r"/roledefinitions", response(503), times=99)

    results = make_manager().grant_group_access(
        grants_df(site, [("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read")]),
    )

    assert actions(results) == ["error"]


# Each site takes several reads before the per-library work, so they run in parallel with max_workers above 1
def on_both(build, a, b):
    """The rows build(site) gives for each of two sites."""
    return pd.concat([build(a), build(b)], ignore_index=True)


TWO_SITE_OPERATIONS = {
    "create_libraries": lambda m, a, b: m.create_libraries(on_both(lambda s: libraries_df(s, "MATH"), a, b)),
    "set_library_metadata": lambda m, a, b: m.set_library_metadata(
        on_both(lambda s: libraries_df(s, "MATH", source_system="MIS"), a, b)),
    "rename_libraries": lambda m, a, b: m.rename_libraries(on_both(lambda s: libraries_df(s, "MATH"), a, b)),
    "unlink_libraries": lambda m, a, b: m.unlink_libraries(on_both(lambda s: libraries_df(s, "MATH"), a, b)),
    "retire_moved_libraries": lambda m, a, b: m.retire_moved_libraries(
        on_both(lambda s: libraries_df(s, "MATH"), a, b)),
    "update_memberships": lambda m, a, b: m.update_memberships(
        on_both(lambda s: memberships_df(s, [("MATH", "jo@contoso.com", "read")]), a, b)),
    "ensure_onedrive_shortcuts": lambda m, a, b: m.ensure_onedrive_shortcuts(
        on_both(lambda s: memberships_df(s, [("MATH", "jo@contoso.com", "read")]), a, b)),
    "grant_group_access": lambda m, a, b: m.grant_group_access(
        on_both(lambda s: grants_df(s, [("MATH", "Staff", "read")]), a, b)),
}


def overlap_site_reads(monkeypatch, first, second):
    """Makes the first site's read end only after the second's, so the two must overlap and finish out of order."""
    import threading
    from alt_fabric_m365.sharepoint import manager as manager_module
    real_get_site_id, second_read = manager_module._get_site_id, threading.Event()

    def get_site_id(token, site_url):
        if site_url == first.url:
            assert second_read.wait(timeout=5), "the second site wasn't read while the first was"
        try:
            return real_get_site_id(token, site_url)
        finally:
            if site_url == second.url:
                second_read.set()
    monkeypatch.setattr(manager_module, "_get_site_id", get_site_id)


@pytest.mark.parametrize("operation", TWO_SITE_OPERATIONS)
def test_sites_are_read_in_parallel_and_their_errors_logged_in_order(make_manager, tenant, monkeypatch, operation):
    from sharepoint_fake import graph_error
    tenant.add_user("jo@contoso.com")
    tenant.add_group("Staff")
    first, second = tenant.add_site("/sites/A"), tenant.add_site("/sites/B")
    tenant.fail("GET", r":/sites/[AB]$", graph_error(403, "accessDenied", "Access denied"), times=2)
    overlap_site_reads(monkeypatch, first, second)

    results = TWO_SITE_OPERATIONS[operation](make_manager(max_workers=2), first, second)

    assert rows(results, "error")["site_url"].tolist() == [first.url, second.url]
    assert tenant.writes() == []


def test_retire_reads_the_new_sites_in_parallel_and_keeps_libraries_it_cant_check(
    make_manager, tenant, site, monkeypatch,
):
    from sharepoint_fake import graph_error
    make_manager().create_libraries(libraries_df(site, "MATH", "ENG"))
    first, second = tenant.add_site("/sites/A"), tenant.add_site("/sites/B")
    tenant.fail("GET", r":/sites/[AB]$", graph_error(403, "accessDenied", "Access denied"), times=2)
    overlap_site_reads(monkeypatch, first, second)
    tenant.calls.clear()

    results = make_manager(max_workers=2).retire_moved_libraries(
        libraries_df(site, "MATH", "ENG").assign(new_site_url=[first.url, second.url]),
    )

    errors = rows(results, "error")
    assert errors["document_library"].tolist() == ["MATH", "ENG"]
    assert all("couldn't be read, so it wasn't retired" in status for status in errors["status"])
    assert tenant.writes() == []
    assert [site.list_by_slug(code)["displayName"] for code in ("MATH", "ENG")] == ["MATH Faculty", "ENG Faculty"]


def test_cancelling_during_the_site_reads_starts_no_more_of_them(make_manager, tenant, monkeypatch):
    import threading
    from concurrent.futures import ThreadPoolExecutor
    from alt_fabric_m365.sharepoint import manager as manager_module
    sites = [tenant.add_site(f"/sites/{name}") for name in "ABCD"]
    real_get_site_id, started, lock = manager_module._get_site_id, [], threading.Lock()
    first_two_running = threading.Barrier(3, timeout=5)  # the two workers and the cancelled cell

    def get_site_id(token, site_url):
        with lock:
            started.append(site_url)
            first_two = len(started) <= 2
        if first_two:
            first_two_running.wait()
            threading.Event().wait(0.5)  # still running when the cell is cancelled
        return real_get_site_id(token, site_url)

    class CancelledExecutor(ThreadPoolExecutor):
        """Interrupted, as by cancelling the notebook cell, while the first two reads run."""

        def map(self, fn, *iterables):
            super().map(fn, *iterables)
            first_two_running.wait()
            raise KeyboardInterrupt
    monkeypatch.setattr(manager_module, "_get_site_id", get_site_id)
    monkeypatch.setattr(manager_module, "ThreadPoolExecutor", CancelledExecutor)

    with pytest.raises(KeyboardInterrupt):
        make_manager(max_workers=2).unlink_libraries(pd.concat([libraries_df(s, "MATH") for s in sites]))

    assert sorted(started) == [sites[0].url, sites[1].url]
    assert tenant.writes() == []


# Regression tests

def test_nested_group_with_an_ampersand_in_its_name(make_manager, tenant, site, library):
    staff = tenant.add_group("Maths & Science Staff")
    nested = pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "group_prefix": [PREFIX],
        "member_group": ["Maths & Science Staff"], "permission": ["read"],
    })

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested)

    assert rows(results, "error").empty
    assert staff["id"] in tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"]


def nested_df(site_url, member_group="Staff"):
    return pd.DataFrame({
        "site_url": [site_url], "library_code": ["MATH"], "group_prefix": [PREFIX],
        "member_group": [member_group], "permission": ["read"],
    })


# Every way a site URL reaches the manager, each called with the URL under test
SITE_URL_OPERATIONS = {
    "create_libraries": lambda m, site, url: m.create_libraries(libraries_df(site, "MATH").assign(site_url=url)),
    "set_library_metadata": lambda m, site, url: m.set_library_metadata(
        libraries_df(site, "MATH", source_system="MIS").assign(site_url=url)),
    "get_library_metadata": lambda m, site, url: m.get_library_metadata(
        libraries_df(site, "MATH").assign(site_url=url)),
    "rename_libraries": lambda m, site, url: m.rename_libraries(libraries_df(site, "MATH").assign(site_url=url)),
    "unlink_libraries": lambda m, site, url: m.unlink_libraries(libraries_df(site, "MATH").assign(site_url=url)),
    "retire_moved_libraries": lambda m, site, url: m.retire_moved_libraries(
        libraries_df(site, "MATH").assign(site_url=url)),
    "retire_moved_libraries new_site_url": lambda m, site, url: m.retire_moved_libraries(
        libraries_df(site, "MATH").assign(new_site_url=url)),
    "update_memberships": lambda m, site, url: m.update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "read")]).assign(site_url=url)),
    "update_memberships libraries_df": lambda m, site, url: m.update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH").assign(site_url=url)),
    "update_memberships nested_groups_df": lambda m, site, url: m.update_memberships(
        memberships_df(site, []), nested_groups_df=nested_df(url)),
    "grant_group_access": lambda m, site, url: m.grant_group_access(grants_df(site, [("MATH", "Staff", "read")]).assign(
        site_url=url)),
    "enforce_site_sharing_lockdown": lambda m, site, url: m.enforce_site_sharing_lockdown([url]),
    "get_site_sharing_settings": lambda m, site, url: m.get_site_sharing_settings(url),
    "get_navigation_link": lambda m, site, url: m.get_navigation_link(url, "MATH"),
}


@pytest.mark.parametrize("url", [
    "http://contoso.sharepoint.com/sites/x",
    "https://evil.example.com/sites/x",
    # urlparse reads these as on contoso.sharepoint.com, but requests sends the first to evil.example.com
    "https://evil.example.com\\@contoso.sharepoint.com/sites/x",
    "https://user@contoso.sharepoint.com/sites/x",
])
@pytest.mark.parametrize("operation", SITE_URL_OPERATIONS)
def test_sharepoint_token_is_never_sent_to_another_host(make_manager, tenant, site, operation, url):
    with pytest.raises(ValueError, match="site_url"):
        SITE_URL_OPERATIONS[operation](make_manager(), site, url)

    assert tenant.calls == []


# Every DataFrame a library code reaches the manager in, each called with the code under test
LIBRARY_CODE_OPERATIONS = {
    "create_libraries": lambda m, site, code: m.create_libraries(libraries_df(site, code)),
    "set_library_metadata": lambda m, site, code: m.set_library_metadata(libraries_df(site, code, source_system="MIS")),
    "get_library_metadata": lambda m, site, code: m.get_library_metadata(libraries_df(site, code)),
    "rename_libraries": lambda m, site, code: m.rename_libraries(libraries_df(site, code)),
    "unlink_libraries": lambda m, site, code: m.unlink_libraries(libraries_df(site, code)),
    "retire_moved_libraries": lambda m, site, code: m.retire_moved_libraries(libraries_df(site, code)),
    "update_memberships": lambda m, site, code: m.update_memberships(
        memberships_df(site, [(code, "jo@contoso.com", "read")])),
    "update_memberships libraries_df": lambda m, site, code: m.update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, code)),
    "update_memberships nested_groups_df": lambda m, site, code: m.update_memberships(
        memberships_df(site, []), nested_groups_df=nested_df(site.url).assign(library_code=code)),
    "grant_group_access": lambda m, site, code: m.grant_group_access(grants_df(site, [(code, "Staff", "read")])),
    "ensure_onedrive_shortcuts": lambda m, site, code: m.ensure_onedrive_shortcuts(
        memberships_df(site, [(code, "jo@contoso.com", "read")])),
}


# SharePoint drops '&' from a library's URL and trims spaces, so no later run would find the library
@pytest.mark.parametrize("code", ["D&T", "MATH "])
@pytest.mark.parametrize("operation", LIBRARY_CODE_OPERATIONS)
def test_library_codes_sharepoint_would_change_are_refused_before_anything_changes(
    make_manager, tenant, site, operation, code,
):
    with pytest.raises(ValueError, match="library_code"):
        LIBRARY_CODE_OPERATIONS[operation](make_manager(), site, code)

    assert tenant.calls == []


def fail_group_creates(tenant, monkeypatch, failure, times=1):
    """Makes the next *times* group creates fail with *failure*, after Graph may or may not have made the group."""
    import requests
    from alt_fabric_m365.sharepoint import _http
    from sharepoint_fake import graph_error, response

    failures = {
        "503": lambda: graph_error(503, "serviceNotAvailable"),
        "403": lambda: graph_error(403, "Authorization_RequestDenied"),
        # Not followed for a POST, and it points at what the POST made
        "303 to the new group": lambda: response(303, headers={"Location": "https://graph.microsoft.com/v1.0/groups/1"}),
    }
    if failure in failures:
        tenant.fail("POST", r"/groups$", failures[failure](), times=times)
        return
    exception = {"connection reset": requests.ConnectionError, "connect timeout": requests.ConnectTimeout}[failure]
    real_request = _http.requests.request
    failed = []

    def fail_group_create(method, url, **kwargs):
        if method == "POST" and url.endswith("/groups") and len(failed) < times:
            failed.append(url)
            raise exception(failure)
        return real_request(method, url, **kwargs)
    monkeypatch.setattr(_http.requests, "request", fail_group_create)


def make_failed_create_visible_only_by_id(tenant, name):
    """Adds the group a failed create made, which Graph's name lookups don't return yet."""
    made = tenant.add_group(name, owners={APP_ID})
    tenant.replicating.add(made["id"])
    return made


# A 5xx, dropped connection or unfollowed redirect can come after Graph made the group, and its name lookups can then
# miss it
@pytest.mark.parametrize("failure", ["503", "connection reset", "303 to the new group"])
def test_a_group_whose_create_may_have_succeeded_is_not_created_again_in_the_run(
    make_manager, tenant, site, monkeypatch, failure,
):
    tenant.add_user("jo@contoso.com")
    manager = make_manager()
    name = "BAA-Sharepoint-MATH-Auto-Read"
    fail_group_creates(tenant, monkeypatch, failure)

    results = manager.create_libraries(libraries_df(site, "MATH"))

    [error] = rows(results, "error").itertuples()
    assert error.document_library == "MATH"
    made = make_failed_create_visible_only_by_id(tenant, name)

    results = manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "read")]))

    [error] = rows(results, "error").itertuples()
    assert "may have been created by an earlier request this run" in error.status
    assert [g["id"] for g in tenant.groups.values() if g["displayName"] == name] == [made["id"]]


@pytest.mark.parametrize("operation", ["create_libraries", "update_memberships"])
def test_a_dry_run_reports_that_a_group_whose_create_may_have_succeeded_wont_be_created(
    make_manager, tenant, site, monkeypatch, operation,
):
    tenant.add_user("jo@contoso.com")
    manager = make_manager()
    fail_group_creates(tenant, monkeypatch, "503")
    manager.create_libraries(libraries_df(site, "MATH"))
    make_failed_create_visible_only_by_id(tenant, "BAA-Sharepoint-MATH-Auto-Read")
    writes = len(tenant.writes())

    results = (
        manager.create_libraries(libraries_df(site, "MATH"), dry_run=True) if operation == "create_libraries"
        else manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "read")]), dry_run=True)
    )

    [error] = rows(results, "error").itertuples()
    assert "may have been created by an earlier request this run" in error.status
    assert len(tenant.writes()) == writes


@pytest.mark.parametrize("failure, times", [
    ("403", 1),  # refused: Graph didn't make the group
    ("connect timeout", 6),  # never arrived, on the first attempt or any retry
])
def test_a_group_whose_create_certainly_failed_is_created_on_the_next_try(
    make_manager, tenant, site, monkeypatch, failure, times,
):
    tenant.add_user("jo@contoso.com")
    manager = make_manager()
    fail_group_creates(tenant, monkeypatch, failure, times=times)

    manager.create_libraries(libraries_df(site, "MATH"))
    manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "read")]))

    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Read") == ["jo@contoso.com"]


def test_a_group_is_created_on_the_next_try_after_its_token_couldnt_be_got(make_manager, tenant, site, monkeypatch):
    import requests
    from alt_fabric_m365.sharepoint import manager as manager_module
    tenant.add_user("jo@contoso.com")
    manager = make_manager()
    # The credential can't be reached for the token the first group create needs, so nothing is sent
    real_find, real_get_token, armed = manager_module._find_group_by_name, manager.credential.get_token, []

    def find_group_by_name(token, name):
        group = real_find(token, name)
        armed.append(group is None and len(armed) == 0)
        return group

    def get_token(*scopes, **kwargs):
        if armed and armed[-1]:
            armed[-1] = False
            raise requests.ConnectionError("Key Vault unreachable")
        return real_get_token(*scopes, **kwargs)
    monkeypatch.setattr(manager_module, "_find_group_by_name", find_group_by_name)
    monkeypatch.setattr(manager.credential, "get_token", get_token)

    results = manager.create_libraries(libraries_df(site, "MATH"))

    [error] = rows(results, "error").itertuples()
    assert "Key Vault unreachable" in error.status
    assert not [call for call in tenant.calls if call[0] == "POST" and call[1].endswith("/groups")]

    manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "read")]))

    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Read") == ["jo@contoso.com"]


@pytest.mark.parametrize("operation", ["update_memberships", "create_libraries"])
def test_a_group_made_earlier_in_the_run_is_not_made_again_while_it_replicates(make_manager, tenant, site, operation):
    manager = make_manager()
    manager.create_libraries(libraries_df(site, "MATH"))
    # Graph's name filters can miss groups created moments ago
    tenant.replicating.update(tenant.groups)
    tenant.add_user("jo@contoso.com")

    if operation == "update_memberships":
        results = manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))
        assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]
    else:
        results = manager.create_libraries(libraries_df(site, "MATH"))

    assert rows(results, "error").empty
    assert len(tenant.groups) == 2


def test_a_group_made_earlier_in_the_run_is_checked_for_duplicates_once_found_by_name(make_manager, tenant, site):
    manager = make_manager()
    manager.create_libraries(libraries_df(site, "MATH"))
    # Since replicated, alongside a group someone else gave the same name
    tenant.add_group("BAA-Sharepoint-MATH-Auto-Contribute", owners={"someone-else"})

    results = manager.create_libraries(libraries_df(site, "MATH"))

    assert "Multiple Entra groups share displayName" in rows(results, "error")["status"].iloc[0]


def test_a_group_made_earlier_in_the_run_follows_a_rename_once_found_by_name(make_manager, tenant, site):
    manager = make_manager()
    manager.create_libraries(libraries_df(site, "MATH"))

    results = manager.create_libraries(libraries_df(site, "MATH", library_display_name="Mathematics"))

    assert len(rows(results, "group_description_updated")) == 2
    assert "'Mathematics'" in tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["description"]


@pytest.mark.parametrize("owners, action", [(set(), "group_owner_added"), ({"someone-else"}, "group_not_owned")])
@pytest.mark.parametrize("operation", ["update_memberships", "create_libraries"])
def test_a_group_made_earlier_without_the_app_as_owner_is_checked_while_it_replicates(
    make_manager, tenant, site, operation, owners, action,
):
    from sharepoint_fake import graph_error
    manager = make_manager()
    manager.create_libraries(libraries_df(site, "MATH", app_owned=False))
    group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    group["owners"] = set(owners)
    tenant.replicating.update(tenant.groups)
    # Reading a group by ID can also miss it for a moment
    tenant.fail("GET", rf"/groups/{group['id']}/owners\?", graph_error(404, "Request_ResourceNotFound"))
    tenant.add_user("jo@contoso.com")

    if operation == "update_memberships":
        results = manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))
        assert tenant.member_upns(group["displayName"]) == ([] if owners else ["jo@contoso.com"])
    else:
        results = manager.create_libraries(libraries_df(site, "MATH"))

    assert group["displayName"] in set(rows(results, action)["group"])
    assert group["owners"] == (owners or {APP_ID})
    assert rows(results, "error").empty
    assert len(tenant.groups) == 2


@pytest.mark.parametrize("second_run, contribute, read", [
    ([("MATH", "sam@contoso.com", "contribute")], ["sam@contoso.com"], []),
    ([("MATH", "jo@contoso.com", "read")], [], ["jo@contoso.com"]),
], ids=["replaced", "moved"])
def test_a_group_made_earlier_in_the_run_is_synced_from_its_members_while_it_replicates(
    make_manager, tenant, site, second_run, contribute, read,
):
    from sharepoint_fake import graph_error
    manager = make_manager()
    manager.create_libraries(libraries_df(site, "MATH"))
    tenant.replicating.update(tenant.groups)
    tenant.add_user("jo@contoso.com")
    tenant.add_user("sam@contoso.com")
    manager.update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))
    writes = len(tenant.writes())

    preview = manager.update_memberships(memberships_df(site, second_run), dry_run=True)
    assert len(tenant.writes()) == writes
    # Reading a group by ID can also miss it for a moment
    tenant.fail("GET", r"/members\?", graph_error(404, "Request_ResourceNotFound"))
    results = manager.update_memberships(memberships_df(site, second_run))
    writes = len(tenant.writes())
    manager.update_memberships(memberships_df(site, second_run))

    for log in (preview, results):
        assert rows(log, "error").empty
    assert list(rows(preview, "remove_from_group (dry run)")["user"]) == ["jo@contoso.com"]
    assert list(rows(results, "removed_from_group")["user"]) == ["jo@contoso.com"]
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == contribute
    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Read") == read
    assert len(tenant.writes()) == writes


def test_group_differing_only_in_case_is_reused_not_duplicated(make_manager, tenant, site, library):
    tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["displayName"] = "baa-sharepoint-MATH-Auto-Contribute"
    tenant.add_user("jo@contoso.com")

    make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    assert len([g for g in tenant.groups.values() if g["displayName"].lower().endswith("-math-auto-contribute")]) == 1



# Code review fixes

def test_group_creation_is_not_repeated_after_a_server_error(make_manager, tenant, site):
    from sharepoint_fake import response
    tenant.fail("POST", r"/v1.0/groups$", response(504))

    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["error", "library_created"]
    assert len(created(tenant, "/v1.0/groups")) == 1  # not retried, so no duplicate can be created
    # The failure stopped the first run before it set the display name, so the re-run sets it
    assert actions(make_manager().create_libraries(libraries_df(site, "MATH"))) == [
        "group_assigned", "group_assigned", "library_renamed",
    ]


def test_unlinking_a_library_that_never_existed_is_a_warning(make_manager, tenant, site):
    jo = tenant.add_user("jo@contoso.com")
    group = tenant.add_group("BAA-Sharepoint-MATH-Auto-Contribute")
    group["members"][jo["id"]] = jo

    results = make_manager().unlink_libraries(libraries_df(site, "MATH"))

    assert actions(results) == ["group_emptied", "group_owner_added", "library_not_found", "manual_groups_preserved"]
    assert not set(results["action"]) & FATAL_ACTIONS
    assert group["members"] == {}


def test_new_code_colliding_with_an_existing_library_is_skipped(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "PE-KS3"))
    tenant.add_user("jo@contoso.com")

    results = make_manager().create_libraries(libraries_df(site, "PEKS3"))

    [row] = results.itertuples()
    assert (row.action, row.document_library) == ("group_name_collision", "PEKS3")
    assert site.list_by_slug("PEKS3") is None

    results = make_manager().update_memberships(memberships_df(site, [("PEKS3", "jo@contoso.com", "contribute")]))

    assert actions(results) == ["group_name_collision"]
    assert tenant.member_upns("BAA-Sharepoint-PEKS3-Auto-Contribute") == []


def test_code_matching_another_librarys_title_is_never_adopted(make_manager, tenant, site):
    artd = site.add_list("Art", slug="ARTD", unique=True)
    tenant.add_user("jo@contoso.com")

    results = make_manager().create_libraries(libraries_df(site, "ART", library_display_name="Art & Design"))

    assert actions(results) == ["library_name_conflict"]
    assert artd["displayName"] == "Art" and artd["role_assignments"] == {}

    results = make_manager().update_memberships(memberships_df(site, [("ART", "jo@contoso.com", "contribute")]))

    assert actions(results) == ["library_not_found"]
    assert artd["role_assignments"] == {}


def test_members_of_a_new_group_are_not_read(make_manager, tenant, site, library):
    tenant.add_user("head@contoso.com")

    make_manager().update_memberships(memberships_df(site, [("MATH", "head@contoso.com", "owner")]))

    owner = tenant.group_named("BAA-Sharepoint-MATH-Auto-Owner")
    assert not any(f"/groups/{owner['id']}/members?" in url for method, url, _ in tenant.calls if method == "GET")


@pytest.mark.parametrize("operation", ["create", "update"])
def test_one_unreadable_site_does_not_stop_the_others(make_manager, tenant, site, operation):
    tenant.add_user("jo@contoso.com")
    gone = "https://contoso.sharepoint.com/sites/Typo"
    manager = make_manager()
    if operation == "create":
        df = pd.concat([libraries_df(site, "MATH"), libraries_df(site, "ENG").assign(site_url=gone)])
        results = manager.create_libraries(df)
        assert site.list_by_slug("MATH") is not None
    else:
        manager.create_libraries(libraries_df(site, "MATH"))
        df = memberships_df(site, [("MATH", "jo@contoso.com", "contribute"), ("ENG", "jo@contoso.com", "read")])
        df.loc[1, "site_url"] = gone
        results = manager.update_memberships(df)
        assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com"]

    [error] = rows(results, "error").itertuples()
    assert error.site_url == gone


def test_one_unresolvable_group_does_not_skip_the_sites_other_grants(make_manager, tenant, site, library, slt_groups):
    from sharepoint_fake import response
    read_group = tenant.group_named("BAA-Sharepoint-SLT-Manual-Read")
    tenant.fail("POST", r"/web/ensureuser$", response(400), times=1)  # the first group looked up
    df = grants_df(site, [
        ("MATH", "BAA-Sharepoint-SLT-Manual-Read", "read"),
        ("MATH", "BAA-Sharepoint-SLT-Manual-Contribute", "contribute"),
    ])
    site.principals.pop(f"c:0t.c|tenant|{read_group['id']}", None)

    results = make_manager().grant_group_access(df)

    assert sorted(actions(results)) == ["error", "group_access_granted"]


def test_a_group_that_cant_be_registered_on_a_site_is_tried_once_for_all_its_libraries(
    make_manager, tenant, site, library, slt_groups,
):
    from sharepoint_fake import response
    make_manager().create_libraries(libraries_df(site, "ENG"))
    tenant.fail("POST", r"/web/ensureuser$", response(400), times=1)
    tenant.calls.clear()

    results = make_manager().grant_group_access(grants_df(site, [
        (code, "BAA-Sharepoint-SLT-Manual-Read", "read") for code in ("MATH", "ENG")
    ]))

    # Each library still gets its own error row, without registering the group again
    assert rows(results, "error")["document_library"].tolist() == ["MATH", "ENG"]
    assert len(created(tenant, "/web/ensureuser")) == 1


def test_each_user_is_resolved_once_per_sync(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH", "ENG", "ART"))
    tenant.add_user("jo.smith@contoso.onmicrosoft.com", mail="jo@contoso.com")
    tenant.calls.clear()

    make_manager().update_memberships(memberships_df(site, [
        (code, "jo@contoso.com", "contribute") for code in ("MATH", "ENG", "ART")
    ]))

    lookups = [url for method, url, _ in tenant.calls if "/v1.0/users" in url and method == "GET"]
    assert len(lookups) == 2  # UPN miss, then mail match — once, not once per library
    assert tenant.member_upns("BAA-Sharepoint-ART-Auto-Contribute") == ["jo.smith@contoso.onmicrosoft.com"]


# Group ownership

def test_new_groups_are_owned_by_the_app(make_manager, tenant, site):
    tenant.add_user("head@contoso.com")
    make_manager().create_libraries(libraries_df(site, "MATH", manual_read=True))
    make_manager().update_memberships(memberships_df(site, [("MATH", "head@contoso.com", "owner")]))

    assert {g["displayName"]: g["owners"] for g in tenant.groups.values()} == {
        "BAA-Sharepoint-MATH-Auto-Read": {APP_ID},
        "BAA-Sharepoint-MATH-Auto-Contribute": {APP_ID},
        "BAA-Sharepoint-MATH-Manual-Read": {APP_ID},
        "BAA-Sharepoint-MATH-Auto-Owner": {APP_ID},
    }


def test_a_same_named_group_someone_else_owns_is_never_granted_access(make_manager, tenant, site):
    squatter = tenant.add_user("student@contoso.com")
    squatted = tenant.add_group("BAA-Sharepoint-DRAMA-Manual-Contribute", owners={squatter["id"]})

    results = make_manager().create_libraries(libraries_df(site, "DRAMA", manual_contribute=True))

    [row] = rows(results, "group_not_owned").itertuples()
    assert row.group == "BAA-Sharepoint-DRAMA-Manual-Contribute"
    assert "group_not_owned" in WARNING_ACTIONS
    lib = site.list_by_slug("DRAMA")
    assert f"c:0t.c|tenant|{squatted['id']}" not in site.principals
    assert len(lib["role_assignments"]) == 2  # the two Auto groups, created by the app
    assert squatted["owners"] == {squatter["id"]}


def test_a_squatted_auto_group_is_not_synced(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    squatter = tenant.add_user("student@contoso.com")
    squatted = tenant.add_group("BAA-Sharepoint-MATH-Auto-Contribute", owners={squatter["id"]})
    tenant.add_user("jo@contoso.com")

    results = make_manager().update_memberships(memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]))

    assert actions(results) == ["group_not_owned"]
    assert squatted["members"] == {}
    assert site.list_by_slug("MATH")["role_assignments"] == {}


def test_an_ownerless_auto_group_with_no_members_to_add_is_adopted(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")  # created app-only by an earlier version
    tenant.add_user("jo@contoso.com")
    df = memberships_df(site, [("MATH", "jo@contoso.com", "contribute")])

    results = make_manager().update_memberships(df)

    assert rows(results, "group_owner_added")["group"].tolist() == ["BAA-Sharepoint-MATH-Auto-Read"]
    assert legacy["owners"] == {APP_ID}
    assert "group_owner_added" not in actions(make_manager().update_memberships(df))


def test_membership_sync_never_empties_a_group_the_app_doesnt_own(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    owner = tenant.add_user("someone@contoso.com")
    member = tenant.add_user("member@contoso.com")
    theirs = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read", owners={owner["id"]})
    theirs["members"][member["id"]] = member
    empty = memberships_df(site, [])
    libraries = libraries_df(site, "MATH").drop(columns="library_display_name")

    for dry_run in (True, False):
        results = make_manager(dry_run=dry_run).update_memberships(empty, libraries_df=libraries)
        assert actions(results) == ["group_not_owned"]

    assert set(theirs["members"]) == {member["id"]}


def test_an_ownerless_group_is_adopted(make_manager, tenant, site):
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")  # created app-only by an earlier version

    results = make_manager().create_libraries(libraries_df(site, "MATH"))

    assert rows(results, "group_owner_added")["group"].tolist() == ["BAA-Sharepoint-MATH-Auto-Read"]
    assert legacy["owners"] == {APP_ID}
    assert f"c:0t.c|tenant|{legacy['id']}" in site.principals
    assert "group_owner_added" not in actions(make_manager().create_libraries(libraries_df(site, "MATH")))


@pytest.mark.parametrize("dry_run", [False, True])
def test_a_duplicated_name_of_a_group_people_own_is_an_error(make_manager, tenant, site, dry_run):
    tenant.add_group("BAA-Sharepoint-SLT-Manual-Read")
    tenant.add_group("BAA-Sharepoint-SLT-Manual-Read")
    df = libraries_df(site, "SLT", auto_groups=False, manual_read=True, app_owned=False)

    results = make_manager(dry_run=dry_run).create_libraries(df)

    assert list(rows(results, "error")["status"].str.contains("Multiple")) == [True]


def test_groups_people_own_are_left_alone_when_not_app_owned(make_manager, tenant, site):
    head = tenant.add_user("head@contoso.com")
    slt_read = tenant.add_group("BAA-Sharepoint-SLT-Manual-Read", owners={head["id"]})
    df = libraries_df(site, "SLT", auto_groups=False, manual_read=True, manual_contribute=True, app_owned=False)

    results = make_manager().create_libraries(df)

    assert not set(results["action"]) & {"group_not_owned", "group_owner_added"}
    assert slt_read["owners"] == {head["id"]}
    assert tenant.group_named("BAA-Sharepoint-SLT-Manual-Contribute")["owners"] == set()
    assert len(site.list_by_slug("SLT")["role_assignments"]) == 2


def test_dry_run_reports_ownership_without_changing_it(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    site.add_list("ENG", unique=True)
    other = tenant.add_user("someone@contoso.com")
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Auto-Contribute")
    tenant.add_group("BAA-Sharepoint-ENG-Auto-Contribute", owners={other["id"]})
    tenant.add_user("jo@contoso.com")
    tenant.calls.clear()

    results = make_manager(dry_run=True).update_memberships(memberships_df(site, [
        ("MATH", "jo@contoso.com", "contribute"), ("ENG", "jo@contoso.com", "contribute"),
    ]))

    assert tenant.writes() == []
    assert legacy["owners"] == set()
    assert results[["document_library", "action"]].values.tolist() == [
        ["MATH", "group_owner_added (dry run)"], ["MATH", "group_assigned (dry run)"],
        ["MATH", "add_to_group (dry run)"], ["ENG", "group_not_owned"],
    ]


def test_create_dry_run_reports_ownership_without_changing_it(make_manager, tenant, site):
    other = tenant.add_user("someone@contoso.com")
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Manual-Read")
    tenant.add_group("BAA-Sharepoint-MATH-Manual-Contribute", owners={other["id"]})
    tenant.calls.clear()

    results = make_manager(dry_run=True).create_libraries(
        libraries_df(site, "MATH", auto_groups=False, manual_read=True, manual_contribute=True),
    )

    assert tenant.writes() == []
    assert legacy["owners"] == set()
    assert results[["group", "action"]].values.tolist() == [
        ["", "library_created (dry run)"],
        ["BAA-Sharepoint-MATH-Manual-Read", "group_owner_added (dry run)"],
        ["BAA-Sharepoint-MATH-Manual-Read", "ensure_manual_read_group (dry run)"],
        # Adopting it also gives it the description a live run sets
        ["BAA-Sharepoint-MATH-Manual-Read", "group_description_updated (dry run)"],
        ["BAA-Sharepoint-MATH-Manual-Contribute", "group_not_owned"],
    ]


def test_dry_run_reports_adopting_a_group_it_would_empty(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")
    jo = tenant.add_user("jo@contoso.com")
    legacy["members"][jo["id"]] = jo
    tenant.calls.clear()

    results = make_manager(dry_run=True).update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH").drop(columns="library_display_name"),
    )

    assert tenant.writes() == []
    assert legacy["owners"] == set()
    assert results["action"].tolist() == ["group_owner_added (dry run)", "remove_from_group (dry run)"]


def test_renaming_a_library_keeps_group_names_and_updates_descriptions(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH", library_display_name="Maths", manual_read=True))
    names = sorted(g["displayName"] for g in tenant.groups.values())

    results = make_manager().create_libraries(
        libraries_df(site, "MATH", library_display_name="Mathematics", manual_read=True),
    )

    assert sorted(g["displayName"] for g in tenant.groups.values()) == names
    assert len(rows(results, "group_description_updated")) == 3
    assert all("'Mathematics' (MATH)" in g["description"] for g in tenant.groups.values())
    assert "NOT synced" in tenant.group_named("BAA-Sharepoint-MATH-Manual-Read")["description"]

    tenant.calls.clear()
    make_manager().create_libraries(libraries_df(site, "MATH", library_display_name="Mathematics", manual_read=True))
    assert created(tenant, "/v1.0/groups/", method="PATCH") == []


def test_a_rename_dry_run_reports_the_descriptions_it_would_update(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH", library_display_name="Maths", manual_read=True))
    tenant.calls.clear()

    renamed = make_manager(dry_run=True).create_libraries(
        libraries_df(site, "MATH", library_display_name="Mathematics", manual_read=True),
    )
    unchanged = make_manager(dry_run=True).create_libraries(
        libraries_df(site, "MATH", library_display_name="Maths", manual_read=True),
    )

    assert tenant.writes() == []
    assert sorted(rows(renamed, "group_description_updated (dry run)")["group"]) == [
        "BAA-Sharepoint-MATH-Auto-Contribute", "BAA-Sharepoint-MATH-Auto-Read", "BAA-Sharepoint-MATH-Manual-Read",
    ]
    assert "group_description_updated (dry run)" not in actions(unchanged)


def test_descriptions_of_groups_people_own_are_left_alone(make_manager, tenant, site):
    head = tenant.add_user("head@contoso.com")
    slt_read = tenant.add_group("BAA-Sharepoint-SLT-Manual-Read", owners={head["id"]})
    slt_read["description"] = "Curated by the headteacher"
    df = libraries_df(site, "SLT", library_display_name="Senior Leadership", auto_groups=False,
                      manual_read=True, app_owned=False)

    results = make_manager().create_libraries(df)

    assert "group_description_updated" not in actions(results)
    assert slt_read["description"] == "Curated by the headteacher"



# Copilot re-review fixes

def test_an_ordinary_list_with_the_code_as_its_url_is_not_adopted(make_manager, tenant, site):
    tasks = site.add_list("English tasks", slug="ENG", template="genericList", unique=False)

    make_manager().create_libraries(libraries_df(site, "ENG", library_display_name="English"))

    assert tasks["displayName"] == "English tasks" and tasks["unique"] is False
    library = site.list_by_slug("ENG")
    assert library["template"] == "documentLibrary" and library["displayName"] == "English"


def test_a_onedrive_shortcut_is_not_repeated_after_a_server_error(make_manager, tenant, site, library):
    from sharepoint_fake import response
    tenant.add_user("jo@contoso.com")
    tenant.fail("POST", r"/drive/root/children$", response(503))

    results = make_manager(create_onedrive_shortcuts=True).update_memberships(
        memberships_df(site, [("MATH", "jo@contoso.com", "contribute")]),
    )

    assert len(created(tenant, "/drive/root/children")) == 1
    assert "onedrive_shortcut_failed" in actions(results)


def test_unlink_never_empties_a_group_the_app_doesnt_own(make_manager, tenant, site):
    owner = tenant.add_user("someone@contoso.com")
    member = tenant.add_user("member@contoso.com")
    theirs = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read", owners={owner["id"]})
    theirs["members"][member["id"]] = member

    for dry_run in (True, False):
        results = make_manager(dry_run=dry_run).unlink_libraries(libraries_df(site, "MATH"))
        assert "group_not_owned" in actions(results)

    assert set(theirs["members"]) == {member["id"]}


def test_retire_waits_for_the_library_on_the_new_site(make_manager, tenant, site):
    new_site = tenant.add_site("/sites/BAANew")
    make_manager().create_libraries(libraries_df(site, "MATH"))
    df = libraries_df(site, "MATH").assign(new_site_url=new_site.url)

    results = make_manager().retire_moved_libraries(df)

    assert actions(results) == ["retire_deferred"]
    assert "retire_deferred" in WARNING_ACTIONS
    assert site.list_by_slug("MATH")["displayName"] == "MATH Faculty"  # left working

    make_manager().create_libraries(libraries_df(new_site, "MATH"))
    results = make_manager().retire_moved_libraries(df)

    assert "library_retired" in actions(results)
    assert site.list_by_slug("MATH")["displayName"] == "MATH_MOVED"


@pytest.mark.parametrize("replacement", ["still inherits", "move-back leftover", "limited access", "duplicate group"])
def test_retire_waits_for_a_working_replacement(make_manager, tenant, site, replacement):
    new_site = tenant.add_site("/sites/BAANew")
    make_manager().create_libraries(libraries_df(site, "MATH"))
    contribute = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    if replacement == "still inherits":
        new_site.add_list("MATH")  # its create failed after the library was made
    elif replacement == "move-back leftover":
        new_site.add_list("MATH_MOVED", slug="MATH", unique=True)  # retired by an earlier move, no group access
    else:
        new_library = new_site.add_list("MATH", unique=True)
        new_site.principals[f"c:0t.c|tenant|{contribute['id']}"] = 77
        if replacement == "limited access":
            # Only from a folder inside it with its own permissions, so no access to the library
            new_library["role_assignments"][77] = {ROLE_DEFINITIONS["Limited Access"]}
        else:
            new_library["role_assignments"][77] = {CONTRIBUTE}
            tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")  # a second group with the library's group name
    tenant.calls.clear()

    results = make_manager().retire_moved_libraries(libraries_df(site, "MATH").assign(new_site_url=new_site.url))

    assert actions(results) == ["retire_deferred"]
    assert tenant.writes() == []
    library = site.list_by_slug("MATH")
    assert (library["displayName"], len(library["role_assignments"])) == ("MATH Faculty", 2)


def test_retire_without_a_new_site_retires_anyway(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))

    results = make_manager().retire_moved_libraries(libraries_df(site, "MATH").assign(new_site_url=None))

    assert "library_retired" in actions(results)
