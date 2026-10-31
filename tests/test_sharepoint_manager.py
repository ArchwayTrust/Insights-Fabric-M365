"""
SharePointLibraryManager operations, run end to end against the in-memory FakeTenant.

Each test checks the action log the operation returns and what it actually changed, and where it matters that
running it again changes nothing and that a dry run makes no writes.
"""
import time

import pandas as pd
import pytest

import alt_fabric_m365
from alt_fabric_m365.auth import Scopes
from alt_fabric_m365.sharepoint import (
    ACTION_LOG_COLUMNS, FATAL_ACTIONS, WARNING_ACTIONS, SharePointLibraryManager,
)

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


# create_libraries

def test_create_library_with_groups_metadata_and_navigation(make_manager, tenant, site):
    df = libraries_df(site, "MATH", library_display_name="Maths", manual_read=True, manual_contribute=True,
                      source_system="Arbor", source_id="F1")

    results = make_manager(add_to_navigation=True).create_libraries(df)

    assert list(results.columns) == ACTION_LOG_COLUMNS
    assert actions(results) == sorted([
        "navigation_label_created", "metadata_set", "metadata_set", "group_assigned", "group_assigned",
        "manual_read_group_ensured", "manual_contribute_group_ensured", "navigation_link_added",
    ])

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
    df = libraries_df(site, "MATH", manual_read=True)
    manager.create_libraries(df)
    tenant.calls.clear()

    results = manager.create_libraries(df)

    assert "navigation_label_created" not in actions(results)
    assert "navigation_link_added" not in actions(results)
    assert actions(results).count("already_assigned") == 3
    assert created(tenant, "/web/ensureuser") == [] and created(tenant, "addroleassignment") == []
    assert created(tenant, "/groups") == []
    assert created(tenant, f"/sites/{site.id}/lists") == []
    assert created(tenant, "breakroleinheritance") == []
    assert len(site.lists) == 1 and len(site.nav[0]["Children"]) == 1


def test_create_library_renames_display_name_and_navigation_link(make_manager, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH", library_display_name="Maths"))

    results = manager.create_libraries(libraries_df(site, "MATH", library_display_name="Mathematics"))

    assert site.list_by_slug("MATH")["displayName"] == "Mathematics"
    assert site.nav[0]["Children"][0]["Title"] == "Mathematics"
    assert "navigation_link_renamed" in actions(results)


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
        {"site_url": site.url, "library_code": "MATH", "source_system": None, "source_id": "F1"},
        {"site_url": site.url, "library_code": "ENG", "source_system": None, "source_id": None},
    ]


def test_set_metadata_needs_a_metadata_column(make_manager, site):
    with pytest.raises(ValueError, match="source_system' or 'source_id"):
        make_manager().set_library_metadata(pd.DataFrame({"site_url": [site.url], "library_code": ["MATH"]}))


def test_set_metadata_on_a_missing_library_is_an_error_row(make_manager, site):
    results = make_manager().set_library_metadata(pd.DataFrame({
        "site_url": [site.url], "library_code": ["MATH"], "source_system": ["Arbor"],
    }))

    assert actions(results) == ["error"]


# rename_libraries

def test_rename_with_a_changed_code_renames_the_groups(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATHS", manual_read=True))
    # The faculty's code changed from MATHS to MATH. The library's URL slug is set at creation and never
    # changes, so here the library is found by the new code (as if it had been created with it).
    site.list_by_slug("MATHS")["slug"] = "MATH"
    df = libraries_df(site, "MATH", library_display_name="Maths", old_library_code="MATHS")

    results = make_manager().rename_libraries(df)

    assert actions(results) == ["group_renamed"] * 3 + ["library_renamed"]
    assert sorted(g["displayName"] for g in tenant.groups.values()) == [
        "BAA-Sharepoint-MATH-Auto-Contribute", "BAA-Sharepoint-MATH-Auto-Read", "BAA-Sharepoint-MATH-Manual-Read",
    ]
    assert site.list_by_slug("MATH")["displayName"] == "Maths"


def test_rename_of_a_missing_library_is_an_error_and_leaves_groups(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATHS"))

    results = make_manager().rename_libraries(libraries_df(site, "MATH", old_library_code="MATHS"))

    assert actions(results) == ["error"]
    assert tenant.group_named("BAA-Sharepoint-MATHS-Auto-Read") is not None


def test_rename_syncs_the_navigation_title(make_manager, site):
    manager = make_manager(add_to_navigation=True)
    manager.create_libraries(libraries_df(site, "MATH"))

    results = manager.rename_libraries(libraries_df(site, "MATH", library_display_name="Mathematics"))

    assert actions(results) == ["library_renamed", "navigation_link_renamed"]
    assert site.nav[0]["Children"][0]["Title"] == "Mathematics"


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

    assert rows(results, "skip")["status"].tolist() == ["unknown permission 'edit'"]


def test_nested_groups_are_added_and_unresolvable_ones_reported(make_manager, tenant, site, library):
    all_staff = tenant.add_group("BAA All Staff")
    old_nested = tenant.add_group("Old Nested")
    reader = tenant.add_user("reader@contoso.com")
    read_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")
    contribute_group = tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")
    read_group["members"] = {reader["id"]: reader}
    contribute_group["members"] = {old_nested["id"]: old_nested}
    nested = pd.DataFrame({
        "site_url": site.url, "library_code": "MATH", "group_prefix": PREFIX,
        "member_group": ["BAA All Staff", "Missing Group"], "permission": ["read", "contribute"],
    })

    results = make_manager().update_memberships(memberships_df(site, []), nested_groups_df=nested)

    assert set(read_group["members"]) == {all_staff["id"]}  # stale user removed, group nested
    # Contribute's nested-group lookup failed, so its stale nested group is kept
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
    assert actions(results) == ["grant_group_access (dry run)"]


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

    assert actions(results) == ["error", "group_assigned", "group_assigned"]
    assert site.nav == []


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


def test_memberships_dry_run_lists_shortcuts_nested_groups_and_reconciles(make_manager, tenant, site, library):
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
        "ensure_group (dry run)", "add_to_group (dry run)", "onedrive_shortcut (dry run)",  # owner
        "ensure_group (dry run)", "add_to_group (dry run)",  # read, nested group
        "reconcile_group (dry run)",  # contribute exists but has no desired members
    ])


def test_rename_dry_run_with_a_changed_code(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATHS"))
    tenant.calls.clear()

    results = make_manager(dry_run=True, add_to_navigation=True).rename_libraries(
        libraries_df(site, "MATH", old_library_code="MATHS"),
    )

    assert tenant.writes() == []
    assert actions(results) == sorted([
        "rename_library (dry run)", "rename_group (dry run)", "rename_group (dry run)",
        "sync_navigation_link (dry run)",
    ])


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

    assert actions(results) == ["error"]
    assert len(created(tenant, "/v1.0/groups")) == 1  # not retried, so no duplicate can be created
    assert actions(make_manager().create_libraries(libraries_df(site, "MATH"))) == ["group_assigned"] * 2


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


def test_rename_never_renames_onto_an_existing_group(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATHS"))
    site.list_by_slug("MATHS")["slug"] = "MATH"
    tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")  # already ensured for the new code

    results = make_manager().rename_libraries(libraries_df(site, "MATH", old_library_code="MATHS"))

    assert sorted(actions(results)) == ["group_name_collision", "group_renamed", "library_renamed"]
    assert tenant.group_named("BAA-Sharepoint-MATHS-Auto-Read") is not None
    assert tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute") is not None


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
        ["MATH", "group_owner_added (dry run)"], ["MATH", "ensure_group (dry run)"], ["MATH", "add_to_group (dry run)"],
        ["ENG", "group_not_owned"],
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
        ["BAA-Sharepoint-MATH-Manual-Read", "group_owner_added (dry run)"],
        ["BAA-Sharepoint-MATH-Manual-Read", "ensure_manual_read_group (dry run)"],
        ["BAA-Sharepoint-MATH-Manual-Contribute", "group_not_owned"],
    ]


def test_dry_run_reports_adopting_a_group_it_would_empty(make_manager, tenant, site):
    site.add_list("MATH", unique=True)
    legacy = tenant.add_group("BAA-Sharepoint-MATH-Auto-Read")
    tenant.calls.clear()

    results = make_manager(dry_run=True).update_memberships(
        memberships_df(site, []), libraries_df=libraries_df(site, "MATH").drop(columns="library_display_name"),
    )

    assert tenant.writes() == []
    assert legacy["owners"] == set()
    assert results["action"].tolist() == ["group_owner_added (dry run)", "reconcile_group (dry run)"]


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


def test_retire_without_a_new_site_retires_anyway(make_manager, tenant, site):
    make_manager().create_libraries(libraries_df(site, "MATH"))

    results = make_manager().retire_moved_libraries(libraries_df(site, "MATH").assign(new_site_url=None))

    assert "library_retired" in actions(results)
