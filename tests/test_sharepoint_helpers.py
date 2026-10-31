"""
The SharePoint module's helpers: naming, input validation, lookups, user and group resolution, the property bag
and navigation URL matching. Requests go to the in-memory FakeTenant.
"""
import threading
import time

import pandas as pd
import pytest
import requests

from alt_fabric_m365.sharepoint import _common, _graph, _groups, _navigation, _sharepoint
from alt_fabric_m365.sharepoint import build_group_name, sanitise_for_group_name
from alt_fabric_m365.sharepoint._groups import GroupResolutionError, UserResolutionError

from sharepoint_fake import GRAPH, FakeTenant, graph_error, response

TOKEN = "token"
SITE = "https://contoso.sharepoint.com/sites/maths"


@pytest.fixture
def tenant(monkeypatch):
    monkeypatch.setattr(time, "sleep", lambda seconds: None)
    return FakeTenant().install(monkeypatch)


# Group naming

def test_sanitise_keeps_only_letters_and_digits():
    assert sanitise_for_group_name("Policy Documents & PE-KS3") == "PolicyDocumentsPEKS3"


def test_build_group_name_with_and_without_site_name():
    assert build_group_name("SP", "https://x.sharepoint.com/sites/finance-team/", "Policy Docs", "Auto-Read") == (
        "SP-financeteam-PolicyDocs-Auto-Read"
    )
    assert build_group_name("SP", SITE, "Policy Docs", "Auto-Read", include_site_name=False) == "SP-PolicyDocs-Auto-Read"


@pytest.mark.parametrize("display_name, nickname", [
    ("BAA-Sharepoint-MATH-Auto-Read", "baasharepointmathautoread"),
    ("---", "group"),
    ("A" * 100, "a" * 64),
])
def test_mail_nickname_is_graph_valid(display_name, nickname):
    assert _groups._build_mail_nickname(display_name) == nickname


def test_odata_escape_doubles_single_quotes():
    assert _groups._odata_escape("O'Neil's") == "O''Neil''s"


def test_group_name_collisions_are_reported_both_ways():
    prefixes = {
        (SITE, "PE-KS3"): "BAA-Sharepoint",
        (SITE, "PEKS3"): "BAA-Sharepoint",
        (SITE, "ENG"): "BAA-Sharepoint",
    }

    collisions = _groups._find_group_name_collisions(prefixes, include_site_name=False)

    assert collisions == {(SITE, "PEKS3"): (SITE, "PE-KS3"), (SITE, "PE-KS3"): (SITE, "PEKS3")}


def test_site_name_separates_otherwise_colliding_codes():
    other_site = "https://contoso.sharepoint.com/sites/science"
    prefixes = {(SITE, "ENG"): "SP", (other_site, "ENG"): "SP"}

    assert _groups._find_group_name_collisions(prefixes, include_site_name=True) == {}
    assert len(_groups._find_group_name_collisions(prefixes, include_site_name=False)) == 2


def test_collisions_are_logged_as_warning_rows():
    log = _common._ThreadSafeLog()
    _groups._log_group_name_collisions({(SITE, "A-1"): (SITE, "A1")}, log)

    [row] = log.to_list()
    assert row["action"] == "group_name_collision"
    assert row["document_library"] == "A-1"


# Validation

def test_required_columns_must_be_present():
    with pytest.raises(ValueError, match="missing required columns"):
        _common._validate_required_columns(pd.DataFrame({"site_url": [SITE]}), {"site_url", "library_code"})


def test_required_columns_must_not_be_null():
    df = pd.DataFrame({"site_url": [SITE, None]})
    with pytest.raises(ValueError, match="1 null value"):
        _common._validate_required_columns(df, {"site_url"})


def test_required_columns_are_cast_to_strings_on_a_copy():
    df = pd.DataFrame({"library_code": [101, 102]})

    validated = _common._validate_required_columns(df, {"library_code"})

    assert list(validated["library_code"]) == ["101", "102"]
    assert list(df["library_code"]) == [101, 102]


@pytest.mark.parametrize("url", ["contoso.sharepoint.com/sites/maths", "ftp://contoso.sharepoint.com/x", "https://"])
def test_malformed_site_urls_are_rejected(url):
    with pytest.raises(ValueError, match="Malformed site_url"):
        _common._validate_site_urls([url])


@pytest.mark.parametrize("max_workers", [1, 4])
def test_run_parallel_logs_failures_and_carries_on(max_workers):
    done = []
    lock = threading.Lock()

    def work(name, *, action_log):
        if name == "bad":
            raise RuntimeError("boom")
        with lock:
            done.append(name)

    log = _common._ThreadSafeLog()
    _common._run_parallel(work, [("a",), ("bad",), ("b",)], max_workers, log)

    assert sorted(done) == ["a", "b"]
    assert [(r["action"], r["status"]) for r in log.to_list()] == [("error", "boom")]


def test_empty_action_log_has_the_standard_columns():
    df = _common._action_log_to_dataframe(_common._ThreadSafeLog())

    assert list(df.columns) == _common.ACTION_LOG_COLUMNS
    assert df.empty


# Libraries

def test_lists_are_keyed_by_url_slug_only(tenant):
    site = tenant.add_site("/sites/maths")
    slt = site.add_list("SLT")
    titled_art = site.add_list("Art", slug="ARTD")  # titled like another faculty's code
    spaced = site.add_list("Year 7 & 8", slug="Y7 & 8")

    lists = _graph._list_existing_lists(TOKEN, site.id)

    assert sorted(lists) == ["artd", "slt", "y7 & 8"]  # slug is percent-decoded
    assert (lists["slt"]["id"], lists["artd"]["id"], lists["y7 & 8"]["id"]) == (slt["id"], titled_art["id"], spaced["id"])
    assert _graph._find_title_conflict(lists, "ART")["id"] == titled_art["id"]
    assert _graph._find_title_conflict(lists, "SLT") is None  # its own title


def test_lists_follow_pages(tenant):
    site = tenant.add_site("/sites/maths")
    for code in ("A", "B", "C", "D", "E"):
        site.add_list(code)
    tenant.list_page_size = 2

    assert sorted(_graph._list_existing_lists(TOKEN, site.id)) == ["a", "b", "c", "d", "e"]


def test_missing_library_raises(tenant):
    site = tenant.add_site("/sites/maths")
    with pytest.raises(ValueError, match="Library 'ENG' not found"):
        _graph._get_list_id(TOKEN, site.id, "ENG")


def test_ensure_library_creates_it_titled_with_its_code_once(tenant):
    site = tenant.add_site("/sites/maths")

    lib = _graph._ensure_document_library(TOKEN, site.id, "ENG")
    again = _graph._ensure_document_library(TOKEN, site.id, "ENG")

    created = site.lists[lib["id"]]
    assert (created["slug"], created["displayName"]) == ("ENG", "ENG")
    assert again["id"] == lib["id"] and len(site.lists) == 1


def test_title_conflict_with_a_display_name(tenant):
    site = tenant.add_site("/sites/maths")
    english = site.add_list("English", slug="EnglishOld")
    site.add_list("Maths", slug="MATH")

    lists = _graph._list_existing_lists(TOKEN, site.id)

    assert _graph._find_title_conflict(lists, "ENG", title="english")["id"] == english["id"]
    assert _graph._find_title_conflict(lists, "MATH", title="Maths") is None  # its own title


# Groups and members

def test_duplicate_group_names_are_refused(tenant):
    tenant.add_group("SP-ENG-Auto-Read")
    tenant.add_group("SP-ENG-Auto-Read")

    with pytest.raises(ValueError, match="Multiple Entra groups share displayName"):
        _groups._find_group_by_name(TOKEN, "SP-ENG-Auto-Read")


def test_prefix_lookup_returns_all_matches_and_unique_lookup_refuses_duplicates(tenant):
    tenant.add_group("SP-ENG-Auto-Read")
    tenant.add_group("SP-ENG-Manual-Read")
    tenant.add_group("SP-ENG-Manual-Read")
    tenant.add_group("SP-MATH-Auto-Read")

    groups = _groups._find_groups_by_name_prefix(TOKEN, "SP-ENG-")

    assert sorted(groups) == ["sp-eng-auto-read", "sp-eng-manual-read"]
    assert _groups._get_unique_group(groups, "SP-ENG-Auto-Read")["displayName"] == "SP-ENG-Auto-Read"
    assert _groups._get_unique_group(groups, "SP-ENG-Auto-Owner") is None
    with pytest.raises(ValueError):
        _groups._get_unique_group(groups, "SP-ENG-Manual-Read")


def test_group_members_are_keyed_by_kind(tenant):
    group = tenant.add_group("SP-ENG-Auto-Read")
    user = tenant.add_user("Jo.Smith@contoso.com")
    nested = tenant.add_group("All Staff")
    service_principal = tenant.add_service_principal()
    group["members"] = {user["id"]: user, nested["id"]: nested, service_principal["id"]: service_principal}

    members = _groups._get_group_members(TOKEN, group["id"])

    assert members == {
        "jo.smith@contoso.com": user["id"],
        f"group:{nested['id']}": nested["id"],
        f"objectid:{service_principal['id']}": service_principal["id"],
    }


def test_adding_an_existing_member_is_success(tenant):
    group = tenant.add_group("SP-ENG-Auto-Read")
    user = tenant.add_user("jo@contoso.com")

    _groups._add_group_member(TOKEN, group["id"], user["id"])
    _groups._add_group_member(TOKEN, group["id"], user["id"])

    assert list(group["members"]) == [user["id"]]


def test_adding_to_a_new_group_waits_for_replication(tenant, monkeypatch):
    sleeps = []
    monkeypatch.setattr(time, "sleep", sleeps.append)
    group = tenant.add_group("SP-ENG-Auto-Read")
    user = tenant.add_user("jo@contoso.com")
    tenant.fail("POST", r"/members/\$ref$", graph_error(404, "Request_ResourceNotFound"), times=2)

    _groups._add_group_member(TOKEN, group["id"], user["id"], is_new_group=True)

    assert user["id"] in group["members"]
    assert sleeps == [5.0, 10.0]


def test_adding_to_an_existing_group_does_not_wait(tenant):
    tenant.fail("POST", r"/members/\$ref$", graph_error(404, "Request_ResourceNotFound"))
    with pytest.raises(requests.HTTPError):
        _groups._add_group_member(TOKEN, "missing-group", "user")


def test_emptying_a_group_removes_every_kind_of_member(tenant):
    group = tenant.add_group("SP-ENG-Auto-Read")
    user = tenant.add_user("jo@contoso.com")
    service_principal = tenant.add_service_principal()
    group["members"] = {user["id"]: user, service_principal["id"]: service_principal}

    assert _groups._empty_group_members(TOKEN, group["id"]) == 2
    assert group["members"] == {}


# Users

def test_user_resolves_directly_by_upn(tenant):
    user = tenant.add_user("Jo.Smith@contoso.com")

    assert _groups._resolve_user_with_fallback(TOKEN, "jo.smith@contoso.com") == (user["id"], "jo.smith@contoso.com")


def test_user_resolves_by_mail_when_the_upn_differs(tenant):
    user = tenant.add_user("jsmith@contoso.onmicrosoft.com", mail="jo.smith@contoso.com")

    assert _groups._resolve_user_with_fallback(TOKEN, "jo.smith@contoso.com") == (
        user["id"], "jsmith@contoso.onmicrosoft.com"
    )


def test_user_resolves_by_proxy_address(tenant):
    user = tenant.add_user("jsmith@contoso.com", proxy_addresses=["jo.smith@oldschool.org"])

    oid, upn = _groups._resolve_user_with_fallback(TOKEN, "Jo.Smith@OldSchool.org")

    assert (oid, upn) == (user["id"], "jsmith@contoso.com")


def test_unknown_user_raises_a_resolution_error(tenant):
    with pytest.raises(UserResolutionError, match="not found by UPN, mail, or proxyAddresses"):
        _groups._resolve_user_with_fallback(TOKEN, "nobody@contoso.com")


def test_malformed_user_is_reported_as_malformed(tenant):
    with pytest.raises(UserResolutionError, match="malformed") as raised:
        _groups._resolve_user_with_fallback(TOKEN, "jo smith@contoso.com")

    assert raised.value.user_principal_name == "jo smith@contoso.com"


def test_empty_user_raises_without_calling_graph(tenant):
    with pytest.raises(UserResolutionError, match="empty"):
        _groups._resolve_user_with_fallback(TOKEN, "  ")

    assert tenant.calls == []


def test_user_lookup_api_failures_are_raised(tenant):
    tenant.fail("GET", r"/users/", graph_error(403, "Authorization_RequestDenied"))

    with pytest.raises(requests.HTTPError):
        _groups._resolve_user_with_fallback(TOKEN, "jo@contoso.com")


def test_user_lookup_quotes_are_escaped(tenant):
    user = tenant.add_user("oneil@contoso.com", mail="o'neil@contoso.com")

    assert _groups._resolve_user_with_fallback(TOKEN, "o'neil@contoso.com")[0] == user["id"]


# Nestable groups

def test_nestable_group_resolves_by_object_id_or_name(tenant):
    group = tenant.add_group("All Staff")

    assert _groups._resolve_nestable_group(TOKEN, group["id"])["id"] == group["id"]
    assert _groups._resolve_nestable_group(TOKEN, " All Staff ")["id"] == group["id"]


@pytest.mark.parametrize("setup, reference, reason", [
    (lambda t: None, "", "empty group reference"),
    (lambda t: None, "00000000-0000-0000-0000-000000000000", "no group has this object ID"),
    (lambda t: None, "Nobody", "no group has this display name"),
    (lambda t: [t.add_group("Dup"), t.add_group("Dup")], "Dup", "shared by more than one group"),
    (lambda t: t.add_group("Team", group_types=["Unified"]), "Team", "not a security group"),
    (lambda t: t.add_group("List", security_enabled=False), "List", "not a security group"),
])
def test_unresolvable_nestable_groups(tenant, setup, reference, reason):
    setup(tenant)

    with pytest.raises(GroupResolutionError, match=reason):
        _groups._resolve_nestable_group(TOKEN, reference)


# SharePoint principals and role assignments

def test_new_group_principal_waits_for_replication(tenant, monkeypatch):
    sleeps = []
    monkeypatch.setattr(time, "sleep", sleeps.append)
    site = tenant.add_site("/sites/maths")
    group = tenant.add_group("SP-ENG-Auto-Read")
    tenant.fail("POST", r"/web/ensureuser$", response(400), times=1)

    principal_id = _sharepoint._resolve_group_principal_id(
        TOKEN, site.url, f"c:0t.c|tenant|{group['id']}", is_new_group=True,
    )

    assert principal_id == site.principals[f"c:0t.c|tenant|{group['id']}"]
    assert sleeps == [5.0]


def test_existing_group_principal_failure_is_raised(tenant):
    site = tenant.add_site("/sites/maths")

    with pytest.raises(requests.HTTPError):
        _sharepoint._resolve_group_principal_id(TOKEN, site.url, "c:0t.c|tenant|missing")


def test_role_definitions_are_read_in_one_request(tenant):
    site = tenant.add_site("/sites/maths")

    role_def_ids = _sharepoint._resolve_role_def_ids(TOKEN, site.url)

    assert role_def_ids == {"read": 1073741826, "contribute": 1073741827, "owner": 1073741829}
    assert len(tenant.calls) == 1


def test_missing_role_definition_is_an_error(tenant):
    site = tenant.add_site("/sites/maths")
    tenant.fail("GET", r"/roledefinitions", response(200, {"d": {"results": [
        {"Id": 1, "Name": "Read", "RoleTypeKind": 2},
    ]}}))

    with pytest.raises(ValueError, match="contribute"):
        _sharepoint._resolve_role_def_ids(TOKEN, site.url)


def test_role_definitions_are_found_by_type_not_name(tenant):
    site = tenant.add_site("/sites/maths")
    tenant.fail("GET", r"/roledefinitions", response(200, {"d": {"results": [
        {"Id": 11, "Name": "Lecture", "RoleTypeKind": 2},
        {"Id": 12, "Name": "Collaboration", "RoleTypeKind": 3},
        {"Id": 13, "Name": "Contrôle total", "RoleTypeKind": 5},
        {"Id": 14, "Name": "Read", "RoleTypeKind": 0},  # a custom level reusing an English name
    ]}}))

    assert _sharepoint._resolve_role_def_ids(TOKEN, site.url) == {"read": 11, "contribute": 12, "owner": 13}


def test_ordinary_lists_are_not_libraries_but_their_titles_count(tenant):
    site = tenant.add_site("/sites/maths")
    tasks = site.add_list("English tasks", slug="ENG", template="genericList")  # /Lists/ENG
    titled = site.add_list("ART", slug="ArtList", template="genericList")

    lists = _graph._list_existing_lists(TOKEN, site.id)

    assert "eng" not in lists and lists["list:eng"]["id"] == tasks["id"]
    assert _graph._find_title_conflict(lists, "ART")["id"] == titled["id"]
    assert _groups._find_existing_library_collisions([(site.url, "LISTENG")], {site.url: lists}) == {}


def test_group_roles_are_keyed_by_claims_login_name(tenant):
    site = tenant.add_site("/sites/maths")
    lst = site.add_list("ENG", unique=True)
    group = tenant.add_group("SP-ENG-Auto-Read")
    principal = _sharepoint._resolve_group_principal_id(TOKEN, site.url, _sharepoint._group_claims(group["id"]))
    _sharepoint._add_role_assignment(TOKEN, site.url, lst["id"], principal, 1073741826)

    roles = _sharepoint._get_library_group_roles(TOKEN, site.url, lst["id"])

    assert roles == {f"c:0t.c|tenant|{group['id']}": {1073741826}}
    tenant.calls.clear()
    assert _sharepoint._ensure_group_assigned_to_library(
        TOKEN, site.url, lst["id"], group["id"], "SP-ENG-Auto-Read", 1073741826, existing_roles=roles,
    ) is False
    assert tenant.calls == []


def test_removing_a_missing_role_assignment_is_not_an_error(tenant):
    site = tenant.add_site("/sites/maths")
    lst = site.add_list("ENG", unique=True)

    assert _sharepoint._remove_library_role_assignment(TOKEN, site.url, lst["id"], 99) is False


# Property bag

def test_property_bag_values_are_xml_escaped_and_read_back(tenant):
    site = tenant.add_site("/sites/maths")
    lst = site.add_list("ENG")

    _sharepoint._set_library_properties(TOKEN, site.url, lst["id"], {"SourceSystem": "A & B <c>", "SourceId": "42"})

    [(_, _, kwargs)] = tenant.calls
    assert b"A &amp; B &lt;c&gt;" in kwargs["data"]
    assert _sharepoint._get_library_properties(TOKEN, site.url, lst["id"], ["SourceSystem", "SourceId", "Other"]) == {
        "SourceSystem": "A & B <c>", "SourceId": "42", "Other": None,
    }


def test_property_bag_csom_errors_are_raised(tenant):
    site = tenant.add_site("/sites/maths")

    with pytest.raises(ValueError, match="List does not exist"):
        _sharepoint._set_library_properties(TOKEN, site.url, "0f0f0f0f-0000-0000-0000-000000000000", {"SourceId": "1"})


# Navigation URLs

@pytest.mark.parametrize("value", [
    "https://contoso.sharepoint.com/sites/maths/Demo%20Library%20%26%20Test",
    "/sites/maths/Demo Library & Test/",
    "/sites/maths/Demo%2520Library%2520%2526%2520Test",
    "/SITES/MATHS/demo library & test",
])
def test_navigation_urls_normalise_to_one_key(value):
    assert _navigation._normalise_nav_url(value) == "/sites/maths/demo library & test"


def test_graph_base_url_matches_the_fake():
    assert _common.GRAPH_BASE_URL == GRAPH


# Regression tests

@pytest.mark.parametrize("name", ["Maths & Science Staff", "Year 7 #1", "C++ Club"])
def test_group_names_with_query_characters_are_found(tenant, name):
    group = tenant.add_group(name)

    assert _groups._find_group_by_name(TOKEN, name)["id"] == group["id"]
    assert _groups._find_groups_by_name_prefix(TOKEN, name[:6])[name.lower()][0]["id"] == group["id"]


def test_group_names_match_case_insensitively_like_graph(tenant):
    group = tenant.add_group("baa-sharepoint-MATH-Auto-Read")

    groups = _groups._find_groups_by_name_prefix(TOKEN, "BAA-Sharepoint-MATH-")

    assert _groups._get_unique_group(groups, "BAA-Sharepoint-MATH-Auto-Read")["id"] == group["id"]


def test_prefix_lookup_sends_the_filter_as_a_parameter_on_the_first_page_only(tenant, monkeypatch):
    pages = [
        response(200, {"value": [], "@odata.nextLink": f"{GRAPH}/groups?$skiptoken=2"}),
        response(200, {"value": []}),
    ]
    sent = []
    monkeypatch.setattr(_common, "GRAPH_BASE_URL", GRAPH)
    monkeypatch.setattr(requests, "request", lambda method, url, **kwargs: sent.append((url, kwargs)) or pages.pop(0))

    _groups._find_groups_by_name_prefix(TOKEN, "SP-")

    assert sent[0][0] == f"{GRAPH}/groups"
    assert sent[0][1]["params"]["$filter"] == "startswith(displayName, 'SP-')"
    assert sent[1][0] == f"{GRAPH}/groups?$skiptoken=2" and sent[1][1]["params"] is None


def test_user_principal_name_cannot_change_the_graph_path(tenant):
    tenant.add_group("Target")

    with pytest.raises(UserResolutionError):
        _groups._resolve_user_with_fallback(TOKEN, "../groups/x@contoso.com")

    direct = [url for method, url, _ in tenant.calls if "/users/" in url]
    assert direct == [f"{GRAPH}/users/..%2Fgroups%2Fx%40contoso.com"]


@pytest.mark.parametrize("url", [
    "http://contoso.sharepoint.com/sites/maths",
    "https://evil.example.com/sites/maths",
    "https://contoso.sharepoint.com.evil.example.com/sites/maths",
])
def test_site_urls_must_be_https_on_the_sharepoint_host(url):
    with pytest.raises(ValueError, match="site_url"):
        _common._validate_site_urls([url], sharepoint_hostname="contoso.sharepoint.com")


def test_site_url_host_check_ignores_case():
    _common._validate_site_urls(["https://Contoso.SharePoint.com/sites/maths"], sharepoint_hostname="contoso.sharepoint.com")


@pytest.mark.parametrize("include_site_name", [False, True])
def test_codes_and_prefixes_differing_only_in_case_collide(include_site_name):
    prefixes = {
        (SITE, "PE-KS3"): "BAA-Sharepoint",
        (SITE, "peks3"): "BAA-Sharepoint",
        (SITE, "ENG"): "BAA-Sharepoint",
        (SITE, "eng"): "baa-sharepoint",
    }

    collisions = _groups._find_group_name_collisions(prefixes, include_site_name=include_site_name)

    assert set(collisions) == set(prefixes)


def test_group_names_differing_only_in_case_are_duplicates(tenant):
    tenant.add_group("SP-ENG-Auto-Read")
    tenant.add_group("sp-eng-auto-read")

    with pytest.raises(ValueError, match="Multiple Entra groups share displayName"):
        _groups._find_group_by_name(TOKEN, "SP-ENG-Auto-Read")


def jwt(claims):
    import base64
    import json
    body = base64.urlsafe_b64encode(json.dumps(claims).encode()).decode().rstrip("=")
    return f"eyJhbGciOiJub25lIn0.{body}.sig"


def test_object_id_is_read_from_the_token():
    object_id = "11111111-2222-3333-4444-555555555555"

    assert _groups._token_object_id(jwt({"oid": object_id, "idtyp": "app"})) == object_id


@pytest.mark.parametrize("token", ["opaque-token", jwt({"appid": "x"}), jwt({"oid": "../users/x"})])
def test_unreadable_object_id_is_an_error(token):
    with pytest.raises(ValueError, match="object ID"):
        _groups._token_object_id(token)
