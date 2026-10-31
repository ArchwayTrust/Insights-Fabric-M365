"""
Runs the example "Manage Faculty Sharepoint Libraries" notebook end to end, outside Fabric.

Each code cell is executed in order, as Fabric runs it, with:
    - spark backed by SQLite, holding small curated faculty tables. The notebook's Spark SQL is translated just
      enough to run (four-part table names, TIMESTAMP literals, regexp_replace, MERGE).
    - notebookutils and the variable library replaced with fakes.
    - the real KeyVaultCertificateCredential, with Entra ID and Key Vault replaced with FakeAzure.
    - Graph and SharePoint replaced with the in-memory FakeTenant.

This checks the notebook's own logic (scope filters, SLT checks, the watermark, the error gate) against the
package, which unit tests of the package alone can't.
"""
import re
import sqlite3
import sys
import time
import types
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import alt_fabric_m365.auth.keyvault
from alt_fabric_m365.auth import Scopes

from conftest import FakeAzure
from sharepoint_fake import APP_ID, HOST, ROLE_DEFINITIONS, FakeTenant

NOTEBOOK = (Path(__file__).resolve().parent.parent / "examples" / "Manage Faculty Sharepoint Libraries.Notebook"
            / "notebook-content.py")
HUB = f"https://{HOST}/sites/BAA"
OLD_HUB = f"https://{HOST}/sites/BAAOld"
NCA_HUB = f"https://{HOST}/sites/NCA"

FACULTY_COLUMNS = [
    "faculty_unique_id", "academy_code", "static_faculty_code", "current_faculty_name", "sharepoint_hub_url",
    "previous_hub_url", "deleted_in_source", "exclude_from_sharepoint", "last_changed_ts",
    "all_staff_security_group", "all_staff_read", "all_staff_contribute",
]


def cells(path):
    """Returns the code of each code cell, with a flag for the parameters cell."""
    result = []
    current = None
    for line in path.read_text(encoding="utf-8").splitlines(keepends=True):
        marker = re.match(r"# (PARAMETERS CELL|CELL|MARKDOWN|METADATA) \*+", line)
        if marker:
            if current is not None:
                result.append(current)
            kind = marker.group(1)
            current = ["", kind == "PARAMETERS CELL"] if kind.endswith("CELL") else None
            continue
        if current is not None:
            current[0] += line
    if current is not None:
        result.append(current)
    return result


class FakeSpark:
    """Runs the notebook's Spark SQL on SQLite, and records the Delta tables it appends to."""

    def __init__(self, faculties, memberships):
        self.db = sqlite3.connect(":memory:")
        self.db.create_function("regexp_replace", 3, lambda s, pattern, repl: re.sub(pattern, repl, s or ""))
        self.db.create_function("UPPER", 1, lambda s: s.upper() if s is not None else None)
        pd.DataFrame(faculties, columns=FACULTY_COLUMNS).to_sql("sharepoint_faculty_list", self.db, index=False)
        pd.DataFrame(memberships, columns=["faculty_unique_id", "member_email"]).to_sql(
            "sharepoint_faculty_membership", self.db, index=False)
        self.db.execute("CREATE TABLE sharepoint_sync_state (process TEXT, last_processed_ts TEXT, updated_at TEXT)")
        self.saved = {}  # table name -> list of DataFrames appended
        self.after_query = None  # called with each SELECT, e.g. to add rows mid-run

    def _translate(self, query):
        query = re.sub(r"\b\w+\.\w+\.it\.(\w+)", r"\1", query)
        return re.sub(r"TIMESTAMP'([^']*)'", r"'\1'", query)

    def sql(self, query):
        query = self._translate(query)
        if re.match(r"\s*CREATE (SCHEMA|TABLE)", query):
            return None
        merge = re.match(r"\s*MERGE INTO .*?'((?:[^']|'')+)' AS process,\s*'([^']+)' AS last_processed_ts", query, re.S)
        if merge:
            process = merge.group(1).replace("''", "'")
            self.db.execute("DELETE FROM sharepoint_sync_state WHERE process = ?", (process,))
            self.db.execute("INSERT INTO sharepoint_sync_state VALUES (?, ?, datetime('now'))", (process, merge.group(2)))
            return None
        frame = FakeSparkFrame(pd.read_sql_query(query, self.db))
        if self.after_query:
            self.after_query(query)
        return frame

    def add_faculties(self, faculties):
        pd.DataFrame(faculties, columns=FACULTY_COLUMNS).to_sql(
            "sharepoint_faculty_list", self.db, index=False, if_exists="append")

    def createDataFrame(self, pdf):
        return types.SimpleNamespace(write=FakeWriter(self, pdf.copy()))

    def watermark(self):
        rows = self.db.execute("SELECT process, last_processed_ts FROM sharepoint_sync_state").fetchall()
        return dict(rows)


class FakeSparkFrame:
    def __init__(self, pdf):
        if "last_changed_ts" in pdf.columns:
            pdf["last_changed_ts"] = pd.to_datetime(pdf["last_changed_ts"])
        self.pdf = pdf

    def toPandas(self):
        return self.pdf.copy()

    def collect(self):
        return [
            {k: datetime.fromisoformat(v) if k.endswith("_ts") and isinstance(v, str) else v for k, v in row.items()}
            for row in self.pdf.to_dict("records")
        ]


class FakeWriter:
    def __init__(self, spark, pdf):
        self.spark, self.pdf = spark, pdf

    def format(self, name):
        assert name == "delta"
        return self

    def mode(self, name):
        assert name == "append"
        return self

    def saveAsTable(self, table):
        self.spark.saved.setdefault(table.rsplit(".", 1)[-1], []).append(self.pdf)


def faculty(fid, code, name, hub=HUB, academy="BAA", deleted=False, previous_hub=None, changed="2026-10-01 09:00:00",
            all_staff_group=None, all_staff_read=0):
    return (fid, academy, code, name, hub, previous_hub, deleted, False, changed, all_staff_group, all_staff_read, 0)


@pytest.fixture
def tenant(monkeypatch):
    tenant = FakeTenant().install(monkeypatch)
    monkeypatch.setattr(time, "sleep", lambda seconds: None)
    tenant.add_site("/sites/BAA")
    tenant.add_site("/sites/BAAOld")
    tenant.add_site("/sites/NCA")
    for upn in ("jo@contoso.com", "sam@contoso.com", "nca.teacher@contoso.com"):
        tenant.add_user(upn)
    tenant.add_group("BAA All Staff")
    return tenant


@pytest.fixture
def azure(monkeypatch):
    """Entra ID and Key Vault, for the notebook's KeyVaultCertificateCredential."""
    fake = FakeAzure(object_id=APP_ID)
    monkeypatch.setattr(alt_fabric_m365.auth.keyvault.requests, "get", fake.get)
    monkeypatch.setattr(alt_fabric_m365.auth.keyvault.requests, "post", fake.post)
    return fake


@pytest.fixture
def run_notebook(monkeypatch, azure):
    """Returns run(spark, **parameters), which executes the notebook and returns its namespace."""
    variables = {"STORE_WORKSPACE_NAME": "Store", "CURATED_LH_NAME": "Curated", "BASE_LH_NAME": "Base",
                 "KEY_VAULT": "https://myvault.vault.azure.net/", "TENANT_ID": "tenant"}
    library = types.SimpleNamespace(getVariable=variables.__getitem__)
    notebookutils = types.ModuleType("notebookutils")
    notebookutils.variableLibrary = types.SimpleNamespace(getLibrary=lambda name: library)
    notebookutils.credentials = types.SimpleNamespace(getToken=lambda audience: f"{audience}-token")
    monkeypatch.setitem(sys.modules, "notebookutils", notebookutils)

    def run(spark, **parameters):
        namespace = {"spark": spark, "notebookutils": notebookutils, "display": lambda *args: None,
                     "__name__": "__main__"}
        # Fabric runs an injected cell with the pipeline's parameters straight after the parameters cell
        overrides = {"SHAREPOINT_HOSTNAME": HOST, "DRY_RUN": False, **parameters}
        for code, is_parameters in cells(NOTEBOOK):
            exec(compile(code, str(NOTEBOOK), "exec"), namespace)
            if is_parameters:
                namespace.update(overrides)
        return namespace

    return run


def standard_data():
    faculties = [
        faculty("F1", "MATH", "Maths", all_staff_group="BAA All Staff", all_staff_read=1),
        faculty("F2", "MATHOLD", "Old Maths", deleted=True),
        faculty("F3", "MATHMOVED", "Moved Maths", previous_hub=OLD_HUB),
        faculty("F4", "ENG", "English"),                    # outside FACULTY_CODE_FILTER
        faculty("F5", "MATHN", "NCA Maths", hub=NCA_HUB, academy="NCA"),  # outside ACADEMY_FILTER
        faculty("F6", "SLT", "Senior Leadership"),          # code reserved for the SLT library
    ]
    memberships = [("F1", "jo@contoso.com"), ("F1", "sam@contoso.com"), ("F5", "nca.teacher@contoso.com"),
                   ("F1", "typo@contoso.com")]
    return faculties, memberships


@pytest.fixture
def spark():
    spark = FakeSpark(*standard_data())
    yield spark
    spark.db.close()


def seed_existing_libraries(tenant):
    """Libraries that earlier runs created: the deleted faculty's, and the moved faculty's on its old hub."""
    from alt_fabric_m365.sharepoint import SharePointLibraryManager

    class Credential:
        def get_token(self, *scopes, **kwargs):
            return types.SimpleNamespace(token="t")

    manager = SharePointLibraryManager(
        Credential(), HOST, include_site_name=False, add_to_navigation=True, service_principal_id=APP_ID,
    )
    for hub, code in ((HUB, "MATHOLD"), (OLD_HUB, "MATHMOVED")):
        manager.create_libraries(pd.DataFrame({
            "site_url": [hub], "library_code": [code], "library_display_name": [code],
            "group_prefix": ["BAA-Sharepoint"],
        }))
    tenant.calls.clear()


def test_live_run_provisions_the_pilot_scope_and_advances_the_watermark(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)

    ns = run_notebook(spark)

    hub, old_hub, nca = (tenant.sites[url.lower()] for url in (HUB, OLD_HUB, NCA_HUB))
    assert sorted(lst["slug"] for lst in hub.lists.values()) == ["MATH", "MATHMOVED", "MATHOLD", "SLT"]
    assert nca.lists == {}  # outside the academy filter
    assert hub.list_by_slug("MATH")["displayName"] == "Maths"
    assert hub.list_by_slug("MATH")["properties"] == {"SourceSystem": "Arbor", "SourceId": "F1"}
    assert hub.list_by_slug("MATHOLD")["displayName"] == "MATHOLD_UNLINKED"
    assert old_hub.list_by_slug("MATHMOVED")["displayName"] == "MATHMOVED_MOVED"

    assert tenant.member_upns("BAA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com", "sam@contoso.com"]
    assert set(tenant.group_named("BAA-Sharepoint-MATH-Auto-Read")["members"]) == {
        tenant.group_named("BAA All Staff")["id"]
    }

    # SLT groups hold Contribute and Read on the faculty library
    math = hub.list_by_slug("MATH")
    for name, role in (("BAA-Sharepoint-SLT-Manual-Contribute", "Contribute"), ("BAA-Sharepoint-SLT-Manual-Read", "Read")):
        principal = hub.principals[f"c:0t.c|tenant|{tenant.group_named(name)['id']}"]
        assert ROLE_DEFINITIONS[role] in math["role_assignments"][principal]

    assert hub.web["MembersCanShare"] is False

    # Faculty groups are owned by the app; the SLT groups are left for the headteacher to own
    assert tenant.group_named("BAA-Sharepoint-MATH-Auto-Contribute")["owners"] == {APP_ID}
    assert tenant.group_named("BAA-Sharepoint-SLT-Manual-Read")["owners"] == set()
    # The unlinked library's navigation link was removed; the others were added under the label
    assert sorted(c["Title"] for c in hub.nav[0]["Children"]) == ["Maths", "Moved Maths", "SLT"]

    # Warnings were surfaced but didn't fail the run, and the action log and unresolved users were persisted
    log = pd.concat(spark.saved["sharepoint_action_log"])
    assert set(log["run_id"]) == {ns["RUN_ID"]} and not log["dry_run"].any()
    [unresolved] = spark.saved["sharepoint_unresolved_users"]
    assert unresolved[["academy", "subject", "user_principal_name"]].values.tolist() == [
        ["BAA", "Maths", "typo@contoso.com"]
    ]
    assert spark.watermark() == {"faculty_sharepoint_sync|BAA|MATH%": "2026-10-01 09:00:00.000000"}


def test_notebook_authenticates_with_the_key_vault_certificate(tenant, run_notebook, spark, azure):
    seed_existing_libraries(tenant)

    ns = run_notebook(spark)

    # The certificate named in the parameters is read from the key vault, which signs each assertion
    assert {url for url, _ in azure.certificate_calls} == {
        f"https://myvault.vault.azure.net/certificates/{ns['CERTIFICATE_NAME']}"
    }
    assert azure.sign_calls
    # One token per API, cached for the whole run
    scopes = [kwargs["data"]["scope"] for _, kwargs in azure.token_calls]
    assert sorted(scopes) == sorted([Scopes.GRAPH, f"https://{HOST}/.default"])
    sharepoint_token = azure.issued[scopes.index(f'https://{HOST}/.default')]
    sharepoint_calls = [kwargs for method, url, kwargs in tenant.calls if "/_api/" in url]
    assert sharepoint_calls
    assert {kwargs["headers"]["Authorization"] for kwargs in sharepoint_calls} == {f"Bearer {sharepoint_token}"}


def test_second_run_finds_no_new_events_and_changes_no_membership(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)
    run_notebook(spark)
    tenant.calls.clear()

    ns = run_notebook(spark)

    assert ns["pdf_library_changes"].empty and ns["pdf_deleted_in_source"].empty and ns["pdf_hub_moves"].empty
    assert not any("/members/$ref" in url for method, url, _ in tenant.calls if method in ("POST", "DELETE"))
    assert not any(url.endswith("/groups") for method, url, _ in tenant.calls if method == "POST")


def test_dry_run_mutates_nothing_and_keeps_the_watermark(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)

    run_notebook(spark, DRY_RUN=True)

    assert tenant.writes() == []
    assert spark.watermark() == {}
    log = pd.concat(spark.saved["sharepoint_action_log"])
    assert log["dry_run"].all()


def test_failed_actions_fail_the_run_and_keep_the_watermark(tenant, run_notebook, spark):
    from sharepoint_fake import graph_error
    seed_existing_libraries(tenant)
    tenant.fail("POST", r"/members/\$ref$", graph_error(403, "Authorization_RequestDenied"), times=99)

    with pytest.raises(RuntimeError, match=r"failed action\(s\) this run"):
        run_notebook(spark)

    assert spark.watermark() == {}


def test_widening_the_scope_replays_events_under_a_new_watermark(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)
    run_notebook(spark)

    ns = run_notebook(spark, ACADEMY_FILTER="BAA,NCA", FACULTY_CODE_FILTER="")

    # Now in scope, the faculty coded SLT collides with the academy SLT library: reported, and skipped
    [warning] = ns["pdf_slt_warnings"].itertuples()
    assert (warning.document_library, warning.action) == ("SLT", "group_name_collision")
    assert [lst["displayName"] for lst in tenant.sites[HUB.lower()].lists.values() if lst["slug"] == "SLT"] == ["SLT"]
    assert sorted(spark.watermark()) == ["faculty_sharepoint_sync|BAA,NCA|*", "faculty_sharepoint_sync|BAA|MATH%"]
    assert tenant.sites[NCA_HUB.lower()].list_by_slug("MATHN") is not None
    assert tenant.sites[HUB.lower()].list_by_slug("ENG") is not None
    assert tenant.member_upns("NCA-Sharepoint-MATHN-Auto-Contribute") == ["nca.teacher@contoso.com"]


def test_notebook_has_no_undefined_names():
    import ast

    from pyflakes.checker import Checker

    source = NOTEBOOK.read_text(encoding="utf-8")
    # spark, notebookutils and display are injected by Fabric
    checker = Checker(ast.parse(source), filename=str(NOTEBOOK), builtins={"spark", "notebookutils", "display"})

    assert [str(m) for m in checker.messages] == []


def test_notebook_only_uses_the_public_package_api():
    source = NOTEBOOK.read_text(encoding="utf-8")

    imported = re.findall(r"from alt_fabric_m365[\w.]* import \(?([^)]*?)\)?\n(?!\s)", source, re.S)
    names = {name.strip() for group in imported for name in group.replace("\n", ",").split(",") if name.strip()}

    assert names and not any(name.startswith("_") for name in names)
    assert "mgr._" not in source
    assert "def _request_with_retry" not in source


# Spark SQL reads 'O''X' as 'OX', and a backslash escapes the next character, so neither can be put in a literal
@pytest.mark.parametrize("parameters", [
    {"ACADEMY_FILTER": "BAA,O'X"},
    {"ACADEMY_FILTER": "BAA\\"},
    {"FACULTY_CODE_FILTER": "MATH'%"},
    {"FACULTY_CODE_FILTER": "MATH\\%"},
])
def test_scope_values_spark_sql_would_misread_are_refused_before_any_sql_runs(
    tenant, run_notebook, spark, parameters,
):
    seed_existing_libraries(tenant)
    queries = []
    run_sql = spark.sql
    spark.sql = lambda query: queries.append(query) or run_sql(query)

    with pytest.raises(ValueError, match=next(iter(parameters))):
        run_notebook(spark, **parameters)

    assert queries == [] and tenant.writes() == []


def test_scope_values_with_spaces_hyphens_and_dots_are_used(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)

    run_notebook(spark, ACADEMY_FILTER="BAA, N-C.A", FACULTY_CODE_FILTER="MATH_%")

    assert spark.watermark() == {"faculty_sharepoint_sync|BAA,N-C.A|MATH_%": "2026-10-01 09:00:00.000000"}


def test_rows_landing_mid_run_are_left_for_the_next_run(tenant, run_notebook, spark):
    seed_existing_libraries(tenant)
    late = [
        faculty("F8", "MATHNEW", "New Maths", changed="2026-10-02 09:00:00"),
        # Newer, and a move: without a shared upper bound the retire pass would see it and advance the
        # watermark past F8, which the create pass had already missed
        faculty("F9", "MATHGONE", "Gone Maths", previous_hub=OLD_HUB, changed="2026-10-03 09:00:00"),
    ]

    def land_after_the_create_read(query):
        if "AS library_display_name" in query and "TRUE AS manual_read" in query:
            spark.after_query = None
            spark.add_faculties(late)

    spark.after_query = land_after_the_create_read
    run_notebook(spark)

    hub = tenant.sites[HUB.lower()]
    assert hub.list_by_slug("MATHNEW") is None
    assert spark.watermark() == {"faculty_sharepoint_sync|BAA|MATH%": "2026-10-01 09:00:00.000000"}

    run_notebook(spark)

    assert hub.list_by_slug("MATHNEW")["displayName"] == "New Maths"
    assert spark.watermark() == {"faculty_sharepoint_sync|BAA|MATH%": "2026-10-03 09:00:00.000000"}


def test_a_moved_faculty_keeps_its_old_library_until_the_new_one_exists(tenant, run_notebook, spark):
    from sharepoint_fake import graph_error
    seed_existing_libraries(tenant)
    hub_id = tenant.sites[HUB.lower()].id
    # Creating libraries on the new hub fails, so MATHMOVED isn't provisioned there
    tenant.fail("POST", rf"/sites/{re.escape(hub_id)}/lists$", graph_error(503, "serviceNotAvailable"), times=99)

    with pytest.raises(RuntimeError, match="failed action"):
        run_notebook(spark)

    old_hub = tenant.sites[OLD_HUB.lower()]
    assert old_hub.list_by_slug("MATHMOVED")["displayName"] == "MATHMOVED"  # not retired
    assert old_hub.list_by_slug("MATHMOVED")["role_assignments"]  # still has its access
    assert spark.watermark() == {}
