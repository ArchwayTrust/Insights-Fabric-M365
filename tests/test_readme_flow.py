"""
Runs the README's "Recommended flow" example against the in-memory FakeTenant, with its suggested input tables
standing in for the lakehouse, so the example can't drift from the package.
"""
import re
import time
import types
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import alt_fabric_m365.auth
import alt_fabric_m365.sharepoint
from alt_fabric_m365.sharepoint import SharePointLibraryManager, sanitise_for_group_name

from sharepoint_fake import APP_ID, HOST, FakeTenant

README = (Path(__file__).resolve().parent.parent / "README.md").read_text(encoding="utf-8")
SECTION = README.split("### Using it in your own setting", 1)[1].split("\n### ", 1)[0]
PARAMETERS, FLOW, SCOPE, LEADERSHIP = re.findall(r"^```python\n(.*?)^```", SECTION, re.S | re.M)


def with_setting_specific_steps(flow, site_url):
    """Inserts the README's setting-specific snippets at the two places the skeleton marks."""
    scope = SCOPE.replace("https://<tenant>.sharepoint.com/sites/Academy", site_url)
    lines = []
    for line in flow.splitlines(keepends=True):
        lines.append(line)
        if line.startswith("# Setting-specific: filter the rows in scope here"):
            lines.append(scope)
        elif line.startswith("# Setting-specific: extra passes go here"):
            lines.append(LEADERSHIP)
    return "".join(lines)


class Frame:
    """A table or query result. It has no collect(), as that gives timestamps in the Python process's time zone
    rather than Spark's session time zone, so the flow mustn't use it."""

    def __init__(self, pdf):
        self.pdf = pdf

    def toPandas(self):
        return self.pdf.copy()


class Spark:
    """The suggested tables, plus just enough SQL for the watermark."""

    def __init__(self, tables):
        self.tables = tables
        self.watermarks = {}  # process key -> last_processed_ts
        self.saved = []
        self.settings = {}
        self.conf = types.SimpleNamespace(set=self.settings.__setitem__)
        self.query_time_zones = []  # the session time zone each table read or query ran in

    @property
    def watermark(self):
        """The only watermark written so far, or None."""
        assert len(self.watermarks) <= 1, self.watermarks
        return next(iter(self.watermarks.values()), None)

    def table(self, name):
        self.query_time_zones.append(self.settings.get("spark.sql.session.timeZone"))
        return Frame(self.tables[name.rsplit(".", 1)[-1]])

    def sql(self, query):
        self.query_time_zones.append(self.settings.get("spark.sql.session.timeZone"))
        if "MERGE INTO" in query:
            process = re.search(r"SELECT '((?:[^']|'')*)' AS process", query).group(1).replace("''", "'")
            self.watermarks[process] = datetime.fromisoformat(re.search(r"TIMESTAMP'([^']+)'", query).group(1))
            return None
        assert "sharepoint_sync_state" in query
        process = re.search(r"process = '((?:[^']|'')*)'", query).group(1).replace("''", "'")
        rows = [{"last_processed_ts": self.watermarks[process]}] if process in self.watermarks else []
        return Frame(pd.DataFrame(rows, columns=["last_processed_ts"]))

    def createDataFrame(self, pdf):
        writer = types.SimpleNamespace()
        writer.format = lambda name: writer
        writer.mode = lambda name: writer
        writer.saveAsTable = lambda name: self.saved.append((name, pdf))
        return types.SimpleNamespace(write=writer)


@pytest.fixture
def setting(monkeypatch):
    tenant = FakeTenant().install(monkeypatch)
    monkeypatch.setattr(time, "sleep", lambda seconds: None)
    site = tenant.add_site("/sites/Academy")
    for upn in ("jo@contoso.com", "sam@contoso.com"):
        tenant.add_user(upn)
    tenant.add_group("All Staff")
    head = tenant.add_user("head@contoso.com")
    tenant.add_group("ACA-Leadership", owners={head["id"]})

    class Credential:
        def __init__(self, *args):
            pass

        def get_token(self, *scopes, **kwargs):
            return types.SimpleNamespace(token="t")

    monkeypatch.setattr(alt_fabric_m365.auth, "KeyVaultCertificateCredential", Credential)
    monkeypatch.setattr(
        alt_fabric_m365.sharepoint, "SharePointLibraryManager",
        lambda *args, **kwargs: SharePointLibraryManager(*args, service_principal_id=APP_ID, **kwargs),
    )

    spark = Spark({
        "sharepoint_libraries": pd.DataFrame({
            "library_id": ["L1", "L2"], "site_url": site.url, "library_code": ["MATH", "ENG"],
            "library_display_name": ["Mathematics", "English"], "group_prefix": "ACA-Sharepoint",
            "manual_read": [True, False], "previous_site_url": [None, None], "deleted": [False, False],
            "last_changed_ts": pd.to_datetime(["2026-10-01", "2026-10-02"]),
        }),
        "sharepoint_library_members": pd.DataFrame({
            "library_id": ["L1", "L1", "L2"],
            "user_principal_name": ["jo@contoso.com", "sam@contoso.com", "jo@contoso.com"],
            "permission": "contribute",
        }),
        "sharepoint_library_nested_groups": pd.DataFrame({
            "library_id": ["L2"], "member_group": ["All Staff"], "permission": ["read"],
        }),
        "sharepoint_library_grants": pd.DataFrame({
            "library_id": ["L1", "L2"], "group_name": "ACA-Leadership", "permission": "read",
        }),
    })

    def run(dry_run, flow=FLOW):
        namespace = {"spark": spark, "display": lambda *args: None, "sanitise_for_group_name": sanitise_for_group_name}
        exec(PARAMETERS, namespace)
        namespace.update(DRY_RUN=dry_run, SHAREPOINT_HOSTNAME=HOST)
        exec(flow, namespace)
        return namespace

    return types.SimpleNamespace(tenant=tenant, site=site, spark=spark, run=run)


def test_readme_flow_provisions_from_the_suggested_tables(setting):
    tenant, site, spark = setting.tenant, setting.site, setting.spark

    setting.run(dry_run=True)

    assert tenant.writes() == []
    assert spark.watermark is None

    setting.run(dry_run=False)

    assert sorted(lst["displayName"] for lst in site.lists.values()) == ["English", "Mathematics"]
    assert site.list_by_slug("MATH")["properties"]["SourceId"] == "L1"
    assert tenant.member_upns("ACA-Sharepoint-MATH-Auto-Contribute") == ["jo@contoso.com", "sam@contoso.com"]
    assert set(tenant.group_named("ACA-Sharepoint-ENG-Auto-Read")["members"]) == {tenant.group_named("All Staff")["id"]}
    leadership = site.principals[f"c:0t.c|tenant|{tenant.group_named('ACA-Leadership')['id']}"]
    assert all(leadership in site.list_by_slug(code)["role_assignments"] for code in ("MATH", "ENG"))
    assert spark.watermark == datetime(2026, 10, 2)
    assert {name.rsplit(".", 1)[-1] for name, _ in spark.saved} == {"sharepoint_action_log"}
    # Each read and the watermark's MERGE use the session time zone, so UTC is set before the first
    assert spark.query_time_zones and set(spark.query_time_zones) == {"UTC"}


def test_readme_flow_treats_a_null_deleted_as_not_deleted(setting):
    # A nullable boolean column, as Spark gives for a row that never had the flag set
    setting.spark.tables["sharepoint_libraries"]["deleted"] = pd.array([False, None], dtype=object)

    setting.run(dry_run=False)

    assert sorted(lst["displayName"] for lst in setting.site.lists.values()) == ["English", "Mathematics"]


def test_readme_flow_rerun_and_delete(setting):
    tenant, site, spark = setting.tenant, setting.site, setting.spark
    setting.run(dry_run=False)
    tenant.calls.clear()

    namespace = setting.run(dry_run=False)

    assert namespace["changed"].empty
    assert not any("/members/$ref" in url or "addroleassignment" in url for _, url, _ in tenant.writes())

    libraries = spark.tables["sharepoint_libraries"]
    libraries.loc[1, ["deleted", "last_changed_ts"]] = [True, pd.Timestamp("2026-10-05")]
    setting.run(dry_run=False)

    assert site.list_by_slug("ENG")["displayName"] == "ENG_UNLINKED"
    assert tenant.member_upns("ACA-Sharepoint-ENG-Auto-Contribute") == []
    assert spark.watermark == datetime(2026, 10, 5)


def test_readme_setting_specific_steps(setting):
    tenant, site, spark = setting.tenant, setting.site, setting.spark
    libraries = spark.tables["sharepoint_libraries"]
    other_site = tenant.add_site("/sites/Other")
    spark.tables["sharepoint_libraries"] = pd.concat([libraries, pd.DataFrame({
        "library_id": ["L3", "L4", "L5"], "site_url": [site.url, site.url, other_site.url],
        "library_code": ["S-L-T", "ART", "PE"], "library_display_name": ["Clash", "Art", "Out of scope"],
        "group_prefix": "ACA-Sharepoint", "manual_read": False, "previous_site_url": None, "deleted": False,
        "exclude": [False, True, False], "last_changed_ts": pd.Timestamp("2026-10-03"),
    })], ignore_index=True)
    flow = with_setting_specific_steps(FLOW, site.url)

    setting.run(dry_run=False, flow=flow)

    # Out of scope, excluded and the clashing code are all left alone
    assert sorted(lst["slug"] for lst in site.lists.values()) == ["ENG", "MATH", "SLT"]
    assert other_site.lists == {}
    # The leadership library's groups are people's, not the app's, and have access to every library on the site
    slt_read = tenant.group_named("ACA-Sharepoint-SLT-Manual-Read")
    assert slt_read["owners"] == set()
    principal = site.principals[f"c:0t.c|tenant|{slt_read['id']}"]
    assert all(principal in site.list_by_slug(code)["role_assignments"] for code in ("MATH", "ENG", "SLT"))
    # The newer rows were all out of scope, so the watermark only covers what this scope processed
    assert spark.watermark == datetime(2026, 10, 2)


def test_a_scoped_rerun_reads_the_watermark_it_wrote(setting):
    flow = with_setting_specific_steps(FLOW, setting.site.url)
    setting.run(dry_run=False, flow=flow)
    [process] = setting.spark.watermarks

    namespace = setting.run(dry_run=False, flow=flow)

    assert process == f"sharepoint_libraries|{setting.site.url}"
    assert namespace["since"] == datetime(2026, 10, 2)
    assert namespace["changed"].empty
