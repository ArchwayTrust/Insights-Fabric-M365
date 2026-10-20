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

import fabric_python_helper.auth
import fabric_python_helper.sharepoint
from fabric_python_helper.sharepoint import SharePointLibraryManager

from sharepoint_fake import APP_ID, HOST, FakeTenant

README = (Path(__file__).resolve().parent.parent / "README.md").read_text(encoding="utf-8")
SECTION = README.split("### Using it in your own setting", 1)[1].split("\n### ", 1)[0]
PARAMETERS, FLOW = re.findall(r"^```python\n(.*?)^```", SECTION, re.S | re.M)


class Frame:
    def __init__(self, pdf):
        self.pdf = pdf

    def toPandas(self):
        return self.pdf.copy()

    def collect(self):
        return self.pdf.to_dict("records")


class Spark:
    """The suggested tables, plus just enough SQL for the watermark."""

    def __init__(self, tables):
        self.tables = tables
        self.watermark = None
        self.saved = []

    def table(self, name):
        return Frame(self.tables[name.rsplit(".", 1)[-1]])

    def sql(self, query):
        if "MERGE INTO" in query:
            self.watermark = datetime.fromisoformat(re.search(r"TIMESTAMP'([^']+)'", query).group(1))
            return None
        assert "sharepoint_sync_state" in query
        rows = [{"last_processed_ts": self.watermark}] if self.watermark else []
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

    monkeypatch.setattr(fabric_python_helper.auth, "KeyVaultCertificateCredential", Credential)
    monkeypatch.setattr(
        fabric_python_helper.sharepoint, "SharePointLibraryManager",
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

    def run(dry_run):
        namespace = {"spark": spark, "display": lambda *args: None}
        exec(PARAMETERS, namespace)
        namespace.update(DRY_RUN=dry_run, SHAREPOINT_HOSTNAME=HOST)
        exec(FLOW, namespace)
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
