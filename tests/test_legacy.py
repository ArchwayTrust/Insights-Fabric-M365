"""
Checks on legacy/, the final release of the old fabric_python_helper name, which only forwards to alt_fabric_m365.
"""
import importlib
import sys
from pathlib import Path

import pytest

import alt_fabric_m365
import alt_fabric_m365.sharepoint._http

LEGACY = Path(__file__).resolve().parent.parent / "legacy"


@pytest.fixture
def legacy(monkeypatch):
    monkeypatch.syspath_prepend(str(LEGACY))
    for name in [n for n in sys.modules if n == "fabric_python_helper" or n.startswith("fabric_python_helper.")]:
        monkeypatch.delitem(sys.modules, name)
    with pytest.warns(FutureWarning, match="renamed alt_fabric_m365"):
        module = importlib.import_module("fabric_python_helper")
    yield module
    for name in [n for n in sys.modules if n == "fabric_python_helper" or n.startswith("fabric_python_helper.")]:
        del sys.modules[name]


def test_public_names_are_the_new_package_objects(legacy):
    assert legacy.__version__ == alt_fabric_m365.__version__
    for name in alt_fabric_m365.__all__:
        assert getattr(legacy, name) is getattr(alt_fabric_m365, name)


def test_submodule_imports_return_the_new_modules(legacy):
    from fabric_python_helper.auth import KeyVaultCertificateCredential
    from fabric_python_helper.sharepoint import _http
    import fabric_python_helper.graph_api as gr

    assert KeyVaultCertificateCredential is alt_fabric_m365.KeyVaultCertificateCredential
    assert _http is alt_fabric_m365.sharepoint._http
    assert gr is alt_fabric_m365.graph_api
    assert importlib.import_module("fabric_python_helper.sharepoint._http") is alt_fabric_m365.sharepoint._http
