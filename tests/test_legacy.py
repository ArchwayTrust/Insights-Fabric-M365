"""
Checks on legacy/, the final release of the old fabric_python_helper name, which only forwards to archwaytrust_fabric.
"""
import importlib
import sys
from pathlib import Path

import pytest

import archwaytrust_fabric
import archwaytrust_fabric.sharepoint._http

LEGACY = Path(__file__).resolve().parent.parent / "legacy"


@pytest.fixture
def legacy(monkeypatch):
    monkeypatch.syspath_prepend(str(LEGACY))
    for name in [n for n in sys.modules if n == "fabric_python_helper" or n.startswith("fabric_python_helper.")]:
        monkeypatch.delitem(sys.modules, name)
    with pytest.warns(FutureWarning, match="renamed archwaytrust_fabric"):
        module = importlib.import_module("fabric_python_helper")
    yield module
    for name in [n for n in sys.modules if n == "fabric_python_helper" or n.startswith("fabric_python_helper.")]:
        del sys.modules[name]


def test_public_names_are_the_new_package_objects(legacy):
    assert legacy.__version__ == archwaytrust_fabric.__version__
    for name in archwaytrust_fabric.__all__:
        assert getattr(legacy, name) is getattr(archwaytrust_fabric, name)


def test_submodule_imports_return_the_new_modules(legacy):
    from fabric_python_helper.auth import KeyVaultCertificateCredential
    from fabric_python_helper.sharepoint import _http
    import fabric_python_helper.graph_api as gr

    assert KeyVaultCertificateCredential is archwaytrust_fabric.KeyVaultCertificateCredential
    assert _http is archwaytrust_fabric.sharepoint._http
    assert gr is archwaytrust_fabric.graph_api
    assert importlib.import_module("fabric_python_helper.sharepoint._http") is archwaytrust_fabric.sharepoint._http
