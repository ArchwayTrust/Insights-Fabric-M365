"""
Checks on the package itself: the version is set in one place, the public names, and that the README's examples are
valid Python.
"""
import ast
import importlib
import re
import subprocess
import sys
from pathlib import Path

import pytest

import alt_fabric_m365

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
PYTHON_BLOCKS = re.findall(r"^```python\n(.*?)^```", README, re.S | re.M)


def test_setup_reads_the_package_version():
    result = subprocess.run([sys.executable, "setup.py", "--version"], cwd=ROOT, capture_output=True, text=True, check=True)

    assert result.stdout.strip().splitlines()[-1] == alt_fabric_m365.__version__


# Every public name, by package. Removing or renaming one breaks notebooks, so needs a major release (see CLAUDE.md);
# adding one needs a minor release. Either way, change this list deliberately.
PUBLIC_NAMES = {
    "alt_fabric_m365": {
        "KeyVaultCertificateCredential", "Scopes", "Emails", "MessageNotFoundError", "SharePointLibraryManager",
    },
    "alt_fabric_m365.auth": {"KeyVaultCertificateCredential", "Scopes"},
    "alt_fabric_m365.exchange": {"Emails", "MessageNotFoundError"},
    "alt_fabric_m365.sharepoint": {
        "SharePointLibraryManager", "UserResolutionError", "GroupResolutionError", "build_group_name",
        "sanitise_for_group_name", "ACTION_LOG_COLUMNS", "ALL_GROUP_SUFFIXES", "AUTO_GROUP_SUFFIXES",
        "MANUAL_GROUP_SUFFIXES", "VALID_ROLES", "FATAL_ACTIONS", "WARNING_ACTIONS",
    },
}


@pytest.mark.parametrize("module_name", PUBLIC_NAMES)
def test_public_names_are_exported(module_name):
    module = importlib.import_module(module_name)

    assert set(module.__all__) == PUBLIC_NAMES[module_name]
    assert len(module.__all__) == len(set(module.__all__))
    for name in module.__all__:
        assert getattr(module, name) is not None


def test_graph_api_was_removed_in_1_0():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("alt_fabric_m365.graph_api")
    assert not hasattr(alt_fabric_m365, "graph_api")


def test_readme_has_python_examples():
    assert len(PYTHON_BLOCKS) > 10


@pytest.mark.parametrize("block", [b for b in PYTHON_BLOCKS if not b.lstrip().startswith("%")])
def test_readme_examples_are_valid_python(block):
    ast.parse(block)
