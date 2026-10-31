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


def test_public_names_are_exported():
    for name in ("KeyVaultCertificateCredential", "Scopes", "Emails", "MessageNotFoundError", "SharePointLibraryManager"):
        assert hasattr(alt_fabric_m365, name)


# Where the README shows the current version: installing it, the files a build makes and the release tag
README_VERSION_PATTERNS = [
    r"choose a version, for example `(\d+\.\d+\.\d+)`",
    r"alt-fabric-m365==(\d+\.\d+\.\d+)",
    r"alt_fabric_m365-(\d+\.\d+\.\d+)(?:-py3|\.tar\.gz)",
    r'__version__="(\d+\.\d+\.\d+)"',
    r"`v(\d+\.\d+\.\d+)`",
]


@pytest.mark.parametrize("pattern", README_VERSION_PATTERNS)
def test_readme_examples_show_the_current_version(pattern):
    # The README is also the PyPI description, so a stale example would ship on the new version's page
    found = re.findall(pattern, README)

    assert found and set(found) == {alt_fabric_m365.__version__}


def test_graph_api_was_removed_in_1_0():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("alt_fabric_m365.graph_api")
    assert not hasattr(alt_fabric_m365, "graph_api")


def test_readme_has_python_examples():
    assert len(PYTHON_BLOCKS) > 10


@pytest.mark.parametrize("block", [b for b in PYTHON_BLOCKS if not b.lstrip().startswith("%")])
def test_readme_examples_are_valid_python(block):
    ast.parse(block)
