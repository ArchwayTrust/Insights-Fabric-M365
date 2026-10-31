"""
Checks on the package itself: the version is set in one place, the public names, and that the README's examples are
valid Python.
"""
import ast
import importlib
import re
import shutil
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


def test_source_distribution_has_what_its_tests_need(tmp_path, monkeypatch):
    import tarfile
    from setuptools import build_meta

    # Build from a clean copy, as setuptools adds the files listed by an earlier build's egg-info
    source = tmp_path / "source"
    shutil.copytree(ROOT, source, ignore=shutil.ignore_patterns(
        ".git", "*.egg-info", "build", "dist", "__pycache__", ".pytest_cache",
    ))
    monkeypatch.chdir(source)

    sdist = tmp_path / build_meta.build_sdist(str(tmp_path))

    with tarfile.open(sdist) as tar:
        shipped = {name.split("/", 1)[1] for name in tar.getnames() if "/" in name}
    # Every test module, fixture and fake, and the example notebook test_faculty_notebook runs
    needed = {path.relative_to(ROOT).as_posix() for path in (ROOT / "tests").glob("*.py")}
    needed |= {path.relative_to(ROOT).as_posix() for path in (ROOT / "examples").rglob("*") if path.is_file()}
    assert needed - shipped == set()


def test_graph_api_was_removed_in_1_0():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("alt_fabric_m365.graph_api")
    assert not hasattr(alt_fabric_m365, "graph_api")


def test_readme_has_python_examples():
    assert len(PYTHON_BLOCKS) > 10


@pytest.mark.parametrize("block", [b for b in PYTHON_BLOCKS if not b.lstrip().startswith("%")])
def test_readme_examples_are_valid_python(block):
    ast.parse(block)
