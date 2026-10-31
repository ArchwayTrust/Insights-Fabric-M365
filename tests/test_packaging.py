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


def python_blocks(markdown):
    """Returns the code of each python fenced block.

    A fence in a list item is indented, and so is its code. Only the fence's indentation is removed, so code indented
    further than its fence still fails to parse.
    """
    return [
        "".join(line.removeprefix(indent) for line in block.splitlines(keepends=True))
        for indent, block in re.findall(r"^([ \t]*)```python\n(.*?)^\1```", markdown, re.S | re.M)
    ]


PYTHON_BLOCKS = python_blocks(README)


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
    needed |= {
        path.relative_to(ROOT).as_posix() for path in (ROOT / "examples").rglob("*")
        # MANIFEST.in leaves out compiled files, which running the example locally can leave behind
        if path.is_file() and "__pycache__" not in path.parts and path.suffix not in (".pyc", ".pyo")
    }
    assert needed - shipped == set()


def test_graph_api_was_removed_in_1_0():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("alt_fabric_m365.graph_api")
    assert not hasattr(alt_fabric_m365, "graph_api")


def test_readme_has_python_examples():
    assert len(PYTHON_BLOCKS) > 10


def test_examples_in_list_items_are_found():
    markdown = "1. Import it:\n\n   ```python\n   import a\n   if a:\n       b()\n   ```\n\n```python\nc()\n```\n"

    assert python_blocks(markdown) == ["import a\nif a:\n    b()\n", "c()\n"]


@pytest.mark.parametrize("markdown", [
    "```python\n    import a\n```\n",
    "1. Import it:\n\n   ```python\n       import a\n   ```\n",
])
def test_code_indented_further_than_its_fence_is_kept_so_it_fails_to_parse(markdown):
    [block] = python_blocks(markdown)

    with pytest.raises(IndentationError):
        ast.parse(block)


@pytest.mark.parametrize("block", [b for b in PYTHON_BLOCKS if not b.lstrip().startswith("%")])
def test_readme_examples_are_valid_python(block):
    ast.parse(block)
