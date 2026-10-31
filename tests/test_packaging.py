"""
Checks on the package itself: the version is set in one place, the public names, what the source distribution ships
and its licence metadata, that the README's examples are valid Python, and that the workflows pin each action to a
commit.
"""
import ast
import email.parser
import importlib
import os
import re
import shutil
import subprocess
import sys
import tarfile
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


@pytest.fixture(scope="module")
def sdist(tmp_path_factory):
    """Builds the source distribution once for this module's tests, and returns its path."""
    from setuptools import build_meta

    # Build from a clean copy, as setuptools adds the files listed by an earlier build's egg-info
    tmp_path = tmp_path_factory.mktemp("sdist")
    source = tmp_path / "source"
    shutil.copytree(ROOT, source, ignore=shutil.ignore_patterns(
        ".git", "*.egg-info", "build", "dist", "__pycache__", ".pytest_cache",
    ))
    cwd = os.getcwd()
    os.chdir(source)
    try:
        return tmp_path / build_meta.build_sdist(str(tmp_path))
    finally:
        os.chdir(cwd)


def test_source_distribution_has_what_its_tests_need(sdist):
    with tarfile.open(sdist) as tar:
        shipped = {name.split("/", 1)[1] for name in tar.getnames() if "/" in name}
    # Every test module, fixture and fake, the workflows whose pins are checked, and the example notebook
    # test_faculty_notebook runs
    needed = {path.relative_to(ROOT).as_posix() for path in (ROOT / "tests").glob("*.py")}
    needed |= {path.relative_to(ROOT).as_posix() for path in (ROOT / ".github" / "workflows").glob("*.yml")}
    needed |= {
        path.relative_to(ROOT).as_posix() for path in (ROOT / "examples").rglob("*")
        # MANIFEST.in leaves out compiled files, which running the example locally can leave behind
        if path.is_file() and "__pycache__" not in path.parts and path.suffix not in (".pyc", ".pyo")
    }
    assert needed - shipped == set()


def test_licence_metadata_is_an_spdx_expression(sdist):
    with tarfile.open(sdist) as tar:
        pkg_info = next(name for name in tar.getnames() if name.count("/") == 1 and name.endswith("/PKG-INFO"))
        metadata = email.parser.HeaderParser().parsestr(tar.extractfile(pkg_info).read().decode("utf-8"))

    assert metadata["License-Expression"] == "BSD-3-Clause"
    assert sorted(metadata.get_all("License-File")) == ["LICENSE", "NOTICE"]
    # PEP 639 deprecates License in favour of License-Expression
    assert metadata["License"] is None


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


def test_readme_action_reference_gives_each_fatal_and_warning_actions_class():
    from alt_fabric_m365.sharepoint import FATAL_ACTIONS, WARNING_ACTIONS
    section = README.split("#### Action reference", 1)[1].split("#### Dry-run actions", 1)[0]
    classed = re.findall(r"^\| `([a-z_]+)` \|.*\| (Fatal|Warning) \|$", section, re.M)

    # A notebook's run gate depends on these, so the reference must agree with them
    assert {name for name, cls in classed if cls == "Fatal"} == set(FATAL_ACTIONS)
    assert {name for name, cls in classed if cls == "Warning"} == set(WARNING_ACTIONS)


def test_readme_dry_run_table_lists_every_dry_run_action():
    section = README.split("#### Dry-run actions", 1)[1].split("\n### ", 1)[0]
    listed = set(re.findall(r"^\| `([a-z_]+ \(dry run\))` \|", section, re.M))
    source = "".join(path.read_text(encoding="utf-8") for path in (ROOT / "alt_fabric_m365").rglob("*.py"))
    # Names built from the live action, such as manual_read_group_ensured (dry run), aren't literals, so aren't checked
    used = set(re.findall(r'"([a-z_]+ \(dry run\))"', source))

    assert used and used - listed == set()


# From 2.0.0, so removing the suffix gives the action a live run logs
def test_each_dry_run_action_is_a_live_action_followed_by_dry_run():
    reference = README.split("#### Action reference", 1)[1].split("#### Dry-run actions", 1)[0]
    live = set(re.findall(r"^\| `([a-z_]+)` \|", reference, re.M))
    section = README.split("#### Dry-run actions", 1)[1].split("\n### ", 1)[0]
    dry = re.findall(r"^\| `([a-z_]+) \(dry run\)` \| `([a-z_]+)`", section, re.M)

    assert dry and all(name in live and first_live == name for name, first_live in dry), dry


def test_graph_api_was_removed_in_1_0():
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module("alt_fabric_m365.graph_api")
    assert not hasattr(alt_fabric_m365, "graph_api")


def test_readme_links_work_on_pypi():
    # The README is the PyPI description, and PyPI doesn't serve the repository's files
    assert re.findall(r"\]\((?!https?://|#|mailto:)([^)\s]+)\)", README) == []


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


# Workflows

# A tag can be moved to other code, so each action is pinned to a commit, with its version in a comment for Dependabot
WORKFLOW_USES = [
    (path.name, match.group(1))
    for path in sorted((ROOT / ".github" / "workflows").glob("*.yml"))
    for match in re.finditer(r"^\s*(?:-\s*)?uses:\s*(.+?)\s*$", path.read_text(encoding="utf-8"), re.MULTILINE)
]


def test_workflows_use_actions():
    assert len(WORKFLOW_USES) >= 5


@pytest.mark.parametrize("workflow, uses", WORKFLOW_USES)
def test_workflow_actions_are_pinned_to_a_commit(workflow, uses):
    if uses.startswith("./"):
        return  # A workflow in this repository runs as it is at the same commit
    assert re.fullmatch(r"[\w.-]+/[\w./-]+@[0-9a-f]{40}  # v\d+(\.\d+)*", uses), f"{workflow}: {uses}"
