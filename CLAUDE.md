# CLAUDE.md

`alt-fabric-m365` (import name `alt_fabric_m365`) is a Python package for calling Microsoft 365 APIs from Microsoft Fabric notebooks as an app registration. It's published to PyPI. Fabric runtimes use Python 3.10 and 3.11, so code must run on 3.10 or later.

## Commands

Run all three before pushing; CI runs the same on Python 3.10, 3.11 and 3.12:

```bash
python -m pip install -e .[test]
python -m pytest
python -m pyflakes alt_fabric_m365 tests setup.py
```

CI also runs `python -m build` and `python -m twine check --strict dist/*`.

## Layout

| Path | What it holds |
| --- | --- |
| `alt_fabric_m365/auth/` | `KeyVaultCertificateCredential` (`keyvault.py`) and `Scopes` (`scopes.py`) |
| `alt_fabric_m365/exchange/` | `Emails` and `MessageNotFoundError` (`emails.py`) |
| `alt_fabric_m365/sharepoint/` | `SharePointLibraryManager` (`manager.py`), with helpers split by the API they call: `_graph.py`, `_sharepoint.py` (REST), `_groups.py`, `_navigation.py`, `_http.py` (retry wrapper), `_common.py` (constants, validation, action log) |
| `examples/`, `docs/` | The faculty SharePoint notebook and its write-up |
| `tests/` | See [Tests](#tests) |

- Each subpackage's `__init__.py` has a docstring explaining the area, re-exports the public names, and lists them in `__all__`. The top-level `alt_fabric_m365/__init__.py` re-exports the main classes.
- Modules a user shouldn't import directly start with `_`, as in `sharepoint/_graph.py`.
- Public names are those in an `__all__`. Don't remove or rename one outside a major release. To rename one, add the new name in a minor release and keep the old one as an alias that warns with `FutureWarning`, then remove the alias in the next major release. Notebooks pin versions, but people upgrade them.

## Functions and methods

- **Method**: when the function reads or changes the object's state (`self`), or is part of the class's interface.
- **`@classmethod`**: when it only needs class attributes, as `KeyVaultCertificateCredential._validate_akv_url` reads `AKV_DNS_SUFFIX`.
- **Module-level function**: everything else. Make it private (`_name`) unless it's meant to be public. Put a single class's helpers above that class in the same file, as `_describe_error` is in `auth/keyvault.py` and `_graph_error` in `exchange/emails.py`. Put helpers shared across a subpackage in their own `_module.py`, as in `sharepoint/`.
- **Don't use `@staticmethod`**, in the package or in the tests. A function that doesn't use `self` or `cls` belongs at module level.
- Helpers that call an API take what they need as arguments (`token`, `site_url`, `group_id`) rather than reaching into a manager, so they can be tested on their own.

## Style

- Docstrings use [Google style](https://google.github.io/styleguide/pyguide.html#38-comments-and-docstrings):
  - The summary sentence starts on the same line as the opening quotes. Module docstrings are the exception: they start on the line after, naming the module's main class or purpose.
  - Sections are `Args:`, `Returns:`, `Raises:` and, for classes, `Attributes:`. Don't use `Parameters:`.
  - Types go in type hints on the signature, not in `Args:` or `Returns:`. `Attributes:` entries give their type in brackets, `tenant_id (str):`, as attributes aren't annotated.
  - Mark an optional argument "Optional." and say what it defaults to.
  - Example:

    ```python
    def _get_certificate(self, akv_headers: dict) -> tuple[str, bytes]:
        """Retrieves the public part of the certificate from Azure Key Vault.

        Args:
            akv_headers: Headers for calling Azure Key Vault, from _akv_headers().

        Returns:
            The key identifier (URL including version) to sign with, and the certificate's DER encoded bytes.
        """
    ```
- Every function and method has type hints for its arguments and return value, using Python 3.10 syntax (`str | None`, `list[dict]`). A credential is left unannotated, as anything with a `get_token` method works.
- Every public class, function and method has a docstring, including what the app registration needs (permissions, roles) to use it.
- Otherwise match the file you're in, for comment density, naming and quote style.
- Comments say why, not what: a constraint of the API, a security reason, a case the code handles that isn't obvious.
- British spelling in names, messages and docs: `sanitise`, `summarise`, `normalise`, `initialise`, `honour`, `licence`. Keep the spelling of names that aren't ours: "Microsoft Entra admin center", "BSD 3-Clause License", `str.capitalize()`, Graph's `conflictBehavior`.
- Keep lines to about 120 characters.
- Constants are `UPPER_CASE`: module-level in `sharepoint/_common.py` and `_http.py`, class attributes in `auth` and `exchange` (`REQUEST_TIMEOUT_SECONDS`).
- The only runtime dependencies are `azure-core`, `pandas` and `requests`. Don't add another without a strong reason, as each one has to be installed in every Fabric environment.
- `notebookutils` only exists in Fabric. Import it inside the function that needs it (see `KeyVaultCertificateCredential._akv_headers`), so the package imports anywhere.

## HTTP and security

- Every request has a timeout. SharePoint code goes through `_request_with_retry` in `sharepoint/_http.py`, which retries throttling and server errors; pass `idempotent=False` for a request that mustn't run twice, such as creating a group.
- Error messages include the status code and the service's error code and message, never request data, tokens or full response bodies from auth endpoints. See `_describe_error`.
- Only send a token to the host it's for. `KeyVaultCertificateCredential` checks the Key Vault URL, and the key identifier Key Vault returns, before sending its token.
- Escape values put into OData filters (`_odata_escape`, or doubling single quotes). Pass query string values with `params=` rather than formatting them into the URL, and quote values put into a URL path (`requests.utils.quote(value, safe='')`).
- `SharePointLibraryManager` operations take a DataFrame of desired state and return an action-log DataFrame. They must be safe to run again (a re-run changes nothing), and a dry run must make no writes.

## Tests

No test calls a real service. `requests` is mocked and `notebookutils` and Spark are replaced with fakes, so the tests run anywhere without credentials.

| File | Covers |
| --- | --- |
| `conftest.py` | Shared fixtures: `fake_notebookutils`, `no_notebookutils`, `credential` (`FakeCredential`), `make_response`, `FakeAzure` (Key Vault and Entra ID) |
| `sharepoint_fake.py` | `FakeTenant`, an in-memory fake of the Graph and SharePoint endpoints that keeps state, so a test can check what an operation changed |
| `test_auth.py` | `KeyVaultCertificateCredential` and its helpers |
| `test_exchange.py` | `Emails` and its helpers |
| `test_sharepoint_http.py` | The retry wrapper |
| `test_sharepoint_helpers.py` | The `sharepoint/_*.py` helpers, imported directly as `_groups._build_mail_nickname` and so on |
| `test_sharepoint_manager.py` | Each `SharePointLibraryManager` operation end to end |
| `test_readme_flow.py` | Runs the README's recommended flow |
| `test_faculty_notebook.py` | Runs the example notebook cell by cell |
| `test_packaging.py` | The version, public names, that removed names stay removed, and that every README Python example parses |

Expectations for a change:

- New behaviour and bug fixes come with tests. A bug fix's test fails without the fix.
- Test through the public interface where you can. Test a private helper directly (`keyvault._describe_error`, `emails_module._graph_error`) when it has cases that are awkward to reach from outside, such as malformed responses.
- Mock at the `requests` boundary (`mock.patch.object(emails_module.requests, "get", ...)`, `FakeTenant().install(monkeypatch)`) so the real request building and response handling run. Build responses with `make_response` or `sharepoint_fake.response`, which return real `requests.Response` objects.
- For a `SharePointLibraryManager` operation, check the action log it returns and the state it changed in `FakeTenant`, that running it again changes nothing, and that a dry run makes no writes (`tenant.writes()`).
- When the fake has to learn a new endpoint, make it behave like the real one, including its errors and its parsing of query strings.
- Patch `time.sleep` in tests that would otherwise wait on retries.
- Use `pytest.mark.parametrize` for cases of one behaviour, and name tests after the behaviour (`test_other_errors_are_raised_not_reported_as_missing`). Group tests in a file under `# Section` comments, and give each test file a module docstring saying what it covers.
- When you add a test file, add it to the "What they cover" table in the README and to the table above.

## Documentation and versions

- `README.md` is the user documentation and the PyPI description. Update it in the same change as the behaviour it describes. Python examples in it must parse (`test_packaging.py`), and the recommended-flow example must keep running (`test_readme_flow.py`).
- The version is set only in `alt_fabric_m365/__init__.py` (`__version__`); `setup.py` reads it from there. Follow the README's [Versioning](README.md#versioning) section: bump the major version for breaking changes (a public name removed or renamed, a changed signature or return value, changed action-log columns), the minor version for new features, and the patch version for fixes and documentation. A release also updates the version shown in the README's install, wheel and TestPyPI examples and its version history paragraph under Installation, plus an upgrading section for a major version.
- Releases are published by a GitHub release tagged `v<version>`; see [Release checklist](README.md#release-checklist). Merging to `main` publishes nothing.

## Pull requests

- Work on a branch and merge into `main` through a pull request; CI must pass.
- `.github/CODEOWNERS` sets who reviews: `@ArchwayTrust/insights` for everything, and `@ArchwayTrust/it-approvers` for `alt_fabric_m365/sharepoint/`, the example notebook and its write-up, and their tests (`tests/test_sharepoint_*.py`, `sharepoint_fake.py`, `test_faculty_notebook.py` and `test_readme_flow.py`). When you add SharePoint files outside those paths, add them there too.
- Keep a pull request to one change. Don't reformat or restyle code it doesn't otherwise touch.
- Titles describe the change, with the version it releases where there is one, for example `Explain scopes in app-only auth (0.7.1)`.
