# fabric-python-helper has been renamed ALT-Fabric-M365

This package is now published as [`alt-fabric-m365`](https://pypi.org/project/alt-fabric-m365/), and imported as `alt_fabric_m365`. There will be no more releases under this name.

To move over:

1. In each Fabric environment, under **Public libraries**, replace `fabric_python_helper` with `alt-fabric-m365`, at the same or a later version.
2. In your notebooks, change `fabric_python_helper` to `alt_fabric_m365` in every import. Nothing else changes:

   ```python
   from alt_fabric_m365.auth import KeyVaultCertificateCredential
   from alt_fabric_m365.graph_api import Emails
   ```

This final release, 0.6.0, contains no code of its own. It installs `alt-fabric-m365` 0.6 and makes `fabric_python_helper` an alias for it, so existing imports keep working but show a `FutureWarning`. Versions up to 0.5.0 are unchanged and stay installable.

See the [documentation](https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper#readme).
