# fabric-python-helper has been renamed archwaytrust-fabric

This package is now published as [`archwaytrust-fabric`](https://pypi.org/project/archwaytrust-fabric/), and imported as `archwaytrust_fabric`. There will be no more releases under this name.

To move over:

1. In each Fabric environment, under **Public libraries**, replace `fabric_python_helper` with `archwaytrust-fabric`, at the same or a later version.
2. In your notebooks, change `fabric_python_helper` to `archwaytrust_fabric` in every import. Nothing else changes:

   ```python
   from archwaytrust_fabric.auth import KeyVaultCertificateCredential
   from archwaytrust_fabric.graph_api import Emails
   ```

This final release, 0.6.0, contains no code of its own. It installs `archwaytrust-fabric` 0.6 and makes `fabric_python_helper` an alias for it, so existing imports keep working but show a `FutureWarning`. Versions up to 0.5.0 are unchanged and stay installable.

See the [documentation](https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper#readme).
