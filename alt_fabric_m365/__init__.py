from .auth import KeyVaultCertificateCredential
from .auth import Scopes
from .exchange import Emails
from .exchange import MessageNotFoundError
from .sharepoint import SharePointLibraryManager

# Deprecated alias for exchange. Imported, as in 0.6, so alt_fabric_m365.graph_api and the fabric_python_helper
# alias keep working without an import of their own.
from . import graph_api
_DEPRECATED_MODULES = (graph_api,)

__all__ = [
    "KeyVaultCertificateCredential", "Scopes", "Emails", "MessageNotFoundError", "SharePointLibraryManager",
]

__version__="0.7.2"
__author__="Ben Dobbs"
__author_email__="bdobbs@archwaytrust.co.uk"
__description__="Collection of useful python code for automation in Microsoft Fabric."
__url__="https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper"
