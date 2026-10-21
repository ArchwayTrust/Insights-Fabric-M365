from .auth import KeyVaultCertificateCredential
from .auth import Scopes
from .graph_api import Emails
from .graph_api import MessageNotFoundError
from .sharepoint import SharePointLibraryManager

__all__ = [
    "KeyVaultCertificateCredential", "Scopes", "Emails", "MessageNotFoundError", "SharePointLibraryManager",
]

__version__="0.6.2"
__author__="Ben Dobbs"
__author_email__="bdobbs@archwaytrust.co.uk"
__description__="Collection of useful python code for automation in Microsoft Fabric."
__url__="https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper"
