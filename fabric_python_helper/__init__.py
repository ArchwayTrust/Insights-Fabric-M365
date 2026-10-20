from .auth import KeyVaultCertificateCredential
from .auth import PemCertificateCredential
from .auth import Scopes
from .graph_api import Emails
from .graph_api import MessageNotFoundError
from .sharepoint import SharePointLibraryManager

__all__ = [
    "KeyVaultCertificateCredential", "PemCertificateCredential", "Scopes",
    "Emails", "MessageNotFoundError", "SharePointLibraryManager",
]

__version__="0.5.0"
__author__="Ben Dobbs"
__author_email__="bdobbs@archwaytrust.co.uk"
__description__="Collection of useful python code for automation in Microsoft Fabric."
__url__="https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper"
