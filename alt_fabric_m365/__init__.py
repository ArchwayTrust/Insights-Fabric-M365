from .auth import KeyVaultCertificateCredential
from .auth import Scopes
from .exchange import Emails
from .exchange import MessageNotFoundError
from .sharepoint import SharePointLibraryManager

__all__ = [
    "KeyVaultCertificateCredential", "Scopes", "Emails", "MessageNotFoundError", "SharePointLibraryManager",
]

__version__="1.3.0"
__author__="Ben Dobbs"
__author_email__="bdobbs@archwaytrust.co.uk"
__description__="Collection of useful python code for automation in Microsoft Fabric."
__url__="https://github.com/ArchwayTrust/Insights-Fabric-M365"
