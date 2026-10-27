"""
Scopes for common Microsoft APIs.
"""


class Scopes:
    """Scopes for common Microsoft APIs, for use with KeyVaultCertificateCredential or any other credential.

    The comment above each says where to give the service principal access. Entra ID issues a token
    whether or not that access exists, but the API rejects requests made without it.
    """

    # Microsoft Graph. Grant application permissions on the app registration, with admin consent.
    # For mailboxes use Exchange RBAC for Applications instead, which can limit the app to specific mailboxes.
    GRAPH = "https://graph.microsoft.com/.default"

    # Power BI REST API. Add the service principal to the workspace, and allow service principals to use
    # the APIs in the Fabric admin portal tenant settings (Developer settings). Access comes from workspace
    # roles, not from API permissions on the app registration.
    POWER_BI = "https://analysis.windows.net/powerbi/api/.default"

    # Fabric REST API. Access is granted the same way as for Power BI.
    FABRIC = "https://api.fabric.microsoft.com/.default"

    # Azure Storage and OneLake (OneLake only accepts this audience). For OneLake give the service principal
    # a workspace role or item permissions, for storage accounts an Azure RBAC role such as Storage Blob Data Reader.
    STORAGE = "https://storage.azure.com/.default"

    # Azure Key Vault. Azure RBAC role or access policy on the key vault.
    KEY_VAULT = "https://vault.azure.net/.default"

    # Azure SQL, and Fabric warehouses, SQL analytics endpoints and SQL databases over TDS. In Fabric give the
    # service principal a workspace role or item permissions, in Azure SQL create a user FROM EXTERNAL PROVIDER.
    SQL = "https://database.windows.net/.default"

    # Azure Data Explorer and Fabric Eventhouse / KQL databases. Give the service principal a workspace role, or add
    # it as a database principal, for example .add database <db> viewers ('aadapp=<client_id>;<tenant_id>').
    KUSTO = "https://api.kusto.windows.net/.default"

    # Azure Resource Manager, for managing Azure resources. Azure RBAC role on the subscription, resource group or resource.
    AZURE_MANAGEMENT = "https://management.azure.com/.default"

    # Azure Monitor Log Analytics query API. Log Analytics Reader on the Log Analytics workspace.
    LOG_ANALYTICS = "https://api.loganalytics.io/.default"
