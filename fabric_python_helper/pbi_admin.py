import msal
import logging
import requests
import json
import time
from notebookutils import mssparkutils
from .auth import KeyVaultCertificateCredential, Scopes

class AccessTokens:
    def __init__(self):
        # Setting up a logger for this class
        self.logger = logging.getLogger(__name__)

    def get_token_as_service_principal(self, tenant_id, client_id, akv_url, akv_secret_name):
        """
        Authenticates as a service principal using client credentials.

        This method retrieves a client secret from Azure Key Vault and uses it
        to authenticate with Microsoft Azure as a service principal, obtaining
        an access token for Power BI.

        Parameters:
            tenant_id (str): Azure tenant ID.
            client_id (str): Azure client ID.
            akv_url (str): Azure Key Vault URL.
            akv_secret_name (str): Name of the secret in Azure Key Vault.

        Returns:
            str: An access token for Power BI API.

        Raises:
            Exception: If the authentication fails or any other error occurs.
        """
        try:
            # Retrieve the client secret from Azure Key Vault
            client_secret = mssparkutils.credentials.getSecret(akv_url, akv_secret_name)

            # Setting the authority URL for Azure AD
            authority_url = f"https://login.microsoftonline.com/{tenant_id}/"
            # The scope for Power BI API
            scope = ["https://analysis.windows.net/powerbi/api/.default"]
            
            # Creating an MSAL application instance
            app = msal.ConfidentialClientApplication(
                client_id,
                authority=authority_url,
                validate_authority=True,
                client_credential=client_secret
            )
            
            # Acquiring the token for the client
            result = app.acquire_token_for_client(scopes=scope)
            
            # Checking if the access token is in the result
            if "access_token" in result:
                return result["access_token"]
            else:
                # Log and raise an error if no access token is found
                error_msg = f"Error in getAccessToken: {result.get('error')}, {result.get('error_description')}"
                self.logger.error(error_msg)
                raise Exception(error_msg)
        except Exception as e:
            # Log and re-raise any exceptions that occur
            self.logger.error(f"Error in as_service_principal: {str(e)}")
            raise

    def get_token_as_service_principal_with_certificate(self, tenant_id, client_id, akv_url, certificate_name, certificate_version=None):
        """
        Authenticates as a service principal using a certificate held in Azure Key Vault.

        Key Vault signs the client assertion, so the certificate's private key never leaves the key vault.
        See fabric_python_helper.auth.KeyVaultCertificateCredential for the permissions required.

        Parameters:
            tenant_id (str): Azure tenant ID.
            client_id (str): Azure client ID.
            akv_url (str): Azure Key Vault URL.
            certificate_name (str): Name of the certificate in Azure Key Vault.
            certificate_version (str): Optional. Version of the certificate to sign with. Defaults to the latest version.

        Returns:
            str: An access token for Power BI API.

        Raises:
            Exception: If the authentication fails or any other error occurs.
        """
        try:
            credential = KeyVaultCertificateCredential(tenant_id, client_id, akv_url, certificate_name, Scopes.POWER_BI, certificate_version)
            return credential.get_access_token()
        except Exception as e:
            # Log and re-raise any exceptions that occur
            self.logger.error(f"Error in get_token_as_service_principal_with_certificate: {str(e)}")
            raise

    def get_token_as_fabric_notebook_owner(self):
        """
        Retrieves an access token as the notebook owner in Fabric.

        This method obtains an access token for Power BI by using the mssparkutils function, 
        assuming that the notebook is running as a Fabric notebook.

        Returns:
            str: An access token for Power BI API.

        Raises:
            Exception: If retrieving the access token fails.
        """
        try:
            # Get the access token using Databricks utilities
            access_token = mssparkutils.credentials.getToken("pbi")
            return access_token
        except Exception as e:
            # Log and re-raise any exceptions that occur
            self.logger.error(f"Error in as_fabric_notebook_owner: {str(e)}")
            raise

class Dataflows:
    # Base URL for Power BI API calls
    BASE_URL = "https://api.powerbi.com/v1.0/myorg/groups/"

    def __init__(self, workspace_id, dataflow_id, access_token, refresh_body={"notifyOption": "NoNotification"}):
        """
        Initializes the dataflow refresher instance with necessary parameters.

        Parameters:
            workspace_id (str): The ID of the Power BI workspace.
            dataflow_id (str): The ID of the Power BI dataflow.
            access_token (str): Access token for Power BI API authentication.
            refresh_body (dict): The request body for the refresh. Defaults to no notification.
        """
        self.workspace_id = workspace_id
        self.dataflow_id = dataflow_id
        self.access_token = access_token
        self.refresh_body = refresh_body

    def _send_refresh_request(self):
        """
        Sends a refresh request to the Power BI API for the specified dataflow.

        Returns:
            Response: The response object from the requests.post call.
        """
        # Constructing the refresh endpoint URL
        refresh_endpoint = f"{self.BASE_URL}{self.workspace_id}/dataflows/{self.dataflow_id}/refreshes"
        
        # Setting up the request headers
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        # Making a POST request to start the refresh
        response = requests.post(refresh_endpoint, headers=headers, json=self.refresh_body)
        response.raise_for_status()  # Raises an HTTPError if the HTTP request returned an unsuccessful status code
        
        return response

    def _handle_refresh_response(self, response):
        """
        Handles the response from the refresh request.

        Parameters:
            response (Response): The response object from the refresh request.

        Returns:
            str: The result of the refresh request handling.
        """
        if response.ok:
            print("Refresh requested successfully.")
            return "Refresh started"
        elif response.status_code == 404:
            error_message = "Resource Not Found. Check the WorkspaceId and DataflowId."
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")
        else:
            error_message = f"Request failed with status code: {response.status_code}"
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")

    def _get_transaction_status(self):
        """
        Gets the transaction status of the dataflow refresh.

        Returns:
            dict: The JSON response from the transaction status request.
        """
        # Constructing the transaction endpoint URL
        transaction_endpoint = f"{self.BASE_URL}{self.workspace_id}/dataflows/{self.dataflow_id}/transactions"
        
        # Setting up the request headers
        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }

        # Making a GET request to retrieve the transaction status
        response = requests.get(transaction_endpoint, headers=headers)
        response.raise_for_status()
        
        return response

    def _handle_transaction_response(self, response):
        """
        Handles the transaction response from the Power BI API.

        Parameters:
            response (Response): The response object from the transaction request.

        Returns:
            dict: The JSON response containing the transaction status.
        """
        try:
            if response.ok:
                return response.json()
            else:
                error_message = f"Transaction Request failed with status code: {response.status_code}"
                print(error_message)
                raise Exception(f"Refresh Request Failed: {error_message}")
        except json.JSONDecodeError:
            error_message = f"Invalid JSON response. Status Code: {response.status_code}. Response Text: {response.text}"
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")

    def refresh_dataflow(self, expected_duration=30, wait_for_completion=True, loop_interval=10):
        """
        Initiates and optionally waits for the completion of a dataflow refresh.

        Parameters:
            expected_duration (int): Expected duration to wait before checking the status. Defaults to 30 seconds.
            wait_for_completion (bool): Whether to wait for the completion of the refresh. Defaults to True.
            loop_interval (int): Interval between status checks if waiting for completion. Defaults to 10 seconds.

        Returns:
            str: The status of the dataflow refresh or error message.
        """
        try:
            # Sending a refresh request
            refresh_response = self._send_refresh_request()
            result = self._handle_refresh_response(refresh_response)
            
            # If not waiting for completion, return the initial result
            if not wait_for_completion:
                print("Refresh has started but completion won't be checked.")
                return result
            
            print("Waiting for expected duration...")
            # Wait for the expected duration before checking the status
            time.sleep(expected_duration)
            
            # Loop to check the transaction status
            while True:
                transaction_status = self._get_transaction_status()
                result = self._handle_transaction_response(transaction_status)
                
                # Check the status of the refresh
                status = result['value'][0]['status']
                if status != "InProgress":
                    print(f"Dataflow refresh completed with status: {status}")
                    return status
                
                print("Refresh is in progress. Waiting to check again...")
                time.sleep(loop_interval)
        except Exception as e:
            return str(e)

class SemanticModels:
    # Base URL for Power BI API calls
    BASE_URL = "https://api.powerbi.com/v1.0/myorg/datasets/"

    def __init__(self, semantic_model_id, access_token, refresh_body={"notifyOption": "NoNotification"}):
        """
        Initializes the RefreshSemanticModel instance.

        Parameters:
            semantic_model_id (str): The ID of the semantic model to be refreshed.
            access_token (str): Access token for Power BI API authentication.
            refresh_body (dict): The request body for the refresh, defaults to no notification.
        """
        self.semantic_model_id = semantic_model_id
        self.access_token = access_token
        self.refresh_body = refresh_body

    def _send_refresh_request(self):
        """
        Sends a refresh request to the Power BI API for the specified semantic model.

        Returns:
            Response: The response object from the requests.post call.
        """
        # Constructing the refresh endpoint URL
        refresh_endpoint = f"{self.BASE_URL}{self.semantic_model_id}/refreshes"
        
        # Setting up the request headers
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        # Making a POST request to start the refresh
        response = requests.post(refresh_endpoint, headers=headers, json=self.refresh_body)
        response.raise_for_status()  # Raises an HTTPError for failed requests
        
        return response

    def _handle_refresh_trigger_response(self, response):
        """
        Handles the response from the refresh request.

        Parameters:
            response (Response): The response object from the refresh request.

        Returns:
            str: The result of handling the refresh request.
        """
        if response.ok:
            print("Refresh requested successfully.")
            return "Refresh started"
        elif response.status_code == 404:
            error_message = "Resource Not Found. Check the Semantic Model ID."
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")
        else:
            error_message = f"Request failed with status code: {response.status_code}"
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")

    def _get_refresh_status(self):
        """
        Gets the refresh status of the semantic model.

        Returns:
            dict: The JSON response from the refresh status request.
        """
        # Constructing the refresh status endpoint URL
        refresh_endpoint = f"{self.BASE_URL}{self.semantic_model_id}/refreshes"
        
        # Setting up the request headers
        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }

        # Making a GET request to retrieve the refresh status
        response = requests.get(refresh_endpoint, headers=headers)
        response.raise_for_status()
        
        return response

    def _handle_refresh_status_response(self, response):
        """
        Handles the refresh status response from the Power BI API.

        Parameters:
            response (Response): The response object from the refresh status request.

        Returns:
            dict: The JSON response containing the refresh status.
        """
        try:
            if response.ok:
                return response.json()
            else:
                error_message = f"Refresh Status Request failed with status code: {response.status_code}"
                print(error_message)
                raise Exception(f"Refresh Request Failed: {error_message}")
        except json.JSONDecodeError:
            error_message = f"Invalid JSON response. Status Code: {response.status_code}. Response Text: {response.text}"
            print(error_message)
            raise Exception(f"Refresh Request Failed: {error_message}")

    def refresh_semantic_model(self, expected_duration=30, wait_for_completion=True, loop_interval=10):
        """
        Initiates and optionally waits for the completion of a semantic model refresh.

        Parameters:
            expected_duration (int): Expected duration to wait before checking the status. Defaults to 30 seconds.
            wait_for_completion (bool): Whether to wait for the completion of the refresh. Defaults to True.
            loop_interval (int): Interval between status checks if waiting for completion. Defaults to 10 seconds.

        Returns:
            str: The status of the semantic model refresh or an error message.
        """
        try:
            # Sending a refresh request
            refresh_response = self._send_refresh_request()
            result = self._handle_refresh_trigger_response(refresh_response)
            
            # If not waiting for completion, return the initial result
            if not wait_for_completion:
                print("Refresh has started but completion won't be checked.")
                return result
            print("Waiting for expected duration...")
            # Wait for the expected duration before checking the status
            time.sleep(expected_duration)
            
            # Loop to check the refresh status
            while True:
                refresh_status = self._get_refresh_status()
                result = self._handle_refresh_status_response(refresh_status)
                
                # Check the status of the refresh
                status = result["value"][0]["status"]
                if status != "Unknown":
                    print(f"Semantic model refresh completed with status: {status}")
                    return status
                
                print("Refresh is in progress. Waiting to check again...")
                time.sleep(loop_interval)
        except Exception as e:
            return str(e)