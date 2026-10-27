import requests
import json
import base64
from ..auth import Scopes


class MessageNotFoundError(Exception):
    """
    Raised when a search finds no emails matching the given criteria.

    Subclasses Exception, so existing code that catches Exception still works.
    """


def _graph_error(response):
    """
    Returns the error code and message from a Microsoft Graph error response.

    Either value is None if the response body isn't a Graph error, such as an empty or non-JSON body.
    """
    try:
        body = response.json()
    except ValueError:
        return None, None

    error = body.get('error') if isinstance(body, dict) else None
    if not isinstance(error, dict):
        return None, None

    return error.get('code'), error.get('message')


class Emails:
    """
    A class to interact with Microsoft Graph API for managing emails in a mailbox.

    Authenticates as an app registration (service principal), using a credential with a get_token method:
        - alt_fabric_m365.auth.KeyVaultCertificateCredential, which keeps the certificate's private key in Azure Key Vault (recommended).
        - azure.identity.ClientSecretCredential, with the client secret read from Azure Key Vault.

    The app needs access to the mailbox. Use Exchange RBAC for Applications to limit it to specific mailboxes, assigning:
        - Application Mail.Read (or Mail.ReadWrite) to search emails and download attachments.
        - Application Mail.ReadWrite to delete emails.
        - Application Mail.Send to send emails.
    Don't grant Mail.Read, Mail.ReadWrite or Mail.Send as Entra ID API permissions, as those apply to every mailbox in the tenant.

    Attributes:
        credential (TokenCredential): Credential used to get Microsoft Graph access tokens.
        mailbox (str or None): Email address of the default mailbox, used unless a method is given a different one.
                               Optional if every call passes shared_mailbox_email.
    """

    GRAPH_URL = "https://graph.microsoft.com/v1.0"

    # Seconds to wait for Microsoft Graph to connect or send data before giving up, so a stalled request doesn't hang the notebook
    REQUEST_TIMEOUT_SECONDS = 30

    def __init__(self, credential, mailbox=None):
        """
        Initializes the Emails class, optionally with a default mailbox.

        Parameters:
            credential (TokenCredential): Credential used to get Microsoft Graph access tokens.
            mailbox (str): Optional. Email address of the default mailbox, used unless a method is given a different one.
                           Can be left out if every call passes shared_mailbox_email.
        """
        self.credential = credential
        self.mailbox = mailbox

    def _headers(self):
        """
        Returns headers for calling Microsoft Graph with a current access token.

        The credential caches tokens, so this is called for each request rather than storing a token
        that would expire during long runs.
        """
        access_token = self.credential.get_token(Scopes.GRAPH).token

        return {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

    def _mailbox_url(self, shared_mailbox_email=None):
        """
        Returns the Microsoft Graph URL for the given mailbox, or for the default mailbox if none is given.

        Raises:
            ValueError: If no mailbox is given and there is no default mailbox.
        """
        mailbox = shared_mailbox_email or self.mailbox
        if not mailbox:
            raise ValueError("No mailbox given. Pass shared_mailbox_email or set a default mailbox when creating Emails.")

        return f"{self.GRAPH_URL}/users/{mailbox}"

    def search_message_by_subject_and_sender(self, subject, sender_email, only_search_inbox=True, only_return_latest=True, shared_mailbox_email=None):
        """
        Searches for an email by its subject and sender's email using Microsoft Graph API with pagination.

        This method queries the Microsoft Graph API to find all emails that match
        the given subject and sender's email address, handling pagination to gather all results.

        Parameters:
            subject (str): The subject of the email to search for.
            sender_email (str): The email address of the sender of the email.
            only_search_inbox (bool): Defaults to True. Limits search to only the main inbox.
            only_return_latest (bool): Defaults to True. If True, returns the ID of the latest message.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to search in.

        Returns:
            str or list: The ID of the latest email or a list of emails that match the search criteria.

        Raises:
            MessageNotFoundError: If no matching messages are found.
            Exception: If the Microsoft Graph request fails.
        """
        mailbox_url = self._mailbox_url(shared_mailbox_email)
        endpoint = f"{mailbox_url}/mailFolders/Inbox/messages" if only_search_inbox else f"{mailbox_url}/messages"

        # Single quotes are escaped by doubling them in OData string literals
        subject_literal = subject.replace("'", "''")
        sender_literal = sender_email.replace("'", "''")

        query_parameters = {
            '$filter': f"subject eq '{subject_literal}' and from/emailAddress/address eq '{sender_literal}'",
            '$select': 'id,subject,from,receivedDateTime,parentFolderId,hasAttachments'
        }

        messages = []
        while endpoint:
            # Get headers for each page, so a token expiring part way through is refreshed
            response = requests.get(endpoint, headers=self._headers(), params=query_parameters, timeout=self.REQUEST_TIMEOUT_SECONDS)
            if response.status_code != 200:
                # Handle unsuccessful API call
                raise Exception(f"Error searching message by subject. Status code: {response.status_code}, Response: {response.text}")

            data = response.json()
            messages.extend(data.get('value', []))

            # Update the endpoint for the next page, if any
            endpoint = data.get('@odata.nextLink')
            query_parameters = None  # Ensure subsequent requests do not repeat the initial query parameters

        if messages:
            sorted_messages = sorted(messages, key=lambda x: x['receivedDateTime'], reverse=True)

            if only_return_latest:
                return sorted_messages[0]['id']
            else:
                return sorted_messages

        raise MessageNotFoundError("No matching messages found.")

    def get_attachment_ids_and_names(self, message_id, shared_mailbox_email=None, include_inline=False):
        """
        Retrieves a list of file attachment IDs and names for a specified email message.

        This method uses the Microsoft Graph API to fetch the attachments for a given
        email message, handling pagination. It returns a list of tuples, each containing
        the ID and name of a file attachment that can be passed to download_attachment.

        Only file attachments are returned. Item attachments (attached emails, events or contacts)
        and reference attachments (links to cloud files, such as OneDrive) are skipped, as they
        can't be downloaded. Inline attachments, such as images in an email signature, are also
        skipped unless include_inline is True.

        Parameters:
            message_id (str): The ID of the email message whose attachments are to be retrieved.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to retrieve attachments from.
            include_inline (bool): Defaults to False. If True, inline file attachments are also returned.

        Returns:
            list of tuple: A list of tuples, where each tuple contains the ID and name of a file attachment.
                           An empty list means the email has no downloadable attachments.

        Raises:
            requests.HTTPError: If the Microsoft Graph request fails.
        """
        # Constructing the URL for the Microsoft Graph API attachments endpoint
        endpoint = f'{self._mailbox_url(shared_mailbox_email)}/messages/{message_id}/attachments'

        # Initializing an empty list to store attachment information
        attachments_info = []
        while endpoint:
            # Performing a GET request to retrieve a page of attachments, raising if unsuccessful
            response = requests.get(endpoint, headers=self._headers(), timeout=self.REQUEST_TIMEOUT_SECONDS)
            response.raise_for_status()

            data = response.json()
            for attachment in data.get('value', []):
                # Skip attached emails and links to cloud files, as only file attachments have content to download
                if attachment.get('@odata.type') != '#microsoft.graph.fileAttachment':
                    continue

                # Skip inline attachments, such as signature logos, unless asked for
                if attachment.get('isInline') and not include_inline:
                    continue

                attachments_info.append((attachment['id'], attachment['name']))

            # Update the endpoint for the next page, if any
            endpoint = data.get('@odata.nextLink')

        return attachments_info

    def download_attachment(self, message_id, attachment_id, is_binary=False, encoding='utf-8', shared_mailbox_email=None):
        """
        Downloads an attachment from an email using the Microsoft Graph API.

        This method fetches an attachment from a specific email. It can handle both
        binary files (like images or PDFs) and text-based files (like CSV or TXT).

        Parameters:
            message_id (str): The ID of the message from which the attachment is to be downloaded.
            attachment_id (str): The ID of the attachment to be downloaded.
            is_binary (bool): Flag indicating whether the attachment is a binary file. Default is False.
            encoding (str): The encoding used to decode the attachment's content. Ignored if is_binary is True. Default is 'utf-8'.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to retrieve the attachment from.

        Returns:
            bytes or str: The content of the attachment, either as a byte string (for binary files) or a decoded string (for text files).
        """
        # Constructing the URL to access the attachment via Microsoft Graph API
        attachment_url = f'{self._mailbox_url(shared_mailbox_email)}/messages/{message_id}/attachments/{attachment_id}'

        # Making a GET request to fetch the attachment and check success
        attachment_response = requests.get(attachment_url, headers=self._headers(), timeout=self.REQUEST_TIMEOUT_SECONDS)
        attachment_response.raise_for_status()

        # Extracting the attachment content from the response
        attachment = attachment_response.json()

        # Only file attachments have content. Attached emails and links to cloud files don't.
        if "contentBytes" not in attachment:
            raise Exception(f"Attachment {attachment.get('name')} is a {attachment.get('@odata.type')}, only file attachments can be downloaded.")

        attachment_content = base64.b64decode(attachment["contentBytes"])

        # Returning the attachment content based on its type
        return attachment_content if is_binary else attachment_content.decode(encoding)

    def delete_email(self, message_id, shared_mailbox_email=None):
        """
        Deletes an email using the Microsoft Graph API.

        This method sends a DELETE request to the Microsoft Graph API to remove an
        email message from the mailbox. It uses the mailbox and the message's ID
        to identify the specific email to be deleted.

        Requires the app to have Application Mail.ReadWrite on the mailbox.

        Parameters:
            message_id (str): The ID of the message to be deleted.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to delete the email from.

        Returns:
            bool: True if the email was successfully deleted. Raises an exception otherwise.
        """
        # Constructing the URL for the Microsoft Graph API delete message endpoint
        delete_url = f'{self._mailbox_url(shared_mailbox_email)}/messages/{message_id}'

        # Sending the DELETE request
        response = requests.delete(delete_url, headers=self._headers(), timeout=self.REQUEST_TIMEOUT_SECONDS)

        # Check if the request was successful
        if response.status_code == 204:
            print("Email deleted successfully.")
            return True
        else:
            raise Exception(f"Failed to delete email. Status code: {response.status_code}, Response: {response.text}")

    def send_email(self, subject, content, email_addresses, shared_mailbox_email=None):
        """
        Sends an email from the mailbox using the Microsoft Graph API.

        Requires the app to have Application Mail.Send on the mailbox.

        Parameters:
            subject (str): Email subject.
            content (str): Email content, as plain text.
            email_addresses (list): A list of recipient email addresses.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to send from.
        """
        url = f'{self._mailbox_url(shared_mailbox_email)}/sendMail'

        # Produce recipients section in correct format.
        recipients = []
        for address in email_addresses:
            recipients.append({"emailAddress": {"address": address}})

        # JSON Representing the email.
        email = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": content
                },
                "toRecipients": recipients
            },
            "saveToSentItems": "true"
        }

        # Post the email.
        response = requests.post(url, headers=self._headers(), data=json.dumps(email), timeout=self.REQUEST_TIMEOUT_SECONDS)

        if response.status_code == 202:
            print("Email sent successfully!")
        else:
            print(f"Failed to send email. Status code: {response.status_code}")
            raise Exception(f"Failed to send email. Status code: {response.status_code}")

    def _check_if_message_exists(self, message_id, shared_mailbox_email=None):
        """
        Checks if an email message exists in the mailbox.

        Only a 404 with the Graph error code ErrorItemNotFound means the message doesn't exist.
        Other errors are raised rather than reported as a missing message, such as a mailbox that
        doesn't exist or can't be reached (404 ErrorInvalidUser or MailboxNotEnabledForRESTAPI),
        a malformed message ID (400) or no access to the mailbox (403).

        Parameters:
            message_id (str): The ID of the message to be checked.
            shared_mailbox_email (str): Optional. The email address of a different mailbox to check the message in.

        Returns:
            bool: True if the message exists, False if it doesn't.

        Raises:
            Exception: If the check fails for any other reason. The message includes the status code and Graph error.
        """
        # Constructing the URL for the Microsoft Graph API message endpoint
        message_url = f'{self._mailbox_url(shared_mailbox_email)}/messages/{message_id}'

        # Making a GET request to the Graph API, only selecting the ID to keep the response small
        response = requests.get(message_url, headers=self._headers(), params={'$select': 'id'}, timeout=self.REQUEST_TIMEOUT_SECONDS)

        if response.ok:
            return True

        error_code, error_message = _graph_error(response)

        # The mailbox was found, but the message wasn't
        if response.status_code == 404 and error_code == 'ErrorItemNotFound':
            return False

        raise Exception(
            f"Failed to check for message existence. Status code: {response.status_code}, "
            f"Error code: {error_code}, Message: {error_message or response.text}"
        )
