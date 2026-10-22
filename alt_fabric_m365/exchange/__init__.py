"""
Work with Exchange Online mailboxes through Microsoft Graph, as an app registration.

Emails finds emails in a mailbox, downloads their attachments, deletes them and sends emails. Give the app
access to specific mailboxes with RBAC for Applications in Exchange Online, see the Emails docstring.
"""
from .emails import Emails, MessageNotFoundError

__all__ = ["Emails", "MessageNotFoundError"]
