import json
from unittest import mock

import pytest
import requests

import alt_fabric_m365
from alt_fabric_m365.exchange import emails as emails_module
from alt_fabric_m365.auth import Scopes
from alt_fabric_m365.exchange import Emails, MessageNotFoundError

from conftest import make_response

GRAPH = "https://graph.microsoft.com/v1.0"
MAILBOX = "reports@contoso.com"
TIMEOUT = Emails.REQUEST_TIMEOUT_SECONDS

FILE = "#microsoft.graph.fileAttachment"
ITEM = "#microsoft.graph.itemAttachment"
REFERENCE = "#microsoft.graph.referenceAttachment"


@pytest.fixture
def emails(credential):
    return Emails(credential, MAILBOX)


def patch_get(**kwargs):
    return mock.patch.object(emails_module.requests, "get", **kwargs)


def file_attachment(attachment_id, name, is_inline=False):
    return {"@odata.type": FILE, "id": attachment_id, "name": name, "isInline": is_inline}


# Credentials and mailboxes

def test_requests_a_graph_token_for_each_call(emails, credential):
    with patch_get(return_value=make_response(200, {"value": []})) as get:
        emails.get_attachment_ids_and_names("m")
        emails.get_attachment_ids_and_names("m")

    assert credential.scopes == [(Scopes.GRAPH,), (Scopes.GRAPH,)]
    assert get.call_args.kwargs["headers"]["Authorization"] == "Bearer graph-token"


def test_mailbox_is_optional_when_each_call_passes_one(credential):
    emails = Emails(credential)

    with patch_get(return_value=make_response(200, {"value": []})) as get:
        emails.get_attachment_ids_and_names("m", shared_mailbox_email="other@contoso.com")

    assert get.call_args.args[0] == f"{GRAPH}/users/other@contoso.com/messages/m/attachments"


@pytest.mark.parametrize("call", [
    lambda e: e.search_message_by_subject_and_sender("subject", "sender@contoso.com"),
    lambda e: e.get_attachment_ids_and_names("m"),
    lambda e: e.download_attachment("m", "a"),
    lambda e: e.delete_email("m"),
    lambda e: e.send_email("subject", "content", ["to@contoso.com"]),
    lambda e: e._check_if_message_exists("m"),
])
def test_no_mailbox_raises_before_any_request(credential, call):
    with mock.patch.multiple(emails_module.requests, get=mock.DEFAULT, post=mock.DEFAULT, delete=mock.DEFAULT) as mocks:
        with pytest.raises(ValueError, match="shared_mailbox_email"):
            call(Emails(credential, None))

    assert not any(m.called for m in mocks.values())


def test_shared_mailbox_overrides_the_default(emails):
    with patch_get(return_value=make_response(200, {"value": []})) as get:
        # Second positional argument is still shared_mailbox_email
        emails.get_attachment_ids_and_names("m", "other@contoso.com")

    assert "/users/other@contoso.com/" in get.call_args.args[0]


# Searching

def test_search_escapes_quotes_and_returns_the_latest(emails):
    page = {"value": [
        {"id": "older", "receivedDateTime": "2026-01-01T00:00:00Z"},
        {"id": "newer", "receivedDateTime": "2026-02-01T00:00:00Z"},
    ]}
    with patch_get(return_value=make_response(200, page)) as get:
        message_id = emails.search_message_by_subject_and_sender("Bob's report", "o'neil@contoso.com")

    assert message_id == "newer"
    assert get.call_args.args[0] == f"{GRAPH}/users/{MAILBOX}/mailFolders/Inbox/messages"
    assert get.call_args.kwargs["params"]["$filter"] == (
        "subject eq 'Bob''s report' and from/emailAddress/address eq 'o''neil@contoso.com'"
    )


def test_search_follows_pages_and_can_return_all(emails):
    pages = [
        make_response(200, {"value": [{"id": "1", "receivedDateTime": "2026-01-01"}], "@odata.nextLink": "https://next"}),
        make_response(200, {"value": [{"id": "2", "receivedDateTime": "2026-02-01"}]}),
    ]
    with patch_get(side_effect=pages) as get:
        messages = emails.search_message_by_subject_and_sender("s", "a@b.c", only_search_inbox=False, only_return_latest=False)

    assert [m["id"] for m in messages] == ["2", "1"]
    assert get.call_args_list[0].args[0] == f"{GRAPH}/users/{MAILBOX}/messages"
    assert get.call_args_list[1].args[0] == "https://next"
    assert get.call_args_list[1].kwargs["params"] is None


def test_search_with_no_match_raises_message_not_found(emails):
    with patch_get(return_value=make_response(200, {"value": []})):
        with pytest.raises(MessageNotFoundError, match="^No matching messages found.$"):
            emails.search_message_by_subject_and_sender("s", "a@b.c")


def test_message_not_found_is_exported_and_still_an_exception():
    assert alt_fabric_m365.MessageNotFoundError is MessageNotFoundError
    assert issubclass(MessageNotFoundError, Exception)


def test_search_error_is_raised(emails):
    with patch_get(return_value=make_response(403, {"error": {"code": "ErrorAccessDenied"}})):
        with pytest.raises(Exception, match="Status code: 403"):
            emails.search_message_by_subject_and_sender("s", "a@b.c")


# Attachments

def test_only_non_inline_file_attachments_are_listed(emails):
    value = [
        file_attachment("f1", "sales.csv"),
        file_attachment("f2", "logo.png", is_inline=True),
        {"@odata.type": ITEM, "id": "i1", "name": "Fwd: email", "isInline": False},
        {"@odata.type": REFERENCE, "id": "r1", "name": "OneDrive link", "isInline": False},
        {"@odata.type": FILE, "id": "f3", "name": "no-inline-flag.pdf"},
    ]
    with patch_get(return_value=make_response(200, {"value": value})):
        assert emails.get_attachment_ids_and_names("m") == [("f1", "sales.csv"), ("f3", "no-inline-flag.pdf")]
        assert emails.get_attachment_ids_and_names("m", include_inline=True) == [
            ("f1", "sales.csv"), ("f2", "logo.png"), ("f3", "no-inline-flag.pdf")
        ]


def test_attachments_follow_pages(emails):
    pages = [
        make_response(200, {"value": [file_attachment("a", "a.csv")], "@odata.nextLink": "https://next"}),
        make_response(200, {"value": [file_attachment("b", "b.csv")]}),
    ]
    with patch_get(side_effect=pages) as get:
        assert emails.get_attachment_ids_and_names("m") == [("a", "a.csv"), ("b", "b.csv")]

    assert get.call_args_list[1].args[0] == "https://next"


def test_attachment_list_error_is_raised(emails):
    with patch_get(return_value=make_response(404, {"error": {"code": "ErrorItemNotFound"}})):
        with pytest.raises(requests.HTTPError):
            emails.get_attachment_ids_and_names("m")


def test_download_decodes_text_or_returns_bytes(emails):
    attachment = {"@odata.type": FILE, "name": "sales.csv", "contentBytes": "Y29sMSxjb2wyCg=="}
    with patch_get(return_value=make_response(200, attachment)) as get:
        assert emails.download_attachment("m", "a") == "col1,col2\n"
        assert emails.download_attachment("m", "a", is_binary=True) == b"col1,col2\n"

    assert get.call_args.args[0] == f"{GRAPH}/users/{MAILBOX}/messages/m/attachments/a"


def test_download_of_non_file_attachment_raises(emails):
    with patch_get(return_value=make_response(200, {"@odata.type": ITEM, "name": "Fwd: email"})):
        with pytest.raises(Exception, match="only file attachments can be downloaded"):
            emails.download_attachment("m", "a")


# Deleting and sending

def test_delete_email(emails):
    with mock.patch.object(emails_module.requests, "delete", return_value=make_response(204)) as delete:
        assert emails.delete_email("m") is True

    assert delete.call_args.args[0] == f"{GRAPH}/users/{MAILBOX}/messages/m"


def test_delete_failure_raises(emails):
    with mock.patch.object(emails_module.requests, "delete", return_value=make_response(403)):
        with pytest.raises(Exception, match="Status code: 403"):
            emails.delete_email("m")


def test_send_email(emails):
    with mock.patch.object(emails_module.requests, "post", return_value=make_response(202)) as post:
        emails.send_email("Subject", "Body", ["a@contoso.com", "b@contoso.com"])

    assert post.call_args.args[0] == f"{GRAPH}/users/{MAILBOX}/sendMail"
    message = json.loads(post.call_args.kwargs["data"])["message"]
    assert message["subject"] == "Subject"
    assert message["body"] == {"contentType": "Text", "content": "Body"}
    assert [r["emailAddress"]["address"] for r in message["toRecipients"]] == ["a@contoso.com", "b@contoso.com"]


def test_send_email_to_one_address_given_as_a_string(emails):
    with mock.patch.object(emails_module.requests, "post", return_value=make_response(202)) as post:
        emails.send_email("Subject", "Body", "a@contoso.com")

    message = json.loads(post.call_args.kwargs["data"])["message"]
    assert message["toRecipients"] == [{"emailAddress": {"address": "a@contoso.com"}}]


@pytest.mark.parametrize("email_addresses", [[], ""])
def test_send_email_with_no_recipients_raises_before_any_request(emails, email_addresses):
    with mock.patch.object(emails_module.requests, "post") as post:
        with pytest.raises(ValueError, match="at least one recipient"):
            emails.send_email("Subject", "Body", email_addresses)

    assert not post.called


def test_send_failure_raises(emails):
    with mock.patch.object(emails_module.requests, "post", return_value=make_response(403)):
        with pytest.raises(Exception, match="Status code: 403"):
            emails.send_email("Subject", "Body", ["a@contoso.com"])


# Checking a message exists

def test_existing_message(emails):
    with patch_get(return_value=make_response(200, {"id": "m"})) as get:
        assert emails._check_if_message_exists("m") is True

    assert get.call_args.args[0] == f"{GRAPH}/users/{MAILBOX}/messages/m"


def test_missing_message(emails):
    body = {"error": {"code": "ErrorItemNotFound", "message": "The specified object was not found in the store."}}
    with patch_get(return_value=make_response(404, body)):
        assert emails._check_if_message_exists("m") is False


@pytest.mark.parametrize("response, expected", [
    (make_response(404, {"error": {"code": "ErrorInvalidUser", "message": "The requested user is invalid."}}), "ErrorInvalidUser"),
    (make_response(404, {"error": {"code": "MailboxNotEnabledForRESTAPI", "message": "Not supported."}}), "MailboxNotEnabledForRESTAPI"),
    (make_response(400, {"error": {"code": "ErrorInvalidIdMalformed", "message": "Id is malformed."}}), "ErrorInvalidIdMalformed"),
    (make_response(403, {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}}), "ErrorAccessDenied"),
    (make_response(503, text="<html>Service Unavailable</html>"), "Service Unavailable"),
    (make_response(404, ["not", "a", "dict"]), "Status code: 404"),
])
def test_other_errors_are_raised_not_reported_as_missing(emails, response, expected):
    with patch_get(return_value=response):
        with pytest.raises(Exception) as error:
            emails._check_if_message_exists("m")

    assert f"Status code: {response.status_code}" in str(error.value)
    assert expected in str(error.value)


@pytest.mark.parametrize("response, expected", [
    (make_response(404, {"error": {"code": "ErrorItemNotFound", "message": "Not found."}}), ("ErrorItemNotFound", "Not found.")),
    (make_response(404, {"error": "a string, not a Graph error"}), (None, None)),
    (make_response(404, ["not", "a", "dict"]), (None, None)),
    (make_response(503, text="<html>Service Unavailable</html>"), (None, None)),
    (make_response(204), (None, None)),
])
def test_graph_error_reads_only_graph_error_bodies(response, expected):
    assert emails_module._graph_error(response) == expected


# Timeouts

def test_every_request_has_a_timeout(emails):
    pages = {"value": [{"id": "1", "receivedDateTime": "2026-01-01"}]}
    attachment = {"@odata.type": FILE, "name": "a.csv", "contentBytes": "aGk="}
    with mock.patch.object(emails_module.requests, "get", side_effect=[
        make_response(200, pages), make_response(200, {"value": []}), make_response(200, attachment), make_response(200, {"id": "m"}),
    ]) as get, mock.patch.object(emails_module.requests, "delete", return_value=make_response(204)) as delete, \
            mock.patch.object(emails_module.requests, "post", return_value=make_response(202)) as post:
        emails.search_message_by_subject_and_sender("s", "a@b.c")
        emails.get_attachment_ids_and_names("m")
        emails.download_attachment("m", "a")
        emails._check_if_message_exists("m")
        emails.delete_email("m")
        emails.send_email("s", "c", ["a@b.c"])

    for call in get.call_args_list + delete.call_args_list + post.call_args_list:
        assert call.kwargs["timeout"] == TIMEOUT
