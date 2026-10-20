"""
An in-memory fake of the Microsoft Graph and SharePoint REST endpoints SharePointLibraryManager calls.

It replaces requests.request, so the real request building, retry wrapper and response handling all run,
and keeps state (groups, members, libraries, role assignments, navigation, sharing settings) so tests can
check what an operation changed, and that running it again changes nothing.

Like the real services, it parses query strings out of the URL it is given, so a value that isn't encoded
properly (for example an '&' in a $filter) is misread here too.
"""
import json
import re
import threading
import uuid
from urllib.parse import parse_qs, unquote, urlparse

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
HOST = "contoso.sharepoint.com"
APP_ID = "aaaaaaaa-0000-0000-0000-00000000a99a"  # the app's service principal object ID

ROLE_DEFINITIONS = {"Read": 1073741826, "Contribute": 1073741827, "Full Control": 1073741829}
# RoleTypeKind (SP.RoleType) of each built-in role definition
ROLE_TYPES = {"Read": 2, "Contribute": 3, "Full Control": 5}

USER = "#microsoft.graph.user"
GROUP = "#microsoft.graph.group"
SERVICE_PRINCIPAL = "#microsoft.graph.servicePrincipal"


def response(status_code, body=None, headers=None):
    resp = requests.Response()
    resp.status_code = status_code
    resp._content = json.dumps(body).encode() if body is not None else b""
    resp.headers.update(headers or {})
    resp.url = "https://example.invalid/"
    return resp


def graph_error(status_code, code, message=""):
    return response(status_code, {"error": {"code": code, "message": message}})


class FakeSite:
    def __init__(self, url):
        self.url = url.rstrip("/")
        self.path = urlparse(self.url).path
        self.id = f"{HOST},{uuid.uuid4()},{uuid.uuid4()}"
        self.lists = {}  # list id -> dict
        self.principals = {}  # claims logon name -> principal id
        self.next_principal_id = 10
        self.web = {"MembersCanShare": True, "RequestAccessEmail": "owner@contoso.com", "UseAccessRequestDefault": False}
        self.member_group = {"Id": 5, "AllowMembersEditMembership": True}
        self.nav = []  # top-level nodes: {"Id", "Title", "Url", "IsExternal", "Children": [...]}
        self.next_nav_id = 1000
        # Set to make sharing writes succeed without changing anything, as SharePoint sometimes does
        self.ignore_sharing_writes = False

    def add_list(self, display_name, slug=None, unique=False, template="documentLibrary"):
        list_id = str(uuid.uuid4())
        slug = slug or display_name
        self.lists[list_id] = {
            "id": list_id,
            "displayName": display_name,
            "slug": slug,
            "template": template,
            "unique": unique,
            "role_assignments": {},  # principal id -> set of role definition ids
            "properties": {},
        }
        return self.lists[list_id]

    def list_by_slug(self, slug):
        for lst in self.lists.values():
            if lst["slug"].lower() == slug.lower() and lst["template"] == "documentLibrary":
                return lst
        return None

    def web_url(self, lst):
        # Ordinary lists live under /Lists/, document libraries at the site root
        folder = "" if lst["template"] == "documentLibrary" else "Lists/"
        return f"{self.url}/{folder}{requests.utils.quote(lst['slug'])}"

    def nav_node(self, title, url, is_external=False):
        node = {"Id": self.next_nav_id, "Title": title, "Url": url, "IsExternal": is_external, "Children": []}
        self.next_nav_id += 1
        return node

    def find_nav_node(self, node_id):
        for node in self.nav:
            if node["Id"] == node_id:
                return node, self.nav
            for child in node["Children"]:
                if child["Id"] == node_id:
                    return child, node["Children"]
        return None, None


class FakeTenant:
    """
    Install with install(monkeypatch). Every request is recorded in calls as (method, url, kwargs).
    """

    def __init__(self):
        self.calls = []
        self.sites = {}
        self.groups = {}  # id -> group dict with a "members" dict of member id -> directory object
        self.users = {}  # id -> user dict
        self.shortcuts = []  # (user id, body)
        self.failures = []  # [method, url regex, response, remaining times]
        self.list_page_size = None  # set to page Graph list results
        self._lock = threading.RLock()  # operations run in parallel with max_workers > 1

    # Setup helpers

    def install(self, monkeypatch):
        from fabric_python_helper.sharepoint import _http
        monkeypatch.setattr(_http.requests, "request", self.request)
        return self

    def add_site(self, path):
        site = FakeSite(f"https://{HOST}{path}")
        self.sites[site.url.lower()] = site
        return site

    def add_user(self, upn, mail=None, proxy_addresses=()):
        user = {
            "@odata.type": USER, "id": str(uuid.uuid4()), "userPrincipalName": upn,
            "mail": mail or upn, "proxyAddresses": [f"SMTP:{mail or upn}", *[f"smtp:{a}" for a in proxy_addresses]],
        }
        self.users[user["id"]] = user
        return user

    def add_group(self, display_name, security_enabled=True, group_types=(), owners=()):
        group = {
            "@odata.type": GROUP, "id": str(uuid.uuid4()), "displayName": display_name, "description": "",
            "mailNickname": "x", "securityEnabled": security_enabled, "groupTypes": list(group_types),
            "members": {}, "owners": set(owners),
        }
        self.groups[group["id"]] = group
        return group

    def add_service_principal(self):
        return {"@odata.type": SERVICE_PRINCIPAL, "id": str(uuid.uuid4())}

    def group_named(self, display_name):
        matches = [g for g in self.groups.values() if g["displayName"] == display_name]
        assert len(matches) <= 1, f"duplicate groups named {display_name}"
        return matches[0] if matches else None

    def member_upns(self, display_name):
        group = self.group_named(display_name)
        return sorted(m.get("userPrincipalName") or m["id"] for m in group["members"].values())

    def fail(self, method, url_pattern, resp, times=1):
        """Makes the next `times` requests matching method and url_pattern return resp."""
        self.failures.append([method, re.compile(url_pattern), resp, times])

    def writes(self):
        """Requests that would change something."""
        return [c for c in self.calls if self._effective_method(c[0], c[2]) != "GET"]

    # Dispatch

    @staticmethod
    def _effective_method(method, kwargs):
        override = (kwargs.get("headers") or {}).get("X-HTTP-Method")
        return override or method.upper()

    def request(self, method, url, **kwargs):
        with self._lock:
            return self._request(method, url, **kwargs)

    def _request(self, method, url, **kwargs):
        self.calls.append((method.upper(), url, kwargs))
        effective = self._effective_method(method, kwargs)

        for failure in self.failures:
            if failure[3] > 0 and failure[0] == effective and failure[1].search(url):
                failure[3] -= 1
                return failure[2]

        parsed = urlparse(url)
        query = {k: v[0] for k, v in parse_qs(parsed.query).items()}
        query.update(kwargs.get("params") or {})
        body = kwargs.get("json")

        if url.startswith(GRAPH):
            return self._graph(effective, unquote(parsed.path[len("/v1.0"):]), query, body, url)

        if parsed.hostname == HOST:
            if "/_vti_bin/client.svc/ProcessQuery" in parsed.path:
                site = self._site_for(parsed.path.split("/_vti_bin/")[0])
                return self._process_query(site, kwargs["data"].decode("utf-8"))
            if "/_api/" in parsed.path:
                site_path, api_path = parsed.path.split("/_api/", 1)
                site = self._site_for(site_path)
                if site is None:
                    return response(404)
                return self._sharepoint(effective, site, unquote(api_path), query, body)

        raise AssertionError(f"Unexpected request: {method} {url}")

    def _site_for(self, path):
        return self.sites.get(f"https://{HOST}{path}".rstrip("/").lower())

    # Microsoft Graph

    def _graph(self, method, path, query, body, url):
        m = re.fullmatch(r"/sites/([^:/]+):(/.*)", path)
        if m and method == "GET":
            site = self.sites.get(f"https://{m.group(1)}{m.group(2)}".rstrip("/").lower())
            return response(200, {"id": site.id}) if site else graph_error(404, "itemNotFound")

        m = re.fullmatch(r"/sites/([^/]+)/lists(?:/([^/]+))?", path)
        if m:
            site = next(s for s in self.sites.values() if s.id == m.group(1))
            if method == "GET" and m.group(2) is None:
                values = [
                    {"id": lst["id"], "displayName": lst["displayName"], "webUrl": site.web_url(lst),
                     "list": {"template": lst["template"]}}
                    for lst in site.lists.values()
                ]
                skip = int(query.get("$skiptoken", 0))
                if self.list_page_size and skip + self.list_page_size < len(values):
                    page = values[skip:skip + self.list_page_size]
                    next_link = f"{GRAPH}/sites/{site.id}/lists?$skiptoken={skip + self.list_page_size}"
                    return response(200, {"value": page, "@odata.nextLink": next_link})
                return response(200, {"value": values[skip:]})
            if method == "POST" and m.group(2) is None:
                assert body["list"] == {"template": "documentLibrary"}
                lst = site.add_list(body["displayName"])
                return response(201, {"id": lst["id"], "displayName": lst["displayName"], "webUrl": site.web_url(lst)})
            if method == "PATCH":
                site.lists[m.group(2)]["displayName"] = body["displayName"]
                return response(200, {})

        if path == "/groups" and method == "GET":
            return self._find_groups(query)

        if path == "/groups" and method == "POST":
            assert body["securityEnabled"] is True and body["mailEnabled"] is False
            assert re.fullmatch(r"[a-z0-9]{1,64}", body["mailNickname"])
            owners = {ref.rsplit("/", 1)[-1] for ref in body.get("owners@odata.bind", [])}
            group = self.add_group(body["displayName"], owners=owners)
            group.update(description=body["description"], mailNickname=body["mailNickname"])
            return response(201, self._public(group))

        m = re.fullmatch(r"/groups/([^/]+)", path)
        if m:
            group = self.groups.get(m.group(1))
            if group is None:
                return graph_error(404, "Request_ResourceNotFound")
            if method == "GET":
                return response(200, self._public(group))
            if method == "PATCH":
                group.update(body)
                return response(204)

        m = re.fullmatch(r"/groups/([^/]+)/owners", path)
        if m and method == "GET":
            return response(200, {"value": [{"id": owner} for owner in sorted(self.groups[m.group(1)]["owners"])]})
        m = re.fullmatch(r"/groups/([^/]+)/owners/\$ref", path)
        if m and method == "POST":
            owner = body["@odata.id"].rsplit("/", 1)[-1]
            if owner in self.groups[m.group(1)]["owners"]:
                return graph_error(400, "Request_BadRequest", "One or more added object references already exist")
            self.groups[m.group(1)]["owners"].add(owner)
            return response(204)

        m = re.fullmatch(r"/groups/([^/]+)/members", path)
        if m and method == "GET":
            members = []
            for member in self.groups[m.group(1)]["members"].values():
                # $select=id,userPrincipalName: Graph returns the type, and the UPN only for users
                members.append({k: v for k, v in member.items() if k in ("@odata.type", "id", "userPrincipalName")})
            return response(200, {"value": members})

        m = re.fullmatch(r"/groups/([^/]+)/members/\$ref", path)
        if m and method == "POST":
            group = self.groups.get(m.group(1))
            if group is None:
                return graph_error(404, "Request_ResourceNotFound")
            member_id = body["@odata.id"].rsplit("/", 1)[-1]
            if member_id in group["members"]:
                return graph_error(400, "Request_BadRequest", "One or more added object references already exist")
            group["members"][member_id] = self.users.get(member_id) or self._public(self.groups[member_id])
            return response(204)

        m = re.fullmatch(r"/groups/([^/]+)/members/([^/]+)/\$ref", path)
        if m and method == "DELETE":
            if self.groups[m.group(1)]["members"].pop(m.group(2), None) is None:
                return graph_error(404, "Request_ResourceNotFound")
            return response(204)

        m = re.fullmatch(r"/users/([^/]+)/drive/root/children", path)
        if m and method == "POST":
            self.shortcuts.append((m.group(1), body))
            return response(201, {"id": "shortcut"})

        m = re.fullmatch(r"/users/(.+)", path)
        if m and method == "GET":
            upn = m.group(1)
            if " " in upn:
                return graph_error(400, "Request_BadRequest", "Invalid object identifier")
            for user in self.users.values():
                if user["userPrincipalName"].lower() == upn.lower():
                    return response(200, {"id": user["id"], "userPrincipalName": user["userPrincipalName"]})
            return graph_error(404, "Request_ResourceNotFound")

        if path == "/users" and method == "GET":
            return self._find_users(query["$filter"])

        raise AssertionError(f"Unexpected Graph request: {method} {url}")

    @staticmethod
    def _public(group):
        return {k: v for k, v in group.items() if k not in ("members", "owners")}

    def _find_groups(self, query):
        flt = query.get("$filter", "")
        m = re.fullmatch(r"displayName eq '((?:[^']|'')*)'", flt)
        if m:
            # Graph compares display names case-insensitively
            name = m.group(1).replace("''", "'").lower()
            return response(200, {"value": [
                self._public(g) for g in self.groups.values() if g["displayName"].lower() == name
            ]})
        m = re.fullmatch(r"startswith\(displayName, '((?:[^']|'')*)'\)", flt)
        if m:
            prefix = m.group(1).replace("''", "'").lower()
            return response(200, {"value": [
                self._public(g) for g in self.groups.values() if g["displayName"].lower().startswith(prefix)
            ]})
        return graph_error(400, "Request_UnsupportedQuery", f"Invalid filter clause: {flt}")

    def _find_users(self, flt):
        m = re.fullmatch(r"mail eq '((?:[^']|'')*)'", flt)
        if m:
            address = m.group(1).replace("''", "'")
            if " " in address:
                return graph_error(400, "Request_BadRequest")
            found = [u for u in self.users.values() if (u["mail"] or "").lower() == address.lower()]
        else:
            m = re.fullmatch(r"proxyAddresses/any\(p:p eq '((?:[^']|'')*)'\)", flt)
            if not m:
                return graph_error(400, "Request_UnsupportedQuery")
            address = m.group(1).replace("''", "'")
            found = [u for u in self.users.values() if address.lower() in (a.lower() for a in u["proxyAddresses"])]
        return response(200, {"value": [{"id": u["id"], "userPrincipalName": u["userPrincipalName"]} for u in found]})

    # SharePoint REST

    def _sharepoint(self, method, site, path, query, body):
        m = re.fullmatch(r"web/lists\(guid'([^']+)'\)(.*)", path)
        if m:
            lst = site.lists.get(m.group(1))
            if lst is None:
                return response(404)
            return self._sharepoint_list(method, site, lst, m.group(2), body)

        if path == "web/roledefinitions" and method == "GET":
            return response(200, {"d": {"results": [
                {"Id": role_id, "Name": name, "RoleTypeKind": ROLE_TYPES[name]}
                for name, role_id in ROLE_DEFINITIONS.items()
            ]}})

        if path == "web/ensureuser" and method == "POST":
            logon_name = body["logonName"]
            group_id = logon_name.rsplit("|", 1)[-1]
            if group_id not in self.groups:
                return response(400, {"error": {"message": {"value": "The specified user could not be found."}}})
            if logon_name not in site.principals:
                site.principals[logon_name] = site.next_principal_id
                site.next_principal_id += 1
            return response(200, {"d": {"Id": site.principals[logon_name]}})

        if path == "web" and method == "GET":
            return response(200, {"d": dict(site.web)})
        if path == "web" and method == "MERGE":
            assert body["__metadata"] == {"type": "SP.Web"}
            if not site.ignore_sharing_writes:
                site.web.update({k: v for k, v in body.items() if k != "__metadata"})
            return response(204)
        if path == "web/setuseaccessrequestdefaultandupdate(useAccessRequestDefault=false)" and method == "POST":
            if not site.ignore_sharing_writes:
                site.web["UseAccessRequestDefault"] = False
            return response(200, {"d": {}})
        if path == "web/associatedmembergroup" and method == "GET":
            return response(200, {"d": dict(site.member_group)}) if site.member_group else response(404)
        m = re.fullmatch(r"web/sitegroups/getbyid\((\d+)\)", path)
        if m and method == "MERGE":
            assert int(m.group(1)) == site.member_group["Id"]
            if not site.ignore_sharing_writes:
                site.member_group["AllowMembersEditMembership"] = body["AllowMembersEditMembership"]
            return response(204)

        if path == "web/navigation/quicklaunch" and method == "GET":
            assert query.get("$expand") == "Children"
            return response(200, {"d": {"results": [
                {**{k: v for k, v in node.items() if k != "Children"}, "Children": {"results": list(node["Children"])}}
                for node in site.nav
            ]}})
        if path == "web/navigation/quicklaunch" and method == "POST":
            node = site.nav_node(body["Title"], body["Url"], body["IsExternal"])
            site.nav.append(node)
            return response(201, {"d": node})
        m = re.fullmatch(r"web/navigation/GetNodeById\((\d+)\)(/Children)?", path)
        if m:
            node, siblings = site.find_nav_node(int(m.group(1)))
            if node is None:
                return response(404)
            if m.group(2) and method == "POST":
                child = site.nav_node(body["Title"], body["Url"], body["IsExternal"])
                node["Children"].append(child)
                return response(201, {"d": child})
            if method == "MERGE":
                node["Title"] = body["Title"]
                return response(204)
            if method == "DELETE":
                siblings.remove(node)
                return response(200)

        raise AssertionError(f"Unexpected SharePoint request: {method} {site.url}/_api/{path}")

    def _sharepoint_list(self, method, site, lst, rest, body):
        if rest == "" and method == "GET":
            return response(200, {"d": {"HasUniqueRoleAssignments": lst["unique"]}})
        if rest.startswith("/breakroleinheritance(") and method == "POST":
            assert "copyRoleAssignments=false" in rest
            lst["unique"] = True
            lst["role_assignments"] = {}
            return response(200, {"d": {}})
        m = re.fullmatch(r"/roleassignments/addroleassignment\(principalid=(\d+), roledefid=(\d+)\)", rest)
        if m and method == "POST":
            if not lst["unique"]:
                return response(400, {"error": {"message": {"value": "inherits permissions"}}})
            lst["role_assignments"].setdefault(int(m.group(1)), set()).add(int(m.group(2)))
            return response(200, {"d": {}})
        if rest == "/roleassignments" and method == "GET":
            login_names = {pid: claims for claims, pid in site.principals.items()}
            return response(200, {"d": {"results": [
                {
                    "PrincipalId": pid,
                    "Member": {"LoginName": login_names.get(pid, f"i:0#.f|membership|user{pid}")},
                    "RoleDefinitionBindings": {"results": [{"Id": r} for r in sorted(roles)]},
                }
                for pid, roles in lst["role_assignments"].items()
            ]}})
        m = re.fullmatch(r"/roleassignments/getbyprincipalid\((\d+)\)", rest)
        if m and method == "DELETE":
            if lst["role_assignments"].pop(int(m.group(1)), None) is None:
                return response(404)
            return response(200)
        if rest == "/rootfolder/properties" and method == "GET":
            return response(200, {"d": dict(lst["properties"])})
        raise AssertionError(f"Unexpected SharePoint list request: {method} {rest}")

    def _process_query(self, site, xml):
        guid = re.search(r'<Parameter Type="Guid">\{([^}]+)\}</Parameter>', xml).group(1)
        lst = site.lists.get(guid)
        if lst is None:
            return response(200, [{"SchemaVersion": "15.0.0.0", "ErrorInfo": {"ErrorMessage": "List does not exist."}}])
        pairs = re.findall(
            r'<Method Name="SetFieldValue"[^>]*><Parameters><Parameter Type="String">(.*?)</Parameter>'
            r'<Parameter Type="String">(.*?)</Parameter></Parameters></Method>',
            xml,
        )
        for key, value in pairs:
            lst["properties"][_xml_unescape(key)] = _xml_unescape(value)
        return response(200, [{"SchemaVersion": "15.0.0.0", "ErrorInfo": None}])


def _xml_unescape(text):
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
