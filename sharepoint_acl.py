import base64
import os
import re
import time
from typing import Dict, Iterable, List, Optional, Set
from urllib.parse import parse_qs, unquote, urlparse

import requests
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.serialization import Encoding, NoEncryption, PrivateFormat, pkcs12
from msal import ConfidentialClientApplication

TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}
GUID_SEARCH_PATTERN = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")


def load_pfx_bytes(pfx_path: Optional[str]) -> bytes:
    if pfx_path:
        with open(pfx_path, "rb") as handle:
            return handle.read()

    encoded = os.getenv("SHAREPOINT_PFX_BASE64", "").strip()
    if encoded:
        return base64.b64decode(encoded, validate=True)
    raise RuntimeError("Set SHAREPOINT_PFX_BASE64 or SHAREPOINT_PFX_PATH.")


def get_sharepoint_access_token(tenant_id: str, client_id: str, site_url: str, pfx_data: bytes, pfx_password: str) -> str:
    private_key, certificate, _ = pkcs12.load_key_and_certificates(
        pfx_data, pfx_password.encode("utf-8") if pfx_password else None
    )
    if private_key is None or certificate is None:
        raise RuntimeError("PFX must contain both the private key and certificate.")

    private_key_pem = private_key.private_bytes(Encoding.PEM, PrivateFormat.PKCS8, NoEncryption()).decode("utf-8")
    thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
    resource = f"{urlparse(site_url).scheme}://{urlparse(site_url).netloc}"
    app = ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential={"private_key": private_key_pem, "thumbprint": thumbprint},
    )
    result = app.acquire_token_for_client(scopes=[f"{resource}/.default"])
    token = result.get("access_token")
    if not token:
        raise RuntimeError(f"SharePoint token acquisition failed: {result.get('error')} - {result.get('error_description')}")
    return token


def normalize_group_ids(values: Iterable[object]) -> Set[str]:
    result = set()
    for value in values:
        normalized = str(value).strip().lower().strip("{}").strip()
        if normalized:
            result.add(normalized)
    return result


def derive_site_url_from_weburi(web_uri: str, fallback: str) -> str:
    parsed = urlparse(web_uri)
    parts = [part for part in parsed.path.split("/") if part]
    if len(parts) >= 2 and parts[0].lower() in {"sites", "teams"}:
        return f"{parsed.scheme}://{parsed.netloc}/{parts[0]}/{parts[1]}"
    return fallback


def is_tenant_principal(login_name: Optional[str]) -> bool:
    lowered = (login_name or "").lower()
    return "|tenant|" in lowered and "|membership|" not in lowered and "sharinglink" not in lowered


def extract_tenant_id(login_name: Optional[str]) -> Optional[str]:
    if not is_tenant_principal(login_name):
        return None
    marker = "|tenant|"
    value = login_name[login_name.lower().find(marker) + len(marker):].strip()
    return value.lower() or None


def extract_rolemanager_tenant_id(login_name: Optional[str]) -> Optional[str]:
    lowered = (login_name or "").lower()
    if "|rolemanager|" not in lowered or "spo-grid-all-users/" not in lowered:
        return None
    match = GUID_SEARCH_PATTERN.search(login_name or "")
    return match.group(0).lower() if match else None


def is_sharing_links_principal(login_name: Optional[str], title: Optional[str]) -> bool:
    return "sharinglink" in (login_name or "").lower() or "sharinglink" in (title or "").lower()


def sp_get(site_url: str, relative_url: str, token: str) -> Dict:
    response = requests.get(
        f"{site_url.rstrip('/')}/{relative_url.lstrip('/')}",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=nometadata"},
        timeout=45,
    )
    response.raise_for_status()
    return response.json()


def sp_get_with_retries(site_url: str, relative_url: str, token: str, max_retries: int = 3) -> requests.Response:
    url = f"{site_url.rstrip('/')}/{relative_url.lstrip('/')}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata=nometadata"}
    for attempt in range(max_retries + 1):
        response = requests.get(url, headers=headers, timeout=45)
        if response.status_code not in TRANSIENT_STATUS_CODES or attempt == max_retries:
            return response
        retry_after = response.headers.get("Retry-After")
        time.sleep(int(retry_after) if retry_after and retry_after.isdigit() else min(2 ** (attempt + 1), 10))
    return response


def get_server_relative_candidates(web_uri: str) -> List[str]:
    encoded = urlparse(web_uri).path
    decoded = unquote(encoded)
    return list(dict.fromkeys(value for value in (encoded, decoded) if value))


def extract_sourcedoc_guid(web_uri: str) -> Optional[str]:
    query = parse_qs(urlparse(web_uri).query)
    values = query.get("sourcedoc") or query.get("SourceDoc")
    if not values:
        return None
    return unquote(values[0]).strip().strip("{}") or None


def get_role_assignments(site_url: str, token: str, web_uri: str) -> List[Dict]:
    sourcedoc = extract_sourcedoc_guid(web_uri)
    if sourcedoc:
        relative = f"_api/web/GetFileById(guid'{sourcedoc}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings"
        response = sp_get_with_retries(site_url, relative, token)
        if response.status_code not in (400, 404):
            response.raise_for_status()
            return response.json().get("value", [])

    for server_relative in get_server_relative_candidates(web_uri):
        escaped = server_relative.replace("'", "''")
        relatives = [
            f"_api/web/GetFileByServerRelativeUrl('{escaped}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings",
            f"_api/web/GetFileByServerRelativePath(decodedurl='{escaped}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings",
            f"_api/web/GetFolderByServerRelativeUrl('{escaped}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings",
            f"_api/web/GetFolderByServerRelativePath(decodedurl='{escaped}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings",
        ]
        for relative in relatives:
            response = sp_get_with_retries(site_url, relative, token)
            if response.status_code in (400, 404):
                continue
            response.raise_for_status()
            return response.json().get("value", [])
    return []


def get_group_member_tenant_ids(site_url: str, token: str, group_id: int, cache: Dict[str, Set[str]]) -> Set[str]:
    key = f"{site_url.lower()}|{group_id}"
    if key in cache:
        return cache[key]
    try:
        users = sp_get(site_url, f"_api/web/sitegroups/getbyid({group_id})/users?$select=LoginName", token).get("value", [])
    except requests.HTTPError:
        cache[key] = set()
        return set()

    result = set()
    for user in users:
        login = user.get("LoginName")
        tenant_id = extract_tenant_id(login) or extract_rolemanager_tenant_id(login)
        if tenant_id:
            result.add(tenant_id)
    cache[key] = result
    return result


def tenant_ids_for_weburi(site_url: str, token: str, web_uri: str, group_cache: Dict[str, Set[str]]) -> Set[str]:
    result = set()
    for assignment in get_role_assignments(site_url, token, web_uri):
        member = assignment.get("Member") or {}
        login = member.get("LoginName")
        title = member.get("Title")
        if is_sharing_links_principal(login, title):
            continue
        tenant_id = extract_tenant_id(login) or extract_rolemanager_tenant_id(login)
        if tenant_id:
            result.add(tenant_id)
        elif int(member.get("PrincipalType") or 0) == 8 and isinstance(member.get("Id"), int):
            result.update(get_group_member_tenant_ids(site_url, token, member["Id"], group_cache))
    return result