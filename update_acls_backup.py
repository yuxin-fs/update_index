import argparse
import base64
import json
import logging
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set
from urllib.parse import parse_qs, unquote, urlparse

import requests
from azure.identity import DefaultAzureCredential
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.serialization import Encoding, NoEncryption, PrivateFormat, pkcs12
from msal import ConfidentialClientApplication

DEFAULT_SITE_URL = "https://fshr.sharepoint.com/teams/SharedServices"
DEFAULT_SEARCH_SERVICE = "srch-ai-kmagent-dev-ccanada-001"
DEFAULT_API_VERSION = "2026-05-01-preview"
READY_POLL_SECONDS = 15
READY_TIMEOUT_SECONDS = 300
READY_STABLE_POLLS = 2
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}
GUID_SEARCH_PATTERN = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")


def load_dotenv_file() -> None:
    env_path = Path(__file__).resolve().parent / ".env"
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        os.environ.setdefault(key, value)


def env_required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Required setting '{name}' is missing or empty.")
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Merge SharePoint ACL tenant IDs into an Azure AI Search chunks index.")
    parser.add_argument("--site-url", default=os.getenv("SHAREPOINT_SITE_URL", DEFAULT_SITE_URL))
    parser.add_argument("--search-service", default=os.getenv("SEARCH_SERVICE", DEFAULT_SEARCH_SERVICE))
    parser.add_argument("--index-name", required=True)
    parser.add_argument("--api-version", default=os.getenv("SEARCH_API_VERSION", DEFAULT_API_VERSION))
    parser.add_argument("--tenant-id", default=os.getenv("TENANT_ID"))
    parser.add_argument("--client-id", default=os.getenv("SHAREPOINT_CLIENT_ID"))
    parser.add_argument("--pfx-path", default=os.getenv("SHAREPOINT_PFX_PATH"))
    parser.add_argument("--pfx-password", default=os.getenv("SHAREPOINT_PFX_PASSWORD", ""))
    parser.add_argument("--query-top", type=int, default=200)
    parser.add_argument("--write-batch-size", type=int, default=200)
    parser.add_argument("--max-chunks", type=int, default=0)
    parser.add_argument("--no-elevated-read", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def get_search_access_token() -> str:
    return DefaultAzureCredential().get_token("https://search.azure.com/.default").token


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


def should_refresh_search_token(response: requests.Response) -> bool:
    if response.status_code == 401:
        return True
    return response.status_code == 403 and "expired" in response.text.lower() and "token" in response.text.lower()


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


def wait_for_index_ready(search_service: str, index_name: str, api_version: str, token: str, elevated_read: bool) -> None:
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs?search=*&$count=true&$top=0&api-version={api_version}"
    token_value = token
    deadline = time.monotonic() + READY_TIMEOUT_SECONDS
    previous_count = None
    stable_hits = 0

    while time.monotonic() < deadline:
        headers = {"Authorization": f"Bearer {token_value}", "Accept": "application/json"}
        if elevated_read:
            headers["x-ms-enable-elevated-read"] = "true"

        response = requests.get(endpoint, headers=headers, timeout=60)
        if should_refresh_search_token(response):
            token_value = get_search_access_token()
            headers["Authorization"] = f"Bearer {token_value}"
            response = requests.get(endpoint, headers=headers, timeout=60)

        response.raise_for_status()
        payload = response.json()
        current_count = int(payload.get("@odata.count", 0))

        print(
            f"[INFO] Readiness check {index_name}: count={current_count} previous={previous_count} stableHits={stable_hits}"
        )

        if current_count > 0 and current_count == previous_count:
            stable_hits += 1
            if stable_hits >= READY_STABLE_POLLS:
                print(f"[INFO] Index {index_name} is ready with stable count {current_count}.")
                return
        else:
            stable_hits = 0

        previous_count = current_count
        time.sleep(READY_POLL_SECONDS)

    raise TimeoutError(f"Index {index_name} did not become ready within {READY_TIMEOUT_SECONDS} seconds.")


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


def derive_site_url_from_weburi(web_uri: str, fallback: str) -> str:
    parsed = urlparse(web_uri)
    parts = [part for part in parsed.path.split("/") if part]
    if len(parts) >= 2 and parts[0].lower() in {"sites", "teams"}:
        return f"{parsed.scheme}://{parsed.netloc}/{parts[0]}/{parts[1]}"
    return fallback


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
        login, title = member.get("LoginName"), member.get("Title")
        if is_sharing_links_principal(login, title):
            continue
        tenant_id = extract_tenant_id(login) or extract_rolemanager_tenant_id(login)
        if tenant_id:
            result.add(tenant_id)
        elif int(member.get("PrincipalType") or 0) == 8 and isinstance(member.get("Id"), int):
            result.update(get_group_member_tenant_ids(site_url, token, member["Id"], group_cache))
    return result


def search_chunks(search_service, index_name, api_version, token, top, max_chunks, elevated_read):
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs/search?api-version={api_version}"
    token_value = token

    def post(body):
        nonlocal token_value
        headers = {"Authorization": f"Bearer {token_value}", "Content-Type": "application/json"}
        if elevated_read:
            headers["x-ms-enable-elevated-read"] = "true"
        response = requests.post(endpoint, headers=headers, json=body, timeout=60)
        if should_refresh_search_token(response):
            token_value = get_search_access_token()
            headers["Authorization"] = f"Bearer {token_value}"
            response = requests.post(endpoint, headers=headers, json=body, timeout=60)
        if response.status_code >= 400:
            print(f"[ERROR] Search request failed for {endpoint}")
            print(f"[ERROR] Request headers: {json.dumps(headers, default=str)}")
            print(f"[ERROR] Request body: {json.dumps(body, default=str)}")
            print(f"[ERROR] Response status: {response.status_code}")
            print(f"[ERROR] Response body: {response.text[:4000]}")
        response.raise_for_status()
        return response.json()

    wait_for_index_ready(search_service, index_name, api_version, token_value, elevated_read)

    total = int(post({"search": "*", "top": 0, "count": True}).get("@odata.count", 0))
    expected = total if max_chunks <= 0 else min(total, max_chunks)
    loaded, seen, previous = [], set(), None
    while len(loaded) < expected:
        body = {"search": "*", "top": min(top, expected - len(loaded)), "orderby": "chunkId asc", "select": "chunkId,url,GroupIds"}
        if previous is not None:
            body["filter"] = f"chunkId gt '{previous.replace(chr(39), chr(39) * 2)}'"
        docs = post(body).get("value", [])
        if not docs:
            raise RuntimeError(f"Search scan stopped at {len(loaded)} of {expected} chunks.")
        for doc in docs:
            key = str(doc.get("chunkId") or "")
            if not key or key in seen or (previous is not None and key <= previous):
                raise RuntimeError(f"Invalid or non-advancing chunkId: {key!r}")
            seen.add(key)
            loaded.append(doc)
            previous = key
            if len(loaded) == expected:
                break
    print(f"[INFO] Search scan validated {len(loaded)} unique chunks.")
    return loaded


def normalize_group_ids(values: Iterable[object]) -> Set[str]:
    result = set()
    for value in values:
        normalized = str(value).strip().lower().strip("{}").strip()
        if normalized:
            result.add(normalized)
    return result


def chunk_actions(chunks, site_url, sp_token, weburi_cache, group_cache):
    actions, failed = [], set()
    stats = {"total_chunks_seen": 0, "missing_chunk_id_or_weburi": 0, "skipped_groupids_all": 0, "unique_weburis": 0, "acl_pull_failures": 0, "unchanged_groupids": 0, "update_candidates": 0}
    for chunk in chunks:
        stats["total_chunks_seen"] += 1
        chunk_id, web_uri = chunk.get("chunkId"), chunk.get("url")
        existing = normalize_group_ids(chunk.get("GroupIds") or [])
        if "all" in existing:
            stats["skipped_groupids_all"] += 1
            stats["unchanged_groupids"] += 1
            continue
        if not chunk_id or not web_uri:
            stats["missing_chunk_id_or_weburi"] += 1
            continue
        if web_uri in failed:
            continue
        if web_uri not in weburi_cache:
            stats["unique_weburis"] += 1
            try:
                derived_site = derive_site_url_from_weburi(web_uri, site_url)
                weburi_cache[web_uri] = tenant_ids_for_weburi(derived_site, sp_token, web_uri, group_cache)
            except requests.HTTPError as exc:
                print(f"[WARN] ACL pull failed for {web_uri}: {exc}")
                stats["acl_pull_failures"] += 1
                failed.add(web_uri)
                continue
        merged = sorted(existing | normalize_group_ids(weburi_cache[web_uri]))
        if merged == sorted(existing):
            stats["unchanged_groupids"] += 1
            continue
        actions.append({"@search.action": "mergeOrUpload", "chunkId": chunk_id, "GroupIds": merged})
        stats["update_candidates"] += 1
    return actions, stats


def send_actions(search_service, index_name, api_version, token, actions, batch_size):
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs/index?api-version={api_version}"
    success = 0
    for start in range(0, len(actions), batch_size):
        pending = actions[start:start + batch_size]
        for attempt in range(1, 5):
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            response = requests.post(endpoint, headers=headers, json={"value": pending}, timeout=60)
            if should_refresh_search_token(response):
                token = get_search_access_token()
                headers["Authorization"] = f"Bearer {token}"
                response = requests.post(endpoint, headers=headers, json={"value": pending}, timeout=60)
            response.raise_for_status()
            items = response.json().get("value", [])
            by_key = {str(action["chunkId"]): action for action in pending}
            retry = []
            for item in items:
                if item.get("status"):
                    success += 1
                elif int(item.get("statusCode") or 0) in TRANSIENT_STATUS_CODES and attempt < 4 and str(item.get("key")) in by_key:
                    retry.append(by_key[str(item.get("key"))])
                else:
                    raise RuntimeError(f"Index update failed: {item}")
            pending = retry
            if not pending:
                break
            time.sleep(min(2 ** attempt, 10))
        if pending:
            raise RuntimeError(f"Gave up retrying {len(pending)} index actions.")
    return success


def main():
    started = time.perf_counter()
    load_dotenv_file()
    args = parse_args()
    if not args.tenant_id or not args.client_id:
        raise RuntimeError("TENANT_ID and SHAREPOINT_CLIENT_ID are required.")
    search_token = get_search_access_token()
    sharepoint_token = get_sharepoint_access_token(
        args.tenant_id, args.client_id, args.site_url, load_pfx_bytes(args.pfx_path), args.pfx_password
    )
    chunks = search_chunks(args.search_service, args.index_name, args.api_version, search_token, args.query_top, args.max_chunks, not args.no_elevated_read)
    actions, stats = chunk_actions(chunks, args.site_url, sharepoint_token, {}, {})
    if args.dry_run:
        print(json.dumps(actions[:3], indent=2))
        updated = len(actions)
    else:
        updated = send_actions(args.search_service, args.index_name, args.api_version, search_token, actions, args.write_batch_size) if actions else 0
    print(f"FINAL SUMMARY: loaded={stats['total_chunks_seen']}, unchanged={stats['unchanged_groupids']}, updated={updated}, failed ACL pulls={stats['acl_pull_failures']}, dry_run={str(args.dry_run).lower()}, elapsed_seconds={time.perf_counter() - started:.2f}")


if __name__ == "__main__":
    main()
