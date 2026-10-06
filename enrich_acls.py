import argparse
import json
import os
import re
import shutil
import subprocess
import time
from typing import Dict, Iterable, List, Optional, Set
from urllib.parse import parse_qs, unquote, urlparse

import requests
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.serialization import Encoding, NoEncryption, PrivateFormat, pkcs12
from msal import ConfidentialClientApplication


DEFAULT_SITE_URL = "https://fshr.sharepoint.com/teams/SharedServices"
DEFAULT_SEARCH_SERVICE = "srch-ai-kmagent-dev-ccanada-001"
DEFAULT_INDEX_NAME = "pt-sharepoint-chunks-index"
DEFAULT_API_VERSION = "2026-05-01-preview"
DEFAULT_PFX_PATH = r"D:\Users\YuXin.Li\Desktop\sharepoint-app.pfx"
DEFAULT_PFX_PASSWORD = "MyStrongPassword123!"
DEFAULT_TENANT_ID = "38c0107d-81d1-419f-8b72-8a491dc9603c"
DEFAULT_CLIENT_ID = "daeec413-d047-4ad3-a1ad-9c85782c319c"
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}
RETRYABLE_INDEX_STATUS_CODES = {429, 500, 502, 503, 504}
GUID_SEARCH_PATTERN = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)


def should_refresh_search_token(response: requests.Response) -> bool:
    if response.status_code == 401:
        return True
    if response.status_code != 403:
        return False

    try:
        body = response.text.lower()
    except Exception:
        body = ""
    return "expired" in body and "token" in body


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "For every chunk in an Azure AI Search index, pull ACL IDs from the chunk's "
            "SharePoint web URI and merge them into GroupIds."
        )
    )
    parser.add_argument("--site-url", default=DEFAULT_SITE_URL)
    parser.add_argument("--search-service", default=DEFAULT_SEARCH_SERVICE)
    parser.add_argument("--index-name", default=DEFAULT_INDEX_NAME)
    parser.add_argument("--api-version", default=DEFAULT_API_VERSION)
    parser.add_argument("--tenant-id", default=os.getenv("TENANT_ID", DEFAULT_TENANT_ID))
    parser.add_argument("--client-id", default=os.getenv("CLIENT_ID", DEFAULT_CLIENT_ID))
    parser.add_argument("--pfx-path", default=os.getenv("PFX_PATH", DEFAULT_PFX_PATH))
    parser.add_argument("--pfx-password", default=os.getenv("PFX_PASSWORD", DEFAULT_PFX_PASSWORD))
    parser.add_argument("--query-top", type=int, default=200)
    parser.add_argument("--write-batch-size", type=int, default=200)
    parser.add_argument("--max-chunks", type=int, default=0, help="0 means no limit")
    parser.add_argument(
        "--no-elevated-read",
        action="store_true",
        help="Disable x-ms-enable-elevated-read when reading chunks from the index.",
    )
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    missing = []
    if not args.tenant_id:
        missing.append("tenant-id/TENANT_ID")
    if not args.client_id:
        missing.append("client-id/CLIENT_ID")
    if not args.pfx_path:
        missing.append("pfx-path/PFX_PATH")
    if missing:
        raise ValueError(f"Missing required values: {', '.join(missing)}")


def get_search_access_token() -> str:
    az_path = shutil.which("az")
    if not az_path:
        raise RuntimeError(
            "Could not find Azure CLI in PATH. This script refreshes Search tokens via "
            "'az account get-access-token', so Azure CLI must be installed and available."
        )

    command = [
        az_path,
        "account",
        "get-access-token",
        "--resource",
        "https://search.azure.com",
        "--query",
        "accessToken",
        "-o",
        "tsv",
    ]
    result = subprocess.run(command, check=True, capture_output=True, text=True)
    token = result.stdout.strip()
    if not token:
        raise RuntimeError("Failed to get Azure Search access token from Azure CLI.")
    return token


def get_sharepoint_access_token(tenant_id: str, client_id: str, site_url: str, pfx_path: str, pfx_password: str) -> str:
    with open(pfx_path, "rb") as pfx_file:
        pfx_data = pfx_file.read()

    private_key, certificate, _ = pkcs12.load_key_and_certificates(
        pfx_data,
        pfx_password.encode("utf-8") if pfx_password else None,
    )

    if private_key is None or certificate is None:
        raise RuntimeError("PFX must contain both private key and certificate.")

    private_key_pem = private_key.private_bytes(
        encoding=Encoding.PEM,
        format=PrivateFormat.PKCS8,
        encryption_algorithm=NoEncryption(),
    ).decode("utf-8")

    thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    parsed = urlparse(site_url)
    resource = f"{parsed.scheme}://{parsed.netloc}"

    app = ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential={
            "private_key": private_key_pem,
            "thumbprint": thumbprint,
        },
    )

    result = app.acquire_token_for_client(scopes=[f"{resource}/.default"])
    token = result.get("access_token")
    if not token:
        raise RuntimeError(f"Failed to acquire SharePoint token: {json.dumps(result, indent=2)}")
    return token


def is_tenant_principal(login_name: Optional[str]) -> bool:
    if not login_name:
        return False
    lowered = login_name.lower()
    return (
        "|tenant|" in lowered
        and "|membership|" not in lowered
        and "sharinglinks" not in lowered
        and "sharinglink" not in lowered
    )


def extract_tenant_id(login_name: Optional[str]) -> Optional[str]:
    if not is_tenant_principal(login_name):
        return None
    marker = "|tenant|"
    lowered = login_name.lower()
    index = lowered.find(marker)
    if index == -1:
        return None
    value = login_name[index + len(marker):].strip()
    if not value:
        return None
    return value.lower()


def extract_rolemanager_tenant_id(login_name: Optional[str]) -> Optional[str]:
    if not login_name:
        return None

    lowered = login_name.lower()
    # Example:
    # c:0-.f|rolemanager|spo-grid-all-users/<tenant-guid>
    if "|rolemanager|" not in lowered or "spo-grid-all-users/" not in lowered:
        return None

    match = GUID_SEARCH_PATTERN.search(login_name)
    if not match:
        return None

    return match.group(0).lower()


def is_sharing_links_principal(login_name: Optional[str], title: Optional[str]) -> bool:
    login = (login_name or "").lower()
    text = (title or "").lower()
    return "sharinglinks" in login or "sharinglink" in login or "sharinglinks" in text or "sharinglink" in text


def sp_get(site_url: str, relative_url: str, token: str) -> Dict:
    url = f"{site_url.rstrip('/')}/{relative_url.lstrip('/')}"
    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=nometadata",
        },
        timeout=45,
    )
    response.raise_for_status()
    return response.json()


def sp_get_with_retries(site_url: str, relative_url: str, token: str, max_retries: int = 3) -> requests.Response:
    url = f"{site_url.rstrip('/')}/{relative_url.lstrip('/')}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=nometadata",
    }

    attempt = 0
    while True:
        response = requests.get(url, headers=headers, timeout=45)
        if response.status_code not in TRANSIENT_STATUS_CODES:
            return response

        attempt += 1
        if attempt > max_retries:
            return response

        retry_after = response.headers.get("Retry-After")
        if retry_after and retry_after.isdigit():
            wait_seconds = int(retry_after)
        else:
            wait_seconds = min(2 ** attempt, 10)
        time.sleep(wait_seconds)


def extract_server_relative_url(web_uri: str) -> str:
    parsed = urlparse(web_uri)
    # Keep URI encoding from metadata_spo_item_weburi so characters like '#'
    # remain %23 and are not treated as URI fragments by downstream calls.
    return parsed.path


def derive_site_url_from_weburi(web_uri: str, fallback_site_url: str) -> str:
    parsed = urlparse(web_uri)
    path_parts = [part for part in parsed.path.split("/") if part]
    if len(path_parts) >= 2 and path_parts[0].lower() in {"sites", "teams"}:
        return f"{parsed.scheme}://{parsed.netloc}/{path_parts[0]}/{path_parts[1]}"
    return fallback_site_url


def get_server_relative_candidates(web_uri: str) -> List[str]:
    encoded = extract_server_relative_url(web_uri)
    decoded = unquote(encoded)
    candidates: List[str] = []
    for value in [encoded, decoded]:
        if value and value not in candidates:
            candidates.append(value)
    return candidates


def extract_sourcedoc_guid(web_uri: str) -> Optional[str]:
    parsed = urlparse(web_uri)
    query = parse_qs(parsed.query)
    values = query.get("sourcedoc") or query.get("SourceDoc")
    if not values:
        return None

    raw_value = unquote(values[0]).strip()
    if raw_value.startswith("{") and raw_value.endswith("}"):
        raw_value = raw_value[1:-1]
    return raw_value or None


def get_role_assignments(site_url: str, token: str, web_uri: str) -> List[Dict]:
    sourcedoc_guid = extract_sourcedoc_guid(web_uri)
    if sourcedoc_guid:
        by_id_relative = (
            f"_api/web/GetFileById(guid'{sourcedoc_guid}')/ListItemAllFields/RoleAssignments"
            "?$expand=Member,RoleDefinitionBindings"
        )
        by_id_response = sp_get_with_retries(site_url, by_id_relative, token)
        if by_id_response.status_code not in (400, 404):
            by_id_response.raise_for_status()
            return by_id_response.json().get("value", [])

    for server_relative in get_server_relative_candidates(web_uri):
        escaped = server_relative.replace("'", "''")
        candidates = [
            (
                "file-url",
                f"_api/web/GetFileByServerRelativeUrl('{escaped}')/ListItemAllFields/RoleAssignments"
                "?$expand=Member,RoleDefinitionBindings",
            ),
            (
                "file-path",
                f"_api/web/GetFileByServerRelativePath(decodedurl='{escaped}')/ListItemAllFields/RoleAssignments"
                "?$expand=Member,RoleDefinitionBindings",
            ),
            (
                "folder-url",
                f"_api/web/GetFolderByServerRelativeUrl('{escaped}')/ListItemAllFields/RoleAssignments"
                "?$expand=Member,RoleDefinitionBindings",
            ),
            (
                "folder-path",
                f"_api/web/GetFolderByServerRelativePath(decodedurl='{escaped}')/ListItemAllFields/RoleAssignments"
                "?$expand=Member,RoleDefinitionBindings",
            ),
        ]

        for _, relative in candidates:
            response = sp_get_with_retries(site_url, relative, token)
            if response.status_code in (400, 404):
                continue
            response.raise_for_status()
            payload = response.json()
            return payload.get("value", [])

    return []


def get_group_member_tenant_ids(site_url: str, token: str, group_id: int, cache: Dict[str, Set[str]]) -> Set[str]:
    cache_key = f"{site_url.lower()}|{group_id}"
    if cache_key in cache:
        return cache[cache_key]

    relative_url = f"_api/web/sitegroups/getbyid({group_id})/users?$select=LoginName"
    try:
        payload = sp_get(site_url, relative_url, token)
    except requests.HTTPError:
        cache[cache_key] = set()
        return set()

    tenant_ids: Set[str] = set()
    for user in payload.get("value", []):
        login_name = user.get("LoginName")
        tenant_id = extract_tenant_id(login_name)
        if tenant_id:
            tenant_ids.add(tenant_id)
            continue

        rolemanager_tenant_id = extract_rolemanager_tenant_id(login_name)
        if rolemanager_tenant_id:
            tenant_ids.add(rolemanager_tenant_id)

    cache[cache_key] = tenant_ids
    return tenant_ids


def tenant_ids_for_weburi(site_url: str, token: str, web_uri: str, group_cache: Dict[str, Set[str]]) -> Set[str]:
    ids: Set[str] = set()
    role_assignments = get_role_assignments(site_url, token, web_uri)

    for assignment in role_assignments:
        member = assignment.get("Member") or {}
        login_name = member.get("LoginName")
        title = member.get("Title")

        if is_sharing_links_principal(login_name, title):
            continue

        direct_tenant_id = extract_tenant_id(login_name)
        if direct_tenant_id:
            ids.add(direct_tenant_id)
            continue

        rolemanager_tenant_id = extract_rolemanager_tenant_id(login_name)
        if rolemanager_tenant_id:
            ids.add(rolemanager_tenant_id)
            continue

        principal_type = int(member.get("PrincipalType") or 0)
        member_id = member.get("Id")

        # PrincipalType 8 is SharePoint group in classic SP principal model.
        if principal_type == 8 and isinstance(member_id, int):
            ids.update(get_group_member_tenant_ids(site_url, token, member_id, group_cache))

    return ids


def search_chunks(
    search_service: str,
    index_name: str,
    api_version: str,
    token: str,
    top: int,
    max_chunks: int,
    elevated_read: bool,
) -> List[Dict]:
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs/search?api-version={api_version}"
    token_value = token

    def _escape_odata(value: str) -> str:
        return value.replace("'", "''")

    def _make_headers() -> Dict[str, str]:
        headers = {
            "Authorization": f"Bearer {token_value}",
            "Content-Type": "application/json",
        }
        if elevated_read:
            headers["x-ms-enable-elevated-read"] = "true"
        return headers

    def _post_search(body: Dict, url: Optional[str] = None) -> requests.Response:
        nonlocal token_value
        target_url = url or endpoint
        headers = _make_headers()
        response = requests.post(target_url, headers=headers, json=body, timeout=60)
        if should_refresh_search_token(response):
            print("[INFO] Search token expired while reading. Refreshing via Azure CLI...")
            token_value = get_search_access_token()
            headers = _make_headers()
            response = requests.post(target_url, headers=headers, json=body, timeout=60)
        try:
            response.raise_for_status()
        except requests.HTTPError as ex:
            raise requests.HTTPError(
                f"{ex}. Response body: {response.text}",
                response=response,
            ) from ex
        return response

    def _parse_sortable_field_error(ex: requests.HTTPError) -> bool:
        response = getattr(ex, "response", None)
        if response is None or response.status_code != 400:
            return False
        body = (response.text or "").lower()
        return "not a sortable field" in body or "only sortable fields can be used in $orderby" in body

    def _scan_with_orderby(orderby: str) -> List[Dict]:
        loaded_docs: List[Dict] = []
        loaded_chunk_ids: Set[str] = set()
        page_number = 0
        previous_chunk_id: Optional[str] = None

        while len(loaded_docs) < expected_total:
            page_number += 1
            body: Dict[str, object] = {
                "search": "*",
                "top": top,
                "orderby": orderby,
                "select": "chunkId,url,GroupIds",
            }

            if previous_chunk_id is not None:
                # Keyset pagination must advance strictly past the last chunkId from the prior page.
                body["filter"] = f"chunkId gt '{_escape_odata(previous_chunk_id)}'"

            payload = _post_search(body).json()
            page_docs = payload.get("value", []) or []
            if not isinstance(page_docs, list):
                raise RuntimeError(f"Search scan page {page_number} returned an invalid payload shape.")

            page_unique_added = 0
            page_last_chunk_id: Optional[str] = None

            for raw_doc in page_docs:
                chunk_id = raw_doc.get("chunkId")
                if chunk_id is None or str(chunk_id).strip() == "":
                    raise RuntimeError(
                        f"Search scan returned a document with a missing or empty chunkId on page {page_number}."
                    )

                chunk_id_text = str(chunk_id)
                if chunk_id_text in loaded_chunk_ids:
                    raise RuntimeError(f"Search scan returned duplicate chunkId {chunk_id_text!r} on page {page_number}.")

                if previous_chunk_id is not None and chunk_id_text <= previous_chunk_id:
                    raise RuntimeError(
                        f"Search scan pagination did not advance on page {page_number}: "
                        f"previous_last_chunkId={previous_chunk_id!r}, current_chunkId={chunk_id_text!r}."
                    )

                loaded_chunk_ids.add(chunk_id_text)
                loaded_docs.append(raw_doc)
                page_unique_added += 1
                page_last_chunk_id = chunk_id_text

                if len(loaded_docs) == expected_total:
                    break

            last_chunk_id_text = page_last_chunk_id or ""
            print(
                f"[INFO] Search scan page {page_number}: records_returned={len(page_docs)}, "
                f"unique_added={page_unique_added}, total_loaded={len(loaded_docs)}, "
                f"expected_total={expected_total}, last_chunkId={last_chunk_id_text}"
            )

            if page_last_chunk_id is None:
                raise RuntimeError(f"Search scan page {page_number} did not return any usable chunkId values.")

            previous_chunk_id = page_last_chunk_id

            if len(loaded_docs) == expected_total:
                break

            if len(page_docs) < top:
                raise RuntimeError(
                    f"Search scan incomplete after page {page_number}: service returned a short final page "
                    f"({len(page_docs)} records < page size {top}) before reaching the expected {expected_total} unique chunks."
                )

            if page_unique_added == 0:
                raise RuntimeError(f"Search scan did not advance on page {page_number}.")

        if len(loaded_chunk_ids) != expected_total:
            raise RuntimeError(
                f"Search scan incomplete: loaded {len(loaded_chunk_ids)} unique chunks but expected {expected_total}."
            )

        print(f"[INFO] Search scan validated {len(loaded_chunk_ids)} unique chunks successfully.")
        return loaded_docs

    count_response = _post_search(
        {
            "search": "*",
            "top": 0,
            "count": True,
        }
    ).json()
    total_available = int(count_response.get("@odata.count", 0))
    expected_total = total_available if max_chunks <= 0 else min(total_available, max_chunks)

    print(
        f"[INFO] Search scan count query reported {total_available} chunks; expected to load {expected_total}."
    )

    if expected_total == 0:
        print("[INFO] Search scan validated 0 unique chunks.")
        return []

    try:
        return _scan_with_orderby("chunkId asc")
    except requests.HTTPError as ex:
        if not _parse_sortable_field_error(ex):
            raise
        raise RuntimeError(
            "The index schema does not permit deterministic chunkId keyset pagination because chunkId is not sortable. "
            "To guarantee a complete scan, make chunkId sortable in the index schema and rebuild/reindex, or use a sortable "
            "unique cursor field."
        )


def normalize_group_id(value: object) -> str:
    normalized = str(value).strip().lower()
    if normalized.startswith("{") and normalized.endswith("}"):
        normalized = normalized[1:-1].strip()
    return normalized


def normalize_group_ids(values: Iterable[object]) -> Set[str]:
    result: Set[str] = set()
    for value in values:
        normalized = normalize_group_id(value)
        if normalized:
            result.add(normalized)
    return result


def chunk_actions(
    chunks: Iterable[Dict],
    site_url: str,
    sp_token: str,
    weburi_cache: Dict[str, Set[str]],
    group_cache: Dict[str, Set[str]],
) -> tuple[List[Dict], Dict[str, int]]:
    actions: List[Dict] = []
    stats = {
        "total_chunks_seen": 0,
        "missing_chunk_id_or_weburi": 0,
        "skipped_groupids_all": 0,
        "unique_weburis": 0,
        "acl_pull_failures": 0,
        "unchanged_groupids": 0,
        "update_candidates": 0,
    }
    failed_weburis: Set[str] = set()

    for chunk in chunks:
        stats["total_chunks_seen"] += 1
        chunk_id = chunk.get("chunkId")
        web_uri = chunk.get("url")
        existing_group_ids = normalize_group_ids(chunk.get("GroupIds") or [])
        # If GroupIds already includes all, ACL expansion is unnecessary.
        if "all" in existing_group_ids:
            stats["skipped_groupids_all"] += 1
            stats["unchanged_groupids"] += 1
            continue

        if not chunk_id or not web_uri:
            stats["missing_chunk_id_or_weburi"] += 1
            continue

        if web_uri in failed_weburis:
            continue

        if web_uri not in weburi_cache:
            stats["unique_weburis"] += 1
            try:
                site_url_for_weburi = derive_site_url_from_weburi(web_uri, site_url)
                weburi_cache[web_uri] = tenant_ids_for_weburi(site_url_for_weburi, sp_token, web_uri, group_cache)
            except requests.HTTPError as ex:
                print(f"[WARN] ACL pull failed for {web_uri}: {ex}")
                stats["acl_pull_failures"] += 1
                failed_weburis.add(web_uri)
                continue

        merged_ids = sorted(existing_group_ids.union(normalize_group_ids(weburi_cache[web_uri])))
        if merged_ids == sorted(existing_group_ids):
            stats["unchanged_groupids"] += 1
            continue

        actions.append(
            {
                "@search.action": "mergeOrUpload",
                "chunkId": chunk_id,
                "GroupIds": merged_ids,
            }
        )
        stats["update_candidates"] += 1

    return actions, stats


def send_actions(search_service: str, index_name: str, api_version: str, token: str, actions: List[Dict], batch_size: int) -> int:
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs/index?api-version={api_version}"
    token_value = token

    success_count = 0
    for i in range(0, len(actions), batch_size):
        pending = actions[i : i + batch_size]
        attempt = 0
        while pending and attempt < 4:
            attempt += 1

            headers = {
                "Authorization": f"Bearer {token_value}",
                "Content-Type": "application/json",
            }
            response = requests.post(endpoint, headers=headers, json={"value": pending}, timeout=60)
            if should_refresh_search_token(response):
                print("[INFO] Search token expired while writing. Refreshing via Azure CLI...")
                token_value = get_search_access_token()
                headers["Authorization"] = f"Bearer {token_value}"
                response = requests.post(endpoint, headers=headers, json={"value": pending}, timeout=60)

            try:
                response.raise_for_status()
            except requests.HTTPError as ex:
                raise requests.HTTPError(
                    f"{ex}. Response body: {response.text}",
                    response=response,
                ) from ex

            payload = response.json()
            result_items = payload.get("value", []) or []
            if len(result_items) != len(pending):
                raise RuntimeError(
                    f"Azure AI Search returned {len(result_items)} result items for {len(pending)} submitted actions."
                )

            pending_by_key = {
                str(action.get("chunkId")): action
                for action in pending
                if action.get("chunkId")
            }
            retry_actions: List[Dict] = []

            for item in result_items:
                if item.get("status"):
                    success_count += 1
                    continue

                key = str(item.get("key") or "")
                status_code = int(item.get("statusCode") or 0)
                if status_code in RETRYABLE_INDEX_STATUS_CODES and attempt < 4 and key in pending_by_key:
                    retry_actions.append(pending_by_key[key])
                    continue

                print(f"[WARN] Failed key={item.get('key')} statusCode={item.get('statusCode')}")

            if retry_actions:
                print(
                    f"[WARN] Retrying {len(retry_actions)} transient indexing failures "
                    f"(attempt {attempt + 1}/4)"
                )
                time.sleep(min(2 ** attempt, 10))

            pending = retry_actions

        if pending:
            print(f"[WARN] Gave up retrying {len(pending)} chunk updates in this batch.")

    return success_count


def print_final_summary(stats: Dict[str, int], updated: int, dry_run: bool, elapsed_seconds: Optional[float] = None) -> None:
    elapsed_text = ""
    if elapsed_seconds is not None:
        elapsed_text = f", elapsed_seconds={elapsed_seconds:.2f}"

    print(
        "FINAL SUMMARY: "
        f"loaded={stats['total_chunks_seen']}, "
        f"unchanged={stats['unchanged_groupids']}, "
        f"updated={updated}, "
        f"failed ACL pulls={stats['acl_pull_failures']}, "
        f"skipped missing weburi/chunkId={stats['missing_chunk_id_or_weburi']}, "
        f"dry_run={str(dry_run).lower()}{elapsed_text}"
    )


def main() -> None:
    start_time = time.perf_counter()
    args = parse_args()
    validate_args(args)

    print("Acquiring tokens...")
    search_token = get_search_access_token()
    sp_token = get_sharepoint_access_token(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        site_url=args.site_url,
        pfx_path=args.pfx_path,
        pfx_password=args.pfx_password,
    )

    print("Loading chunks from search index...")
    chunks = search_chunks(
        search_service=args.search_service,
        index_name=args.index_name,
        api_version=args.api_version,
        token=search_token,
        top=args.query_top,
        max_chunks=args.max_chunks,
        elevated_read=not args.no_elevated_read,
    )
    print(f"Loaded {len(chunks)} chunks")

    weburi_cache: Dict[str, Set[str]] = {}
    group_cache: Dict[str, Set[str]] = {}

    print("Building merge actions...")
    actions, stats = chunk_actions(
        chunks=chunks,
        site_url=args.site_url,
        sp_token=sp_token,
        weburi_cache=weburi_cache,
        group_cache=group_cache,
    )
    print(f"Prepared {len(actions)} chunk updates")
    print(
        "Run stats: "
        f"total_chunks_seen={stats['total_chunks_seen']}, "
        f"skipped_groupids_all={stats['skipped_groupids_all']}, "
        f"unique_weburis={stats['unique_weburis']}, "
        f"missing_chunk_id_or_weburi={stats['missing_chunk_id_or_weburi']}, "
        f"acl_pull_failures={stats['acl_pull_failures']}, "
        f"unchanged_groupids={stats['unchanged_groupids']}, "
        f"update_candidates={stats['update_candidates']}"
    )

    elapsed_seconds = time.perf_counter() - start_time
    if not actions:
        print("No GroupIds changes required.")
        print_final_summary(stats=stats, updated=0, dry_run=args.dry_run, elapsed_seconds=elapsed_seconds)
        return

    if args.dry_run:
        sample = actions[:3]
        print("Dry run enabled. Sample updates:")
        print(json.dumps(sample, indent=2))
        print_final_summary(stats=stats, updated=len(actions), dry_run=True, elapsed_seconds=elapsed_seconds)
        return

    print("Writing updates to Azure AI Search...")
    success = send_actions(
        search_service=args.search_service,
        index_name=args.index_name,
        api_version=args.api_version,
        token=search_token,
        actions=actions,
        batch_size=args.write_batch_size,
    )
    print(f"Updated {success} chunk documents")
    elapsed_seconds = time.perf_counter() - start_time
    print_final_summary(stats=stats, updated=success, dry_run=False, elapsed_seconds=elapsed_seconds)


if __name__ == "__main__":
    main()
