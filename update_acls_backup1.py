import argparse
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Dict, List, Set

import requests
from azure.identity import DefaultAzureCredential
from sharepoint_acl import (
    derive_site_url_from_weburi,
    get_sharepoint_access_token,
    load_pfx_bytes,
    normalize_group_ids,
    tenant_ids_for_weburi,
)

DEFAULT_SITE_URL = "https://fshr.sharepoint.com/teams/SharedServices"
DEFAULT_SEARCH_SERVICE = "srch-ai-kmagent-prod-ccanada-001"
DEFAULT_API_VERSION = "2026-05-01-preview"
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}


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


def should_refresh_search_token(response: requests.Response) -> bool:
    if response.status_code == 401:
        return True
    return response.status_code == 403 and "expired" in response.text.lower() and "token" in response.text.lower()


def search_chunks(search_service, index_name, api_version, token, top, max_chunks, elevated_read):
    endpoint = f"https://{search_service}.search.windows.net/indexes/{index_name}/docs/search?api-version={api_version}"
    token_value = token

    def post(body):
        nonlocal token_value
        headers = {"Authorization": f"Bearer {token_value}", "Content-Type": "application/json"}
        if elevated_read:
            logging.warning(
                "elevated-read is disabled for this managed-identity-based ACL update. Using the active token directly instead."
            )
        response = requests.post(endpoint, headers=headers, json=body, timeout=120)
        if should_refresh_search_token(response):
            token_value = get_search_access_token()
            headers["Authorization"] = f"Bearer {token_value}"
            response = requests.post(endpoint, headers=headers, json=body, timeout=120)
        if response.status_code >= 400:
            logging.error(
                "Azure AI Search request failed status=%s url=%s payload=%s response=%s",
                response.status_code,
                response.url,
                json.dumps(body, default=str),
                response.text[:4000],
            )
        response.raise_for_status()
        return response.json()

    total_response = post({"search": "*", "top": 1, "count": True})
    total = int(total_response.get("@odata.count", 0))
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
