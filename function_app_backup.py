import json
import logging
import os
from datetime import datetime, timezone
from typing import Any, Dict, List

import azure.functions as func
import requests
from azure.identity import DefaultAzureCredential
from sharepoint_acl import (
    derive_site_url_from_weburi,
    get_sharepoint_access_token,
    load_pfx_bytes,
    normalize_group_ids,
    tenant_ids_for_weburi,
)

app = func.FunctionApp()

SEARCH_API_VERSION = os.getenv("SEARCH_API_VERSION", "2026-05-01-preview")
INDEXER_POLL_SECONDS = int(os.getenv("INDEXER_POLL_SECONDS", "30"))
INDEXER_TIMEOUT_MINUTES = int(os.getenv("INDEXER_TIMEOUT_MINUTES", "180"))


def _required_setting(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Required app setting '{name}' is missing or empty.")
    return value


def _index_prefixes() -> List[str]:
    raw = _required_setting("INDEX_PREFIXES")
    try:
        parsed = json.loads(raw)
        if not isinstance(parsed, list):
            raise ValueError("INDEX_PREFIXES JSON must be an array.")
        values = [str(item).strip() for item in parsed if str(item).strip()]
    except json.JSONDecodeError:
        values = [item.strip() for item in raw.split(",") if item.strip()]
    if not values:
        raise RuntimeError("INDEX_PREFIXES contains no values.")
    return values


def _search_endpoint() -> str:
    service = _required_setting("SEARCH_SERVICE")
    if service.startswith("https://"):
        return service.rstrip("/")
    return f"https://{service}.search.windows.net"


def _search_token() -> str:
    return DefaultAzureCredential().get_token("https://search.azure.com/.default").token


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _run_indexer(endpoint: str, indexer_name: str, token: str) -> datetime:
    url = f"{endpoint}/indexers/{indexer_name}/run?api-version={SEARCH_API_VERSION}"
    requested_at = datetime.now(timezone.utc)
    response = requests.post(url, headers=_headers(token), timeout=60)
    if response.status_code not in (202, 204):
        raise RuntimeError(
            f"Could not start indexer '{indexer_name}'. HTTP {response.status_code}: {response.text[:2000]}"
        )
    logging.info("Started indexer %s (HTTP %s).", indexer_name, response.status_code)
    return requested_at


def _parse_time(value):
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def _wait_for_indexer(endpoint: str, indexer_name: str, token: str, requested_at: datetime) -> None:
    import time

    url = f"{endpoint}/indexers/{indexer_name}/status?api-version={SEARCH_API_VERSION}"
    deadline = time.monotonic() + INDEXER_TIMEOUT_MINUTES * 60

    while time.monotonic() < deadline:
        response = requests.get(url, headers=_headers(token), timeout=60)
        if response.status_code == 401:
            token = _search_token()
            response = requests.get(url, headers=_headers(token), timeout=60)
        response.raise_for_status()
        payload = response.json()
        last = payload.get("lastResult") or {}
        started = _parse_time(last.get("startTime"))
        ended = _parse_time(last.get("endTime"))
        result_status = str(last.get("status", "")).lower()

        logging.info(
            "Indexer %s: serviceStatus=%s resultStatus=%s start=%s end=%s processed=%s failed=%s",
            indexer_name,
            payload.get("status"),
            result_status,
            last.get("startTime"),
            last.get("endTime"),
            last.get("itemsProcessed"),
            last.get("itemsFailed"),
        )

        # Ignore a previous execution briefly returned after the run request.
        is_requested_run = started is not None and started >= requested_at
        if is_requested_run and ended is not None:
            if result_status in ("success", "warning"):
                return
            raise RuntimeError(
                f"Indexer '{indexer_name}' ended with status '{result_status}'. "
                f"Errors: {json.dumps(last.get('errors') or [])[:4000]}"
            )
        time.sleep(INDEXER_POLL_SECONDS)

    raise TimeoutError(f"Indexer '{indexer_name}' exceeded {INDEXER_TIMEOUT_MINUTES} minutes.")


@app.route(
    route="enrich-sharepoint-acls",
    methods=["POST"],
    auth_level=func.AuthLevel.FUNCTION,
)
def enrich_sharepoint_acls(req: func.HttpRequest) -> func.HttpResponse:
    try:
        payload = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "Request body must be valid JSON."}),
            status_code=400,
            mimetype="application/json",
        )

    values = payload.get("values")
    if not isinstance(values, list):
        return func.HttpResponse(
            json.dumps({"error": "Request must contain a values array."}),
            status_code=400,
            mimetype="application/json",
        )

    try:
        tenant_id = _required_setting("TENANT_ID")
        client_id = _required_setting("SHAREPOINT_CLIENT_ID")
        default_site_url = _required_setting("SHAREPOINT_SITE_URL")
        pfx_data = load_pfx_bytes(os.getenv("SHAREPOINT_PFX_PATH"))
        sharepoint_token = get_sharepoint_access_token(
            tenant_id=tenant_id,
            client_id=client_id,
            site_url=default_site_url,
            pfx_data=pfx_data,
            pfx_password=os.getenv("SHAREPOINT_PFX_PASSWORD", ""),
        )
    except Exception as exc:
        logging.exception("SharePoint authentication failed: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "SharePoint authentication failed."}),
            status_code=500,
            mimetype="application/json",
        )

    response_values: List[Dict[str, Any]] = []
    weburi_cache: Dict[str, set[str]] = {}
    group_cache: Dict[str, set[str]] = {}

    for record in values:
        record_id = str(record.get("recordId") or "")
        data = record.get("data") or {}
        web_uri = str(data.get("url") or "").strip()

        existing_user_ids = normalize_group_ids(data.get("existingUserIds") or [])
        existing_group_ids = normalize_group_ids(data.get("existingGroupIds") or [])

        result: Dict[str, Any] = {
            "recordId": record_id,
            "data": {},
            "errors": [],
            "warnings": [],
        }

        if not record_id:
            result["errors"].append({"message": "The input record has no recordId."})
            response_values.append(result)
            continue

        if not web_uri:
            result["data"]["mergedUserIds"] = sorted(existing_user_ids)
            result["data"]["mergedGroupIds"] = sorted(existing_group_ids)
            result["warnings"].append(
                {"message": "No SharePoint web URI was supplied. Native ACL metadata was preserved."}
            )
            response_values.append(result)
            continue

        try:
            if web_uri not in weburi_cache:
                site_url = derive_site_url_from_weburi(web_uri, default_site_url)
                resolved_ids = tenant_ids_for_weburi(
                    site_url=site_url,
                    token=sharepoint_token,
                    web_uri=web_uri,
                    group_cache=group_cache,
                )
                weburi_cache[web_uri] = normalize_group_ids(resolved_ids)

            merged_group_ids = existing_group_ids | weburi_cache[web_uri]
            result["data"]["mergedUserIds"] = sorted(existing_user_ids)
            result["data"]["mergedGroupIds"] = sorted(merged_group_ids)
        except requests.HTTPError as exc:
            response = getattr(exc, "response", None)
            status = response.status_code if response is not None else None
            response_text = response.text[:2000] if response is not None else str(exc)
            logging.exception(
                "ACL enrichment failed. recordId=%s url=%s status=%s response=%s",
                record_id,
                web_uri,
                status,
                response_text,
            )
            result["data"]["mergedUserIds"] = sorted(existing_user_ids)
            result["data"]["mergedGroupIds"] = sorted(existing_group_ids)
            result["warnings"].append(
                {"message": "Custom ACL resolution failed. Native SharePoint ACL metadata was preserved."}
            )
        except Exception as exc:
            logging.exception(
                "Unexpected ACL enrichment failure. recordId=%s url=%s error=%s",
                record_id,
                web_uri,
                exc,
            )
            result["data"]["mergedUserIds"] = sorted(existing_user_ids)
            result["data"]["mergedGroupIds"] = sorted(existing_group_ids)
            result["warnings"].append(
                {"message": "Unexpected custom ACL resolution failure. Native SharePoint ACL metadata was preserved."}
            )

        response_values.append(result)

    return func.HttpResponse(
        json.dumps({"values": response_values}),
        status_code=200,
        mimetype="application/json",
    )


@app.timer_trigger(
    schedule="0 10 20 * * *",
    arg_name="timer",
    run_on_startup=False,
    use_monitor=True,
)
def daily_acl_refresh(timer: func.TimerRequest) -> None:
    if timer.past_due:
        logging.warning("The daily ACL refresh timer is past due.")

    endpoint = _search_endpoint()
    failures = []

    for prefix in _index_prefixes():
        indexer_name = f"{prefix}-pages-indexer"
        indexer_names = [indexer_name]
        if indexer_name == "spa-sharepoint-pages-indexer":
            indexer_names.append("spa-sharepoint-docs-indexer")

        try:
            token = _search_token()
            for current_indexer in indexer_names:
                requested_at = _run_indexer(endpoint, current_indexer, token)
                _wait_for_indexer(endpoint, current_indexer, token, requested_at)
            logging.info("Completed ingestion with ACL enrichment for prefix '%s'.", prefix)
        except Exception as exc:
            logging.exception("Prefix '%s' failed: %s", prefix, exc)
            failures.append({"prefix": prefix, "error": str(exc)})

    if failures:
        raise RuntimeError(f"One or more prefixes failed: {json.dumps(failures)}")
