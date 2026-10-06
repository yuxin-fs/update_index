import json
import logging
import os
from datetime import datetime, timezone
from typing import List

import azure.functions as func
import requests
from azure.identity import DefaultAzureCredential

import update_acls

app = func.FunctionApp()

SEARCH_API_VERSION = os.getenv("SEARCH_API_VERSION", "2026-05-01-preview")
INDEXER_POLL_SECONDS = int(os.getenv("INDEXER_POLL_SECONDS", "30"))
INDEXER_TIMEOUT_MINUTES = int(os.getenv("INDEXER_TIMEOUT_MINUTES", "180"))
UPDATE_ACLS_TIMEOUT_MINUTES = int(os.getenv("UPDATE_ACLS_TIMEOUT_MINUTES", "180"))


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


def _run_acl_update(index_name: str) -> None:
    logging.info("Running ACL update in-process for index '%s'.", index_name)
    update_acls.run_for_index(
        index_name=index_name,
        site_url=os.getenv("SHAREPOINT_SITE_URL", update_acls.DEFAULT_SITE_URL),
        search_service=os.getenv("SEARCH_SERVICE", update_acls.DEFAULT_SEARCH_SERVICE),
        api_version=os.getenv("SEARCH_API_VERSION", update_acls.DEFAULT_API_VERSION),
        tenant_id=os.getenv("TENANT_ID"),
        client_id=os.getenv("SHAREPOINT_CLIENT_ID"),
        pfx_path=os.getenv("SHAREPOINT_PFX_PATH"),
        pfx_password=os.getenv("SHAREPOINT_PFX_PASSWORD", ""),
        query_top=200,
        write_batch_size=200,
        max_chunks=0,
        no_elevated_read=False,
        dry_run=False,
    )


@app.timer_trigger(
    schedule="0 12 20 * * *",
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

        chunks_index = f"{prefix}-chunks-index"
        try:
            token = _search_token()
            for current_indexer in indexer_names:
                requested_at = _run_indexer(endpoint, current_indexer, token)
                _wait_for_indexer(endpoint, current_indexer, token, requested_at)
            _run_acl_update(chunks_index)
            logging.info("Completed indexer and ACL update for prefix '%s'.", prefix)
        except Exception as exc:
            logging.exception("Prefix '%s' failed: %s", prefix, exc)
            failures.append({"prefix": prefix, "error": str(exc)})

    if failures:
        raise RuntimeError(f"One or more prefixes failed: {json.dumps(failures)}")
