"""
Batch text-extraction + downstream-service caller.

For each row in an Excel file this tool:
  1. Reads: storage account name, container name, and the full blob path.
  2. Downloads the blob from Azure Storage (auth via `az login` / DefaultAzureCredential).
  3. Sends the file bytes to a Tika server to extract the text.
  4. POSTs the extracted text to the GuardPII 'recognize' endpoint for PII detection.
  5. Saves a summary Excel + one JSON file per row (the full response, ready to share).

Authentication
--------------
* Storage: uses a user-assigned **Managed Identity** when running on the VM
  (set USE_MANAGED_IDENTITY=true and MANAGED_IDENTITY_CLIENT_ID in .env). The MI needs the
  DATA-plane role "Storage Blob Data Reader" on the account/container. (Management-plane
  "Reader"/"Owner" does NOT grant access to blob *data*.)
  For local testing, set USE_MANAGED_IDENTITY=false to fall back to your `az login` identity.
* Tika + downstream service: plain HTTP; add any required key/token via .env.
"""

from __future__ import annotations

import argparse
import json
import logging
import mimetypes
import os
import re
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from dotenv import load_dotenv

from azure.core.exceptions import ClientAuthenticationError, HttpResponseError, ServiceRequestError
from azure.identity import DefaultAzureCredential, ManagedIdentityCredential
from azure.storage.blob import BlobServiceClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("ocr-tool")
# Quiet the very verbose Azure SDK HTTP/identity logging (IMDS token requests, etc.)
logging.getLogger("azure").setLevel(logging.WARNING)


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    input_file: str
    output_excel: str
    sheet_name: Optional[str]
    col_storage: str
    col_container: str
    col_fullpath: str
    tika_endpoint: str
    service_endpoint: str
    service_functions_key: Optional[str]
    detection_model: str
    output_dir: str = "response"
    allowed_extensions: tuple = (
        "pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif", "webp", "heic", "heif",
    )
    process_extensionless: bool = False
    max_workers: int = 5
    resume: bool = True

    @classmethod
    def from_env(cls) -> "Config":
        load_dotenv()
        return cls(
            input_file=os.getenv("INPUT_FILE", os.getenv("INPUT_EXCEL", "input_sample.csv")),
            output_excel=os.getenv("OUTPUT_EXCEL", "output.xlsx"),
            sheet_name=os.getenv("SHEET_NAME") or None,
            col_storage=os.getenv("COL_STORAGE", "Source"),
            col_container=os.getenv("COL_CONTAINER", "Container"),
            col_fullpath=os.getenv("COL_FULLPATH", "FullPath"),
            tika_endpoint=os.getenv("TIKA_ENDPOINT", "https://tr-tika.safesendwebsites.com/tika").strip(),
            service_endpoint=os.getenv(
                "SERVICE_ENDPOINT",
                "https://tr-orion-dev-guardpii.safesendwebsites.com/api/recognize",
            ).strip(),
            service_functions_key=os.getenv("SERVICE_FUNCTIONS_KEY") or None,
            detection_model=os.getenv("DETECTION_MODEL", "llm").strip(),
            output_dir=os.getenv("OUTPUT_DIR", "response"),
            allowed_extensions=tuple(
                e.strip().lstrip(".").lower()
                for e in os.getenv(
                    "ALLOWED_EXTENSIONS",
                    "pdf,png,jpg,jpeg,tif,tiff,bmp,gif,webp,heic,heif",
                ).split(",")
                if e.strip()
            ),
            process_extensionless=os.getenv("PROCESS_EXTENSIONLESS", "false").strip().lower()
            in ("1", "true", "yes", "y"),
            max_workers=int(os.getenv("MAX_WORKERS", "5") or "5"),
            resume=os.getenv("RESUME", "true").strip().lower() in ("1", "true", "yes", "y"),
        )


# --------------------------------------------------------------------------- #
# Small retry helper for transient failures (best practice: backoff + retry)
# --------------------------------------------------------------------------- #
def with_retry(fn, *, attempts: int = 3, base_delay: float = 2.0):
    last_exc = None
    for i in range(attempts):
        try:
            return fn()
        except ClientAuthenticationError:
            # Auth / identity errors (e.g. "Identity not found") won't fix themselves --
            # fail fast instead of burning ~14s of backoff per blob.
            raise
        except (ServiceRequestError, HttpResponseError, requests.RequestException) as exc:
            last_exc = exc
            wait = base_delay * (2 ** i)
            log.warning(
                "Transient error (attempt %d/%d): %s -- retrying in %.0fs",
                i + 1, attempts, exc, wait,
            )
            time.sleep(wait)
    raise last_exc  # type: ignore[misc]


# --------------------------------------------------------------------------- #
# Azure clients (one shared credential)
# --------------------------------------------------------------------------- #
_credential = None
_blob_clients: dict[str, BlobServiceClient] = {}
_blob_clients_lock = threading.Lock()


def build_credential():
    """
    Choose how to authenticate to Azure.

    On the VM (default): use the user-assigned Managed Identity.
        USE_MANAGED_IDENTITY=true
        MANAGED_IDENTITY_CLIENT_ID=<the MI's client id>
        (or MANAGED_IDENTITY_OBJECT_ID=<the MI's object/principal id> if the client id
         is unknown)

    Locally: set USE_MANAGED_IDENTITY=false to fall back to `az login`.
    """
    load_dotenv()  # make sure .env is loaded before we read auth settings
    use_mi = os.getenv("USE_MANAGED_IDENTITY", "true").strip().lower() in ("1", "true", "yes", "y")
    client_id = os.getenv("MANAGED_IDENTITY_CLIENT_ID", "").strip() or None
    object_id = os.getenv("MANAGED_IDENTITY_OBJECT_ID", "").strip() or None
    resource_id = os.getenv("MANAGED_IDENTITY_RESOURCE_ID", "").strip() or None

    if use_mi:
        if client_id:
            log.info("Auth: user-assigned Managed Identity (client_id=%s)", client_id)
            return ManagedIdentityCredential(client_id=client_id)
        if object_id:
            log.info("Auth: user-assigned Managed Identity (object_id=%s)", object_id)
            return ManagedIdentityCredential(identity_config={"object_id": object_id})
        if resource_id:
            log.info("Auth: user-assigned Managed Identity (resource_id=%s)", resource_id)
            return ManagedIdentityCredential(identity_config={"resource_id": resource_id})
        log.info("Auth: system-assigned Managed Identity")
        return ManagedIdentityCredential()

    log.info("Auth: DefaultAzureCredential (local dev / az login)")
    if client_id:
        return DefaultAzureCredential(managed_identity_client_id=client_id)
    return DefaultAzureCredential()


def get_credential():
    """Build the credential once and reuse it."""
    global _credential
    if _credential is None:
        _credential = build_credential()
    return _credential


def verify_auth() -> None:
    """
    Acquire a token once, up front, so authentication problems fail fast with a
    clear, actionable message instead of erroring on every single blob.
    """
    log.info("Checking Azure authentication...")
    try:
        get_credential().get_token("https://storage.azure.com/.default")
    except ClientAuthenticationError as exc:
        raise SystemExit(
            "\nAzure authentication FAILED -- could not get a token for the managed identity.\n"
            f"  Detail: {exc}\n\n"
            "This is an IDENTITY problem, NOT the storage account name: token acquisition\n"
            "happens before any storage call. 'Identity not found' from IMDS means the id in\n"
            ".env is not a user-assigned identity attached to THIS VM. Fix one of:\n"
            "  1. Attach Abyss-04 to this VM:  Portal > this VM > Identity > 'User assigned' tab.\n"
            "  2. Use the correct CLIENT ID:   Portal > Managed Identities > Abyss-04 > 'Client ID',\n"
            "     then set MANAGED_IDENTITY_CLIENT_ID to it.\n"
            "  3. If your GUID is the OBJECT (principal) id, set MANAGED_IDENTITY_OBJECT_ID instead\n"
            "     (clear MANAGED_IDENTITY_CLIENT_ID). Or set MANAGED_IDENTITY_RESOURCE_ID to the\n"
            "     identity's full ARM resource id.\n"
        )
    log.info("Auth OK.")


def blob_service(storage_name: str) -> BlobServiceClient:
    """Return a cached BlobServiceClient for the given storage account (thread-safe)."""
    with _blob_clients_lock:
        client = _blob_clients.get(storage_name)
        if client is None:
            url = f"https://{storage_name}.blob.core.windows.net"
            client = BlobServiceClient(account_url=url, credential=get_credential())
            _blob_clients[storage_name] = client
        return client


# --------------------------------------------------------------------------- #
# Core steps
# --------------------------------------------------------------------------- #
def _accounts_match(segment: str, account: str) -> bool:
    """
    True if a path segment names the given storage account.

    Tolerates a spurious leading 'o' seen in some FullPath values (e.g.
    'oqassrcontentstore' where the real account is 'qassrcontentstore' -- the
    'o'-prefixed name doesn't even resolve in DNS).
    """
    s, a = segment.strip().lower(), account.strip().lower()
    if not a:
        return False
    return s == a or s == "o" + a or "o" + s == a


def parse_full_path(full_path: str, col_storage: str = "", col_container: str = ""):
    """
    Work out (storage_account, container, blob_name) for one row.

    The Source/Container columns are AUTHORITATIVE for the storage account and
    container: the FullPath's own first segment has proven unreliable (an extra
    leading 'o' was seen, e.g. 'oqassrcontentstore' vs the real
    'qassrcontentstore'). FullPath is used only to derive the blob name -- we strip
    any leading storage-account and container segments (matched by NAME, tolerating
    the 'o' quirk) and treat the rest as the blob path. When a column is missing we
    fall back to the value parsed from the path itself.

    Examples (with Source='qassrcontentstore', Container='documentation'):
      /oqassrcontentstore/documentation/ssr/v1/Doc.pdf
          -> ("qassrcontentstore", "documentation", "ssr/v1/Doc.pdf")
      /qassrcontentstore/documentation/ssr/v1/Doc.pdf
          -> ("qassrcontentstore", "documentation", "ssr/v1/Doc.pdf")
    """
    p = str(full_path).strip().replace("\\", "/")

    host_account = ""
    if "://" in p:                       # URL form -> reduce to the path after the host
        host, _, rest = p.split("://", 1)[1].partition("/")
        host_account = host.split(".", 1)[0]
        p = rest

    segments = [s for s in p.split("/") if s]

    # Storage account: prefer the Source column, then the URL host, then the 1st
    # path segment (only when the path clearly carries storage + container + blob).
    positional_storage = segments[0] if len(segments) >= 3 else ""
    storage = (col_storage or host_account or positional_storage).strip()

    # Drop a leading storage-account segment if the path repeats it (handles the
    # bogus 'o' prefix too) so it isn't mistaken for part of the blob name.
    if segments and _accounts_match(segments[0], storage):
        segments = segments[1:]

    # Container: prefer the Container column, else the next path segment.
    container = (col_container or (segments[0] if segments else "")).strip()

    # Drop a leading container segment if the path repeats it.
    if segments and container and segments[0].lower() == container.lower():
        segments = segments[1:]

    return storage, container, "/".join(segments)


def blob_to_json_name(blob_name: str) -> str:
    """
    Build a safe .json filename from a blob name (same base name as the blob).

    The blob's folder path is flattened so identically named files in different
    folders don't overwrite each other (e.g. 'copy/Metadata.png' vs 'Metadata.png').
    """
    name = blob_name.strip("/")
    stem, _ext = os.path.splitext(name)          # drop the original extension
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem or name)
    return f"{stem or 'response'}.json"


def download_blob(storage_name: str, container: str, blob_name: str) -> bytes:
    client = blob_service(storage_name).get_blob_client(container=container, blob=blob_name)
    return with_retry(lambda: client.download_blob().readall())


def run_tika(cfg: Config, file_bytes: bytes, blob_name: str) -> str:
    """
    Send the file bytes to the Tika server and return the extracted plain text.

    Mirrors:
      curl -X PUT '<TIKA_ENDPOINT>' -H 'Accept: text/plain' \\
           -H 'Content-Type: application/pdf' --data-binary '@file.pdf'

    We request `text/plain` so Tika returns clean text -- the JSON variant wraps the
    text in HTML markup. The Content-Type is guessed from the file extension; when it
    can't be guessed (e.g. an extensionless blob) we send `application/octet-stream`
    so Tika auto-detects the real type from the bytes.
    """
    content_type, _ = mimetypes.guess_type(blob_name)
    headers = {
        "Accept": "text/plain",
        "Content-Type": content_type or "application/octet-stream",
    }

    def _put():
        resp = requests.put(cfg.tika_endpoint, data=file_bytes, headers=headers, timeout=180)
        resp.raise_for_status()
        return resp.text.strip()

    return with_retry(_put)


def call_recognize(cfg: Config, text: str):
    """
    POST the extracted text to the GuardPII 'recognize' endpoint and return its JSON.

    Mirrors:
      curl '<SERVICE_ENDPOINT>' -H 'Content-Type: application/json' \\
           -H 'x-functions-key: <key>' -H 'x-operation-id: <fresh-guid>' \\
           --data '{ "texts": ["..."], "detection_model": "llm" }'

    IMPORTANT: x-operation-id must be UNIQUE per request -- reusing one makes the
    service return HTTP 500. We generate a fresh UUID for every call.
    """
    headers = {
        "Content-Type": "application/json",
        "x-operation-id": str(uuid.uuid4()),
    }
    if cfg.service_functions_key:
        headers["x-functions-key"] = cfg.service_functions_key

    payload = {
        "texts": [text],
        "detection_model": cfg.detection_model,
    }

    def _post():
        resp = requests.post(cfg.service_endpoint, json=payload, headers=headers, timeout=120)
        resp.raise_for_status()
        try:
            return resp.json()
        except ValueError:
            return {"raw": resp.text}

    return with_retry(_post)


# --------------------------------------------------------------------------- #
# Input reading
# --------------------------------------------------------------------------- #
def read_input_table(path: str, sheet_name: Optional[str]) -> pd.DataFrame:
    """Read the input list from CSV or Excel, chosen by the file extension."""
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    return pd.read_excel(path, sheet_name=sheet_name or 0, dtype=str)


# --------------------------------------------------------------------------- #
# Per-blob worker (runs on a thread-pool thread)
# --------------------------------------------------------------------------- #
def process_one(cfg: Config, row_num: int, storage: str, container: str, blob_name: str) -> dict:
    """
    Handle a single blob end-to-end and return a result record.

    Safe to run concurrently: it only touches per-call state plus the thread-safe
    shared blob-client cache, and each row writes its own uniquely named JSON file.
    Any failure is captured in the record (status='error') so one bad blob never
    stops the batch.
    """
    rec: dict = {
        "row": row_num,
        "storage": storage,
        "container": container,
        "blob_name": blob_name,
        "status": "ok",
        "text_chars": 0,
        "output_json": "",
        "error": "",
    }

    # Extension gate: process allowed types (pdf/images); optionally try extensionless.
    ext = os.path.splitext(blob_name)[1].lstrip(".").lower()
    if not ext:
        if not cfg.process_extensionless:
            rec["status"] = "skipped"
            rec["error"] = "extensionless blob (PROCESS_EXTENSIONLESS is off)"
            return rec
    elif cfg.allowed_extensions and ext not in cfg.allowed_extensions:
        rec["status"] = "skipped"
        rec["error"] = f"extension '{ext}' not in {list(cfg.allowed_extensions)}"
        return rec

    # Resume: if this blob's JSON already exists, don't redo it (lets you re-run a
    # huge batch after a crash and only pick up what's left).
    out_json = Path(cfg.output_dir) / blob_to_json_name(blob_name)
    if cfg.resume and out_json.exists():
        rec["status"] = "exists"
        rec["output_json"] = str(out_json)
        return rec

    try:
        data = download_blob(storage, container, blob_name)
        text = run_tika(cfg, data, blob_name)
        rec["text_chars"] = len(text)

        response = call_recognize(cfg, text)
        rec["service_response"] = response

        out_json.write_text(
            json.dumps(response, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        rec["output_json"] = str(out_json)
    except Exception as exc:  # keep going on a single-row failure
        rec["status"] = "error"
        rec["error"] = str(exc)
        log.error("[row %d] failed: %s", row_num, exc)

    return rec


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description="Batch Tika text extraction + GuardPII recognize")
    parser.add_argument("--input", help="Input file path, CSV or Excel (overrides .env)")
    parser.add_argument("--output", help="Output Excel path (overrides .env)")
    parser.add_argument("--limit", type=int, default=0, help="Only process the first N rows (0 = all)")
    parser.add_argument(
        "--workers", type=int, default=0,
        help="How many blobs to process in parallel (0 = use MAX_WORKERS from .env, default 5)",
    )
    parser.add_argument(
        "--process-extensionless", dest="process_extensionless",
        action="store_true", default=None,
        help="Also process blobs with NO extension (Tika auto-detects the real type). "
             "Overrides PROCESS_EXTENSIONLESS from .env for this run.",
    )
    args = parser.parse_args()

    cfg = Config.from_env()
    if args.input:
        cfg.input_file = args.input
    if args.output:
        cfg.output_excel = args.output
    if args.workers:
        cfg.max_workers = args.workers
    cfg.max_workers = max(1, cfg.max_workers)
    if args.process_extensionless is not None:
        cfg.process_extensionless = args.process_extensionless

    if not cfg.tika_endpoint:
        raise SystemExit("TIKA_ENDPOINT is not set. Copy .env.example to .env and fill it in.")
    if not cfg.service_endpoint:
        raise SystemExit("SERVICE_ENDPOINT is not set. Copy .env.example to .env and fill it in.")
    if not cfg.service_functions_key:
        raise SystemExit("SERVICE_FUNCTIONS_KEY is not set. Add the x-functions-key to your .env.")

    verify_auth()

    Path(cfg.output_dir).mkdir(exist_ok=True)

    log.info("Reading %s", cfg.input_file)
    df = read_input_table(cfg.input_file, cfg.sheet_name)
    if cfg.col_fullpath not in df.columns:
        raise SystemExit(
            f"Column '{cfg.col_fullpath}' not found. Available columns: {list(df.columns)}. "
            f"Set COL_FULLPATH in .env. (Storage and container are parsed from this column; "
            f"COL_STORAGE/COL_CONTAINER are optional fallbacks.)"
        )
    has_storage_col = cfg.col_storage in df.columns
    has_container_col = cfg.col_container in df.columns

    if args.limit:
        df = df.head(args.limit)

    # Build the work list first (storage/container from the authoritative columns;
    # blob name from FullPath).
    tasks: list[tuple] = []
    for i, row in df.iterrows():
        full_path = str(row[cfg.col_fullpath]).strip()
        fb_storage = str(row[cfg.col_storage]).strip() if has_storage_col else ""
        fb_container = str(row[cfg.col_container]).strip() if has_container_col else ""
        storage, container, blob_name = parse_full_path(full_path, fb_storage, fb_container)
        tasks.append((int(i) + 1, storage, container, blob_name))

    total = len(tasks)
    log.info("Processing %d row(s) with %d parallel worker(s)...", total, cfg.max_workers)

    # Fan out across a small thread pool. The steps are all network I/O (blob
    # download, Tika, recognize), so threads give real speed-up despite the GIL.
    results: list[dict] = []
    done = 0
    with ThreadPoolExecutor(max_workers=cfg.max_workers) as pool:
        futures = {
            pool.submit(process_one, cfg, row_num, storage, container, blob_name): row_num
            for (row_num, storage, container, blob_name) in tasks
        }
        for fut in as_completed(futures):
            rec = fut.result()
            results.append(rec)
            done += 1
            status, blob = rec["status"], rec["blob_name"]
            if status == "ok":
                log.info("[%d/%d] ok: %s (%d chars) -> %s",
                         done, total, blob, rec["text_chars"], rec["output_json"])
            elif status == "exists":
                log.info("[%d/%d] already done, skipping: %s", done, total, blob)
            elif status == "skipped":
                log.info("[%d/%d] skipped (%s): %s", done, total, rec["error"], blob)
            else:
                log.info("[%d/%d] ERROR: %s (%s)", done, total, blob, rec["error"])

    # Keep the output readable: sort back into input row order (the pool finishes
    # out of order).
    results.sort(key=lambda r: r["row"])

    # Write a readable summary Excel (nested JSON flattened to a string).
    summary = pd.DataFrame(results)
    if "service_response" in summary.columns:
        summary["service_response"] = summary["service_response"].apply(
            lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v
        )
    summary.to_excel(cfg.output_excel, index=False)
    n_ok = sum(1 for r in results if r["status"] == "ok")
    n_exists = sum(1 for r in results if r["status"] == "exists")
    n_skipped = sum(1 for r in results if r["status"] == "skipped")
    n_error = sum(1 for r in results if r["status"] == "error")
    log.info(
        "Done. ok=%d already_done=%d skipped=%d error=%d | Summary -> %s | Per-blob responses -> %s/",
        n_ok, n_exists, n_skipped, n_error, cfg.output_excel, cfg.output_dir,
    )


if __name__ == "__main__":
    main()
