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
import time
import uuid
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
    allowed_extensions: tuple = ("png", "pdf")
    process_extensionless: bool = True

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
                for e in os.getenv("ALLOWED_EXTENSIONS", "png,pdf").split(",")
                if e.strip()
            ),
            process_extensionless=os.getenv("PROCESS_EXTENSIONLESS", "true").strip().lower()
            in ("1", "true", "yes", "y"),
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
    """Return a cached BlobServiceClient for the given storage account."""
    if storage_name not in _blob_clients:
        url = f"https://{storage_name}.blob.core.windows.net"
        _blob_clients[storage_name] = BlobServiceClient(account_url=url, credential=get_credential())
    return _blob_clients[storage_name]


# --------------------------------------------------------------------------- #
# Core steps
# --------------------------------------------------------------------------- #
def parse_full_path(full_path: str, fallback_storage: str = "", fallback_container: str = ""):
    """
    Parse a FullPath value into (storage_account, container, blob_name).

    The path is split on '/': the 1st segment is the storage account, the 2nd is
    the container, and everything after is the blob name (the file path inside the
    container). URL-form paths are supported too. The column values are used only
    as a fallback when the path itself doesn't carry enough segments.

    Examples:
      /devssecontentstore/abc/copy/Metadata.png
          -> ("devssecontentstore", "abc", "copy/Metadata.png")
      devssecontentstore/abc/file.pdf
          -> ("devssecontentstore", "abc", "file.pdf")
      https://devssecontentstore.blob.core.windows.net/abc/file.pdf
          -> ("devssecontentstore", "abc", "file.pdf")
    """
    p = str(full_path).strip().replace("\\", "/")

    # URL form: the storage account lives in the host name.
    if "://" in p:
        host, _, rest = p.split("://", 1)[1].partition("/")
        storage = host.split(".", 1)[0] or fallback_storage
        segments = [s for s in rest.split("/") if s]
        container = segments[0] if segments else fallback_container
        return storage, container, "/".join(segments[1:])

    segments = [s for s in p.split("/") if s]
    if len(segments) >= 3:
        return segments[0], segments[1], "/".join(segments[2:])
    if len(segments) == 2:                       # no storage in the path
        return fallback_storage, segments[0], segments[1]
    return fallback_storage, fallback_container, segments[0] if segments else ""


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
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description="Batch Tika text extraction + GuardPII recognize")
    parser.add_argument("--input", help="Input file path, CSV or Excel (overrides .env)")
    parser.add_argument("--output", help="Output Excel path (overrides .env)")
    parser.add_argument("--limit", type=int, default=0, help="Only process the first N rows (0 = all)")
    args = parser.parse_args()

    cfg = Config.from_env()
    if args.input:
        cfg.input_file = args.input
    if args.output:
        cfg.output_excel = args.output

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

    results: list[dict] = []
    for i, row in df.iterrows():
        full_path = str(row[cfg.col_fullpath]).strip()
        fb_storage = str(row[cfg.col_storage]).strip() if has_storage_col else ""
        fb_container = str(row[cfg.col_container]).strip() if has_container_col else ""
        # Storage, container and blob name are all taken from the FullPath column
        # (with the storage/container columns as fallback).
        storage, container, blob_name = parse_full_path(full_path, fb_storage, fb_container)

        rec: dict = {
            "row": int(i) + 1,
            "storage": storage,
            "container": container,
            "blob_name": blob_name,
            "status": "ok",
            "text_chars": 0,
            "output_json": "",
            "error": "",
        }
        log.info("[%d] %s / %s / %s", rec["row"], storage, container, blob_name)

        # Decide whether to process this blob based on its extension:
        #   * allowed extensions (e.g. png, pdf) -> process
        #   * no extension -> try it too (often just a missing extension); Tika
        #     auto-detects the real type, and if it can't parse it we skip on the
        #     error handler below and move on
        #   * anything else (.msi, .zip, ...) -> skip
        ext = os.path.splitext(blob_name)[1].lstrip(".").lower()
        if not ext:
            if not cfg.process_extensionless:
                rec["status"] = "skipped"
                rec["error"] = "extensionless blob (PROCESS_EXTENSIONLESS is off)"
                log.info("[%d] skipped (%s)", rec["row"], rec["error"])
                results.append(rec)
                continue
            log.info("[%d] extensionless -- trying Tika auto-detect", rec["row"])
        elif cfg.allowed_extensions and ext not in cfg.allowed_extensions:
            rec["status"] = "skipped"
            rec["error"] = f"extension '{ext}' not in {list(cfg.allowed_extensions)}"
            log.info("[%d] skipped (%s)", rec["row"], rec["error"])
            results.append(rec)
            continue

        try:
            data = download_blob(storage, container, blob_name)
            text = run_tika(cfg, data, blob_name)
            rec["text_chars"] = len(text)

            response = call_recognize(cfg, text)
            rec["service_response"] = response

            # Save the recognize response as JSON, named after the blob.
            out_json = Path(cfg.output_dir) / blob_to_json_name(blob_name)
            out_json.write_text(
                json.dumps(response, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            rec["output_json"] = str(out_json)
        except Exception as exc:  # keep going on a single-row failure
            rec["status"] = "error"
            rec["error"] = str(exc)
            log.error("[%d] failed: %s", rec["row"], exc)

        results.append(rec)

    # Write a readable summary Excel (nested JSON flattened to a string).
    summary = pd.DataFrame(results)
    if "service_response" in summary.columns:
        summary["service_response"] = summary["service_response"].apply(
            lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v
        )
    summary.to_excel(cfg.output_excel, index=False)
    n_ok = sum(1 for r in results if r["status"] == "ok")
    n_skipped = sum(1 for r in results if r["status"] == "skipped")
    n_error = sum(1 for r in results if r["status"] == "error")
    log.info(
        "Done. processed=%d skipped=%d error=%d | Summary -> %s | Per-blob responses -> %s/",
        n_ok, n_skipped, n_error, cfg.output_excel, cfg.output_dir,
    )


if __name__ == "__main__":
    main()
