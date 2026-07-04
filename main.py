"""
Batch text-extraction + downstream-service caller.

For each row in an Excel file this tool:
  1. Reads: storage account name, container name, and the full blob path.
  2. Downloads the blob from Azure Storage (auth via `az login` / DefaultAzureCredential).
  3. Sends the file bytes to a Tika server to extract the text.
  4. POSTs the extracted text as JSON to your downstream service.
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
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from dotenv import load_dotenv

from azure.core.exceptions import HttpResponseError, ServiceRequestError
from azure.identity import DefaultAzureCredential, ManagedIdentityCredential
from azure.storage.blob import BlobServiceClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("ocr-tool")


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    input_excel: str
    output_excel: str
    sheet_name: Optional[str]
    col_storage: str
    col_container: str
    col_fullpath: str
    tika_endpoint: str
    service_endpoint: str
    service_api_key: Optional[str]
    output_dir: str = "responses"

    @classmethod
    def from_env(cls) -> "Config":
        load_dotenv()
        return cls(
            input_excel=os.getenv("INPUT_EXCEL", "input.xlsx"),
            output_excel=os.getenv("OUTPUT_EXCEL", "output.xlsx"),
            sheet_name=os.getenv("SHEET_NAME") or None,
            col_storage=os.getenv("COL_STORAGE", "storage_name"),
            col_container=os.getenv("COL_CONTAINER", "container"),
            col_fullpath=os.getenv("COL_FULLPATH", "full_path"),
            tika_endpoint=os.getenv("TIKA_ENDPOINT", "https://tr-tika.safesendwebsites.com/tika").strip(),
            service_endpoint=os.getenv("SERVICE_ENDPOINT", "").strip(),
            service_api_key=os.getenv("SERVICE_API_KEY") or None,
        )


# --------------------------------------------------------------------------- #
# Small retry helper for transient failures (best practice: backoff + retry)
# --------------------------------------------------------------------------- #
def with_retry(fn, *, attempts: int = 3, base_delay: float = 2.0):
    last_exc = None
    for i in range(attempts):
        try:
            return fn()
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

    if use_mi:
        if client_id:
            log.info("Auth: user-assigned Managed Identity (client_id=%s)", client_id)
            return ManagedIdentityCredential(client_id=client_id)
        if object_id:
            log.info("Auth: user-assigned Managed Identity (object_id=%s)", object_id)
            return ManagedIdentityCredential(identity_config={"object_id": object_id})
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


def blob_service(storage_name: str) -> BlobServiceClient:
    """Return a cached BlobServiceClient for the given storage account."""
    if storage_name not in _blob_clients:
        url = f"https://{storage_name}.blob.core.windows.net"
        _blob_clients[storage_name] = BlobServiceClient(account_url=url, credential=get_credential())
    return _blob_clients[storage_name]


# --------------------------------------------------------------------------- #
# Core steps
# --------------------------------------------------------------------------- #
def extract_blob_name(full_path: str, storage_name: str, container: str) -> str:
    """
    Turn the 'full path' cell into the blob name (everything after the container).

    Handles values such as:
      * storageacct/mycontainer/folder/file.pdf
      * mycontainer/folder/file.pdf
      * https://storageacct.blob.core.windows.net/mycontainer/folder/file.pdf
    """
    p = str(full_path).strip().replace("\\", "/")
    if "://" in p:
        p = p.split("://", 1)[1]          # drop the https:// scheme
    p = p.lstrip("/")

    marker = f"{container}/"
    idx = p.find(marker)
    if idx != -1:
        return p[idx + len(marker):]

    # Fallback: assume "<storage>/<container>/<blob...>" or "<container>/<blob...>"
    segments = p.split("/")
    if segments and segments[0] == storage_name and len(segments) > 2:
        return "/".join(segments[2:])
    if len(segments) > 1:
        return "/".join(segments[1:])
    return p


def download_blob(storage_name: str, container: str, blob_name: str) -> bytes:
    client = blob_service(storage_name).get_blob_client(container=container, blob=blob_name)
    return with_retry(lambda: client.download_blob().readall())


def _text_from_tika_json(data) -> str:
    """Pull the extracted text out of a Tika JSON response (handles a few shapes)."""
    if isinstance(data, str):
        return data
    if isinstance(data, list):  # /rmeta-style: a list of metadata dicts
        return "\n".join(_text_from_tika_json(d) for d in data)
    if isinstance(data, dict):
        for key in ("X-TIKA:content", "content", "text"):
            if data.get(key):
                return str(data[key])
        return json.dumps(data, ensure_ascii=False)
    return str(data)


def run_tika(cfg: Config, file_bytes: bytes, blob_name: str) -> str:
    """
    Send the file bytes to the Tika server and return the extracted text.

    Mirrors:
      curl -X PUT '<TIKA_ENDPOINT>' -H 'Accept: application/json' \\
           -H 'Content-Type: application/pdf' --data-binary '@file.pdf'

    The Content-Type is guessed from the file extension so non-PDF files work too.
    """
    content_type, _ = mimetypes.guess_type(blob_name)
    headers = {
        "Accept": "application/json",
        "Content-Type": content_type or "application/pdf",
    }

    def _put():
        resp = requests.put(cfg.tika_endpoint, data=file_bytes, headers=headers, timeout=180)
        resp.raise_for_status()
        # Depending on the server, /tika returns plain text or JSON.
        if "json" in resp.headers.get("Content-Type", "").lower():
            return _text_from_tika_json(resp.json())
        return resp.text

    return with_retry(_put)


def call_service(cfg: Config, text: str, meta: dict) -> dict:
    """
    POST the extracted text to your downstream service as JSON.

    >>> Customize the `payload` below to match what your service expects. <<<
    """
    headers = {"Content-Type": "application/json"}
    if cfg.service_api_key:
        headers["Authorization"] = f"Bearer {cfg.service_api_key}"

    payload = {
        "file": meta.get("blob_name"),
        "content": text,
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
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description="Batch OCR + downstream service caller")
    parser.add_argument("--input", help="Input Excel path (overrides .env)")
    parser.add_argument("--output", help="Output Excel path (overrides .env)")
    parser.add_argument("--limit", type=int, default=0, help="Only process the first N rows (0 = all)")
    args = parser.parse_args()

    cfg = Config.from_env()
    if args.input:
        cfg.input_excel = args.input
    if args.output:
        cfg.output_excel = args.output

    if not cfg.tika_endpoint:
        raise SystemExit("TIKA_ENDPOINT is not set. Copy .env.example to .env and fill it in.")
    if not cfg.service_endpoint:
        raise SystemExit("SERVICE_ENDPOINT is not set. Copy .env.example to .env and fill it in.")

    Path(cfg.output_dir).mkdir(exist_ok=True)

    log.info("Reading %s", cfg.input_excel)
    df = pd.read_excel(cfg.input_excel, sheet_name=cfg.sheet_name or 0)
    for col in (cfg.col_storage, cfg.col_container, cfg.col_fullpath):
        if col not in df.columns:
            raise SystemExit(
                f"Column '{col}' not found. Available columns: {list(df.columns)}. "
                f"Adjust the COL_* values in .env."
            )

    if args.limit:
        df = df.head(args.limit)

    results: list[dict] = []
    for i, row in df.iterrows():
        storage = str(row[cfg.col_storage]).strip()
        container = str(row[cfg.col_container]).strip()
        full_path = str(row[cfg.col_fullpath]).strip()
        blob_name = extract_blob_name(full_path, storage, container)

        rec: dict = {
            "row": int(i) + 1,
            "storage": storage,
            "container": container,
            "blob_name": blob_name,
            "status": "ok",
            "text_chars": 0,
            "error": "",
        }
        log.info("[%d] %s / %s / %s", rec["row"], storage, container, blob_name)

        try:
            data = download_blob(storage, container, blob_name)
            text = run_tika(cfg, data, blob_name)
            rec["text_chars"] = len(text)

            response = call_service(cfg, text, rec)
            rec["service_response"] = response

            # Persist the full response so it can be shared.
            out_json = Path(cfg.output_dir) / f"row{rec['row']:04d}.json"
            out_json.write_text(
                json.dumps({"meta": rec, "response": response}, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
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
    log.info("Done. Summary -> %s | Full responses -> %s/", cfg.output_excel, cfg.output_dir)


if __name__ == "__main__":
    main()
