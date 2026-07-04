"""
Standalone diagnostic: download ONE blob via a RAW REST call.

This deliberately does NOT use the azure-storage-blob SDK. It gets a token straight
from the VM's IMDS endpoint (the same way your earlier curl test did) and then does a
plain HTTP GET on the blob. That isolates the HTTP 500 "InternalError" you're seeing:

  * If this script ALSO returns 500 InternalError  -> the storage account / network is
    the problem (NOT the tool). Hand that to whoever owns 'devssecontentstore'.
  * If this script returns 200 and downloads bytes -> it's something in how the SDK talks
    to storage, and the main tool can be adjusted.

Needs nothing beyond what's already installed (requests, python-dotenv).

Usage (PowerShell, all one line). Defaults to the blob you want to check:
  python diag.py
  python diag.py qassrcontentstore event-data "gtr/.../Filled Organizer.pdf"
"""

from __future__ import annotations

import os
import re
import sys
from urllib.parse import quote

import requests
from dotenv import load_dotenv

load_dotenv()

# Defaults to the blob you want to check so you can just: python diag.py
# (storage account taken from the Source column -- the leading "o" in the FullPath is ignored)
DEFAULT_ACCOUNT = "qassrcontentstore"
DEFAULT_CONTAINER = "event-data"
DEFAULT_BLOB = "gtr/00f6141d-18ef-42ee-b270-482a5fa18b8b/filledorganizer/Filled Organizer.pdf"

IMDS_URL = "http://169.254.169.254/metadata/identity/oauth2/token"
STORAGE_RESOURCE = "https://storage.azure.com/"


def get_token(client_id: str | None) -> str:
    """Get a storage token directly from IMDS (bypasses azure-identity too)."""
    params = {"api-version": "2018-02-01", "resource": STORAGE_RESOURCE}
    if client_id:
        params["client_id"] = client_id
    resp = requests.get(IMDS_URL, headers={"Metadata": "true"}, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _names(xml_text: str, tag: str = "Name") -> list[str]:
    return re.findall(rf"<{tag}>(.*?)</{tag}>", xml_text or "", flags=re.DOTALL)


def _auth_headers(token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}", "x-ms-version": "2021-08-06"}


def list_containers(account: str, token: str) -> None:
    """List the containers that actually exist on the account."""
    url = f"https://{account}.blob.core.windows.net/?comp=list"
    resp = requests.get(url, headers=_auth_headers(token), timeout=60)
    print(f"\n3) List containers on '{account}': HTTP {resp.status_code}")
    if resp.status_code == 200:
        names = _names(resp.text)
        if names:
            print("   Containers that exist on this account:")
            for n in names:
                print(f"     - {n}")
        else:
            print("   (account has no containers, or none visible to this identity)")
    else:
        print("   " + (resp.text or "(empty)")[:600].replace("\n", "\n   "))
        if resp.status_code == 403:
            print("   (this identity can read blobs but not LIST containers -- needs account-scope read)")


def list_blobs(account: str, container: str, token: str, prefix: str = "", limit: int = 25) -> None:
    """List up to `limit` blob names in a container (optionally under a prefix)."""
    url = (
        f"https://{account}.blob.core.windows.net/{container}"
        f"?restype=container&comp=list&maxresults={limit}"
    )
    if prefix:
        url += f"&prefix={quote(prefix, safe='/')}"
    resp = requests.get(url, headers=_auth_headers(token), timeout=60)
    print(f"\n4) List blobs in '{container}' (prefix='{prefix}'): HTTP {resp.status_code}")
    if resp.status_code == 200:
        names = _names(resp.text)
        if names:
            print(f"   First {len(names)} blob name(s):")
            for n in names:
                print(f"     - {n}")
        else:
            print("   (no blobs match this prefix)")
    else:
        print("   " + (resp.text or "(empty)")[:600].replace("\n", "\n   "))


def main() -> None:
    account = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_ACCOUNT
    container = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_CONTAINER
    blob = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_BLOB

    client_id = os.getenv("MANAGED_IDENTITY_CLIENT_ID", "").strip() or None
    print(f"1) Getting storage token from IMDS (client_id={client_id}) ...")
    token = get_token(client_id)
    print("   Token OK.\n")

    blob_enc = quote(blob, safe="/")
    url = f"https://{account}.blob.core.windows.net/{container}/{blob_enc}"
    print(f"2) Raw REST GET (no azure-storage-blob SDK):\n   {url}\n")

    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}", "x-ms-version": "2021-08-06"},
        timeout=60,
    )

    print(f"   HTTP {resp.status_code}")
    x_req_id = resp.headers.get("x-ms-request-id")
    if x_req_id:
        print(f"   x-ms-request-id: {x_req_id}")

    if resp.status_code == 200:
        print(f"\nSUCCESS -- downloaded {len(resp.content)} bytes.")
        print("=> Raw REST works, so the 500 is SDK-specific. Tell the assistant to adjust the tool.")
    else:
        print("\n   Response body:")
        print("   " + (resp.text or "(empty)")[:1200].replace("\n", "\n   "))
        if resp.status_code == 500 and "InternalError" in (resp.text or ""):
            print(
                "\n=> Raw REST ALSO returns 500 InternalError. This confirms the problem is the\n"
                "   STORAGE ACCOUNT / network, NOT this tool. Share the x-ms-request-id above with\n"
                "   whoever owns 'devssecontentstore' (or Azure support) to investigate."
            )
        elif resp.status_code == 403:
            print(
                "\n=> 403: token is fine but this identity lacks DATA access. It needs the\n"
                "   'Storage Blob Data Reader' role on the account/container."
            )
        elif resp.status_code == 404:
            body = resp.text or ""
            if "ContainerNotFound" in body:
                print(
                    f"\n=> 404 ContainerNotFound: account '{account}' is reachable and auth works,\n"
                    f"   but it has no container named '{container}'. Listing the REAL containers..."
                )
                list_containers(account, token)
            else:
                print(
                    "\n=> 404 BlobNotFound: the container exists but the blob path is wrong.\n"
                    "   Listing blobs under the leading path to help you find the right name..."
                )
                prefix = blob.split("/")[0] if "/" in blob else ""
                list_blobs(account, container, token, prefix=prefix)


if __name__ == "__main__":
    main()
