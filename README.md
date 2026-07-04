# Batch Text Extraction + Service Caller

Small Python tool that, for each row in an Excel file:

1. Reads **storage account**, **container**, and **full blob path**.
2. Downloads the blob from **Azure Storage**.
3. Sends the file to a **Tika server** (`PUT /tika`) to extract text.
4. **POSTs** the extracted text as JSON to your downstream service.
5. Saves a summary Excel + one JSON file per row under `responses/`.

## Authentication (important)

On the VM the tool authenticates with a **user-assigned Managed Identity** — no keys, no `az login`.

| Setting | Value |
|---|---|
| MI name | **Abyss-04** |
| MI client id | `9e11a244-d9b5-44a3-8234-07d2141b3f69` (set as `MANAGED_IDENTITY_CLIENT_ID`) |

| Service | Auth used | Role the MI needs |
|---|---|---|
| Azure Storage | `ManagedIdentityCredential` (Abyss-04) | **Storage Blob Data Reader** on the account/container (data-plane) |
| Tika server | Plain HTTP (no auth by default) | — |

> **Prerequisites on the VM (one-time):**
> 1. The managed identity **Abyss-04** must be **assigned to the VM** (VM → Identity → User assigned).
> 2. **Abyss-04** must have the **Storage Blob Data Reader** role on the storage account/container.
>    Management-plane "Reader"/"Owner" does **not** grant blob *data* access.
>
> ⚠️ The value `9e11a244-...` was given as "MI Id". The code uses it as the **client id**.
> If auth fails with an identity-not-found error, it may be the **object (principal) id** instead —
> in that case clear `MANAGED_IDENTITY_CLIENT_ID` and set `MANAGED_IDENTITY_OBJECT_ID` to that value in `.env`.

### Local testing (optional)

To run on your own machine instead of the VM, set `USE_MANAGED_IDENTITY=false` in `.env` and
`az login` — it will use your signed-in identity (which then needs the Storage Blob Data Reader role).

## Setup on the VM

After cloning this repo onto the VM:

### Linux VM (Ubuntu/Debian)

```bash
cd tools

# 1. create & activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 2. install dependencies
pip install -r requirements.txt

# 3. configure
cp .env.example .env
nano .env      # set SERVICE_ENDPOINT and the COL_* names; MI details are pre-filled

# 4. run
python main.py --limit 3      # test on 3 rows first
python main.py                # then the full run
```

### Windows VM (PowerShell)

```powershell
cd tools

# 1. create & activate a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. install dependencies
pip install -r requirements.txt

# 3. configure
Copy-Item .env.example .env
notepad .env   # set SERVICE_ENDPOINT and the COL_* names; MI details are pre-filled

# 4. run
python main.py --limit 3      # test on 3 rows first
python main.py                # then the full run
```

> No `az login` is needed on the VM — the Managed Identity **Abyss-04** is used automatically.
> The log line `Auth: user-assigned Managed Identity (client_id=9e11a244-...)` confirms it's active.

### Verify the identity works (optional)

```bash
# lists blobs using the VM's managed identity
az storage blob list --account-name <acct> --container-name <container> \
  --auth-mode login --login-experience-v2 false 2>/dev/null || \
az login --identity --username 9e11a244-d9b5-44a3-8234-07d2141b3f69
```

## Expected Excel format

Three columns (names are configurable in `.env` via `COL_*`):

| storage_name | container | full_path |
|---|---|---|
| mystorageacct | invoices | mystorageacct/invoices/2026/07/file1.pdf |
| mystorageacct | invoices | https://mystorageacct.blob.core.windows.net/invoices/2026/07/file2.pdf |

The tool automatically derives the blob name (everything after the container), so both a
plain path and a full `https://...` URL work.

## Run options

```bash
python main.py                 # process everything in input.xlsx
python main.py --limit 3       # test on the first 3 rows first
python main.py --input other.xlsx --output result.xlsx
```

## Output

- **output.xlsx** — one row per file: `status`, `text_chars`, `error`, and the service response (as JSON text).
- **responses/rowNNNN.json** — the full response for each file, ready to share.

## Customize the service payload

The downstream call is in [`call_service`](main.py). Edit the `payload` dict to match what
your endpoint expects (it currently sends `{ "file": <blob name>, "content": <extracted text> }`).
