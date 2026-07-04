# Batch Text Extraction + Service Caller

Small Python tool that, for each row in an Excel file:

1. Reads **storage account**, **container**, and **full blob path**.
2. Downloads the blob from **Azure Storage**.
3. Sends the file to a **Tika server** (`PUT /tika`, `Accept: text/plain`) to extract text.
4. **POSTs** the extracted text to the **GuardPII `recognize`** endpoint for PII detection.
5. Saves each response as `response/<blob-name>.json`, plus a summary Excel.

## Authentication (important)

On the VM the tool authenticates with a **Managed Identity** — no keys, no `az login`.

| Setting | Value |
|---|---|
| VM name | **Abyss-04** (its **system-assigned** identity is used) |
| Client id | `7458e3e2-adbf-4703-9fde-f477693018f7` (set as `MANAGED_IDENTITY_CLIENT_ID`) |
| Object id | `9e11a244-d9b5-44a3-8234-07d2141b3f69` (the "MI Id" originally provided — used for RBAC) |

| Service | Auth used | Role the MI needs |
|---|---|---|
| Azure Storage | `ManagedIdentityCredential` (Abyss-04) | **Storage Blob Data Reader** on the account/container (data-plane) |
| Tika server | Plain HTTP (no auth by default) | — |
| GuardPII `recognize` | `x-functions-key` header (secret in `.env`) | — |

> **Prerequisites on the VM (one-time):**
> 1. The VM **Abyss-04** has a **system-assigned** managed identity enabled (it does).
> 2. That identity must have the **Storage Blob Data Reader** role on the storage account/container.
>    Management-plane "Reader"/"Owner" does **not** grant blob *data* access.
>
> ℹ️ **About the two GUIDs:** a managed identity has a **Client ID** (`7458e3e2-...`, used for auth)
> and an **Object/Principal ID** (`9e11a244-...`, used for role assignments). The `9e11a244-...`
> value originally provided as "MI Id" is the **object id**, which is why using it as a client id
> failed with `Identity not found`. The tool is now configured with the correct client id.
>
> If auth ever fails, alternatives in `.env`: clear `MANAGED_IDENTITY_CLIENT_ID` and set
> `MANAGED_IDENTITY_OBJECT_ID=9e11a244-...`, **or** leave all MI ids empty to use the VM's
> default (system-assigned) identity.

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
nano .env      # set SERVICE_FUNCTIONS_KEY (secret); columns, endpoints & MI are pre-filled

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
notepad .env   # set SERVICE_FUNCTIONS_KEY (secret); columns, endpoints & MI are pre-filled

# 4. run
python main.py --limit 3      # test on 3 rows first
python main.py                # then the full run
```

> No `az login` is needed on the VM — the managed identity of **Abyss-04** is used automatically.
> The log line `Auth: user-assigned Managed Identity (client_id=7458e3e2-...)` confirms it's active,
> followed by `Auth OK.` once a token is obtained.

### Verify the identity works (optional)

```bash
# lists blobs using the VM's managed identity
az storage blob list --account-name <acct> --container-name <container> \
  --auth-mode login 2>/dev/null || \
az login --identity --username 7458e3e2-adbf-4703-9fde-f477693018f7
```

## Input file (CSV or Excel)

`INPUT_FILE` in `.env` points to a `.csv` or `.xlsx`. A ready-to-use **input_sample.csv**
(10 rows) ships with the repo. Expected columns (configurable via `COL_*`):

| Source | Container | FullPath |
|---|---|---|
| devssecontentstore | abc | /devssecontentstore/abc/Metadata.png |
| devssecontentstore | afd-test | /devssecontentstore/afd-test/GA_Taxdue - tpr.pdf |

(The real export also has a `SizeBytes` column, which the tool ignores.)

**Storage, container and blob name are all parsed from `FullPath`** by splitting on `/`
(1st segment = storage account, 2nd = container, the rest = blob path). The `Source`/`Container`
columns are used only as a fallback. A leading-slash path like `/Source/Container/folder/file.pdf`
and a full `https://...` URL both work.

**Only PDF and image blobs are processed** — `pdf, png, jpg, jpeg, tif, tiff, bmp, gif, webp,
heic, heif`. Everything else (`.msi`, `.zip`, …) is skipped and marked `skipped` in the summary.
Change the list with `ALLOWED_EXTENSIONS` in `.env` (comma-separated).

**Extensionless blobs are skipped** by default. If you want the tool to also try them (Tika
auto-detects the real type; unparseable files are then skipped), set `PROCESS_EXTENSIONLESS=true`.

## Run options

```bash
python main.py                 # process everything in input_sample.csv
python main.py --limit 3       # test on the first 3 rows first
python main.py --input other.csv --output result.xlsx
```

## Output

- **response/&lt;blob-name&gt;.json** — the recognize response for each file, named after the blob
  (folders in the blob path are flattened with `_`, e.g. `copy/Metadata.png` → `copy_Metadata.json`).
  This is the main output.
- **output.xlsx** — a summary: one row per file with `status`, `text_chars`, `output_json`, `error`.

## The downstream call (GuardPII recognize)

The call is in [`call_recognize`](main.py). For each file it sends:

```json
{ "texts": ["<extracted text>"], "detection_model": "llm" }
```

with headers `x-functions-key` (from `.env`) and a **fresh `x-operation-id` GUID per request**.

> ⚠️ Reusing an `x-operation-id` makes the service return **HTTP 500**
> (`'NoneType' object has no attribute 'set_supportive_context_word'`). The tool generates a
> new GUID for every call, so this is handled automatically.

The response is a JSON array of detected PII (e.g. `PERSON`, `FIRST_NAME`, `US_SSN`) and is saved
to `response/<blob-name>.json`:

```json
[{ "PERSON": { "Hardik": "Ashley" }, "US_SSN": { "123-00-3224": "375-96-0000" } }]
```

Change the model via `DETECTION_MODEL` in `.env`.
