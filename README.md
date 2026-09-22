[README-ITBoost-To-ITGlue.md](https://github.com/user-attachments/files/32540778/README-ITBoost-To-ITGlue.md)
# ITBoost to IT Glue Migration Script

**File:** `Prepare-ITBoostForITGlue.ps1`

**Platform:** Windows PowerShell 5.1+

**Purpose:** Migrates document exports from ITBoost to IT Glue by fixing the UID-named HTML files, patching document titles, and either uploading directly via the IT Glue API or producing a ZIP for manual import.

---

## Background

ITBoost exports documents in a nested folder structure where each document folder has a meaningful name (e.g. `DOC-1263-How to Contact Help Desk.docx`) but the HTML file inside is named with a random UID (e.g. `7d7751db-7bac-4aca-a70d-197214dc869a.html`). When imported into IT Glue as-is, documents either have no title or inherit the UID as their name.

This script:
1. Scans the export ZIP at any folder depth to find all document folders
2. Renames each HTML file to match its parent folder name
3. Patches the `<title>` tag inside each HTML file to match
4. Delivers the result either via the IT Glue API (recommended) or as a re-zipped file for manual import

---

## Requirements

- Windows only (ZIP re-packing uses `Shell.Application` COM object)
- PowerShell 5.1 or later
- IT Glue API key (for API upload mode only)
- The ITBoost export ZIP file

---

## Usage

### Standard run

```powershell
.\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\itboost-export.zip"
```

### With output folder pre-specified (ZIP mode only)

```powershell
.\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\itboost-export.zip" -OutputFolder "C:\Downloads\itglue-output"
```

### Diagnostic mode (no files written, inspect structure only)

```powershell
.\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\itboost-export.zip" -DiagnosticOnly
```

> Run diagnostic first on any new export to confirm the script is finding the correct folders and HTML files before processing.

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-ZipPath` | Yes | - | Full path to the ITBoost export ZIP file |
| `-OutputFolder` | No | Prompted if needed | Output folder for ZIP files (ZIP mode only) |
| `-DiagnosticOnly` | No | Off | Scan and report structure without writing anything |
| `-WorkDir` | No | `%TEMP%\itboost_prep` | Temporary working directory used during processing |

---

## Interactive Prompts

The script is fully interactive after the ZIP is extracted. Below is a summary of every prompt in order.

---

### All runs

#### Prompt 1 - Delivery method

```
How would you like to deliver the documents?
  [A] Upload directly to IT Glue via API
  [Z] Create ZIP file(s) for manual import

Enter A or Z:
```

Enter `A` to upload directly to IT Glue via the REST API.
Enter `Z` to produce one or more ZIP files for manual import through the IT Glue UI.

---

### API upload path (A)

#### Prompt 2 - Region

```
Select your IT Glue region:
  [1] North America  (api.itglue.com)
  [2] Europe         (api.eu.itglue.com)
  [3] Australia      (api.au.itglue.com)

Enter 1, 2 or 3:
```

Select the region that matches your IT Glue instance. This sets the API base URL for all subsequent calls.

#### Prompt 3 - API key

```
Enter your IT Glue API key: ********
```

Your API key is entered as a masked secure string and is not echoed to the console or written to disk.

To generate an API key: IT Glue > Account > Settings > API Keys > New API Key.

#### Prompt 4 - Organisation mapping (repeated per category)

After fetching all organisations from IT Glue, the script prompts you to map each top-level category from the ITBoost export to an IT Glue organisation.

```
--------------------------------------------
Category : global  (47 documents)
  Search org (or Enter to skip):
```

- Type part of the organisation name (case-insensitive search)
- If one match is found it is selected automatically
- If multiple matches are found you are shown a numbered list to pick from
- Press Enter with no input to skip the entire category (those documents will not be uploaded)

If there are more than 10 matches, refine your search term.

#### Prompt 5 - Upload confirmation

```
============================================
 UPLOAD PLAN
============================================
  global
    -> Acme Corp (ID: 12345)  [47 docs]
  Client Knowledgebase
    -> Acme Holdings (ID: 67890)  [12 docs]

  Skipped (no org assigned):
    - _Public facing T&Cs

Proceed with upload? (Y/N):
```

Review the full plan before anything is sent to IT Glue. Enter `Y` to proceed or `N` to abort with no changes made.

---

### ZIP output path (Z)

#### Prompt 2 - Output folder (if not passed as parameter)

```
Enter output folder path for ZIP file(s):
```

Only shown if `-OutputFolder` was not provided on the command line.

#### Prompt 3 - ZIP structure

```
How would you like to structure the ZIPs?
  [M] Merge all into one ZIP
  [S] Separate ZIP per top-level category

Enter M or S:
```

`M` produces a single `itglue-import-all.zip` containing all documents.
`S` produces one ZIP per top-level category, named `itglue-import-<category>.zip`. Useful when importing into different IT Glue organisations.

---

## Output files

### API mode

No output files are produced locally. Any documents that fail to upload are logged to:
```
<same folder as the input ZIP>\itglue-upload-failures.txt
```

### ZIP mode

One or more ZIP files are written to the `-OutputFolder`:

| Mode | Output |
|---|---|
| Merge (M) | `itglue-import-all.zip` |
| Separate (S) | `itglue-import-<category name>.zip` per category |

---

## Importing ZIPs into IT Glue (ZIP mode only)

1. Go to **Organizations** > select the target org > **Documents**
2. Click **Import** > **Documents**
3. Set **Document Importer Type** to **HTML**
4. Upload the ZIP file
5. Enable **Allow partial imports**
6. Click **Continue** - IT Glue will email you when the import is complete

> IT Glue has a **300 MB limit per ZIP**. If your ZIP exceeds this, split the export into smaller batches before running the script.

> Documents imported via ZIP land **flat** at the top level of the organisation's document list. You will need to move them into folders manually after import.

---

## Notes

- The script searches for document folders at **any depth** inside the ZIP - it is not limited to a fixed number of folder levels
- HTML files are identified by extension (case-insensitive), so both `.html` and `.HTML` are handled
- Folder names ending in `.docx` or `.doc` have that extension stripped when used as the document title
- Characters invalid in filenames (`\ / : * ? " < > |`) are replaced with a hyphen in the output file name but the original name is preserved in the `<title>` tag and API document name
- The script uses `Shell.Application` (not `Compress-Archive`) to produce ZIP files, which avoids compatibility issues with the IT Glue importer
- API uploads are rate-limited to approximately 6-7 documents per second (150ms delay) to stay within IT Glue API limits
- The temporary working directory (`%TEMP%\itboost_prep`) is automatically deleted after the script completes
