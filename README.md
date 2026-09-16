Run it like this:

powershell
.\Upload-ITBoostToITGlue.ps1 `
  -ZipPath "C:\Downloads\6aa32c1eec08039e4bd008f6_1789078560596.zip" `
  -ApiKey "your-api-key-here"

The interactive flow looks like this:

Fetching organisations from IT Glue...
Found 47 organisations.

══════════════════════════════════════════
 CATEGORY → ORGANISATION MAPPING
══════════════════════════════════════════

──────────────────────────────────────────
Category : global  (47 documents)
  Search org (or Enter to skip): acme

  Multiple matches:
    [1] Acme Corp
    [2] Acme Holdings
  Enter number (or 0 to search again): 1
  Selected: Acme Corp (ID: 12345)

──────────────────────────────────────────
Category : Client Knowledgebase  (12 documents)
  Search org (or Enter to skip):        <- Enter to skip
  Skipping this category.

Then shows you the full upload plan and asks for a Y/N confirmation before touching anything in IT Glue.

A few things to be aware of:

The script uploads at ~6-7 docs/second (150ms delay) to stay within IT Glue's API rate limits
Any failures are logged to a itglue-upload-failures.txt file next to your ZIP so you can retry them
Documents will land in the org but not in any folder — IT Glue doesn't support setting a folder via the documents API at creation time, so you'd still need to organise them manually after upload
