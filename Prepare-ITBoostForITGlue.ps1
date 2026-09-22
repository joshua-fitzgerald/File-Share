# ITBoost-To-ITGlue.ps1
# Extracts an ITBoost export ZIP, renames HTML files to match their folder names,
# patches <title> tags, then either:
#   (A) Uploads directly to IT Glue via API, or
#   (B) Re-zips for manual import via the IT Glue document importer
#
# USAGE:
#   .\ITBoost-To-ITGlue.ps1 -ZipPath "C:\Downloads\export.zip"
#   .\ITBoost-To-ITGlue.ps1 -ZipPath "C:\Downloads\export.zip" -DiagnosticOnly
#
# Handles any folder depth - finds doc folders by presence of .html files

param(
    [Parameter(Mandatory)]
    [string]$ZipPath,

    [string]$OutputFolder = "",

    [switch]$DiagnosticOnly,

    [string]$WorkDir = "$env:TEMP\itboost_prep"
)

# ============================================================
# SECTION 1: EXTRACT AND SCAN
# ============================================================

if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
New-Item -ItemType Directory -Path $WorkDir | Out-Null
$extractPath = Join-Path $WorkDir "extracted"
New-Item -ItemType Directory -Path $extractPath | Out-Null

Write-Host ""
Write-Host "Extracting ZIP..." -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $extractPath)

# Find documents root
$docsRoot = Get-ChildItem -Path $extractPath -Directory -Recurse |
    Where-Object { $_.Name -eq "documents" } |
    Select-Object -First 1
$scanRoot = if ($docsRoot) { $docsRoot.FullName } else { $extractPath }

Write-Host "Scan root: $scanRoot" -ForegroundColor DarkGray
Write-Host "Scanning for document folders..." -ForegroundColor Cyan

# Find all folders that directly contain an .html file (any depth)
$allDocFolders = Get-ChildItem -Path $scanRoot -Directory -Recurse | Where-Object {
    $folder = $_
    (Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -ieq ".html" }).Count -gt 0
}

Write-Host "Found $($allDocFolders.Count) document folders." -ForegroundColor Green

# ============================================================
# SECTION 2: DIAGNOSTIC MODE
# ============================================================

if ($DiagnosticOnly) {
    Write-Host ""
    Write-Host "=== DIAGNOSTIC MODE - no files will be written ===" -ForegroundColor Yellow
    foreach ($doc in $allDocFolders) {
        $relativePath = $doc.FullName.Replace($scanRoot, "").TrimStart("\").TrimStart("/")
        Write-Host ""
        Write-Host "  DOC FOLDER : $relativePath" -ForegroundColor Yellow
        foreach ($f in (Get-ChildItem -Path $doc.FullName -File)) {
            Write-Host "    FILE: $($f.Name)  [ext: $($f.Extension)]" -ForegroundColor White
        }
    }
    Write-Host ""
    Write-Host "Diagnostic complete. Re-run without -DiagnosticOnly to process." -ForegroundColor Green
    Remove-Item $WorkDir -Recurse -Force
    exit
}

# ============================================================
# SECTION 3: GROUP BY TOP-LEVEL CATEGORY
# ============================================================

$categoryGroups = $allDocFolders | Group-Object {
    $relative = $_.FullName.Substring($scanRoot.Length).TrimStart("\").TrimStart("/")
    $relative.Split([char]92)[0].Split([char]47)[0]
}

Write-Host ""
Write-Host "Found $($categoryGroups.Count) top-level categories:" -ForegroundColor Cyan
Write-Host ""
foreach ($g in $categoryGroups) {
    Write-Host "  $($g.Name)  ($($g.Count) documents)"
}

# ============================================================
# SECTION 4: CHOOSE DELIVERY METHOD
# ============================================================

Write-Host ""
Write-Host "How would you like to deliver the documents?" -ForegroundColor Yellow
Write-Host "  [A] Upload directly to IT Glue via API"
Write-Host "  [Z] Create ZIP file(s) for manual import"
Write-Host ""

do {
    $deliveryMode = (Read-Host "Enter A or Z").Trim().ToUpper()
} while ($deliveryMode -ne "A" -and $deliveryMode -ne "Z")

# ============================================================
# SECTION 5: SHARED - BUILD DOC LIST (rename + patch title)
# ============================================================

# Returns array of objects: @{ DocName, SafeName, HtmlContent, Folder }
function Build-DocList {
    param([object[]]$Folders)

    $docs    = [System.Collections.Generic.List[hashtable]]::new()
    $total   = $Folders.Count
    $count   = 0
    $skipped = 0

    foreach ($folder in $Folders) {
        $count++
        $rawName = $folder.Name
        $docName = $rawName -replace "\.docx$", "" -replace "\.doc$", ""

        $allFiles  = Get-ChildItem -Path $folder.FullName -File
        $htmlFiles = $allFiles | Where-Object { $_.Extension -ieq ".html" }

        if ($htmlFiles.Count -eq 0) {
            Write-Warning "[$count/$total] No HTML in: $rawName"
            $skipped++
            continue
        }

        $htmlFile = $htmlFiles[0]
        $safeName = $docName -replace '[\\/:*?"<>|]', "-"

        $content = Get-Content -Path $htmlFile.FullName -Raw -Encoding UTF8
        if ($content -match "(?i)<title>") {
            $content = $content -replace "(?i)<title>.*?</title>", "<title>$docName</title>"
        } elseif ($content -match "(?i)<head>") {
            $content = $content -replace "(?i)<head>", "<head>`n<title>$docName</title>"
        } else {
            $content = "<html><head><title>$docName</title></head>`n" + $content
        }

        $docs.Add(@{
            DocName     = $docName
            SafeName    = $safeName
            HtmlContent = $content
            Folder      = $folder
            AllFiles    = $allFiles
        })

        Write-Host "  [$count/$total] Ready: $docName" -ForegroundColor Green
    }

    Write-Host "  Skipped: $skipped" -ForegroundColor DarkGray
    return $docs
}

# ============================================================
# SECTION 6A: API UPLOAD PATH
# ============================================================

if ($deliveryMode -eq "A") {

    # Prompt for region
    Write-Host ""
    Write-Host "Select your IT Glue region:" -ForegroundColor Yellow
    Write-Host "  [1] North America  (api.itglue.com)"
    Write-Host "  [2] Europe         (api.eu.itglue.com)"
    Write-Host "  [3] Australia      (api.au.itglue.com)"
    Write-Host ""

    do {
        $regionPick = (Read-Host "Enter 1, 2 or 3").Trim()
    } while ($regionPick -notin @("1","2","3"))

    $ApiBase = switch ($regionPick) {
        "1" { "https://api.itglue.com" }
        "2" { "https://api.eu.itglue.com" }
        "3" { "https://api.au.itglue.com" }
    }
    Write-Host "Using: $ApiBase" -ForegroundColor DarkGray

    # Prompt for API key securely
    Write-Host ""
    $apiKeySecure = Read-Host "Enter your IT Glue API key" -AsSecureString
    $apiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKeySecure)
    )

    $headers = @{
        "x-api-key"    = $apiKey
        "Content-Type" = "application/vnd.api+json"
    }

    # Fetch all orgs
    Write-Host ""
    Write-Host "Fetching organisations from IT Glue..." -ForegroundColor Cyan
    $allOrgs = [System.Collections.Generic.List[object]]::new()
    $page = 1
    do {
        $resp = Invoke-RestMethod -Uri "$ApiBase/organizations?page[size]=100&page[number]=$page" -Headers $headers -Method Get
        foreach ($org in $resp.data) { $allOrgs.Add($org) }
        $page++
    } while ($resp.data.Count -eq 100)
    Write-Host "Found $($allOrgs.Count) organisations." -ForegroundColor Green

    # Map each category to an org
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " CATEGORY TO ORGANISATION MAPPING" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Type part of an org name to search."
    Write-Host " Press Enter with no input to skip a category."
    Write-Host ""

    $categoryMap = @{}

    foreach ($group in $categoryGroups) {
        $docCount = $group.Count
        Write-Host "--------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Category : $($group.Name)  ($docCount documents)" -ForegroundColor Yellow

        $selectedOrg = $null
        while ($null -eq $selectedOrg) {
            $search = (Read-Host "  Search org (or Enter to skip)").Trim()
            if ($search -eq "") { Write-Host "  Skipping." -ForegroundColor DarkGray; break }

            $matches = $allOrgs | Where-Object { $_.attributes.name -like "*$search*" }

            if ($matches.Count -eq 0) {
                Write-Host "  No match for '$search'. Try again." -ForegroundColor Red
                continue
            }
            if ($matches.Count -eq 1) {
                $selectedOrg = $matches[0]
                Write-Host "  Matched: $($selectedOrg.attributes.name) (ID: $($selectedOrg.id))" -ForegroundColor Green
                break
            }

            Write-Host "  Multiple matches:" -ForegroundColor Cyan
            $limit = [Math]::Min($matches.Count, 10)
            for ($i = 0; $i -lt $limit; $i++) {
                Write-Host "    [$($i+1)] $($matches[$i].attributes.name)"
            }
            if ($matches.Count -gt 10) { Write-Host "    ... $($matches.Count - 10) more. Refine search." -ForegroundColor DarkGray; continue }

            do {
                $pick = Read-Host "  Enter number (or 0 to search again)"
                $pickInt = 0
                [int]::TryParse($pick, [ref]$pickInt) | Out-Null
            } while ($pickInt -lt 0 -or $pickInt -gt $matches.Count)

            if ($pickInt -eq 0) { continue }
            $selectedOrg = $matches[$pickInt - 1]
            Write-Host "  Selected: $($selectedOrg.attributes.name) (ID: $($selectedOrg.id))" -ForegroundColor Green
        }

        if ($null -ne $selectedOrg) {
            $categoryMap[$group.Name] = @{ OrgId = $selectedOrg.id; OrgName = $selectedOrg.attributes.name }
        }
    }

    # Confirm plan
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " UPLOAD PLAN" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    foreach ($cat in $categoryMap.Keys) {
        $cnt = ($categoryGroups | Where-Object { $_.Name -eq $cat }).Count
        Write-Host "  $cat" -ForegroundColor Yellow
        Write-Host "    -> $($categoryMap[$cat].OrgName) (ID: $($categoryMap[$cat].OrgId))  [$cnt docs]"
    }
    $skippedCats = $categoryGroups | Where-Object { -not $categoryMap.ContainsKey($_.Name) }
    if ($skippedCats) {
        Write-Host ""
        Write-Host "  Skipped (no org assigned):" -ForegroundColor DarkGray
        foreach ($s in $skippedCats) { Write-Host "    - $($s.Name)" -ForegroundColor DarkGray }
    }

    Write-Host ""
    $confirm = (Read-Host "Proceed with upload? (Y/N)").Trim().ToUpper()
    if ($confirm -ne "Y") {
        Write-Host "Aborted." -ForegroundColor Red
        Remove-Item $WorkDir -Recurse -Force
        exit
    }

    # Upload
    $totalOk     = 0
    $totalFailed = 0
    $failedDocs  = [System.Collections.Generic.List[string]]::new()

    foreach ($group in $categoryGroups) {
        if (-not $categoryMap.ContainsKey($group.Name)) { continue }
        $orgId   = $categoryMap[$group.Name].OrgId
        $orgName = $categoryMap[$group.Name].OrgName

        Write-Host ""
        Write-Host "[$($group.Name)] -> $orgName" -ForegroundColor Cyan

        $docs = Build-DocList -Folders $group.Group

        $docTotal = $docs.Count
        $docCount = 0

        foreach ($doc in $docs) {
            $docCount++
            $body = @{
                data = @{
                    type       = "documents"
                    attributes = @{
                        "organization-id" = [int]$orgId
                        name              = $doc.DocName
                        content           = $doc.HtmlContent
                        "public"          = $false
                    }
                }
            } | ConvertTo-Json -Depth 10

            try {
                Invoke-RestMethod -Uri "$ApiBase/documents" -Headers $headers -Method Post -Body $body | Out-Null
                Write-Host "  [$docCount/$docTotal] OK : $($doc.DocName)" -ForegroundColor Green
                $totalOk++
            } catch {
                $detail = $_.ErrorDetails.Message
                Write-Host "  [$docCount/$docTotal] FAIL: $($doc.DocName) -- $detail" -ForegroundColor Red
                $failedDocs.Add("[$($group.Name)] $($doc.DocName)")
                $totalFailed++
            }

            Start-Sleep -Milliseconds 150
        }
    }

    # Summary
    Write-Host ""
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "Upload complete!" -ForegroundColor Green
    Write-Host "  Uploaded : $totalOk"
    Write-Host "  Failed   : $totalFailed"

    if ($failedDocs.Count -gt 0) {
        $logPath = Join-Path (Split-Path $ZipPath) "itglue-upload-failures.txt"
        $failedDocs | Out-File -FilePath $logPath -Encoding UTF8
        Write-Host ""
        Write-Host "Failed docs logged to: $logPath" -ForegroundColor Yellow
    }
}

# ============================================================
# SECTION 6B: ZIP OUTPUT PATH
# ============================================================

if ($deliveryMode -eq "Z") {

    if ($OutputFolder -eq "") {
        $OutputFolder = Read-Host "Enter output folder path for ZIP file(s)"
    }
    $OutputFolder = $OutputFolder.TrimEnd("\").TrimEnd("/")
    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

    Write-Host ""
    Write-Host "How would you like to structure the ZIPs?" -ForegroundColor Yellow
    Write-Host "  [M] Merge all into one ZIP"
    Write-Host "  [S] Separate ZIP per top-level category"
    Write-Host ""

    do {
        $zipMode = (Read-Host "Enter M or S").Trim().ToUpper()
    } while ($zipMode -ne "M" -and $zipMode -ne "S")

    function Create-ZipFromFolder {
        param([string]$SourceFolder, [string]$ZipOutputPath)

        if (Test-Path $ZipOutputPath) { Remove-Item $ZipOutputPath -Force }
        $zipBytes = [byte[]](80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
        [System.IO.File]::WriteAllBytes($ZipOutputPath, $zipBytes)

        $shell    = New-Object -ComObject Shell.Application
        $zipShell = $shell.NameSpace($ZipOutputPath)
        $srcItems = $shell.NameSpace($SourceFolder).Items()

        foreach ($item in $srcItems) {
            $zipShell.CopyHere($item, 0x14)
            $timeout = 120; $elapsed = 0
            do { Start-Sleep -Milliseconds 500; $elapsed += 0.5 } while ($elapsed -lt $timeout)
        }
        Start-Sleep -Seconds 2
    }

    function Write-DocsToFolder {
        param([object[]]$Docs, [string]$StagingPath)
        New-Item -ItemType Directory -Path $StagingPath -Force | Out-Null
        foreach ($doc in $Docs) {
            $destFolder = Join-Path $StagingPath $doc.SafeName
            New-Item -ItemType Directory -Path $destFolder | Out-Null
            $doc.AllFiles | Where-Object { $_.Extension -ine ".html" } |
                ForEach-Object { Copy-Item -Path $_.FullName -Destination $destFolder }
            [System.IO.File]::WriteAllText(
                (Join-Path $destFolder "$($doc.SafeName).html"),
                $doc.HtmlContent,
                [System.Text.Encoding]::UTF8
            )
        }
    }

    $totalProcessed = 0
    $totalSkipped   = 0

    if ($zipMode -eq "M") {
        Write-Host ""
        Write-Host "Building document list..." -ForegroundColor Cyan
        $docs        = Build-DocList -Folders $allDocFolders
        $stagingPath = Join-Path $WorkDir "merged"
        Write-DocsToFolder -Docs $docs -StagingPath $stagingPath
        $outputZip = Join-Path $OutputFolder "itglue-import-all.zip"
        Write-Host ""
        Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
        Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip
        $totalProcessed = $docs.Count
    } else {
        foreach ($group in $categoryGroups) {
            Write-Host ""
            Write-Host "Processing category: $($group.Name)" -ForegroundColor Cyan
            $docs        = Build-DocList -Folders $group.Group
            $safeCatName = $group.Name -replace '[\\/:*?"<>|]', "-"
            $stagingPath = Join-Path $WorkDir "cat_$safeCatName"
            Write-DocsToFolder -Docs $docs -StagingPath $stagingPath
            $outputZip = Join-Path $OutputFolder "itglue-import-$safeCatName.zip"
            Write-Host ""
            Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
            Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip
            $totalProcessed += $docs.Count
        }
    }

    Write-Host ""
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "Complete!" -ForegroundColor Green
    Write-Host "  Processed : $totalProcessed documents"
    Write-Host "  Output    : $OutputFolder"
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. IT Glue > Organizations > [Org] > Documents > Import > Documents"
    Write-Host "  2. Set type to HTML, upload ZIP, enable Allow partial imports"
    Write-Host "  3. IT Glue has a 300 MB limit per ZIP"
    Write-Host "  4. Docs land flat after import - move into folders manually"
}

Remove-Item $WorkDir -Recurse -Force
