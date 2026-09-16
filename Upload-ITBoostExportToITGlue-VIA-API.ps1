# Upload-ITBoostToITGlue.ps1
# Extracts ITBoost export ZIP, prompts for IT Glue org mapping per category,
# then uploads each document directly via the IT Glue API.
#
# USAGE:
#   .\Upload-ITBoostToITGlue.ps1 -ZipPath "C:\Downloads\itboost-export.zip" -ApiKey "your-api-key"

param(
    [Parameter(Mandatory)]
    [string]$ZipPath,

    [Parameter(Mandatory)]
    [string]$ApiKey,

    [string]$ApiBase = "https://api.au.itglue.com",

    [string]$WorkDir = "$env:TEMP\itboost_upload"
)

$ErrorActionPreference = "Stop"

# ── Headers ───────────────────────────────────────────────────────────────────
$headers = @{
    "x-api-key"    = $ApiKey
    "Content-Type" = "application/vnd.api+json"
}

# ── Helper: IT Glue API GET ────────────────────────────────────────────────────
function Invoke-ITGlueGet {
    param([string]$Endpoint)
    $uri = "$ApiBase$Endpoint"
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        return $response
    } catch {
        Write-Error "API GET failed for $uri : $_"
    }
}

# ── Helper: IT Glue API POST ───────────────────────────────────────────────────
function Invoke-ITGluePost {
    param([string]$Endpoint, [hashtable]$Body)
    $uri  = "$ApiBase$Endpoint"
    $json = $Body | ConvertTo-Json -Depth 10
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $json
        return $response
    } catch {
        $detail = $_.ErrorDetails.Message
        Write-Warning "API POST failed: $detail"
        return $null
    }
}

# ── 1. Extract ZIP ─────────────────────────────────────────────────────────────
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
New-Item -ItemType Directory -Path $WorkDir | Out-Null
$extractPath = Join-Path $WorkDir "extracted"
New-Item -ItemType Directory -Path $extractPath | Out-Null

Write-Host ""
Write-Host "Extracting ZIP..." -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $extractPath)

# ── 2. Find documents root ─────────────────────────────────────────────────────
$docsRoot = Get-ChildItem -Path $extractPath -Directory -Recurse |
    Where-Object { $_.Name -eq "documents" } |
    Select-Object -First 1
$scanRoot = if ($docsRoot) { $docsRoot.FullName } else { $extractPath }
$categoryFolders = Get-ChildItem -Path $scanRoot -Directory

# ── 3. Fetch all IT Glue orgs ──────────────────────────────────────────────────
Write-Host "Fetching organisations from IT Glue..." -ForegroundColor Cyan

$allOrgs = [System.Collections.Generic.List[object]]::new()
$page    = 1
do {
    $resp = Invoke-ITGlueGet -Endpoint "/organizations?page[size]=100&page[number]=$page"
    foreach ($org in $resp.data) { $allOrgs.Add($org) }
    $page++
} while ($resp.data.Count -eq 100)

Write-Host "Found $($allOrgs.Count) organisations." -ForegroundColor Green

# ── 4. Per-category org mapping ────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " CATEGORY → ORGANISATION MAPPING" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " For each category you can:"
Write-Host "   • Type part of an org name to search"
Write-Host "   • Press Enter with no input to skip the category"
Write-Host ""

$categoryMap = @{}  # category folder name -> org id

foreach ($category in $categoryFolders) {
    $docCount = (Get-ChildItem -Path $category.FullName -Directory).Count
    Write-Host "──────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "Category : $($category.Name)  ($docCount documents)" -ForegroundColor Yellow

    $selectedOrg = $null

    while ($null -eq $selectedOrg) {
        $search = (Read-Host "  Search org (or Enter to skip)").Trim()

        if ($search -eq "") {
            Write-Host "  Skipping this category." -ForegroundColor DarkGray
            break
        }

        $matches = $allOrgs | Where-Object { $_.attributes.name -like "*$search*" }

        if ($matches.Count -eq 0) {
            Write-Host "  No organisations matched '$search'. Try again." -ForegroundColor Red
            continue
        }

        if ($matches.Count -eq 1) {
            $selectedOrg = $matches[0]
            Write-Host "  Matched: $($selectedOrg.attributes.name) (ID: $($selectedOrg.id))" -ForegroundColor Green
            break
        }

        # Multiple matches — let user pick
        Write-Host "  Multiple matches:" -ForegroundColor Cyan
        for ($i = 0; $i -lt [Math]::Min($matches.Count, 10); $i++) {
            Write-Host "    [$($i+1)] $($matches[$i].attributes.name)"
        }
        if ($matches.Count -gt 10) {
            Write-Host "    ... and $($matches.Count - 10) more. Refine your search." -ForegroundColor DarkGray
            continue
        }

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
        $categoryMap[$category.Name] = @{
            OrgId   = $selectedOrg.id
            OrgName = $selectedOrg.attributes.name
        }
    }
}

# Confirm before uploading
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " UPLOAD PLAN" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
foreach ($cat in $categoryMap.Keys) {
    $docCount = (Get-ChildItem -Path (Join-Path $scanRoot $cat) -Directory).Count
    Write-Host "  $cat" -ForegroundColor Yellow
    Write-Host "    -> $($categoryMap[$cat].OrgName) (ID: $($categoryMap[$cat].OrgId))  [$docCount docs]"
}

$skippedCategories = $categoryFolders | Where-Object { -not $categoryMap.ContainsKey($_.Name) }
if ($skippedCategories) {
    Write-Host ""
    Write-Host "  Skipped categories (no org assigned):" -ForegroundColor DarkGray
    foreach ($s in $skippedCategories) { Write-Host "    - $($s.Name)" -ForegroundColor DarkGray }
}

Write-Host ""
$confirm = (Read-Host "Proceed with upload? (Y/N)").Trim().ToUpper()
if ($confirm -ne "Y") {
    Write-Host "Aborted." -ForegroundColor Red
    Remove-Item $WorkDir -Recurse -Force
    exit
}

# ── 5. Upload documents ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Starting upload..." -ForegroundColor Cyan

$totalOk      = 0
$totalSkipped = 0
$totalFailed  = 0
$failedDocs   = [System.Collections.Generic.List[string]]::new()

foreach ($category in $categoryFolders) {
    if (-not $categoryMap.ContainsKey($category.Name)) { continue }

    $orgId      = $categoryMap[$category.Name].OrgId
    $orgName    = $categoryMap[$category.Name].OrgName
    $docFolders = Get-ChildItem -Path $category.FullName -Directory
    $catTotal   = $docFolders.Count
    $catCount   = 0

    Write-Host ""
    Write-Host "[$($category.Name)] -> $orgName  ($catTotal documents)" -ForegroundColor Cyan

    foreach ($folder in $docFolders) {
        $catCount++
        $rawName = $folder.Name
        $docName = $rawName -replace '\.docx$', '' -replace '\.doc$', ''

        $htmlFiles = Get-ChildItem -Path $folder.FullName -Filter "*.html" -File
        if ($htmlFiles.Count -eq 0) {
            Write-Warning "  [$catCount/$catTotal] No HTML: $rawName — skipping"
            $totalSkipped++
            continue
        }

        $htmlContent = Get-Content -Path $htmlFiles[0].FullName -Raw -Encoding UTF8

        $body = @{
            data = @{
                type       = "documents"
                attributes = @{
                    "organization-id" = [int]$orgId
                    name              = $docName
                    content           = $htmlContent
                    "public"          = $false
                }
            }
        }

        $result = Invoke-ITGluePost -Endpoint "/documents" -Body $body

        if ($null -ne $result) {
            Write-Host "  [$catCount/$catTotal] OK : $docName" -ForegroundColor Green
            $totalOk++
        } else {
            Write-Host "  [$catCount/$catTotal] FAIL: $docName" -ForegroundColor Red
            $failedDocs.Add("[$($category.Name)] $docName")
            $totalFailed++
        }

        # Rate limiting — IT Glue API allows ~10 req/sec, stay conservative
        Start-Sleep -Milliseconds 150
    }
}

# ── 6. Summary ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Upload complete!" -ForegroundColor Green
Write-Host "  Uploaded : $totalOk"
Write-Host "  Skipped  : $totalSkipped"
Write-Host "  Failed   : $totalFailed"

if ($failedDocs.Count -gt 0) {
    Write-Host ""
    Write-Host "Failed documents:" -ForegroundColor Red
    $failedDocs | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }

    $logPath = Join-Path (Split-Path $ZipPath) "itglue-upload-failures.txt"
    $failedDocs | Out-File -FilePath $logPath -Encoding UTF8
    Write-Host ""
    Write-Host "Failure log saved to: $logPath" -ForegroundColor Yellow
}

Remove-Item $WorkDir -Recurse -Force
