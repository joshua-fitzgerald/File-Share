# Prepare-ITBoostForITGlue.ps1
# Renames ITBoost HTML export files to match their parent folder names,
# patches the <title> tag, and re-zips for IT Glue bulk import.
#
# USAGE:
#   .\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\itboost-export.zip" -OutputFolder "C:\Downloads\itglue-output"
#
# ITBoost ZIP structure handled:
#   documents/
#     global/
#       DOC-1263-How to Contact Help Desk.docx/
#         7d7751db-uid.html
#     Client Knowledgebase/
#       DOC-xxx-Some Doc.docx/
#         uid.html

param(
    [Parameter(Mandatory)]
    [string]$ZipPath,

    [Parameter(Mandatory)]
    [string]$OutputFolder,

    [string]$WorkDir = "$env:TEMP\itboost_prep"
)

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

# ── 3. Discover categories ─────────────────────────────────────────────────────
$categoryFolders = Get-ChildItem -Path $scanRoot -Directory

Write-Host ""
Write-Host "Found $($categoryFolders.Count) categories in the export:" -ForegroundColor Cyan
Write-Host ""
for ($i = 0; $i -lt $categoryFolders.Count; $i++) {
    $docCount = (Get-ChildItem -Path $categoryFolders[$i].FullName -Directory).Count
    Write-Host "  [$($i+1)] $($categoryFolders[$i].Name)  ($docCount documents)"
}

# ── 4. Prompt for merge preference ────────────────────────────────────────────
Write-Host ""
Write-Host "How would you like to output the ZIPs?" -ForegroundColor Yellow
Write-Host "  [M] Merge all categories into one ZIP"
Write-Host "  [S] Separate ZIP per category"
Write-Host ""

do {
    $mode = (Read-Host "Enter M or S").Trim().ToUpper()
} while ($mode -ne "M" -and $mode -ne "S")

# ── 5. Helper: process a set of document folders into an output path ───────────
function Process-DocumentFolders {
    param(
        [System.IO.DirectoryInfo[]]$Folders,
        [string]$OutputPath
    )

    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

    $total   = $Folders.Count
    $count   = 0
    $skipped = 0

    foreach ($folder in $Folders) {
        $count++
        $rawName = $folder.Name
        $docName = $rawName -replace '\.docx$', '' -replace '\.doc$', ''

        $htmlFiles = Get-ChildItem -Path $folder.FullName -Filter "*.html" -File
        if ($htmlFiles.Count -eq 0) {
            Write-Warning "[$count/$total] No HTML in: $rawName — skipping"
            $skipped++
            continue
        }
        $htmlFile = $htmlFiles[0]

        $safeName    = $docName -replace '[\\/:*?"<>|]', '-'
        $newHtmlName = "$safeName.html"

        # Patch <title>
        $content = Get-Content -Path $htmlFile.FullName -Raw -Encoding UTF8
        if ($content -match '(?i)<title>') {
            $content = $content -replace '(?i)<title>.*?</title>', "<title>$docName</title>"
        } elseif ($content -match '(?i)<head>') {
            $content = $content -replace '(?i)<head>', "<head>`n<title>$docName</title>"
        } else {
            $content = "<html><head><title>$docName</title></head>`n" + $content
        }

        $destFolder = Join-Path $OutputPath $safeName
        New-Item -ItemType Directory -Path $destFolder | Out-Null

        # Copy attachments/images
        Get-ChildItem -Path $folder.FullName -File |
            Where-Object { $_.Extension -ne ".html" } |
            ForEach-Object { Copy-Item -Path $_.FullName -Destination $destFolder }

        [System.IO.File]::WriteAllText((Join-Path $destFolder $newHtmlName), $content, [System.Text.Encoding]::UTF8)
        Write-Host "  [$count/$total] $newHtmlName" -ForegroundColor Green
    }

    return @{ Processed = $count - $skipped; Skipped = $skipped }
}

# ── 6. Helper: ZIP a folder using Shell.Application ───────────────────────────
function Create-ZipFromFolder {
    param(
        [string]$SourceFolder,
        [string]$ZipOutputPath
    )

    if (Test-Path $ZipOutputPath) { Remove-Item $ZipOutputPath -Force }

    # Seed empty valid ZIP
    $zipBytes = [byte[]](80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
    [System.IO.File]::WriteAllBytes($ZipOutputPath, $zipBytes)

    $shell    = New-Object -ComObject Shell.Application
    $zipShell = $shell.NameSpace($ZipOutputPath)
    $srcShell = $shell.NameSpace($SourceFolder)

    foreach ($item in $srcShell.Items()) {
        $zipShell.CopyHere($item, 0x14)
        $timeout = 120; $elapsed = 0
        do {
            Start-Sleep -Milliseconds 500
            $elapsed += 0.5
        } while ($elapsed -lt $timeout)
    }
    Start-Sleep -Seconds 2
}

# ── 7. Ensure output folder exists ────────────────────────────────────────────
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

# ── 8. Execute based on mode ──────────────────────────────────────────────────
$totalProcessed = 0
$totalSkipped   = 0

if ($mode -eq "M") {
    # Merge all into one ZIP
    Write-Host ""
    Write-Host "Merging all categories..." -ForegroundColor Cyan

    $allDocFolders = $categoryFolders | ForEach-Object { Get-ChildItem -Path $_.FullName -Directory }
    $stagingPath   = Join-Path $WorkDir "merged"
    $result        = Process-DocumentFolders -Folders $allDocFolders -OutputPath $stagingPath

    $outputZip = Join-Path $OutputFolder "itglue-import-all.zip"
    Write-Host ""
    Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
    Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip

    $totalProcessed = $result.Processed
    $totalSkipped   = $result.Skipped

} else {
    # Separate ZIP per category
    foreach ($category in $categoryFolders) {
        Write-Host ""
        Write-Host "Processing category: $($category.Name)" -ForegroundColor Cyan

        $docFolders  = Get-ChildItem -Path $category.FullName -Directory
        $stagingPath = Join-Path $WorkDir "cat_$($category.Name)"
        $result      = Process-DocumentFolders -Folders $docFolders -OutputPath $stagingPath

        $safeCatName = $category.Name -replace '[\\/:*?"<>|]', '-'
        $outputZip   = Join-Path $OutputFolder "itglue-import-$safeCatName.zip"
        Write-Host ""
        Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
        Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip

        $totalProcessed += $result.Processed
        $totalSkipped   += $result.Skipped
    }
}

# ── 9. Summary ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "Complete!" -ForegroundColor Green
Write-Host "  Processed : $totalProcessed documents"
Write-Host "  Skipped   : $totalSkipped"
Write-Host "  Output    : $OutputFolder"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. IT Glue > Organizations > [Org] > Documents > Import > Documents"
Write-Host "  2. Set type to 'HTML', upload ZIP, enable 'Allow partial imports'"
Write-Host "  3. Note: IT Glue has a 300 MB limit per ZIP"
Write-Host "  4. Docs land flat after import — move into folders manually"

Remove-Item $WorkDir -Recurse -Force
