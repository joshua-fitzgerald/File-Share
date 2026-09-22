# Prepare-ITBoostForITGlue.ps1
# Renames ITBoost HTML export files to match their parent folder names,
# patches the <title> tag, and re-zips for IT Glue bulk import.
#
# USAGE:
#   .\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\export.zip" -OutputFolder "C:\Downloads\output"
#   .\Prepare-ITBoostForITGlue.ps1 -ZipPath "C:\Downloads\export.zip" -OutputFolder "C:\Downloads\output" -DiagnosticOnly
#
# Handles any depth of nesting under documents/ - finds doc folders by presence of .html files

param(
    [Parameter(Mandatory)]
    [string]$ZipPath,

    [Parameter(Mandatory)]
    [string]$OutputFolder,

    [switch]$DiagnosticOnly,

    [string]$WorkDir = "$env:TEMP\itboost_prep"
)

# 1. Extract ZIP
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
New-Item -ItemType Directory -Path $WorkDir | Out-Null
$extractPath = Join-Path $WorkDir "extracted"
New-Item -ItemType Directory -Path $extractPath | Out-Null

Write-Host ""
Write-Host "Extracting ZIP..." -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $extractPath)

# 2. Find documents root
$docsRoot = Get-ChildItem -Path $extractPath -Directory -Recurse |
    Where-Object { $_.Name -eq "documents" } |
    Select-Object -First 1
$scanRoot = if ($docsRoot) { $docsRoot.FullName } else { $extractPath }

Write-Host "Scan root: $scanRoot" -ForegroundColor DarkGray

# 3. Find all doc folders by detecting which folders contain an .html file directly inside
Write-Host "Scanning for document folders..." -ForegroundColor Cyan

$allDocFolders = Get-ChildItem -Path $scanRoot -Directory -Recurse | Where-Object {
    $folder = $_
    $htmlCount = (Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -ieq ".html" }).Count
    $htmlCount -gt 0
}

Write-Host "Found $($allDocFolders.Count) document folders." -ForegroundColor Green

# 4. Diagnostic mode
if ($DiagnosticOnly) {
    Write-Host ""
    Write-Host "=== DIAGNOSTIC MODE - no files will be written ===" -ForegroundColor Yellow
    foreach ($doc in $allDocFolders) {
        $relativePath = $doc.FullName.Replace($scanRoot, "").TrimStart("\").TrimStart("/")
        Write-Host ""
        Write-Host "  DOC FOLDER : $relativePath" -ForegroundColor Yellow
        $allFiles = Get-ChildItem -Path $doc.FullName -File
        foreach ($f in $allFiles) {
            Write-Host "    FILE: $($f.Name)  [ext: $($f.Extension)]" -ForegroundColor White
        }
    }
    Write-Host ""
    Write-Host "Diagnostic complete. Re-run without -DiagnosticOnly to process." -ForegroundColor Green
    Remove-Item $WorkDir -Recurse -Force
    exit
}

# 5. Group by top-level category (first path segment under scanRoot)
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

Write-Host ""
Write-Host "How would you like to output the ZIPs?" -ForegroundColor Yellow
Write-Host "  [M] Merge all into one ZIP"
Write-Host "  [S] Separate ZIP per top-level category"
Write-Host ""

do {
    $mode = (Read-Host "Enter M or S").Trim().ToUpper()
} while ($mode -ne "M" -and $mode -ne "S")

# 6. Process function
function Process-DocumentFolders {
    param(
        [object[]]$Folders,
        [string]$OutputPath
    )

    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

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

        $htmlFile    = $htmlFiles[0]
        $safeName    = $docName -replace '[\\/:*?"<>|]', "-"
        $newHtmlName = "$safeName.html"

        $content = Get-Content -Path $htmlFile.FullName -Raw -Encoding UTF8
        if ($content -match "(?i)<title>") {
            $content = $content -replace "(?i)<title>.*?</title>", "<title>$docName</title>"
        } elseif ($content -match "(?i)<head>") {
            $content = $content -replace "(?i)<head>", "<head>`n<title>$docName</title>"
        } else {
            $content = "<html><head><title>$docName</title></head>`n" + $content
        }

        $destFolder = Join-Path $OutputPath $safeName
        New-Item -ItemType Directory -Path $destFolder | Out-Null

        $allFiles | Where-Object { $_.Extension -ine ".html" } |
            ForEach-Object { Copy-Item -Path $_.FullName -Destination $destFolder }

        [System.IO.File]::WriteAllText((Join-Path $destFolder $newHtmlName), $content, [System.Text.Encoding]::UTF8)
        Write-Host "  [$count/$total] $newHtmlName" -ForegroundColor Green
    }

    return @{ Processed = $count - $skipped; Skipped = $skipped }
}

# 7. ZIP function
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
        $timeout = 120
        $elapsed = 0
        do { Start-Sleep -Milliseconds 500; $elapsed += 0.5 } while ($elapsed -lt $timeout)
    }
    Start-Sleep -Seconds 2
}

# 8. Ensure output folder exists
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

$totalProcessed = 0
$totalSkipped   = 0

if ($mode -eq "M") {
    Write-Host ""
    Write-Host "Merging all categories..." -ForegroundColor Cyan
    $stagingPath = Join-Path $WorkDir "merged"
    $result      = Process-DocumentFolders -Folders $allDocFolders -OutputPath $stagingPath
    $outputZip   = Join-Path $OutputFolder "itglue-import-all.zip"
    Write-Host ""
    Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
    Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip
    $totalProcessed = $result.Processed
    $totalSkipped   = $result.Skipped
} else {
    foreach ($group in $categoryGroups) {
        Write-Host ""
        Write-Host "Processing category: $($group.Name)" -ForegroundColor Cyan
        $safeCatName = $group.Name -replace '[\\/:*?"<>|]', "-"
        $stagingPath = Join-Path $WorkDir "cat_$safeCatName"
        $result      = Process-DocumentFolders -Folders $group.Group -OutputPath $stagingPath
        $outputZip   = Join-Path $OutputFolder "itglue-import-$safeCatName.zip"
        Write-Host ""
        Write-Host "Creating ZIP: $outputZip" -ForegroundColor Cyan
        Create-ZipFromFolder -SourceFolder $stagingPath -ZipOutputPath $outputZip
        $totalProcessed += $result.Processed
        $totalSkipped   += $result.Skipped
    }
}

# 9. Summary
Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "Complete!" -ForegroundColor Green
Write-Host "  Processed : $totalProcessed documents"
Write-Host "  Skipped   : $totalSkipped"
Write-Host "  Output    : $OutputFolder"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. IT Glue > Organizations > [Org] > Documents > Import > Documents"
Write-Host "  2. Set type to HTML, upload ZIP, enable Allow partial imports"
Write-Host "  3. IT Glue has a 300 MB limit per ZIP"
Write-Host "  4. Docs land flat after import - move into folders manually"

Remove-Item $WorkDir -Recurse -Force
