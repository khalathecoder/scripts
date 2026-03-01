# ============================================================
#  Amazon_Organizer.ps1
#  Organizes Amazon export files from Downloads into
#  Documents\Business\Amazon\{Year}\{Month}
#  Handles: Order history CSVs, Invoices/PDFs, Spreadsheets
# ============================================================

# ---- CONFIGURATION -----------------------------------------
$DocsFolder      = [Environment]::GetFolderPath("MyDocuments")
$DownloadsFolder = Join-Path $env:USERPROFILE "Downloads"
$AmazonBase      = Join-Path $DocsFolder "Business\Amazon"
$DryRun          = $true  # Set to $false to actually move files
# ------------------------------------------------------------

# Amazon filename patterns to match
$AmazonPatterns = @(
    "*amazon*",
    "*order*history*",
    "*order-history*",
    "*digital-order*",
    "*invoice*amazon*",
    "*amazon*invoice*",
    "*amazon*receipt*",
    "*amazon*export*",
    "*amazon*report*"
)

# Subfolder by file type
$AmazonTypeMap = @{
    ".csv"  = "Order History"
    ".xlsx" = "Spreadsheets"
    ".xls"  = "Spreadsheets"
    ".pdf"  = "Invoices"
}

$moved   = 0
$skipped = 0

Write-Host ""
Write-Host "---- Amazon Export Organizer --------------------" -ForegroundColor White

foreach ($file in Get-ChildItem -Path $DownloadsFolder -File -ErrorAction SilentlyContinue) {
    $name = $file.Name.ToLower()
    $ext  = $file.Extension.ToLower()

    # Check if file matches any Amazon pattern
    $isAmazon = $false
    foreach ($pattern in $AmazonPatterns) {
        if ($name -like $pattern) { $isAmazon = $true; break }
    }

    if (-not $isAmazon) {
        continue
    }

    # Check file type is recognised
    if (-not $AmazonTypeMap.ContainsKey($ext)) {
        Write-Host "SKIP  (unrecognised type) : $($file.Name)" -ForegroundColor DarkGray
        $skipped++
        continue
    }

    $typeFolder = $AmazonTypeMap[$ext]
    $year       = $file.LastWriteTime.ToString("yyyy")
    $month      = $file.LastWriteTime.ToString("MM-MMMM")

    $destDir  = Join-Path $AmazonBase "$typeFolder\$year\$month"
    $destFile = Join-Path $destDir $file.Name

    if (Test-Path $destFile) {
        $stamp    = $file.LastWriteTime.ToString("yyyyMMdd_HHmmss")
        $destFile = Join-Path $destDir "$($file.BaseName)_$stamp$ext"
    }

    if ($DryRun) {
        Write-Host "DRY   $($file.Name)" -ForegroundColor Cyan
        Write-Host "   ->  $destFile" -ForegroundColor Yellow
    } else {
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Move-Item -Path $file.FullName -Destination $destFile
        Write-Host "MOVED $($file.Name)  ->  $destFile" -ForegroundColor Green
    }
    $moved++
}

Write-Host ""
Write-Host "============================================" -ForegroundColor White
if ($DryRun) {
    Write-Host " DRY RUN - no files were moved." -ForegroundColor Yellow
    Write-Host " Change DryRun to false to apply changes." -ForegroundColor Yellow
} else {
    Write-Host " Done!" -ForegroundColor Green
}
Write-Host " Amazon files moved  : $moved"
Write-Host " Skipped             : $skipped"
Write-Host "============================================" -ForegroundColor White
