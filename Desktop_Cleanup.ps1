# ============================================================
#  Desktop_Cleanup.ps1
#  Moves recognised files from Desktop into Documents\Business
#  or Documents\Personal (by type and date).
#  Asks what to do with unrecognised files.
# ============================================================

# ---- CONFIGURATION -----------------------------------------
$DocsFolder    = [Environment]::GetFolderPath("MyDocuments")
$DesktopFolder = [Environment]::GetFolderPath("Desktop")
$DryRun        = $true  # Set to $false to actually move files
# ------------------------------------------------------------

$TypeMap = @{
    ".pdf"  = "PDFs"
    ".doc"  = "Word Docs"
    ".docx" = "Word Docs"
    ".xls"  = "Spreadsheets"
    ".xlsx" = "Spreadsheets"
    ".csv"  = "Spreadsheets"
    ".ppt"  = "Presentations"
    ".pptx" = "Presentations"
    ".jpg"  = "Images"
    ".jpeg" = "Images"
    ".png"  = "Images"
    ".gif"  = "Images"
    ".bmp"  = "Images"
    ".txt"  = "Text Files"
    ".md"   = "Text Files"
    ".rtf"  = "Text Files"
    ".zip"  = "Archives"
    ".rar"  = "Archives"
    ".7z"   = "Archives"
}

$BusinessKeywords = @("invoice", "report", "contract", "proposal", "meeting",
                      "budget", "project", "client", "brief", "agenda",
                      "statement", "receipt", "tax", "expense")
$PersonalKeywords = @("photo", "holiday", "vacation", "family", "personal",
                      "medical", "insurance", "diary", "letter", "recipe")

function Get-Category ($file) {
    $name = $file.BaseName.ToLower()
    foreach ($kw in $BusinessKeywords) { if ($name -like "*$kw*") { return "Business" } }
    foreach ($kw in $PersonalKeywords) { if ($name -like "*$kw*") { return "Personal" } }
    return "Personal"
}

function Move-ToDocuments ($file) {
    $ext       = $file.Extension.ToLower()
    $category  = Get-Category $file
    $typFolder = $TypeMap[$ext]
    $year      = $file.LastWriteTime.ToString("yyyy")
    $month     = $file.LastWriteTime.ToString("MM-MMMM")

    $destDir  = Join-Path $DocsFolder "$category\$typFolder\$year\$month"
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
}

$moved   = 0
$left    = 0
$skipped = 0

Write-Host ""
Write-Host "---- Desktop Cleanup ----------------------------" -ForegroundColor White

foreach ($file in Get-ChildItem -Path $DesktopFolder -File -ErrorAction SilentlyContinue) {
    $ext = $file.Extension.ToLower()

    # Skip .lnk shortcut files silently
    if ($ext -eq ".lnk" -or $ext -eq ".url") {
        $left++
        continue
    }

    if ($TypeMap.ContainsKey($ext)) {
        Move-ToDocuments $file
        $moved++
    } else {
        # Ask the user what to do
        Write-Host ""
        Write-Host "UNKNOWN: $($file.Name)" -ForegroundColor Yellow
        Write-Host "  [M] Move to Documents\Personal\Misc" -ForegroundColor Cyan
        Write-Host "  [B] Move to Documents\Business\Misc" -ForegroundColor Cyan
        Write-Host "  [L] Leave it on the Desktop" -ForegroundColor Cyan
        Write-Host "  [D] Delete it" -ForegroundColor Red

        if (-not $DryRun) {
            $choice = Read-Host "  Your choice"
            switch ($choice.ToUpper()) {
                "M" {
                    $destDir  = Join-Path $DocsFolder "Personal\Misc"
                    $destFile = Join-Path $destDir $file.Name
                    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Move-Item -Path $file.FullName -Destination $destFile
                    Write-Host "  MOVED to Personal\Misc" -ForegroundColor Green
                    $moved++
                }
                "B" {
                    $destDir  = Join-Path $DocsFolder "Business\Misc"
                    $destFile = Join-Path $destDir $file.Name
                    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Move-Item -Path $file.FullName -Destination $destFile
                    Write-Host "  MOVED to Business\Misc" -ForegroundColor Green
                    $moved++
                }
                "D" {
                    Remove-Item -Path $file.FullName -Force
                    Write-Host "  DELETED" -ForegroundColor Red
                    $skipped++
                }
                default {
                    Write-Host "  LEFT on Desktop" -ForegroundColor DarkGray
                    $left++
                }
            }
        } else {
            Write-Host "  (DRY RUN - will prompt when run for real)" -ForegroundColor DarkGray
            $skipped++
        }
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor White
if ($DryRun) {
    Write-Host " DRY RUN - no files were moved." -ForegroundColor Yellow
    Write-Host " Change DryRun to false to apply changes." -ForegroundColor Yellow
} else {
    Write-Host " Done!" -ForegroundColor Green
}
Write-Host " Files moved         : $moved"
Write-Host " Files left          : $left  (shortcuts or your choice)"
Write-Host " Files deleted       : $skipped"
Write-Host "============================================" -ForegroundColor White
