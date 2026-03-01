# ============================================================
#  Organize_Folders_script.ps1
#  1. Sorts your Documents folder into Business/Personal
#     subfolders, then by file type, then by Year\Month.
#  2. Cleans up your Downloads folder:
#       - Deletes .exe files older than N days
#       - Moves recognised file types to Documents\Business
#         or Documents\Personal
#       - Leaves everything else alone
# ============================================================

# ---- CONFIGURATION -----------------------------------------
$DocsFolder      = [Environment]::GetFolderPath("MyDocuments")
$DownloadsFolder = Join-Path $env:USERPROFILE "Downloads"
$ExeAgeDays      = 7      # Delete .exe files older than this many days
$DryRun          = $false  # Set to $false to actually move/delete files
# ------------------------------------------------------------

# File-type buckets -> subfolder name
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

# Keywords in filenames -> Business or Personal bucket
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

function Move-FileToDocuments ($file) {
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
        Write-Host "DRY   $($file.FullName)" -ForegroundColor Cyan
        Write-Host "   ->  $destFile" -ForegroundColor Yellow
    } else {
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Move-Item -Path $file.FullName -Destination $destFile
        Write-Host "MOVED $($file.Name)  ->  $destFile" -ForegroundColor Green
    }
}

# ============================================================
#  SECTION 1 - Organize Documents folder
# ============================================================
Write-Host ""
Write-Host "---- Documents Folder ---------------------------" -ForegroundColor White

$docsMoved   = 0
$docsSkipped = 0

foreach ($file in Get-ChildItem -Path $DocsFolder -File -ErrorAction SilentlyContinue) {
    $ext = $file.Extension.ToLower()
    if (-not $TypeMap.ContainsKey($ext)) {
        Write-Host "SKIP  (unknown type) : $($file.Name)" -ForegroundColor DarkGray
        $docsSkipped++
        continue
    }
    Move-FileToDocuments $file
    $docsMoved++
}

# ============================================================
#  SECTION 2 - Clean up Downloads folder
# ============================================================
Write-Host ""
Write-Host "---- Downloads Folder ---------------------------" -ForegroundColor White

$exeDeleted = 0
$dlMoved    = 0
$dlLeft     = 0
$cutoff     = (Get-Date).AddDays(-$ExeAgeDays)

foreach ($file in Get-ChildItem -Path $DownloadsFolder -File -ErrorAction SilentlyContinue) {
    $ext = $file.Extension.ToLower()

    if ($ext -eq ".exe") {
        if ($file.LastWriteTime -lt $cutoff) {
            if ($DryRun) {
                Write-Host "DRY   DELETE (exe older than $ExeAgeDays days) : $($file.Name)" -ForegroundColor Magenta
            } else {
                Remove-Item -Path $file.FullName -Force
                Write-Host "DELETED $($file.Name)" -ForegroundColor Red
            }
            $exeDeleted++
        } else {
            Write-Host "KEEP  (exe, recent)  : $($file.Name)" -ForegroundColor DarkGray
            $dlLeft++
        }
        continue
    }

    if ($TypeMap.ContainsKey($ext)) {
        Move-FileToDocuments $file
        $dlMoved++
        continue
    }

    Write-Host "LEAVE (unknown type) : $($file.Name)" -ForegroundColor DarkGray
    $dlLeft++
}

# ============================================================
#  SUMMARY
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor White
if ($DryRun) {
    Write-Host " DRY RUN - no files were moved or deleted." -ForegroundColor Yellow
    Write-Host " Change DryRun to false to apply changes." -ForegroundColor Yellow
} else {
    Write-Host " Done!" -ForegroundColor Green
}
Write-Host " Documents organised : $docsMoved  (skipped: $docsSkipped)"
Write-Host " Downloads moved     : $dlMoved"
Write-Host " Downloads deleted   : $exeDeleted  (exe files older than $ExeAgeDays days)"
Write-Host " Downloads left      : $dlLeft  (recent or unrecognised)"
Write-Host "============================================" -ForegroundColor White