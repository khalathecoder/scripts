# ============================================================
#  Duplicate_Finder.ps1
#  Scans Documents and Downloads for duplicate files.
#  Duplicates are identified by file hash (exact content match).
#  Keeps the OLDEST copy, deletes the rest.
#  Always do a DRY RUN first to review before deleting!
# ============================================================

# ---- CONFIGURATION -----------------------------------------
$ScanFolders = @(
    [Environment]::GetFolderPath("MyDocuments"),
    (Join-Path $env:USERPROFILE "Downloads")
)
$DryRun = $true  # Set to $false to actually delete duplicates
# ------------------------------------------------------------

Write-Host ""
Write-Host "---- Duplicate File Finder ----------------------" -ForegroundColor White
Write-Host " Scanning folders..." -ForegroundColor DarkGray

$hashTable  = @{}
$totalFiles = 0
$dupCount   = 0
$savedBytes = 0

# Collect all files and hash them
foreach ($folder in $ScanFolders) {
    foreach ($file in Get-ChildItem -Path $folder -File -Recurse -ErrorAction SilentlyContinue) {
        $totalFiles++
        try {
            $hash = (Get-FileHash -Path $file.FullName -Algorithm MD5).Hash
            if (-not $hashTable.ContainsKey($hash)) {
                $hashTable[$hash] = @()
            }
            $hashTable[$hash] += $file
        } catch {
            Write-Host "SKIP  (could not hash) : $($file.Name)" -ForegroundColor DarkGray
        }
    }
}

Write-Host " Scanned $totalFiles files. Checking for duplicates..." -ForegroundColor DarkGray
Write-Host ""

# Process duplicates
foreach ($hash in $hashTable.Keys) {
    $group = $hashTable[$hash]
    if ($group.Count -lt 2) { continue }

    # Sort by creation time - keep the oldest, delete the rest
    $sorted = $group | Sort-Object CreationTime
    $keeper = $sorted[0]
    $dupes  = $sorted[1..($sorted.Count - 1)]

    Write-Host "DUPLICATE GROUP:" -ForegroundColor White
    Write-Host "  KEEP   $($keeper.FullName)  [$($keeper.CreationTime.ToString('yyyy-MM-dd'))]" -ForegroundColor Green

    foreach ($dupe in $dupes) {
        $size = [math]::Round($dupe.Length / 1KB, 1)
        $savedBytes += $dupe.Length

        if ($DryRun) {
            Write-Host "  DRY DELETE  $($dupe.FullName)  [$($dupe.CreationTime.ToString('yyyy-MM-dd'))]  ($size KB)" -ForegroundColor Magenta
        } else {
            try {
                Remove-Item -Path $dupe.FullName -Force
                Write-Host "  DELETED  $($dupe.FullName)  ($size KB)" -ForegroundColor Red
            } catch {
                Write-Host "  FAILED to delete $($dupe.FullName) : $_" -ForegroundColor Yellow
            }
        }
        $dupCount++
    }
    Write-Host ""
}

$savedMB = [math]::Round($savedBytes / 1MB, 2)

Write-Host "============================================" -ForegroundColor White
if ($DryRun) {
    Write-Host " DRY RUN - no files were deleted." -ForegroundColor Yellow
    Write-Host " Change DryRun to false to apply changes." -ForegroundColor Yellow
} else {
    Write-Host " Done!" -ForegroundColor Green
}
Write-Host " Total files scanned : $totalFiles"
Write-Host " Duplicates found    : $dupCount"
Write-Host " Space to recover    : $savedMB MB"
Write-Host "============================================" -ForegroundColor White
