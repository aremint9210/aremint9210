# orchestrate_all.ps1
param(
    [string]$choices,  # The space-separated numbers (e.g. "1 2 5") or "all"
    [string]$pages = "all"
)

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$generateScript = Join-Path $scriptDir "generate_report.py"
$auditScript = Join-Path $scriptDir "check_missing_reports.py"
$mergeScript = Join-Path $scriptDir "dynamic_merge_tool.ps1"
$MAPPING_DIR = Join-Path $scriptDir "..\EXCELL MAPPING"

# 1. Get List of CSV Files
$csvFiles = Get-ChildItem -Path $MAPPING_DIR -Filter "*.csv" | Where-Object { 
    $_.Name -notin "image number.csv", "Master List.csv" 
} | Sort-Object Name

# 2. Determine which files to process
$targetFiles = @()
if ($choices -eq "all") {
    $targetFiles = $csvFiles
} else {
    $indices = $choices.Split(" ") | ForEach-Object { [int]$_ - 1 }
    foreach ($idx in $indices) {
        if ($idx -ge 0 -and $idx -lt $csvFiles.Count) {
            $targetFiles += $csvFiles[$idx]
        }
    }
}

if ($targetFiles.Count -eq 0) {
    Write-Host "No valid CSV files selected." -ForegroundColor Red
    return
}

Write-Host "`nFound $($targetFiles.Count) substations to process." -ForegroundColor Yellow

# 3. Process each file one-by-one
foreach ($file in $targetFiles) {
    $csvName = $file.Name
    
    Write-Host "`n===============================================" -ForegroundColor Yellow
    Write-Host "   NOW PROCESSING: $csvName" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Yellow

    # STEP 1: GENERATION
    Write-Host "`n>>> STEP 1: GENERATING REPORTS ($csvName) <<<" -ForegroundColor Cyan
    # Using & and waiting is the most robust way in PowerShell
    & python $generateScript $csvName $pages "--no-merge"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Generator reported errors for $csvName" -ForegroundColor Yellow
    }

    # STEP 2: AUDIT & AUTO-FIX
    Write-Host "`n>>> STEP 2: AUDITING & AUTO-FIXING ($csvName) <<<" -ForegroundColor Cyan
    & python $auditScript $csvName "--auto-fix"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Audit tool reported errors for $csvName" -ForegroundColor Yellow
    }

    # STEP 3: MERGING
    Write-Host "`n>>> STEP 3: MERGING GROUPS ($csvName) <<<" -ForegroundColor Cyan
    # We call the merge tool script directly
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mergeScript -csvName $csvName -targetGroupsRaw "all"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Merger reported errors for $csvName" -ForegroundColor Yellow
    }

    Write-Host "`n[COMPLETED] Substation: $csvName" -ForegroundColor Green
    Write-Host "-----------------------------------------------"
}

Write-Host "`n===============================================" -ForegroundColor Green
Write-Host "   ALL SUBSTATIONS PROCESSED SUCCESSFULLY" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
