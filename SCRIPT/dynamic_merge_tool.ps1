param(
    [string]$csvName,
    [string]$targetGroupsRaw
)

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$SCRIPT_PARENT = Split-Path -Path $scriptDir -Parent  # This is the "99. Generate Report" folder
$BASE_DIR = Split-Path -Path $SCRIPT_PARENT -Parent  # This is the project root (PO TIANU JB)

$MAPPING_DIR = Join-Path $SCRIPT_PARENT "EXCELL MAPPING"
$OUTPUT_ROOT = Join-Path $SCRIPT_PARENT "OUTPUT FILE\MERGED GROUPS"
$outputRoot_Main = Join-Path $SCRIPT_PARENT "OUTPUT FILE" # For input search

# 1. Resolve CSV Path
if ([string]::IsNullOrWhiteSpace($csvName)) {
    Write-Host "Error: No CSV name provided." -ForegroundColor Red
    return
}

if (-not $csvName.EndsWith(".csv")) { $csvName += ".csv" }
$csvPath = Join-Path $MAPPING_DIR $csvName

if (-not (Test-Path $csvPath)) {
    Write-Host "Error: CSV not found at $csvPath" -ForegroundColor Red
    return
}

# 2. Determine Input Directory
$prefix = $csvName.Split(".")[0]
$outputRoot = Join-Path $SCRIPT_PARENT "OUTPUT FILE"
$individualRoot = Join-Path $outputRoot "INDIVIDUAL"

# Priority 1: INDIVIDUAL\Prefix
$csvSpecificFolder = Join-Path $individualRoot $prefix
if (Test-Path $csvSpecificFolder) {
    $inputDir = $csvSpecificFolder
    Write-Host "Found CSV-specific folder in INDIVIDUAL: $prefix" -ForegroundColor Green
} else {
    # Priority 2: Root\Prefix (Legacy)
    $csvLegacyFolder = Join-Path $outputRoot $prefix
    if (Test-Path $csvLegacyFolder) {
        $inputDir = $csvLegacyFolder
        Write-Host "Found legacy CSV-specific folder: $prefix" -ForegroundColor Green
    } else {
        # Priority 3: Fallback to INDIVIDUAL root
        Write-Host "No specific folder found for prefix $prefix. Falling back to INDIVIDUAL folder." -ForegroundColor Gray
        $inputDir = $individualRoot
    }
}

if (-not (Test-Path $inputDir)) {
    Write-Host "Error: Could not find input folder at $inputDir" -ForegroundColor Red
    return
}

# 3. Setup Output Directory
$targetOutputDir = $inputDir
if (!(Test-Path $targetOutputDir)) {
    New-Item -ItemType Directory -Force -Path $targetOutputDir | Out-Null
}

# 4. Load CSV and Determine Groups
$data = Import-Csv $csvPath
$availableGroups = $data | Where-Object { $_.Group -match '^\d+$' } | Select-Object -ExpandProperty Group -Unique | Sort-Object { [int]$_ }

$groupsToProcess = @()
if ([string]::IsNullOrWhiteSpace($targetGroupsRaw) -or $targetGroupsRaw -eq "all") {
    $groupsToProcess = $availableGroups
} else {
    $groupsToProcess = $targetGroupsRaw.Split(",") | ForEach-Object { $_.Trim() }
}

# 5. Word Automation
Write-Host "Merging groups from: $($csvName)" -ForegroundColor Cyan
Write-Host "Input Folder: $(Split-Path $inputDir -Leaf)" -ForegroundColor Cyan
Write-Host "Output Folder: $targetOutputDir" -ForegroundColor Cyan

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

foreach ($groupName in $groupsToProcess) {
    if ($availableGroups -notcontains $groupName) {
        Write-Host "Warning: Group $groupName not found in CSV. Skipping." -ForegroundColor Yellow
        continue
    }

    Write-Host "Processing Group $groupName..." -NoNewline
    
    $pages = $data | Where-Object { $_.Group -eq $groupName } | Select-Object -ExpandProperty "Output File Name" -Unique
    
    if ($pages.Count -eq 0) {
        Write-Host " [EMPTY]" -ForegroundColor Gray
        continue
    }

    $foundFirst = $false
    $doc = $null

    foreach ($pageName in $pages) {
        $filePath = Join-Path $inputDir "$pageName.docx"
        
        if (Test-Path $filePath) {
            if (-not $foundFirst) {
                # This is the first available page, use it as the base
                $outputFileName = "Group $groupName.docx"
                $outputFilePath = Join-Path $targetOutputDir $outputFileName
                Copy-Item $filePath $outputFilePath -Force
                $doc = $word.Documents.Open($outputFilePath)
                $selection = $word.Selection
                $foundFirst = $true
            } else {
                # Append subsequent pages
                $selection.EndKey(6) | Out-Null # wdStory
                $selection.InsertBreak(7)      # wdPageBreak
                $selection.InsertFile($filePath)
            }
        } else {
            Write-Host " [Missing: $pageName.docx]" -NoNewline
        }
    }

    if ($foundFirst) {
        $doc.Save()
        $doc.Close()
        Write-Host " [SAVED]" -ForegroundColor Green
    } else {
        Write-Host " [ERROR: No files found]" -ForegroundColor Red
    }
}

$word.Quit()
Write-Host "`nAll requested groups merged successfully." -ForegroundColor Yellow
