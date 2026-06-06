param(
    [string]$CsvPath,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$baseDir = Split-Path -Path $scriptDir -Parent
$templatePath = Join-Path $baseDir "TEMPLATE\00. EXECUTIVE SUMMARY - RMU.docx"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $csvName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    # Target: INDIVIDUAL\SubstationName\Group 0.docx
    $outputDir = Join-Path $baseDir "OUTPUT FILE\INDIVIDUAL\$csvName"
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir }
    $outputPath = Join-Path $outputDir "Group 0.docx"
}

Write-Host "Starting Summary Generation (PowerShell)..."
Write-Host "CSV Path: $CsvPath"

if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

if (-not (Test-Path $templatePath)) {
    Write-Error "Template not found: $templatePath"
    exit 1
}

# Load Data
$data = Import-Csv -Path $CsvPath
Write-Host "Read $($data.Count) rows from CSV."

# Determine Template based on Switchgear Type
$swgType = $data[0]."{{ swg.type }}".Trim()
Write-Host "Detected Switchgear Type: $swgType"

if ($swgType -eq "VCB") {
    $templatePath = Join-Path $baseDir "TEMPLATE\00. EXECUTIVE SUMMARY - VCB.docx"
} else {
    $templatePath = Join-Path $baseDir "TEMPLATE\00. EXECUTIVE SUMMARY - RMU.docx"
}

if (-not (Test-Path $templatePath)) {
    Write-Error "Template not found: $templatePath"
    exit 1
}

# Initialize Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false

try {
    $doc = $word.Documents.Open($templatePath)
    
    # Pre-filter data in memory to avoid redundant Word calls
    Write-Host "Filtering data..."
    $processedCombos = @{}
    $finalRows = New-Object System.Collections.Generic.List[PSObject]
    $itemSequence = 0
    $lastItemId = $null

    foreach ($row in $data) {
        $typeVal = $row.Type.Trim()
        if ([string]::IsNullOrWhiteSpace($typeVal)) { continue }

        $area = $row."{{ swg.area }}".Trim()
        if ([string]::IsNullOrWhiteSpace($area) -or $area -eq "nan") {
            $area = $row.Attribute.Trim()
        }

        $comboKey = "$typeVal|$area"
        if ($processedCombos.ContainsKey($comboKey)) { continue }

        # Image check
        $folder = $row.Folder
        $file = $row."Image File Name"
        if ($folder -and $file) {
            $imgPath = Join-Path $folder $file
            if (-not (Test-Path $imgPath)) { continue }
        } else { continue }

        # Sequence
        $rawNo = $row.No.Trim()
        $itemId = "$typeVal|$rawNo"
        if ($itemId -ne $lastItemId) {
            $itemSequence++
            $lastItemId = $itemId
        }

        # Equipment Description
        $equipDesc = ""
        if ($typeVal -like "*Overview*") {
            $equipDesc = "$($row."{{ swg.rating }}") $($row."{{ swg.type }}") $($row."{{ swg.manufacturer }}")"
        } elseif ($typeVal -eq "POWER TX1") {
            $equipDesc = "PTX1 - $($row."{{ tx1.rating }}") $($row."{{ tx1.manufacturer }}")"
        } elseif ($typeVal -eq "POWER TX2") {
            $r = $row."{{ tx2.rating }}"
            $m = $row."{{ tx2.manufacturer }}"
            if (-not $r -or $r -eq "nan") { $r = $row."{{ tx1.rating }}" }
            if (-not $m -or $m -eq "nan") { $m = $row."{{ tx1.manufacturer }}" }
            $equipDesc = "PTX2 - $r $m"
        } elseif ($typeVal -eq "LOCALT TX") {
            $equipDesc = "LTX - $($row."{{ fp1.rating }}") $($row."{{ fp1.manufacturer }}")"
        } elseif ($typeVal -eq "FEEDER PILLAR") {
            $equipDesc = "FEEDER PILLAR - $($row."{{ fp1.manufacturer }}")"
        } elseif ($typeVal -like "*panel*") {
            $pNo = $row."{{ panel.linknumber }}".Trim()
            $pName = $row."{{ panel.name }}".Trim()
            if ($pNo -and $pName -and $pNo -ne "-" -and $pName -ne "-") {
                $equipDesc = "PANEL $pNo`n$pName"
            } elseif ($pNo -and $pNo -ne "-") {
                $equipDesc = "PANEL $pNo"
            } elseif ($pName -and $pName -ne "-") {
                $equipDesc = $pName
            } else {
                $equipDesc = $typeVal.ToUpper()
            }
        } else {
            $equipDesc = $typeVal.ToUpper()
        }
        $equipDesc = $equipDesc.Replace("nan", "").Replace(" - ", " ").Trim()

        $finalRows.Add([PSCustomObject]@{
            Seq = $itemSequence
            Equip = $equipDesc
            Area = $area
        })
        $processedCombos[$comboKey] = $true
    }

    Write-Host "Found $($finalRows.Count) unique inspection items. Generating table..."

    # Recreate the table with all rows at once
    try {
        $existingTable = $doc.Tables.Item(1)
        $range = $existingTable.Range
        $existingTable.Delete()
        $table = $doc.Tables.Add($range, ($finalRows.Count + 1), 6)
        try { $table.Style = "Table Grid" } catch { Write-Host "Warning: 'Table Grid' style not found. Using default." }
    } catch {
        $table = $doc.Tables.Add($doc.Range($doc.Content.End - 1), ($finalRows.Count + 1), 6)
        try { $table.Style = "Table Grid" } catch { Write-Host "Warning: 'Table Grid' style not found. Using default." }
    }

    # Bulk Table Formatting
    $table.Range.Font.Name = "Calibri"
    $table.Range.Font.Size = 10
    $table.Range.ParagraphFormat.Alignment = 1 # Center
    $table.Range.Cells.VerticalAlignment = 1 # Center
    $table.Range.ParagraphFormat.SpaceBefore = 0
    $table.Range.ParagraphFormat.SpaceAfter = 0
    $table.Range.ParagraphFormat.LineSpacingRule = 4 # wdLineSpaceExactly
    $table.Range.ParagraphFormat.LineSpacing = 12

    # Column Widths
    $colWidths = @(28.35, 127.575, 141.75, 70.875, 70.875, 70.875)
    for ($i=1; $i -le 6; $i++) {
        $table.Columns.Item($i).Width = $colWidths[$i-1]
    }

    # Header
    $headers = @("NO.", "EQUIPMENT", "AREA", "IR (Abs T/∆T)", "US (dB)", "DEFECT")
    for ($i=1; $i -le 6; $i++) {
        $c = $table.Cell(1, $i)
        $c.Range.Text = $headers[$i-1]
        $c.Range.Font.Bold = $true
        $c.Shading.BackgroundPatternColor = 12632256 # Light Gray
    }

    # Fill Data
    $rowIndex = 2
    $lastEquip = ""
    foreach ($row in $finalRows) {
        $table.Cell($rowIndex, 1).Range.Text = [string]$row.Seq
        
        if ($row.Equip -ne $lastEquip) {
            $table.Cell($rowIndex, 2).Range.Text = $row.Equip
            $lastEquip = $row.Equip
        }
        
        $table.Cell($rowIndex, 3).Range.Text = $row.Area
        $table.Cell($rowIndex, 4).Range.Text = "-"
        $table.Cell($rowIndex, 5).Range.Text = "-"
        
        if ($row.Area.ToUpper().Contains("OVERVIEW")) {
            $table.Cell($rowIndex, 6).Range.Text = "-"
        } else {
            $table.Cell($rowIndex, 6).Shading.BackgroundPatternColor = 5287936 # Green
        }
        $rowIndex++
    }

    # Merging Pass
    Write-Host "Merging cells..."
    $startMergeRow = 2
    for ($r = 2; $r -le $table.Rows.Count; $r++) {
        $val = $table.Cell($r, 2).Range.Text -replace "[\x07\x0D]",""
        if ($r -gt 2) {
            $prevVal = $table.Cell($r-1, 2).Range.Text -replace "[\x07\x0D]",""
            if ($val -ne "" -and $val -ne $prevVal) {
                if ($r - 1 -gt $startMergeRow) {
                    $table.Cell($startMergeRow, 2).Merge($table.Cell($r - 1, 2))
                }
                $startMergeRow = $r
            }
        }
    }
    if ($table.Rows.Count -gt $startMergeRow) {
        $table.Cell($startMergeRow, 2).Merge($table.Cell($table.Rows.Count, 2))
    }

    # Final Save
    $word.ScreenUpdating = $true
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
    $doc.SaveAs([string]$outputPath)
    Write-Host "Summary generated successfully at: $outputPath"

} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    if ($null -ne $doc) { $doc.Close($false) }
    $word.Quit()
}
