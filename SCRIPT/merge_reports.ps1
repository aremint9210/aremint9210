param(
    [string]$CsvPath = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\EXCELL MAPPING\Master List.csv",
    [string]$IndividualDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\OUTPUT FILE\INDIVIDUAL",
    [string]$GroupOutputDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\OUTPUT FILE\GROUP"
)

if (-not (Test-Path $GroupOutputDir)) { New-Item -ItemType Directory -Path $GroupOutputDir }

$csv = Import-Csv $csvPath
$groups = $csv | Where-Object { [string]::IsNullOrWhiteSpace($_."Group") -eq $false } | Group-Object "Group"

function Kill-Word {
    Get-Process word -ErrorAction SilentlyContinue | Stop-Process -Force
}

# Conversion factor: 1 cm = 28.35 points
$topMarginPoints = 1.5 * 28.35
$bottomMarginPoints = 1.0 * 28.35

Write-Host "Found $($groups.Count) groups in CSV."

foreach ($group in $groups) {
    $groupName = $group.Name
    $outputPath = [string](Join-Path $groupOutputDir "group$($groupName).docx")
    
    if (Test-Path $outputPath) {
        Write-Host "--- Skipping Group: $($groupName) (Already exists) ---" -ForegroundColor Gray
        continue
    }

    $pageNames = $group.Group | Select-Object -ExpandProperty "Output File Name" -Unique | Sort-Object { [int]($_ -replace "\D", "") }
    
    Write-Host "`n--- Merging Group: $($groupName) ($($pageNames.Count) pages) ---" -ForegroundColor Cyan
    
    try {
        Write-Host "  Starting Word..."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true # Make visible so you can see if it's stuck on a prompt
        $word.DisplayAlerts = 0 
        
        Write-Host "  Creating new document..."
        $mainDoc = $word.Documents.Add()
        
        Write-Host "  Setting margins (Top 1.5cm, Bottom 1cm)..."
        $mainDoc.PageSetup.TopMargin = $topMarginPoints
        $mainDoc.PageSetup.BottomMargin = $bottomMarginPoints
        $mainDoc.PageSetup.LeftMargin = 2.0 * 28.35
        $mainDoc.PageSetup.RightMargin = 2.0 * 28.35

        $selection = $word.Selection

        $first = $true
        foreach ($pageName in $pageNames) {
            $filePath = Join-Path $individualDir "$pageName.docx"
            
            if (Test-Path $filePath) {
                Write-Host "  Inserting $pageName..." -NoNewline
                try {
                    if (-not $first) {
                        $selection.InsertBreak(7) 
                    }
                    $selection.InsertFile($filePath)
                    $first = $false
                    Write-Host " [OK]" -ForegroundColor Green
                } catch {
                    Write-Host " [FAILED]" -ForegroundColor Red
                    Write-Error "Error inserting $($pageName): $($_.Exception.Message)"
                }
            } else {
                Write-Warning "  File not found: $($pageName).docx"
            }
        }

        Write-Host "  Removing headers and footers..."
        foreach ($section in $mainDoc.Sections) {
            foreach ($header in $section.Headers) { $header.Range.Delete() }
            foreach ($footer in $section.Footers) { $footer.Range.Delete() }
        }

        Write-Host "  Saving group$($groupName).docx..."
        [object]$savePath = $outputPath
        [object]$format = 12 
        
        $mainDoc.SaveAs([ref]$savePath, [ref]$format)
        $mainDoc.Close([ref]0) 
        Write-Host "[SUCCESS] Saved group$($groupName).docx" -ForegroundColor Green
    } catch {
        Write-Error "  Failed to process group $($groupName). Error: $($_.Exception.Message)"
    } finally {
        if ($null -ne $word) {
            try { $word.Quit([ref]0) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        Kill-Word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-Host "`nAll groups processed!" -ForegroundColor Yellow
Pause
