$csvPath = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\EXCELL MAPPING\Master List.csv"
$individualDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\OUTPUT FILE\INDIVIDUAL"
$groupOutputDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\OUTPUT FILE\GROUP"

if (-not (Test-Path $groupOutputDir)) { New-Item -ItemType Directory -Path $groupOutputDir }

$csv = Import-Csv $csvPath
$availableGroups = $csv | Where-Object { [string]::IsNullOrWhiteSpace($_."Group") -eq $false } | Select-Object -ExpandProperty "Group" -Unique | Sort-Object

Write-Host "Available Groups:" -ForegroundColor Yellow
$availableGroups | ForEach-Object { Write-Host " - $_" }

$targetGroup = Read-Host "`nEnter the Group name you want to merge"

if ($availableGroups -notcontains $targetGroup) {
    Write-Error "Group '$targetGroup' not found in CSV."
    Pause
    exit
}

$groupData = $csv | Where-Object { $_."Group" -eq $targetGroup }

function Kill-Word {
    Get-Process word -ErrorAction SilentlyContinue | Stop-Process -Force
}

# Conversion factor: 1 cm = 28.35 points
$topMarginPoints = 1.5 * 28.35
$bottomMarginPoints = 1.0 * 28.35

$outputPath = [string](Join-Path $groupOutputDir "group$($targetGroup).docx")

if (Test-Path $outputPath) {
    $choice = Read-Host "File 'group$($targetGroup).docx' already exists. Overwrite? (Y/N)"
    if ($choice -ne "Y" -and $choice -ne "y") {
        Write-Host "Operation cancelled."
        Pause
        exit
    }
    Remove-Item $outputPath -Force
}

$pageNames = $groupData | Select-Object -ExpandProperty "Output File Name" -Unique | Sort-Object { [int]($_ -replace "\D", "") }

Write-Host "`n--- Merging Group: $($targetGroup) ($($pageNames.Count) pages) ---" -ForegroundColor Cyan

try {
    Write-Host "  Starting Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0 
    
    Write-Host "  Creating new document..."
    $mainDoc = $word.Documents.Add()
    
    Write-Host "  Setting margins..."
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
                # Check if Word is still responsive
                $test = $word.Name 

                if (-not $first) {
                    $selection.InsertBreak(7) 
                }
                $selection.InsertFile($filePath)
                $first = $false
                Write-Host " [OK]" -ForegroundColor Green
            } catch {
                Write-Host " [FAILED]" -ForegroundColor Red
                $errMsg = $_.Exception.Message
                Write-Error "Error inserting $($pageName): $($errMsg)"
                
                if ($errMsg -like "*RPC server is unavailable*" -or $errMsg -like "*remote procedure call failed*") {
                    Write-Host "`n[CRITICAL] Word has crashed or disconnected. Cannot continue this group." -ForegroundColor Red
                    break
                }
            }
        } else {
            Write-Warning "  File not found: $($pageName).docx"
        }
    }

    # Only attempt to save if Word is still alive
    try {
        $test = $word.Name
        Write-Host "  Removing headers and footers..."
        foreach ($section in $mainDoc.Sections) {
            foreach ($header in $section.Headers) { try { $header.Range.Delete() } catch {} }
            foreach ($footer in $section.Footers) { try { $footer.Range.Delete() } catch {} }
        }

        Write-Host "  Saving group$($targetGroup).docx..."
        [object]$savePath = $outputPath
        [object]$format = 12 
        
        $mainDoc.SaveAs([ref]$savePath, [ref]$format)
        $mainDoc.Close([ref]0) 
        Write-Host "[SUCCESS] Saved group$($targetGroup).docx" -ForegroundColor Green
    } catch {
        Write-Error "  Could not save document because Word is no longer responsive."
    }
} catch {
    Write-Error "  Failed to process group $($targetGroup). Error: $($_.Exception.Message)"
} finally {
    if ($null -ne $word) {
        try { $word.Quit([ref]0) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    Kill-Word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`nDone!" -ForegroundColor Yellow
Pause
