$root = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026"
$scriptDir = Join-Path $root "GENERATE REPORT\SCRIPT"

if (!(Test-Path $scriptDir)) { New-Item -ItemType Directory -Force -Path $scriptDir }

# Get all files in root that are NOT .bat files
$filesToMove = Get-ChildItem -Path $root -File | Where-Object { $_.Extension -ne ".bat" }

foreach ($file in $filesToMove) {
    Write-Host "Moving $($file.Name)..."
    Move-Item -Path $file.FullName -Destination $scriptDir -Force -ErrorAction SilentlyContinue
}

Write-Host "Cleanup complete."
