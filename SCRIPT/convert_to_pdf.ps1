param(
    [string]$SourceFolder,
    [string]$DestinationFolder
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $SourceFolder)) {
    Write-Error "Source folder does not exist: $SourceFolder"
    exit 1
}

if (-not (Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder | Out-Null
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

# Word Constants
$wdExportFormatPDF = 17
$wdExportOptimizeForPrint = 0 # High quality for printing (vs 1 for on-screen)
$wdExportAllDocument = 0
$wdExportDocumentContent = 0
$wdExportCreateWordBookmarks = 1

try {
    $files = Get-ChildItem -Path $SourceFolder -Filter "*.docx"
    Write-Host "Found $($files.Count) Word documents to convert." -ForegroundColor Yellow

    foreach ($file in $files) {
        $pdfPath = Join-Path $DestinationFolder ($file.BaseName + ".pdf")
        
        Write-Host "Converting: $($file.Name)..." -NoNewline
        
        $doc = $word.Documents.Open($file.FullName, $false, $true) # Open ReadOnly
        
        # Disable image compression for this document to preserve resolution (300+ PPI)
        # Note: Specific PPI settings are not directly exposed in Word's COM object, 
        # but are controlled via 'OptimizeFor' and disabling compression.
        try { $doc.DoNotCompressPictures = $true } catch { }
        
        # Export as PDF with high quality and advanced settings
        $doc.ExportAsFixedFormat(
            $pdfPath,
            $wdExportFormatPDF,
            $false,                   # OpenAfterExport
            $wdExportOptimizeForPrint,
            $wdExportAllDocument,
            1,                        # From
            1,                        # To
            $wdExportDocumentContent,
            $true,                    # IncludeDocProps
            $true,                    # KeepIRM
            $wdExportCreateWordBookmarks,
            $true,                    # DocStructureTags
            $true,                    # BitmapMissingFonts
            $false                    # UseISO19005_1 (Set to $true for PDF/A)
        )
        
        $doc.Close($false)
        Write-Host " [DONE]" -ForegroundColor Green
    }
} catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    # Clean up any stuck Word processes
    Get-Process word -ErrorAction SilentlyContinue | Stop-Process -Force
}
