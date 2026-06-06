param(
    [string]$SubFolder
)

$ErrorActionPreference = "Stop"

# Configuration
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$BaseDir = Join-Path $ScriptDir "..\OUTPUT FILE\INDIVIDUAL"
$TargetDir = Join-Path $BaseDir $SubFolder

if (-not (Test-Path $TargetDir)) {
    Write-Host "Error: Folder not found: $TargetDir" -ForegroundColor Red
    return
}

$Files = Get-ChildItem -Path $TargetDir -Filter "*.docx"
if ($Files.Count -eq 0) {
    Write-Host "No .docx files found in $TargetDir" -ForegroundColor Yellow
    return
}

Write-Host "Starting Cleanup in: $TargetDir" -ForegroundColor Cyan
Write-Host "Processing $($Files.Count) files..."

$Word = New-Object -ComObject Word.Application
$Word.Visible = $false
$Word.DisplayAlerts = 0

try {
    foreach ($File in $Files) {
        Write-Host "  Processing: $($File.Name)"
        $Doc = $Word.Documents.Open($File.FullName)

        # 1. Remove Headers and Footers
        foreach ($Section in $Doc.Sections) {
            foreach ($Header in $Section.Headers) {
                try { $Header.Range.Delete() } catch {}
            }
            foreach ($Footer in $Section.Footers) {
                try { $Footer.Range.Delete() } catch {}
            }
        }

        # 2. Replace ALL humidity patterns (e.g., 75.0%, 82.0% -> 75%, 82%)
        # Using Word Wildcards: ([0-9]@) matches one or more digits
        $FindText = "([0-9]@).0%"
        $ReplaceText = "\1%"
        
        $Selection = $Word.Selection
        $Selection.Find.ClearFormatting()
        $Selection.Find.Replacement.ClearFormatting()
        
        # Use MatchWildcards for pattern matching
        $Selection.Find.Execute(
            $FindText,      # FindText
            $false,         # MatchCase
            $false,         # MatchWholeWord
            $true,          # MatchWildcards
            $false,         # MatchSoundsLike
            $false,         # MatchAllWordForms
            $true,          # Forward
            1,              # wdFindContinue
            $false,         # Format
            $ReplaceText,   # ReplaceWith
            2               # wdReplaceAll
        ) | Out-Null

        $Doc.Save()
        $Doc.Close()
        Write-Host "  [DONE]" -ForegroundColor Green
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    $Word.Quit()
    # Cleanup COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`nCleanup Process Complete!" -ForegroundColor Green
