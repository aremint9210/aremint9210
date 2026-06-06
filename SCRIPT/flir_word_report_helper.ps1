param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"

function Resolve-FlirBinPath {
    $registryPath = "Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{0A235AB1-0B24-4808-BC7F-B631374B5E13}\InprocServer32"
    $defaultBinPath = "C:\Program Files (x86)\FLIR Systems\FLIR Word Reporting\bin"
    try {
        $registeredDll = (Get-ItemProperty -Path $registryPath -ErrorAction Stop)."(default)"
        if ($registeredDll -and (Test-Path $registeredDll)) {
            return Split-Path -Path $registeredDll -Parent
        }
    } catch {}
    if (Test-Path $defaultBinPath) { return $defaultBinPath }
    throw "FLIR Word Reporting installation was not found."
}

function Apply-FieldOverrides {
    param($Document, $Manifest)
    $overrides = @{}
    foreach ($page in $Manifest.pages) {
        $properties = $page.field_overrides.PSObject.Properties
        if ($null -ne $properties) {
            foreach ($property in $properties) {
                $overrides[$property.Name] = [string]$property.Value      
            }
        }
    }
    $selection = $Document.Content
    foreach ($key in $overrides.Keys) {
        if ($key -eq "{{visual_image}}") { continue }
        $val = $overrides[$key]
        $selection.Find.ClearFormatting()
        $selection.Find.Replacement.ClearFormatting()
        $selection.Find.Execute($key, $false, $true, $false, $false, $false, $true, 1, $false, $val, 2) | Out-Null
    }
}

function Replace-TextWithImage {
    param($Document, $Manifest)
    foreach ($page in $Manifest.pages) {
        if ([string]::IsNullOrWhiteSpace($page.visual_path) -or -not (Test-Path $page.visual_path)) { continue }
        $range = $Document.Content
        $range.Find.ClearFormatting()
        if ($range.Find.Execute("{{visual_image}}")) {
            $imageRange = $range.Duplicate
            $imageRange.ParagraphFormat.Alignment = 1 # Center
            $range.Text = "" 
            $shape = $imageRange.InlineShapes.AddPicture($page.visual_path, $false, $true)
            if ($shape.Width -gt 400) {
                $ratio = 400 / $shape.Width
                $shape.Width = 400
                $shape.Height = $shape.Height * $ratio
            }
        }
    }
}

function Remove-PageBreaksAndMerge {
    param($Document)
    $range = $Document.Content
    $range.Find.ClearFormatting()
    $range.Find.Replacement.ClearFormatting()
    $range.Find.Execute("^m", $false, $false, $false, $false, $false, $true, 1, $false, "", 2) | Out-Null
    $range.Find.Execute("^b", $false, $false, $false, $false, $false, $true, 1, $false, "", 2) | Out-Null
    $paragraphs = $Document.Paragraphs
    for ($i = $paragraphs.Count; $i -gt 1; $i--) {
        $currPara = $paragraphs.Item($i)
        $prevPara = $paragraphs.Item($i - 1)
        if ($currPara.Range.Text.Trim().Length -eq 0 -and $currPara.Range.Tables.Count -eq 0) {
            if ($prevPara.Range.Tables.Count -gt 0) {
                if ($i -lt $paragraphs.Count) {
                    $nextPara = $paragraphs.Item($i + 1)
                    if ($nextPara.Range.Tables.Count -gt 0) { $currPara.Range.Delete() }
                }
            }
        }
    }
}

function New-ProReport {
    param($Manifest, $Word)
    $totalTemplatePages = 1
    try {
        $tempDoc = $Word.Documents.Open($Manifest.template_path, $false, $true)
        $totalTemplatePages = $tempDoc.ComputeStatistics(2)
        $tempDoc.Close($false)
    } catch {
        Write-Warning "Could not count template pages. Defaulting to 1."
    }
    $report = New-Object WizardClassLibrary.ProReport
    $report.Template = [string]$Manifest.template_path
    $report.IntroPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.SummaryPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.ReportPages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProReportPage]"
    $report.ReportProperties = New-Object "System.Collections.Generic.List[WizardClassLibrary.ReportProperty]"
    for ($i = 1; $i -le $totalTemplatePages; $i++) {
        $reportPage = New-Object WizardClassLibrary.ProReportPage
        $reportPage.TemplatePageNr = $i
        $reportPage.IRImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProIRImage]"
        $reportPage.DCImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProDCImage]"
        if ($i -eq 1) {
            $firstPage = $Manifest.pages[0]
            if (-not [string]::IsNullOrWhiteSpace($firstPage.ir_path)) {
                $irImage = New-Object WizardClassLibrary.ProIRImage
                $irImage.FileName = [string]$firstPage.ir_path
                $irImage.Sketch = 1
                $irImage.TextComments = New-Object "System.Collections.Generic.List[WizardClassLibrary.TextComment]"
                $reportPage.IRImages.Add($irImage)
            }
            if (-not [string]::IsNullOrWhiteSpace($firstPage.visual_path)) {
                $dcImage = New-Object WizardClassLibrary.ProDCImage
                $dcImage.FileName = [string]$firstPage.visual_path
                $dcImage.Sketch = 0
                $reportPage.DCImages.Add($dcImage)
            }
        }
        $report.ReportPages.Add($reportPage)
    }
    return $report
}

function Serialize-ProReport {
    param($Report, [string]$OutputPath)
    $serializer = New-Object System.Xml.Serialization.XmlSerializer([WizardClassLibrary.ProReport])
    $writer = New-Object System.IO.StreamWriter($OutputPath, $false, [System.Text.Encoding]::Unicode)
    try { $serializer.Serialize($writer, $Report) } finally { $writer.Dispose() }
}

$document = $null
try {
    $manifest = Get-Content -Raw -Path $ManifestPath | ConvertFrom-Json   
    $flirBinPath = Resolve-FlirBinPath
    $wizardDll = Join-Path $flirBinPath "WizardClassLibrary.dll"
    $wordAddinDll = Join-Path $flirBinPath "WordAddin_O2k7.dll"
    [Reflection.Assembly]::LoadFrom($wizardDll) | Out-Null
    [WizardClassLibrary.WordApp]::CheckAndRegisterWordAddin($wordAddinDll)
    if (-not [WizardClassLibrary.WordApp]::StartApp()) { throw "FLIR WordApp could not start." }
    
    $word = [WizardClassLibrary.WordApp]::GetWordApp()
    $analysisPath = Join-Path $env:TEMP ("analysis-" + [guid]::NewGuid().ToString() + ".xml")
    $reportXmlPath = Join-Path $env:TEMP ("proreport-" + [guid]::NewGuid().ToString() + ".xml")
    [WizardClassLibrary.WordApp]::AnalyzeTemplate([string]$manifest.template_path, $analysisPath, [ref]$false) | Out-Null
    
    $report = New-ProReport -Manifest $manifest -Word $word
    Serialize-ProReport -Report $report -OutputPath $reportXmlPath        
    [WizardClassLibrary.WordApp]::CreateReport($reportXmlPath, [ref]$false) | Out-Null
    
    $document = $word.ActiveDocument
    $word.DisplayAlerts = 0
    
    Apply-FieldOverrides -Document $document -Manifest $manifest
    Replace-TextWithImage -Document $document -Manifest $manifest
    Remove-PageBreaksAndMerge -Document $document
    
    $outputPath = [string]$manifest.output_path
    if (Test-Path $outputPath) { Remove-Item -Path $outputPath -Force }
    $document.SaveAs2($outputPath)
    Write-Output "Generated: $outputPath"
    
    $document.Close($false)
    $document = $null
} catch {
    Write-Error $_.Exception.Message
    exit 1
} finally {
    if ($null -ne $document) { try { $document.Close($false) } catch {} }
    try { [WizardClassLibrary.WordApp]::CloseWordApp() } catch {}
}
