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
    } catch {
    }

    if (Test-Path $defaultBinPath) {
        return $defaultBinPath
    }

    throw "FLIR Word Reporting installation was not found."
}

function Normalize-Label {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    $normalized = $Text.Replace([char]7, " ").Replace("`r", " ").Replace("`n", " ")
    $normalized = [regex]::Replace($normalized, "\s+", " ").Trim().ToLowerInvariant()
    return $normalized.TrimEnd(":")
}

function Get-CellText {
    param($Cell)

    return [string]$Cell.Range.Text
}

function Set-CellText {
    param(
        $Cell,
        [string]$Value
    )

    $range = $Cell.Range
    $range.End = $range.End - 1
    $range.Text = $Value
}

function Apply-FieldOverrides {
    param(
        $Document,
        $Manifest
    )

    # We use global find and replace for text placeholders
    $findText = ""
    $replaceText = ""
    
    # Merge all overrides from all pages (usually there is only 1 page in individual templates)
    $overrides = @{}
    foreach ($page in $Manifest.pages) {
        $properties = $page.field_overrides.PSObject.Properties
        if ($null -ne $properties) {
            foreach ($property in $properties) {
                # Keep original key (e.g. {{substation_name}})
                $overrides[$property.Name] = [string]$property.Value
            }
        }
    }

    $selection = $Document.Content
    foreach ($key in $overrides.Keys) {
        if ($key -eq "{{visual_image}}") { continue } # Handled separately
        
        $val = $overrides[$key]
        $selection.Find.ClearFormatting()
        $selection.Find.Replacement.ClearFormatting()
        $selection.Find.Execute(
            $key,              # FindText
            $true,             # MatchCase
            $true,             # MatchWholeWord
            $false,            # MatchWildcards
            $false,            # MatchSoundsLike
            $false,            # MatchAllWordForms
            $true,             # Forward
            1,                 # Wrap (wdFindContinue)
            $false,            # Format
            $val,              # ReplaceWith
            2                  # Replace (wdReplaceAll)
        ) | Out-Null
    }
}

function Replace-TextWithImage {
    param(
        $Document,
        $Manifest
    )

    foreach ($page in $Manifest.pages) {
        if ($null -eq $page.visual_path -or -not (Test-Path $page.visual_path)) {
            continue
        }

        $range = $Document.Content
        $range.Find.ClearFormatting()
        if ($range.Find.Execute("{{visual_image}}")) {
            $imageRange = $range.Duplicate
            $range.Text = "" # Clear the placeholder text
            
            # Add the picture at the range
            $shape = $imageRange.InlineShapes.AddPicture($page.visual_path, $false, $true)
            
            # Optional: Resize image to fit nicely (e.g., max width 400 points)
            if ($shape.Width -gt 400) {
                $ratio = 400 / $shape.Width
                $shape.Width = 400
                $shape.Height = $shape.Height * $ratio
            }
        }
    }
}

function New-ProReport {
    param($Manifest)

    $report = New-Object WizardClassLibrary.ProReport
    $report.Template = [string]$Manifest.template_path
    $report.IntroPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.SummaryPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.ReportPages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProReportPage]"
    $report.ReportProperties = New-Object "System.Collections.Generic.List[WizardClassLibrary.ReportProperty]"

    foreach ($page in $Manifest.pages) {
        $reportPage = New-Object WizardClassLibrary.ProReportPage
        $reportPage.TemplatePageNr = [int]$page.template_page_no
        $reportPage.IRImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProIRImage]"
        $reportPage.DCImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProDCImage]"

        $irImage = New-Object WizardClassLibrary.ProIRImage
        $irImage.FileName = [string]$page.ir_path
        $irImage.Sketch = 1
        $irImage.TextComments = New-Object "System.Collections.Generic.List[WizardClassLibrary.TextComment]"
        $reportPage.IRImages.Add($irImage)

        $dcImage = New-Object WizardClassLibrary.ProDCImage
        $dcImage.FileName = [string]$page.visual_path
        $dcImage.Sketch = 0
        $reportPage.DCImages.Add($dcImage)

        $report.ReportPages.Add($reportPage)
    }

    return $report
}

function Serialize-ProReport {
    param(
        $Report,
        [string]$OutputPath
    )

    $serializer = New-Object System.Xml.Serialization.XmlSerializer([WizardClassLibrary.ProReport])
    $writer = New-Object System.IO.StreamWriter($OutputPath, $false, [System.Text.Encoding]::Unicode)
    try {
        $serializer.Serialize($writer, $Report)
    } finally {
        $writer.Dispose()
    }
}

function Validate-TemplatePages {
    param(
        [xml]$AnalysisXml,
        $Manifest
    )

    $templatePages = @($AnalysisXml.TemplateFile.TemplatePages.TemplatePage)
    if ($templatePages.Count -eq 0) {
        throw "FLIR template analysis found no template pages."
    }

    $maxPage = 0
    foreach ($page in $Manifest.pages) {
        $maxPage = [Math]::Max($maxPage, [int]$page.template_page_no)
    }

    if ($maxPage -gt $templatePages.Count) {
        throw "Manifest references template page $maxPage, but FLIR only reported $($templatePages.Count) template pages."
    }
}

$analysisPath = Join-Path $env:TEMP ("tnb-flir-analysis-" + [guid]::NewGuid().ToString() + ".xml")
$reportXmlPath = Join-Path $env:TEMP ("tnb-flir-proreport-" + [guid]::NewGuid().ToString() + ".xml")
$document = $null
$word = $null
$success = $false

try {
    if (-not (Test-Path $ManifestPath)) {
        throw "Manifest file not found: $ManifestPath"
    }

    $manifest = Get-Content -Raw -Path $ManifestPath | ConvertFrom-Json
    $flirBinPath = Resolve-FlirBinPath
    $wizardDll = Join-Path $flirBinPath "WizardClassLibrary.dll"
    $wordAddinDll = Join-Path $flirBinPath "WordAddin_O2k7.dll"

    if (-not (Test-Path $wizardDll)) {
        throw "WizardClassLibrary.dll not found: $wizardDll"
    }
    if (-not (Test-Path $wordAddinDll)) {
        throw "WordAddin_O2k7.dll not found: $wordAddinDll"
    }

    [Reflection.Assembly]::LoadFrom($wizardDll) | Out-Null
    [WizardClassLibrary.WordApp]::CheckAndRegisterWordAddin($wordAddinDll)

    $started = [WizardClassLibrary.WordApp]::StartApp()
    if (-not $started) {
        throw "FLIR WordApp could not start Word."
    }

    [WizardClassLibrary.WordApp]::HideApp()
    
    # Disable alerts immediately to prevent hangs on dialog boxes
    $word = [WizardClassLibrary.WordApp]::GetWordApp()
    if ($null -ne $word) {
        $word.DisplayAlerts = 0 # wdAlertsNone
    }

    $analysisErrored = $false
    $analysisOk = [WizardClassLibrary.WordApp]::AnalyzeTemplate(
        [string]$manifest.template_path,
        $analysisPath,
        [ref]$analysisErrored
    )
    if (-not $analysisOk -or -not (Test-Path $analysisPath)) {
        throw "WordApp.AnalyzeTemplate failed. Analysis XML was not created."
    }

    [xml]$analysisXml = Get-Content -Raw -Path $analysisPath
    Validate-TemplatePages -AnalysisXml $analysisXml -Manifest $manifest

    $report = New-ProReport -Manifest $manifest
    Serialize-ProReport -Report $report -OutputPath $reportXmlPath

    $createErrored = $false
    $createOk = [WizardClassLibrary.WordApp]::CreateReport($reportXmlPath, [ref]$createErrored)
    $word = [WizardClassLibrary.WordApp]::GetWordApp()
    if (-not $createOk -or $null -eq $word -or $word.Documents.Count -lt 1) {
        throw "WordApp.CreateReport did not produce a Word document."
    }

    $word.DisplayAlerts = 0
    $document = $word.ActiveDocument

    # Remove Headers and Footers from all sections
    foreach ($section in $document.Sections) {
        foreach ($h in $section.Headers) {
            try { $h.Range.Delete() } catch {}
        }
        foreach ($f in $section.Footers) {
            try { $f.Range.Delete() } catch {}
        }
    }

    Replace-TextWithImage -Document $document -Manifest $manifest
    Apply-FieldOverrides -Document $document -Manifest $manifest

    $outputPath = [string]$manifest.output_path
    $outputDir = Split-Path -Path $outputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    if (Test-Path $outputPath) {
        Remove-Item -Path $outputPath -Force
    }

    $document.SaveAs2($outputPath)
    $success = $true
    Write-Output "Generated FLIR report: $outputPath"
    if ($analysisErrored -or $createErrored) {
        Write-Output "FLIR reported warning flags during automation; output document was still generated."
    }
} catch {
    $details = @(
        $_.Exception.Message
        "Template analysis XML: $analysisPath"
        "ProReport XML: $reportXmlPath"
    ) -join [Environment]::NewLine
    Write-Error $details
    exit 1
} finally {
    if ($document -ne $null) {
        try {
            $document.Close($false)
        } catch {
        }
    }

    try {
        [WizardClassLibrary.WordApp]::CloseWordApp()
    } catch {
    }

    if ($success) {
        Remove-Item -Path $analysisPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $reportXmlPath -Force -ErrorAction SilentlyContinue
    }
}
