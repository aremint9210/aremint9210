$baseDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026"
$templatePath = Join-Path $baseDir "GENERATE REPORT\TEMPLATE\00. EXECUTIVE SUMMARY - Copy.docx"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($templatePath)
    Write-Host "Tables count: $($doc.Tables.Count)"
    if ($doc.Tables.Count -ge 1) {
        $t = $doc.Tables.Item(1)
        Write-Host "Table 1 rows: $($t.Rows.Count)"
        Write-Host "Table 1 columns: $($t.Columns.Count)"
    }
} finally {
    if ($null -ne $doc) { $doc.Close($false) }
    $word.Quit()
}
