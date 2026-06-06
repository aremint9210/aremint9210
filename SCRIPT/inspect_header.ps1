$baseDir = "C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026"
$templatePath = Join-Path $baseDir "GENERATE REPORT\TEMPLATE\00. EXECUTIVE SUMMARY - Copy.docx"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($templatePath)
    $t = $doc.Tables.Item(1)
    Write-Host "Header Row 1 Cells:"
    for ($i=1; $i -le 10; $i++) {
        try {
            $c = $t.Cell(1, $i)
            $txt = $c.Range.Text.Trim().Replace([char]7, '')
            Write-Host "Cell $i : $txt"
        } catch { break }
    }
    Write-Host "Header Row 2 Cells:"
    for ($i=1; $i -le 10; $i++) {
        try {
            $c = $t.Cell(2, $i)
            $txt = $c.Range.Text.Trim().Replace([char]7, '')
            Write-Host "Cell $i : $txt"
        } catch { break }
    }
} finally {
    if ($null -ne $doc) { $doc.Close($false) }
    $word.Quit()
}
