param([Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference = 'Stop'
$j = Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$excel=$null; $book=$null; $word=$null; $doc=$null
try {
    if ($j.mode -eq 'read') {
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
        $excel.AutomationSecurity=3
        $book=$excel.Workbooks.Open($j.source.path,0,$true)
        $available=@($book.Worksheets | ForEach-Object { $_.Name })
        if ($available -notcontains [string]$j.source.sheet) {
            throw "Worksheet not found: $($j.source.sheet). Available worksheets: $($available -join ', '). No automatic fallback."
        }
        $sheet=$book.Worksheets.Item([string]$j.source.sheet)
        $records=@()
        foreach ($row in $j.source.data_rows) {
            $record=@{row=[int]$row}
            foreach ($field in $j.source.columns.PSObject.Properties) {
                $cell=$sheet.Range(([string]$field.Value)+[string]$row)
                if ($cell.MergeCells) { throw "Merged mapped cell: $($cell.Address())" }
                $shown=[string]$cell.Text
                if ($shown.StartsWith('#')) { throw "Error or clipped value: $($cell.Address())" }
                if ($field.Name -eq 'id') { $record[$field.Name]=$shown }
                else { $record[$field.Name]=$cell.Value2 }
            }
            $records+=,$record
        }
        ConvertTo-Json -InputObject $records -Depth 8 | Set-Content -LiteralPath $j.result -Encoding UTF8
    } elseif ($j.mode -eq 'verify') {
        $word=New-Object -ComObject Word.Application
        $word.Visible=$false; $word.DisplayAlerts=0; $word.AutomationSecurity=3
        $doc=$word.Documents.Open($j.candidate,$false,$false)
        $doc.Save(); $doc.Close(0)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null
        $doc=$null
        $doc=$word.Documents.Open($j.candidate,$false,$true)
        foreach ($body in $j.bodies) {
            if (-not $doc.Content.Text.Contains([string]$body)) { throw 'Word reopen: certificate body mismatch' }
        }
        if ($doc.Tables.Count -ne [int]$j.count) { throw 'Word reopen: signing table count mismatch' }
        $doc.Repaginate()
        if ($doc.ComputeStatistics(2) -lt ([int]$j.count+1)) { throw 'Word pagination: missing cover or certificate page' }
        $doc.ExportAsFixedFormat($j.pdf,17)
        if (-not (Test-Path -LiteralPath $j.pdf)) { throw 'Word PDF preview was not created' }
    } else { throw 'Unknown Office operation' }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    if ($null -ne $doc) { $doc.Close(0); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null }
    if ($null -ne $word) { $word.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null }
}
