param([Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference = 'Stop'
$j = Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    if ($j.mode -eq 'read') {
        $book = $excel.Workbooks.Open($j.input, 0, $true)
        if ($j.sheet) { $s = $book.Worksheets.Item($j.sheet) }
        elseif ($book.Worksheets.Count -eq 1) { $s = $book.Worksheets.Item(1) }
        else { throw 'Multiple worksheets: specify --sheet explicitly' }
        $used = $s.UsedRange
        if ($used.Row -ne 1 -or $used.Column -ne 1) { throw 'Headers must start at A1' }
        $headers = @(); $rows = @()
        for ($c=1; $c -le $used.Columns.Count; $c++) { $headers += [string]$s.Cells.Item(1,$c).Value2 }
        for ($r=2; $r -le $used.Rows.Count; $r++) {
            $row = @()
            for ($c=1; $c -le $used.Columns.Count; $c++) {
                $cell=$s.Cells.Item($r,$c); $value=$cell.Value2
                # Preserve identifiers formatted with leading zeroes; numbers stay numeric.
                if ($headers[$c-1] -in @('LIN_BAN','XIAO_BAN','小班编')) { if ($null -ne $value) { $value=[string]$cell.Text } }
                $row += $value
            }
            $rows += ,$row
        }
        @{ headers=$headers; rows=$rows } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $j.result -Encoding UTF8
    } elseif ($j.mode -eq 'write') {
        $book=$excel.Workbooks.Open($j.template,0,$true)
        $s=$book.Worksheets.Item('A.1')
        $n=$j.rows.Count
        if ($n -lt 1) { throw 'No rows' }
        $height=$s.Rows.Item(4).RowHeight
        $headerHeight=$s.Rows.Item(3).RowHeight
        if ($n -gt 1) { $s.Range('A4:AG4').Copy($s.Range("A5:AG$($n+3)")) }
        $s.Range("A4:AG$($n+3)").RowHeight=$height
        $s.Range("E4:G$($n+3)").NumberFormat='@'
        for ($r=0; $r -lt $n; $r++) {
            for ($c=0; $c -lt 33; $c++) {
                $v=$j.rows[$r][$c]
                if ($null -ne $v) {
                    if ($v -is [string]) { $s.Cells.Item($r+4,$c+1).NumberFormat='@'; $s.Cells.Item($r+4,$c+1).Value2=[string]$v }
                    else { $s.Cells.Item($r+4,$c+1).Value2=[double]$v }
                }
            }
        }
        $s.Cells.Interior.Pattern=-4142
        $s.Range("AA4:AA$($n+3)").NumberFormat='0"°"'
        # Fit only headers and data, excluding the merged title. Temporarily
        # disable wrapping so long species names contribute their full width.
        for ($c=1; $c -le 33; $c++) {
            $range=$s.Range($s.Cells.Item(3,$c),$s.Cells.Item($n+3,$c))
            $wrap=@(); for ($r=3; $r -le $n+3; $r++) { $wrap += $s.Cells.Item($r,$c).WrapText }
            $range.WrapText=$false
            $range.Columns.AutoFit() | Out-Null
            $s.Columns.Item($c).ColumnWidth=[Math]::Min(255,$s.Columns.Item($c).ColumnWidth+1)
            for ($r=3; $r -le $n+3; $r++) { $s.Cells.Item($r,$c).WrapText=$wrap[$r-3] }
        }
        $s.Rows.Item(3).RowHeight=$headerHeight
        $s.Range("A4:AG$($n+3)").RowHeight=$height
        $widths=@(); for ($c=1; $c -le 33; $c++) { $widths += $s.Columns.Item($c).ColumnWidth }
        $s.Range('A1').Select() | Out-Null
        $book.SaveAs($j.output,51) | Out-Null
        $book.Close($false); $book=$null
        # Reopen with normal loading; do not request repair or use DisplayAlerts to accept repair.
        $book=$excel.Workbooks.Open($j.output,0,$true)
        if ($book.Worksheets.Count -ne 1) { throw 'Unexpected worksheets after save' }
        $s=$book.Worksheets.Item('A.1')
        for ($r=0; $r -lt $n; $r++) {
            for ($c=0; $c -lt 33; $c++) {
                $expected=$j.rows[$r][$c]; $actual=$s.Cells.Item($r+4,$c+1).Value2
                if ($null -eq $expected) { if ($null -ne $actual) { throw "Unexpected value at $r,$c" } }
                elseif ($expected -is [string]) { if ([string]$actual -cne $expected) { throw "Text mismatch at $r,$c" } }
                elseif ([Math]::Abs([double]$actual-[double]$expected) -gt 1e-10) { throw "Number mismatch at $r,$c" }
            }
        }
        for ($c=1; $c -le 33; $c++) { if ($s.Columns.Item($c).ColumnWidth -ne $widths[$c-1]) { throw 'Column width changed' } }
        if ($s.Rows.Item(3).RowHeight -ne $headerHeight) { throw 'Header height changed' }
        if ($s.Rows.Item(4).RowHeight -ne $height -or $s.Range('A3').Borders.Item(9).LineStyle -eq -4142) { throw 'Excel failed to preserve template styles' }
    } else { throw 'Unknown operation' }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
}
