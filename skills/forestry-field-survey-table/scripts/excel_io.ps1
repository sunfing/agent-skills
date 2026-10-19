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
        foreach($d in $j.sheets){
            $s=$book.Worksheets.Item($d.name);$n=$d.rows.Count;$nc=$d.rows[0].Count
            if($n -lt 1){throw 'No rows'}
            $last=$n+3;$oldlast=$s.UsedRange.Rows.Count
            $height=$s.Rows.Item(4).RowHeight
            $headerHeight=$s.Rows.Item(3).RowHeight
            $s.Range($s.Cells.Item(4,1),$s.Cells.Item([Math]::Max($last,$oldlast),$nc)).ClearContents()
            if($last -gt 4){$s.Range($s.Cells.Item(4,1),$s.Cells.Item(4,$nc)).Copy($s.Range($s.Cells.Item(5,1),$s.Cells.Item($last,$nc)))}
            if($oldlast -gt $last){$s.Range($s.Cells.Item($last+1,1),$s.Cells.Item($oldlast,$nc)).Clear()}
            $s.Range($s.Cells.Item(4,1),$s.Cells.Item($last,$nc)).RowHeight=$height
            for($r=0;$r -lt $n;$r++){for($c=0;$c -lt $nc;$c++){
                $v=$d.rows[$r][$c];$cell=$s.Cells.Item($r+4,$c+1)
                if($null -eq $v){continue}
                if($v -is [string]){$cell.NumberFormat='@';$cell.Value2=[string]$v}else{$cell.Value2=[double]$v}
            }}
            $s.Cells.Item(1,1).Value2=[string]$d.title
            # Fit headers and data only, excluding merged title/subtitle rows.
            # Excel AutoFit ignores wrapped cells; restore wrapping afterwards.
            for($c=1;$c -le $nc;$c++){
                $range=$s.Range($s.Cells.Item(3,$c),$s.Cells.Item($last,$c))
                $wrap=$range.WrapText
                $wraps=@()
                if($null -eq $wrap){for($r=3;$r -le $last;$r++){$wraps+= [bool]$s.Cells.Item($r,$c).WrapText}}
                $range.WrapText=$false
                $range.Columns.AutoFit() | Out-Null
                if($null -eq $wrap){for($r=3;$r -le $last;$r++){$s.Cells.Item($r,$c).WrapText=$wraps[$r-3]}}
                else{$range.WrapText=$wrap}
            }
            if($d.name -eq '外业调查表'){
                $s.Columns.Item(4).ColumnWidth=10
                $s.Columns.Item(14).ColumnWidth=10
                $s.Columns.Item(37).ColumnWidth=15
            }
            # Wrapping can change row height; keep the template heights.
            $s.Rows.Item(3).RowHeight=$headerHeight
            $s.Range($s.Cells.Item(4,1),$s.Cells.Item($last,$nc)).RowHeight=$height
            $s.PageSetup.PrintArea=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($last,$nc)).Address()
        }
        $book.Worksheets.Item(1).Activate()
        $book.SaveAs($j.output,51) | Out-Null
        $book.Close($false);$book=$null
        $book=$excel.Workbooks.Open($j.output,0,$true)
        if($book.Worksheets.Count -ne 2){throw 'Expected source and field sheets'}
        foreach($d in $j.sheets){
            $s=$book.Worksheets.Item($d.name)
            if([string]$s.Cells.Item(1,1).Value2 -cne $d.title){throw 'Title mismatch'}
            if($d.name -eq '外业调查表'){
                foreach($pair in @(@(4,10),@(14,10),@(37,15))){
                    if([Math]::Abs($s.Columns.Item($pair[0]).ColumnWidth-$pair[1]) -gt 0.1){throw 'Fixed column width mismatch'}
                }
            }
            for($r=0;$r -lt $d.rows.Count;$r++){for($c=0;$c -lt $d.rows[$r].Count;$c++){
                $expected=$d.rows[$r][$c];$cell=$s.Cells.Item($r+4,$c+1);$actual=$cell.Value2
                if($cell.HasFormula){throw 'Source and field sheets must be independent'}
                if($null -eq $expected){if($null -ne $actual){throw "Unexpected value $($d.name) $r,$c"}}
                elseif($expected -is [string]){if([string]$actual -cne $expected){throw "Text mismatch $r,$c"}}
                elseif([Math]::Abs([double]$actual-[double]$expected) -gt 1e-10){throw "Number mismatch $r,$c"}
            }}
        }
    } else { throw 'Unknown operation' }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
}
