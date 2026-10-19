param([Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference='Stop'
$j=Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$e=$null;$b=$null;$t=$null;$fontChecks=@()
try {
 $e=New-Object -ComObject Excel.Application
 $e.Visible=$false;$e.DisplayAlerts=$false;$e.EnableEvents=$false;$e.ScreenUpdating=$false;$e.AutomationSecurity=3
 $b=$e.Workbooks.Open($j.source,0,$true)
 $t=$e.Workbooks.Open($j.template,0,$true)
 $e.Calculation=-4135
 foreach($ws in @($b.Worksheets)){if($ws.Name -in $j.replace_names){$ws.Delete()}}
 $anchor=$b.Worksheets.Item($b.Worksheets.Count)
 foreach($d in $j.sheets){$s=$b.Worksheets.Add([Type]::Missing,$anchor);$s.Name=$d.name;$anchor=$s}
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name);$ref=$t.Worksheets.Item($d.name)
  $nc=[int]$d.cols;$head=[int]$d.head;$last=[int]$d.last
  [void]$ref.Range($ref.Cells.Item(1,1),$ref.Cells.Item($head,$nc)).Copy($s.Range('A1'))
  if($d.name -eq 'B.5'){$s.UsedRange.UnMerge();$s.UsedRange.ClearContents()}
  for($c=1;$c -le $nc;$c++){
   $sc=if($d.name -eq 'B.5' -and $c -gt 3){4}else{$c}
   $s.Columns.Item($c).ColumnWidth=$ref.Columns.Item($sc).ColumnWidth
  }
  for($r=1;$r -le $head;$r++){$s.Rows.Item($r).RowHeight=$ref.Rows.Item($r).RowHeight}
  if($last -gt $head){
   $sr=if($d.name -eq 'A.1'){9}else{$head+1}
   for($r=$head+1;$r -le $last;$r++){
    [void]$ref.Range($ref.Cells.Item($sr,1),$ref.Cells.Item($sr,$nc)).Copy($s.Cells.Item($r,1))
    $s.Rows.Item($r).RowHeight=$ref.Rows.Item($sr).RowHeight
   }
   $body=$s.Range($s.Cells.Item($head+1,1),$s.Cells.Item($last,$nc))
   $body.UnMerge();$body.ClearContents();$body.NumberFormat='0.########;-0.########;;@'
  }
  foreach($p in $d.cells.PSObject.Properties){
   if($null -eq $p.Value){continue}
   $cell=$s.Range($p.Name)
   # Keep each merged cell's numeric/date format independent.
   if($p.Value -is [string] -and $p.Value.StartsWith('=')){$cell.MergeArea.NumberFormat='0.########';$cell.Formula=$p.Value}
   elseif($p.Value -is [string]){$cell.Value2=$p.Value}
   else{$cell.Value2=[double]$p.Value}
  }
  foreach($m in $d.merges){[void]$s.Range($m).Merge()}
  foreach($f in $d.formats){$s.Range($f[0]).NumberFormat=$f[1]}
  foreach($p in $d.heights.PSObject.Properties){$s.Rows.Item([int]$p.Name).RowHeight=[double]$p.Value}
  if(-not $d.preserve_body){
   $first=if($d.name -in @('A.1','B.6')){3}else{4}
   $table=$s.Range($s.Cells.Item($first,1),$s.Cells.Item([int]$d.table_last,$nc))
   $table.Borders.LineStyle=1;$table.Borders.Weight=2
  }
  $all=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($last,$nc));$all.Interior.Pattern=-4142
  foreach($prop in @('PaperSize','Orientation','Zoom','FitToPagesWide','FitToPagesTall','LeftMargin','RightMargin','TopMargin','BottomMargin','HeaderMargin','FooterMargin','CenterHorizontally','CenterVertically','PrintTitleRows','PrintTitleColumns','LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter')){$s.PageSetup.$prop=$ref.PageSetup.$prop}
  $s.PageSetup.PrintArea=$all.Address()
  $addresses=if($d.name -eq '现状调查表'){@('A1','D3','D7')}else{@('A1')}
  foreach($address in $addresses){
   $font=$ref.Range($address).Font
   $fontChecks+=@{sheet=$d.name;address=$address;name=$font.Name;size=$font.Size;bold=$font.Bold}
  }
  $b.Activate();$s.Activate();$e.ActiveWindow.View=2;$e.ActiveWindow.Zoom=100
 }
 $order=@($j.input_names)+@($j.sheets | ForEach-Object {$_.name})
 for($i=$order.Count-1;$i -ge 0;$i--){$s=$b.Worksheets.Item($order[$i]);if($s.Index -ne 1){$s.Move($b.Worksheets.Item(1))}}
 # Named-style font IDs can collide when copying between native workbooks.
 # Apply template fonts explicitly after all copies, before export/save.
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  foreach($f in $d.font_formats){$font=$s.Range($f[0]).Font;$font.Name=$f[1];$font.Size=[double]$f[2];$font.Bold=[bool]$f[3];$font.Italic=[bool]$f[4]}
 }
 # Native AutoFit cannot size merged title cells. Measure the same text/font
 # in a disposable unmerged cell, then grow only the title's final row.
 $measure=$b.Worksheets.Add()
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  $titles=if($d.name -eq '现状调查表'){@('A1','A16')}else{@('A1')}
  foreach($address in $titles){
   $cell=$s.Range($address);$area=$cell.MergeArea;$probe=$measure.Range('A1')
   $probe.Value2=$cell.Value2;$probe.WrapText=$true
   foreach($prop in @('Name','Size','Bold','Italic')){$probe.Font.$prop=$cell.Font.$prop}
   $lo=0.1;$hi=255.0
   for($i=0;$i -lt 12;$i++){
    $mid=($lo+$hi)/2;$probe.ColumnWidth=$mid
    if($probe.Width -gt $area.Width){$hi=$mid}else{$lo=$mid}
   }
   $probe.ColumnWidth=$lo;[void]$probe.EntireRow.AutoFit()
   $cell.WrapText=$true;$cell.ShrinkToFit=$false
   $extra=$probe.RowHeight-$area.Height
   if($extra -gt 0){$row=$s.Rows.Item($area.Row+$area.Rows.Count-1);$row.RowHeight=$row.RowHeight+$extra}
  }
 }
 $measure.Delete()
 $t.Close($false);$t=$null;$e.Calculation=-4105;$e.CalculateFullRebuild()
 foreach($d in $j.sheets){$b.Worksheets.Item($d.name).ExportAsFixedFormat(0,(Join-Path $j.stage "$($d.name).pdf"))}
 $b.Worksheets.Item('内业处理表').Activate()
 $b.SaveAs($j.candidate,51);$b.Close($false);$b=$null
 $b=$e.Workbooks.Open($j.candidate,0,$true)
 foreach($check in $fontChecks){
  $font=$b.Worksheets.Item($check.sheet).Range($check.address).Font
  if($font.Name -ne $check.name -or $font.Size -ne $check.size -or $font.Bold -ne $check.bold){throw "Template font changed: $($check.sheet)!$($check.address)"}
 }
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  $range=$s.Range($s.Cells.Item(1,1),$s.Cells.Item([int]$d.last,[int]$d.cols))
  $cached=$range.Value2;$formulas=$range.Formula
  foreach($p in $d.checks.PSObject.Properties){
   $match=[regex]::Match($p.Name,'^([A-Z]+)([0-9]+)$');$c=0
   foreach($ch in $match.Groups[1].Value.ToCharArray()){$c=$c*26+[int]$ch-64}
   $v=$cached[[int]$match.Groups[2].Value,$c]
   if($v -isnot [double] -or [Math]::Abs($v-[double]$p.Value) -gt 0.00000001){throw "Excel total mismatch: $($d.name)!$($p.Name): $v vs $($p.Value)"}
  }
  foreach($p in $d.cells.PSObject.Properties){
   if($p.Value -is [string] -and $p.Value.StartsWith('=')){
    $match=[regex]::Match($p.Name,'^([A-Z]+)([0-9]+)$');$c=0
    foreach($ch in $match.Groups[1].Value.ToCharArray()){$c=$c*26+[int]$ch-64}
    $actual=$formulas[[int]$match.Groups[2].Value,$c] -replace "'(?<name>A\.1|B\.7|内业处理表)'!",'${name}!'
    $expected=$p.Value -replace "'(?<name>A\.1|B\.7|内业处理表)'!",'${name}!'
    if($actual -ne $expected){throw "Formula mismatch: $($d.name)!$($p.Name)"}
   }
  }
 }
 Write-Output 'Excel save, reopen, formula and subtotal verification passed'
}catch{
 Write-Output "Sheet: $($d.name); cell: $($p.Name)"
 Write-Output $_.InvocationInfo.PositionMessage
 throw
}finally{
 if($t){$t.Close($false)};if($b){$b.Close($false)}
 if($e){$e.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($e)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
