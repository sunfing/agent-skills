param([Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference='Stop'
$j=Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$e=$null;$b=$null;$t=$null
try {
 $e=New-Object -ComObject Excel.Application
 $e.Visible=$false;$e.DisplayAlerts=$false;$e.EnableEvents=$false;$e.AutomationSecurity=3
 $b=$e.Workbooks.Open($j.source,0,$true)
 $t=$e.Workbooks.Open($j.template,0,$true)
 # Replace the whole report set before writing formulas. Deleting an old sheet
 # detail after populating summaries would convert their new links to #REF!.
 foreach($ws in @($b.Worksheets)){if($ws.Name -in $j.replace_names){$ws.Delete()}}
 $anchor=$b.Worksheets.Item('A.1')
 if($anchor.Index -ne 1){$anchor.Move($b.Worksheets.Item(1))}
 foreach($d in $j.sheets){$new=$b.Worksheets.Add([Type]::Missing,$anchor);$new.Name=$d.name;$anchor=$new}
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  $n=if($d.name -eq '小班明细'){8}else{[int]$d.name.Substring(2)};$last=[int]$d.last;$nc=[int]$d.cols;$head=[int]$d.head
  $ref=$t.Worksheets.Item($d.name)
  if($n -eq 5){
   # B.5 expands the template's ownership blocks for all actual land classes.
   $ref.Range('A1:O6').Copy($s.Range('A1'))
   $s.Range('A1:O6').UnMerge();$s.Range('A1:O6').ClearContents()
   for($c=1;$c -le $nc;$c++){
    $sc=if($c -le 3){$c}else{4}
    $s.Columns.Item($c).ColumnWidth=$ref.Columns.Item($sc).ColumnWidth
   }
  }elseif($n -ne 7){
   $ref.Range($ref.Cells.Item(1,1),$ref.Cells.Item($head,$nc)).Copy($s.Range('A1'))
   for($c=1;$c -le $nc;$c++){$s.Columns.Item($c).ColumnWidth=$ref.Columns.Item($c).ColumnWidth}
  }else{
   # B.7 is a real template too. Keep the fixed administration columns and
   # repeat its first three-column fee block for the configured categories.
   $ref.Range('A1:AB7').Copy($s.Range('A1'))
   $s.Range('A1:AB7').UnMerge();$s.Range('A1:AB7').ClearContents()
   if($nc -lt 28){$s.Range($s.Cells.Item(1,$nc+1),$s.Cells.Item(7,28)).Clear()}
   for($c=1;$c -le $nc;$c++){
    $sc=if($c -le 7){$c}else{8+(($c-8)%3)}
    $s.Columns.Item($c).ColumnWidth=$ref.Columns.Item($sc).ColumnWidth
   }
  }
  $all=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($last,$nc))
  $all.Font.Name='宋体';$all.Font.Size=8;$all.Font.Bold=$false
  $all.Interior.Pattern=-4142;$all.HorizontalAlignment=-4108;$all.VerticalAlignment=-4108;$all.WrapText=$true
  $all.RowHeight=18
  # Text-formatted template cells would display formula strings. Set a numeric
  # format before writing .Formula; text results still preserve leading zeros.
  $all.NumberFormat='0.########'
  foreach($p in $d.cells.PSObject.Properties){
   if($null -eq $p.Value){continue}
   $cell=$s.Range($p.Name)
   if($p.Value -is [string] -and $p.Value.StartsWith('=')){$cell.Formula=$p.Value}
   elseif($p.Value -is [int] -or $p.Value -is [long]){$cell.Value2=[double]$p.Value}
   else{$cell.Value2=$p.Value}
  }
  foreach($m in $d.merges){[void]$s.Range($m).Merge()}
  foreach($f in $d.formats){$s.Range($f[0]).NumberFormat=$f[1]}
  $s.Range('A1').MergeArea.Font.Size=11;$s.Range('A1').MergeArea.Font.Bold=$true
  if($n -le 4){
   for($r=3;$r -le $head;$r++){$s.Rows.Item($r).RowHeight=$ref.Rows.Item($r).RowHeight}
   $s.Rows.Item(1).RowHeight=20;$s.Rows.Item(2).RowHeight=20
   # The asset includes the correctly merged unit row.
  }elseif($n -eq 5){
   for($r=1;$r -le $head;$r++){$s.Rows.Item($r).RowHeight=$ref.Rows.Item($r).RowHeight}
   $s.Range($s.Cells.Item(7,1),$s.Cells.Item($last,$nc)).RowHeight=24
   $s.Range('A3').MergeArea.HorizontalAlignment=-4152
  }elseif($n -eq 6){
   $s.Rows.Item(1).RowHeight=36;$s.Rows.Item(2).RowHeight=18
   $s.Range('A3:K5').RowHeight=20;$s.Range("A6:K$last").RowHeight=24
  }elseif($n -eq 7){
   for($r=1;$r -le $head;$r++){$s.Rows.Item($r).RowHeight=$ref.Rows.Item($r).RowHeight}
   # Restore template typography after the common data-cell defaults. Header
   # merges/labels come from the project payload; formatting comes from Excel.
   for($r=1;$r -le $head;$r++){for($c=1;$c -le $nc;$c++){
    $sc=if($c -le 7){$c}else{8+(($c-8)%3)}
    $from=$ref.Cells.Item($r,$sc);$to=$s.Cells.Item($r,$c)
    $to.Font.Name=$from.Font.Name;$to.Font.Size=$from.Font.Size;$to.Font.Bold=$from.Font.Bold
   }}
   $s.Range('A1').MergeArea.Font.Name=$ref.Range('A1').Font.Name
   $s.Range('A1').MergeArea.Font.Size=$ref.Range('A1').Font.Size
   $s.Range('A1').MergeArea.Font.Bold=$ref.Range('A1').Font.Bold
   $s.Range('A3').MergeArea.HorizontalAlignment=-4152
  }else{
   $s.Columns.Item('F').ColumnWidth=7;$s.Columns.Item('L').ColumnWidth=8
   $s.Columns.Item('P').ColumnWidth=5;$s.Columns.Item('I').ColumnWidth=8
   $s.Columns.Item('J').ColumnWidth=7;$s.Columns.Item('AB').ColumnWidth=5
   $s.Rows.Item(1).RowHeight=34;$s.Rows.Item(2).RowHeight=8;$s.Rows.Item(3).RowHeight=42
   $s.Range("A4:AB$last").RowHeight=28
  }
  foreach($p in $d.heights.PSObject.Properties){$s.Rows.Item([int]$p.Name).RowHeight=[double]$p.Value}
  $first=if($n -in 6,8){3}else{4}
  $table=$s.Range($s.Cells.Item($first,1),$s.Cells.Item([int]$d.table_last,$nc))
  $table.Borders.LineStyle=1;$table.Borders.Weight=2
  $s.PageSetup.PaperSize=9;$s.PageSetup.Orientation=$(if($n -le 4){1}else{2})
  $s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1;$s.PageSetup.FitToPagesTall=$false
  $s.PageSetup.PrintArea=$all.Address()
  $s.PageSetup.PrintTitleRows='$1:$'+$head
  $s.PageSetup.LeftMargin=$e.CentimetersToPoints(1);$s.PageSetup.RightMargin=$e.CentimetersToPoints(1)
  $s.PageSetup.TopMargin=$e.CentimetersToPoints(1.2);$s.PageSetup.BottomMargin=$e.CentimetersToPoints(1.2)
  $s.PageSetup.CenterHorizontally=$true;$s.PageSetup.CenterVertically=$false
  $b.Activate();$s.Activate();$e.ActiveWindow.View=2;$e.ActiveWindow.Zoom=100
 }
 $t.Close($false);$t=$null
 $e.CalculateFullRebuild()
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name);$s.Activate();$s.DisplayPageBreaks=$true
  $cuts=@($s.HPageBreaks | ForEach-Object {$_.Location.Row})
  # A vertically merged label can be clipped or disappear on the next printed
  # page. Split it at actual Excel page boundaries and repeat the same label.
  foreach($cut in $cuts){
   for($c=1;$c -le [int]$d.cols;$c++){
    $cell=$s.Cells.Item([int]$cut,$c)
    if(-not $cell.MergeCells){continue}
    $m=$cell.MergeArea;$top=$m.Row;$left=$m.Column;$bottom=$top+$m.Rows.Count-1;$right=$left+$m.Columns.Count-1
    if($top -lt $cut -and $bottom -ge $cut){
     $label=$m.Cells.Item(1,1).Value2;$m.UnMerge()
     foreach($part in @(@($top,($cut-1)),@($cut,$bottom))){
      $q=$s.Range($s.Cells.Item([int]$part[0],$left),$s.Cells.Item([int]$part[1],$right))
      [void]$q.Merge();$q.Cells.Item(1,1).Value2=$label
     }
    }
    $c=$right
   }
  }
  $s.ExportAsFixedFormat(0,(Join-Path $j.stage "$($d.name).pdf"))
 }
 $b.Worksheets.Item('A.1').Activate()
 $b.SaveAs($j.candidate,51);$b.Close($false);$b=$null
 $b=$e.Workbooks.Open($j.candidate,0,$true)
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  foreach($p in $d.checks.PSObject.Properties){
   $v=$s.Range($p.Name).Value2
   if($v -isnot [double] -or [Math]::Abs($v-[double]$p.Value) -gt 0.00000001){throw "Excel total mismatch: $($d.name)!$($p.Name)"}
  }
 }
 Write-Output 'Excel save, reopen, and subtotal verification passed'
}finally{
 if($t){$t.Close($false)};if($b){$b.Close($false)}
 if($e){$e.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($e)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
