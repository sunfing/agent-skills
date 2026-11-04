param([Parameter(Mandatory=$true)][string]$Job, [switch]$ValidateOnly, [switch]$CleanupOnly)
$ErrorActionPreference='Stop'
$j=Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$sessionPath=Join-Path $j.stage 'excel-session.json'
if($CleanupOnly){
 if(Test-Path -LiteralPath $sessionPath){
  $owned=Get-Content -LiteralPath $sessionPath -Raw | ConvertFrom-Json
  $tracked=@{};$tracked[$owned.pid.ToString()]=[long]$owned.started
  $exitPath=Join-Path $j.stage 'excel-exit.json';$normalExit=$false
  if(Test-Path -LiteralPath $exitPath){$exitInfo=Get-Content -LiteralPath $exitPath -Raw | ConvertFrom-Json;$normalExit=($exitInfo.pid -eq $owned.pid -and $exitInfo.code -eq 0)}
  $deadline=(Get-Date).AddSeconds($(if($normalExit){2}else{8}))
  do {
   $children=@(Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE' OR Name = 'WerFault.exe'" | Where-Object {$_.ThreadCount -gt 0})
   foreach($child in $children){
    $parent=$child.ParentProcessId.ToString();$started=$child.CreationDate.ToUniversalTime().Ticks
    if($tracked.ContainsKey($parent) -and $started -ge $tracked[$parent] -and $child.ProcessId -notin $owned.existing){$tracked[$child.ProcessId.ToString()]=$started}
   }
   Start-Sleep -Milliseconds 200
  }while((Get-Date) -lt $deadline)
  # Stop only live descendants whose PID and creation time still match this run.
  foreach($child in $children){
   $key=$child.ProcessId.ToString()
   if($tracked.ContainsKey($key) -and $child.CreationDate.ToUniversalTime().Ticks -eq $tracked[$key]){
    $process=Get-Process -Id $child.ProcessId -ErrorAction SilentlyContinue
    if($process -and -not $process.HasExited){
     Stop-Process -Id $process.Id -ErrorAction Stop
     if(-not $process.WaitForExit(5000)){throw 'Owned Excel descendant did not exit'}
    }
   }
  }
 }
 Write-Output 'Owned Excel session and descendant exits verified'
 return
}
$e=$null;$b=$null;$t=$null;$fontChecks=@();$ownedExcel=$false
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ExcelWindowOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$existingExcel=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
try {
 $e=New-Object -ComObject Excel.Application
 [uint32]$excelProcessId=0
 [void][ExcelWindowOwner]::GetWindowThreadProcessId([IntPtr]$e.Hwnd,[ref]$excelProcessId)
 if(-not $excelProcessId -or $excelProcessId -in $existingExcel){throw 'Refusing to use a pre-existing Excel process'}
 $ownedExcel=$true;$process=Get-Process -Id $excelProcessId;[void]$process.Handle
 $session=@{pid=$excelProcessId;started=$process.StartTime.ToUniversalTime().Ticks.ToString();existing=$existingExcel}
 $session | ConvertTo-Json | Set-Content -LiteralPath $sessionPath -Encoding UTF8
 $session | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $j.stage "excel-session-$excelProcessId.json") -Encoding UTF8
 $e.Visible=$false;$e.DisplayAlerts=$false;$e.EnableEvents=$false;$e.ScreenUpdating=$false;$e.AutomationSecurity=3
 $b=$e.Workbooks.Open($j.source,0,$true)
 $t=$e.Workbooks.Open($j.template,0,$true)
 $e.Calculation=-4135
 foreach($name in $j.input_names){
  $s=$b.Worksheets.Item($name)
  if($s.Range('A1').Value2 -ne $null){
   $font=$s.Range('A1').Font
   $fontChecks+=@{sheet=$name;address='A1';name=$font.Name;size=$font.Size;bold=$font.Bold;italic=$font.Italic}
  }
 }
 foreach($d in $j.sheets){
  $ref=$t.Worksheets.Item($d.name)
  $addresses=if($d.name -eq '现状调查表'){@('A1','A16','D3','D7')}else{@('A1')}
  foreach($address in $addresses){
   $font=$ref.Range($address).Font
   $fontChecks+=@{sheet=$d.name;address=$address;name=$font.Name;size=$font.Size;bold=$font.Bold;italic=$font.Italic}
  }
 }
 if(-not $ValidateOnly){
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
    $styleRow=$sr
    if($d.style_rows -and $d.style_rows.PSObject.Properties["$r"]){$styleRow=[int]$d.style_rows.PSObject.Properties["$r"].Value}
    [void]$ref.Range($ref.Cells.Item($styleRow,1),$ref.Cells.Item($styleRow,$nc)).Copy($s.Cells.Item($r,1))
    $s.Rows.Item($r).RowHeight=$ref.Rows.Item($styleRow).RowHeight
   }
   $body=$s.Range($s.Cells.Item($head+1,1),$s.Cells.Item($last,$nc))
   [void]$body.UnMerge();[void]$body.ClearContents()
  }
  if($d.name -eq 'B.5'){
   $expanded=$s.Range($s.Cells.Item(4,4),$s.Cells.Item($last,$nc))
   $expanded.HorizontalAlignment=-4108;$expanded.VerticalAlignment=-4108;$expanded.WrapText=$true
  }
  foreach($p in $d.cells.PSObject.Properties){
   if($null -eq $p.Value){continue}
   $cell=$s.Range($p.Name)
   # Keep each merged cell's numeric/date format independent.
   if($p.Value -is [string] -and $p.Value.StartsWith('=')){$cell.Formula=$p.Value}
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
  $all=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($last,$nc))
  foreach($prop in @('PaperSize','Orientation','Zoom','FitToPagesWide','FitToPagesTall','LeftMargin','RightMargin','TopMargin','BottomMargin','HeaderMargin','FooterMargin','CenterHorizontally','CenterVertically','PrintTitleRows','PrintTitleColumns','LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter')){$s.PageSetup.$prop=$ref.PageSetup.$prop}
  $s.PageSetup.PrintArea=$all.Address()
  if($d.name -eq 'B.5' -and $nc -gt 15){$s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1;$s.PageSetup.FitToPagesTall=$false}
  foreach($boundary in $d.manual_breaks){[void]$s.HPageBreaks.Add($s.Cells.Item(([int]$boundary+1),1))}
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
 $e.Calculation=-4105;$e.CalculateFullRebuild()
 $e.Calculation=-4135
 # Measure wrapped text using its actual font and width, without shrinking fonts.
 $measure=$b.Worksheets.Add()
 $widthCache=@{};$heightCache=@{}
 function Get-TextHeight($cell,$area,$value){
   $widthKey=$area.Width.ToString('F2');$font=$cell.Font
   if($font.Name -is [DBNull] -or $font.Size -is [DBNull] -or $font.Bold -is [DBNull] -or $font.Italic -is [DBNull]){
    # Mixed rich text has no single Range.Font. Measure with the largest run.
    $font=$cell.Characters(1,1).Font
    for($charIndex=2;$charIndex -le $value.Length;$charIndex++){
     $runFont=$cell.Characters($charIndex,1).Font
     if($runFont.Size -gt $font.Size -or ($runFont.Size -eq $font.Size -and $runFont.Bold)){$font=$runFont}
    }
   }
   $key="$widthKey|$($font.Name)|$($font.Size)|$($font.Bold)|$($font.Italic)|$value"
   if(-not $heightCache.ContainsKey($key)){
    $probe=$measure.Range('A1');$probe.Value2=$value;$probe.WrapText=$true
    foreach($prop in @('Name','Size','Bold','Italic')){$probe.Font.$prop=$font.$prop}
    if(-not $widthCache.ContainsKey($widthKey)){
     $lo=0.1;$hi=255.0
     for($i=0;$i -lt 12;$i++){
      $mid=($lo+$hi)/2;$probe.ColumnWidth=$mid
      if($probe.Width -gt $area.Width){$hi=$mid}else{$lo=$mid}
     }
     $widthCache[$widthKey]=$lo
    }
    $probe.ColumnWidth=$widthCache[$widthKey];[void]$probe.EntireRow.AutoFit()
    if($probe.RowHeight -ge 409){throw "Text exceeds a printable row: $($d.name)!$($cell.Address())"}
    $heightCache[$key]=$probe.RowHeight+2
   }
  return $heightCache[$key]
 }
 foreach($d in $j.sheets){
  $s=$b.Worksheets.Item($d.name)
  $range=$s.Range($s.Cells.Item(1,1),$s.Cells.Item([int]$d.last,[int]$d.cols));$values=$range.Value2
  for($r=1;$r -le [int]$d.last;$r++){for($c=1;$c -le [int]$d.cols;$c++){
   $value=$values[$r,$c]
   if($value -isnot [string] -or $value.Length -eq 0){continue}
   $cell=$s.Cells.Item($r,$c);$area=$cell.MergeArea
   if($area.Rows.Count -gt 1 -and $r -gt [int]$d.head){continue}
   if(-not $cell.WrapText){continue}
   $required=Get-TextHeight $cell $area $value
   $cell.ShrinkToFit=$false;$extra=$required-$area.Height
   if($extra -gt 0){$row=$s.Rows.Item($area.Row+$area.Rows.Count-1);$row.RowHeight=$row.RowHeight+$extra}
  }}
  # Repeat labels across pages and ensure even a short final segment fits.
  $s.DisplayPageBreaks=$true;$labels=@()
  foreach($merged in $d.merges){
   $area=$s.Range($merged)
   if($area.Rows.Count -gt 1 -and $area.Row -gt [int]$d.head){$labels+=@{address=$merged;value=$area.Cells.Item(1,1).Value2}}
  }
  for($layoutPass=0;$layoutPass -lt 4;$layoutPass++){
   foreach($label in $labels){
    $area=$s.Range($label.address);[void]$area.UnMerge();[void]$area.ClearContents();[void]$area.Merge();$area.Cells.Item(1,1).Value2=$label.value
   }
   $boundaries=@($s.HPageBreaks | ForEach-Object {$_.Location.Row});$grew=$false
   foreach($label in $labels){
    $area=$s.Range($label.address);$start=$area.Row;$end=$start+$area.Rows.Count-1
    $cuts=@($boundaries | Where-Object {$_ -gt $start -and $_ -le $end} | Sort-Object -Unique)
    $left=$area.Column;$right=$left+$area.Columns.Count-1;[void]$area.UnMerge();[void]$area.ClearContents()
    foreach($stop in @($cuts)+@($end+1)){
     $segment=$s.Range($s.Cells.Item($start,$left),$s.Cells.Item(($stop-1),$right))
     [void]$segment.Merge();$cell=$segment.Cells.Item(1,1);$cell.Value2=$label.value
     foreach($edge in @(7,8,9,10)){$segment.Borders.Item($edge).LineStyle=1;$segment.Borders.Item($edge).Weight=2}
     if($label.value -is [string] -and $label.value.Length -gt 0){
      $required=Get-TextHeight $cell $segment $label.value
      if($required -gt $segment.Height){
       $row=$s.Rows.Item(($stop-1));$newHeight=$row.RowHeight+$required-$segment.Height
       if($newHeight -ge 409){throw "Page label exceeds printable row height: $($d.name)!$($label.address)"}
       $row.RowHeight=$newHeight;$grew=$true
      }
     }
     $start=$stop
    }
   }
   if(-not $grew){break}
   if($layoutPass -eq 3){throw "Page label layout did not converge: $($d.name)"}
  }
 }
 $measure.Delete()
 $t.Close($false);$t=$null;$e.Calculation=-4105;$e.CalculateFullRebuild()
 $b.Worksheets.Item('内业处理表').Activate()
 $b.SaveAs($j.candidate,51);$b.Close($false);$b=$null
 Write-Output 'Excel candidate saved; title restoration and native validation pending'
 return
 }
 $b.Close($false);$b=$null;$t.Close($false);$t=$null
 # Reopen the repaired package read-only; do not Save it through Excel again.
 $b=$e.Workbooks.Open($j.candidate,0,$true)
 foreach($check in $fontChecks){
  $font=$b.Worksheets.Item($check.sheet).Range($check.address).Font
  if($font.Name -ne $check.name -or $font.Size -ne $check.size -or $font.Bold -ne $check.bold -or $font.Italic -ne $check.italic){throw "Preserved font changed: $($check.sheet)!$($check.address)"}
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
  foreach($p in $d.display_checks.PSObject.Properties){
   if($s.Range($p.Name).Text -ne [string]$p.Value){throw "Display mismatch: $($d.name)!$($p.Name)"}
  }
  foreach($boundary in $d.manual_breaks){
   if($s.Rows.Item(([int]$boundary+1)).PageBreak -ne -4135){throw "Manual page break missing: $($d.name)"}
  }
  $boundaries=@($s.HPageBreaks | ForEach-Object {$_.Location.Row})
  foreach($merged in $d.merges){
   $original=$s.Range($merged)
   foreach($boundary in $boundaries){
    if($boundary -le $original.Row -or $boundary -ge ($original.Row+$original.Rows.Count)){continue}
    $segment=$s.Cells.Item($boundary,$original.Column).MergeArea
    if($segment.Row -lt $boundary){throw "Merged label crosses a printed page: $($d.name)!$merged"}
   }
  }
 }
 foreach($d in $j.sheets){$b.Worksheets.Item($d.name).ExportAsFixedFormat(0,(Join-Path $j.stage "$($d.name).pdf"))}
 Write-Output 'Excel reopen, title fonts, formula and subtotal verification passed'
}catch{
 Write-Output "Sheet: $($d.name); cell: $($p.Name)"
 Write-Output $_.InvocationInfo.PositionMessage
 throw
}finally{
  # Release child RCWs before closing their workbooks. Keeping Range/Font
  # wrappers alive across Quit can crash Office and launch a recovery window.
  foreach($referenceName in @('s','ref','font','cell','area','probe','range','row','all','body','table','measure','segment','original','expanded','runFont','ws','anchor')){
   Set-Variable -Name $referenceName -Value $null -Scope Script
  }
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
  try {if($t){$t.Close($false)}} finally {
   try {if($b){$b.Close($false)}} finally {
    $t=$null;$b=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    try {if($e -and $ownedExcel){$e.Quit()}} finally {
     if($e){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($e)}
     [GC]::Collect();[GC]::WaitForPendingFinalizers()
     if($ownedExcel -and $process.WaitForExit(5000)){
      @{pid=$excelProcessId;code=$process.ExitCode} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $j.stage 'excel-exit.json') -Encoding UTF8
      if($process.ExitCode -ne 0){throw "Excel exited abnormally: $($process.ExitCode)"}
     }
     elseif($ownedExcel){throw 'Excel did not exit after releasing document references'}
    }
   }
  }
}
