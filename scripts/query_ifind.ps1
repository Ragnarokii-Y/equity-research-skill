param(
 [Parameter(Mandatory=$true)][string]$QueryPath,
 [Parameter(Mandatory=$true)][string]$OutputDir,
 [ValidateRange(20,600)][int]$TimeoutSeconds=180
)
$ErrorActionPreference='Stop'
$queries=@(Get-Content -LiteralPath $QueryPath -Raw -Encoding utf8 | ConvertFrom-Json)
if (-not $queries.Count) { throw 'Empty query list' }
foreach($q in $queries){if ($q.Formula -notmatch '^=thsiFinD\('){throw ('Unsupported formula: '+$q.Id)}}
$resolvedOutput=[IO.Path]::GetFullPath($OutputDir)
$allowedOutput=[IO.Path]::GetFullPath('G:\Codex').TrimEnd('\')
if (-not $resolvedOutput.StartsWith($allowedOutput+'\',[StringComparison]::OrdinalIgnoreCase)) {throw 'OutputDir must be a project folder under G:\Codex.'}
New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
$taskExcel=New-Object -ComObject Excel.Application
$initialBooks=$taskExcel.Workbooks.Count
$ownsApp=($initialBooks -eq 0)
$priorAlerts=$taskExcel.DisplayAlerts
$book=$null
try {
 if($ownsApp){$taskExcel.Visible=$false}
 $taskExcel.DisplayAlerts=$false
 $addins=@($taskExcel.COMAddIns | ForEach-Object {[pscustomobject]@{ProgId=$_.ProgId;Description=$_.Description;Connected=[bool]$_.Connect}})
 $addins | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $resolvedOutput 'Excel加载项检查.json') -Encoding utf8
 $plugin=@($addins | Where-Object {$_.ProgId -match 'ThsXlsSlug' -and $_.Connected})
 if(-not $plugin.Count){throw 'iFinD COM add-in is not connected. No settings changed.'}
 $book=$taskExcel.Workbooks.Add()
 $sheet=$book.Worksheets.Item(1);$sheet.Name='原始公式'
 $headers=@('ID','证券','报告期或年度','指标','单位','截至日','公式文本','查询结果','状态','取数时间')
 for($c=0;$c -lt $headers.Count;$c++){$sheet.Cells.Item(1,$c+1).Value2=$headers[$c]}
 for($i=0;$i -lt $queries.Count;$i++){
  $q=$queries[$i];$row=$i+2
  $sheet.Cells.Item($row,1).Value2=[string]$q.Id
  $sheet.Cells.Item($row,2).Value2=[string]$q.Code
  $sheet.Cells.Item($row,3).Value2=[string]$q.Period
  $sheet.Cells.Item($row,4).Value2=[string]$q.Field
  $sheet.Cells.Item($row,5).Value2=[string]$q.Unit
  $sheet.Cells.Item($row,6).Value2=[string]$q.AsOf
  $sheet.Cells.Item($row,7).NumberFormat='@';$sheet.Cells.Item($row,7).Value2=[string]$q.Formula
  $sheet.Cells.Item($row,8).Formula=[string]$q.Formula
 }
 $taskExcel.Calculate()
 $watch=[Diagnostics.Stopwatch]::StartNew()
 do {
  $pending=0
  for($i=0;$i -lt $queries.Count;$i++){
   $cell=$sheet.Cells.Item($i+2,8);$v=$cell.Value2;$tx=$cell.Text
   if($null -eq $v -or $tx -match '抓取|获取|加载|连接中|loading|fetching'){$pending++}
  }
  if($pending -eq 0){break}
  Write-Output ('Waiting: '+$pending+' queries, '+[int]$watch.Elapsed.TotalSeconds+'s')
  Start-Sleep -Seconds 5
 } while($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds)
 $snapshot=@()
 $queriedAt=(Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
 for($i=0;$i -lt $queries.Count;$i++){
  $q=$queries[$i];$row=$i+2;$cell=$sheet.Cells.Item($row,8);$v=$cell.Value2;$tx=[string]$cell.Text
  $isNumber=($v -is [double] -or $v -is [int] -or $v -is [decimal])
  $status=if($tx.StartsWith('#')){'error'}elseif($isNumber){if([double]$v -eq 0){'zero_review'}else{'numeric'}}elseif($tx -match '抓取|获取|加载|loading|fetching'){'timeout'}elseif($null -eq $v -or $tx -eq ''){'missing'}else{'text_review'}
  $sheet.Cells.Item($row,9).Value2=$status;$sheet.Cells.Item($row,10).Value2=$queriedAt
  $snapshot += [pscustomobject]@{Id=$q.Id;Code=$q.Code;Period=$q.Period;Field=$q.Field;Unit=$q.Unit;AsOf=$q.AsOf;Window=$q.Window;Formula=$q.Formula;Value=$v;Display=$tx;Status=$status;QueriedAt=$queriedAt}
 }
 $sheet.UsedRange.Columns.AutoFit() | Out-Null
 $sheet.Columns.Item(7).ColumnWidth=65
 $book.SaveAs((Join-Path $resolvedOutput 'iFinD_原始公式.xlsx'),51)
 $snapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $resolvedOutput 'iFinD_固定快照.json') -Encoding utf8
 $snapshot | Group-Object Status | Select-Object Name,Count | ConvertTo-Json
} finally {
 if($null -ne $book){$book.Close($false)}
 $taskExcel.DisplayAlerts=$priorAlerts
 if($ownsApp){$taskExcel.Quit()}
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) | Out-Null
}
