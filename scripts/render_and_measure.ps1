param(
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [int]$PreviewWidth=1326,
    [int]$PreviewHeight=746,
    [double]$TolerancePt=1.0,
    [switch]$Overwrite
)
$ErrorActionPreference='Stop'
# All generated files remain in the user-authorized root; input may be elsewhere.
$taskInput=(Resolve-Path -LiteralPath $ReportPath).ProviderPath
$taskOutput=[System.IO.Path]::GetFullPath($OutputDir)
$allowedRoot=[System.IO.Path]::GetFullPath('G:\Codex').TrimEnd('\')
if (-not $taskOutput.StartsWith($allowedRoot+'\',[System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDir must be a task directory within G:\Codex.'
}
if ([System.IO.Path]::GetExtension($taskInput) -ne '.pptx') { throw 'ReportPath must be an existing PPTX.' }
if ($PreviewWidth -lt 100 -or $PreviewHeight -lt 100) { throw 'Preview dimensions are too small.' }
$previewDir=Join-Path $taskOutput 'previews'
$pdfPath=Join-Path $taskOutput ([System.IO.Path]::GetFileNameWithoutExtension($taskInput)+'.pdf')
$measurePath=Join-Path $taskOutput 'render_measurements.json'
$summaryPath=Join-Path $taskOutput 'render_summary.json'
foreach($taskTarget in @($previewDir,$pdfPath,$measurePath,$summaryPath)) {
    if ((Test-Path -LiteralPath $taskTarget) -and -not $Overwrite) {
        throw ('Output already exists; use a new task directory or -Overwrite: '+$taskTarget)
    }
}
New-Item -ItemType Directory -Path $taskOutput,$previewDir -Force | Out-Null
$measurementRows=[System.Collections.Generic.List[object]]::new()
$renderErrors=[System.Collections.Generic.List[string]]::new()
$presentation=$null
$pptApplication=$null

function Add-TextMeasurement {
    param($TextShape,[int]$Page,[string]$ObjectPath,[Nullable[int]]$Row,[Nullable[int]]$Column)
    if (-not $TextShape.HasTextFrame -or -not $TextShape.TextFrame.HasText) { return }
    $frame=$TextShape.TextFrame2
    $range=$frame.TextRange
    $availableWidth=[Math]::Max(0.0,[double]$TextShape.Width-[double]$frame.MarginLeft-[double]$frame.MarginRight)
    $availableHeight=[Math]::Max(0.0,[double]$TextShape.Height-[double]$frame.MarginTop-[double]$frame.MarginBottom)
    $boundWidth=[double]$range.BoundWidth
    $boundHeight=[double]$range.BoundHeight
    $fontSize=$null
    try {$fontSize=[double]$range.Font.Size} catch {}
    $measurementRows.Add([pscustomobject]@{
        Page=$Page;ObjectPath=$ObjectPath;Row=$Row;Column=$Column;Text=$range.Text;
        X=[double]$TextShape.Left;Y=[double]$TextShape.Top;
        Width=[double]$TextShape.Width;Height=[double]$TextShape.Height;
        MarginLeft=[double]$frame.MarginLeft;MarginRight=[double]$frame.MarginRight;
        MarginTop=[double]$frame.MarginTop;MarginBottom=[double]$frame.MarginBottom;
        AvailableWidth=$availableWidth;AvailableHeight=$availableHeight;
        BoundLeft=[double]$range.BoundLeft;BoundTop=[double]$range.BoundTop;
        BoundWidth=$boundWidth;BoundHeight=$boundHeight;Font=$fontSize;
        HeightOverflow=($boundHeight -gt ($availableHeight+$TolerancePt));
        WidthOverflow=($boundWidth -gt ($availableWidth+$TolerancePt));
        AutoSize=[int]$frame.AutoSize;WordWrap=[int]$frame.WordWrap
    })
}

function Measure-Shape {
    param($Shape,[int]$Page,[string]$ObjectPath)
    if ([int]$Shape.Type -eq 6) { # msoGroup
        for($childIndex=1;$childIndex -le $Shape.GroupItems.Count;$childIndex++) {
            Measure-Shape -Shape $Shape.GroupItems.Item($childIndex) -Page $Page -ObjectPath ($ObjectPath+'/'+$childIndex)
        }
        return
    }
    if ($Shape.HasTable) {
        for($rowIndex=1;$rowIndex -le $Shape.Table.Rows.Count;$rowIndex++) {
            for($columnIndex=1;$columnIndex -le $Shape.Table.Columns.Count;$columnIndex++) {
                Add-TextMeasurement -TextShape $Shape.Table.Cell($rowIndex,$columnIndex).Shape -Page $Page -ObjectPath $ObjectPath -Row $rowIndex -Column $columnIndex
            }
        }
    } else {
        Add-TextMeasurement -TextShape $Shape -Page $Page -ObjectPath $ObjectPath -Row $null -Column $null
    }
}

try {
    $pptApplication=New-Object -ComObject PowerPoint.Application
    # ReadOnly=-1, Untitled=0, WithWindow=0; no visible document window.
    $presentation=$pptApplication.Presentations.Open($taskInput,-1,0,0)
    $pageCount=[int]$presentation.Slides.Count
    $slideWidth=[double]$presentation.PageSetup.SlideWidth
    $slideHeight=[double]$presentation.PageSetup.SlideHeight
    $presentation.Export($previewDir,'PNG',$PreviewWidth,$PreviewHeight)
    for($pageIndex=1;$pageIndex -le $pageCount;$pageIndex++) {
        $slide=$presentation.Slides.Item($pageIndex)
        for($shapeIndex=1;$shapeIndex -le $slide.Shapes.Count;$shapeIndex++) {
            try { Measure-Shape -Shape $slide.Shapes.Item($shapeIndex) -Page $pageIndex -ObjectPath ([string]$shapeIndex) }
            catch { $renderErrors.Add(('Page '+$pageIndex+' object '+$shapeIndex+': '+$_.Exception.Message)) }
        }
    }
    $presentation.SaveAs($pdfPath,32) # ppSaveAsPDF
    $measurementRows.ToArray() | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $measurePath -Encoding utf8
    $heightFlags=@($measurementRows | Where-Object {$_.HeightOverflow})
    $widthFlags=@($measurementRows | Where-Object {$_.WidthOverflow})
    $pngFiles=@(Get-ChildItem -LiteralPath $previewDir -File -Filter '*.PNG')
    $summary=[pscustomobject]@{
        ReportPath=$taskInput;OutputDir=$taskOutput;PageCount=$pageCount;
        WidthPt=$slideWidth;HeightPt=$slideHeight;
        PreviewWidth=$PreviewWidth;PreviewHeight=$PreviewHeight;PNGCount=$pngFiles.Count;
        PDFPath=$pdfPath;MeasurementFile=$measurePath;TextObjectCount=$measurementRows.Count;
        HeightOverflowCount=$heightFlags.Count;WidthOverflowCount=$widthFlags.Count;
        MeasurementErrorCount=$renderErrors.Count;MeasurementErrors=$renderErrors.ToArray();
        AllPagesRendered=($pngFiles.Count -eq $pageCount);
        VisualReviewRequired=$true;
        Note='Bounds flag possible clipping; rendered pages still require visual inspection for overlap, chart labels and cropping.'
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding utf8
    $summary | ConvertTo-Json -Depth 3 -Compress | Write-Output
    if($pngFiles.Count -ne $pageCount) { throw 'Preview count does not match slide count.' }
    if($renderErrors.Count -gt 0) { throw 'Some objects could not be measured; review render_summary.json.' }
} finally {
    if($null -ne $presentation) {
        try {$presentation.Close()} catch {}
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null
    }
    if($null -ne $pptApplication) {
        # Never Quit: another user presentation may share this application instance.
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pptApplication) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
