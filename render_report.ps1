param([string]$InputPath, [string]$OutputDir)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
$null=New-Item -ItemType Directory -Force -Path $OutputDir
$word=$null
$document=$null
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false
    $word.DisplayAlerts=0
    $document=$word.Documents.Open($InputPath, $false, $true)
    $document.Repaginate()
    $document.ActiveWindow.View.Type=3
    $pdfPath=Join-Path $OutputDir 'report.pdf'
    $document.ExportAsFixedFormat($pdfPath,17)
    $pages=$document.ActiveWindow.Panes.Item(1).Pages
    $pageCount=$document.ComputeStatistics(2)
    Write-Output "Pages: $pageCount"
    for($i=1;$i -le $pages.Count;$i++) {
        $bytes=$pages.Item($i).EnhMetaFileBits
        $stream=New-Object System.IO.MemoryStream(,$bytes)
        $metafile=[System.Drawing.Image]::FromStream($stream)
        $scale=1.5
        $bitmap=New-Object System.Drawing.Bitmap([int]($metafile.Width*$scale),[int]($metafile.Height*$scale))
        $graphics=[System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.DrawImage($metafile,0,0,$bitmap.Width,$bitmap.Height)
        $pngPath=Join-Path $OutputDir ("page-{0}.png" -f $i)
        $bitmap.Save($pngPath,[System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose();$bitmap.Dispose();$metafile.Dispose();$stream.Dispose()
        Write-Output $pngPath
    }
} finally {
    if($null -ne $document){$document.Close(0)}
    if($null -ne $word){$word.Quit()}
}
