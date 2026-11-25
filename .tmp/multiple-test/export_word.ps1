param([string]$InputFile, [string]$OutputFile)
$ErrorActionPreference = 'Stop'
$app = $null
$document = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $document = $app.Documents.Open($InputFile, $false, $true, $false)
    $document.Repaginate()
    $document.ExportAsFixedFormat($OutputFile, 17)
    Write-Output ('PAGES=' + $document.ComputeStatistics(2))
} finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $app) { $app.Quit(0) }
}
