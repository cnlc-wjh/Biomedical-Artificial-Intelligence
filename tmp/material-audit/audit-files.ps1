param([string]$root)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$sha=[System.Security.Cryptography.SHA256]::Create()
$archive=[System.IO.Compression.ZipFile]::OpenRead("$root\2. Data.zip")
try {
foreach($e in $archive.Entries){
if($e.Length -eq 0){continue}
$s=$e.Open();$zipHash=[BitConverter]::ToString($sha.ComputeHash($s));$s.Dispose()
$local=Join-Path $root $e.FullName
if(Test-Path -LiteralPath $local){
$s=[System.IO.File]::Open($local,'Open','Read','ReadWrite');$localHash=[BitConverter]::ToString($sha.ComputeHash($s));$s.Dispose()
if($zipHash -ne $localHash){Write-Output "DIFFERENT: $($e.FullName)"}
}else{Write-Output "ZIP-ONLY: $($e.FullName)"}
}
Write-Output "ZIP entries fully read: $($archive.Entries.Count)"
}finally{$archive.Dispose();$sha.Dispose()}
$rows=Import-Csv -LiteralPath "$root\V3\data\heart-disease-four-site.csv"
Write-Output "DATA records: $($rows.Count)"
$cols=$rows[0].PSObject.Properties.Name
foreach($col in $cols){
$values=@($rows | ForEach-Object {$_.PSObject.Properties[$col].Value})
$valid=@($values | Where-Object {$_ -ne ''})
$distinct=@($valid | Sort-Object -Unique)
$range=if($col -notin 'record_id','site'){($valid | ForEach-Object {[double]$_} | Measure-Object -Minimum -Maximum) }else{$null}
[pscustomobject]@{Field=$col;Nonempty=$valid.Count;Missing=$values.Count-$valid.Count;Distinct=$distinct.Count;Min=$range.Minimum;Max=$range.Maximum;Values=if($distinct.Count -lt 15){$distinct -join ','}else{'many'}} | ConvertTo-Json -Compress
}
$rows | Group-Object site | Select-Object Name,Count | Format-Table
Get-ChildItem -LiteralPath "$root\V3" -Recurse -Filter '*.csv' | Where-Object FullName -NotLike '*\data\*' | ForEach-Object {Write-Output $_.FullName;Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8}
Get-Content -LiteralPath "$root\V3\review\verified-results.json" -Raw -Encoding UTF8
Get-Content -LiteralPath "$root\V3\review\analysis-design.json" -Raw -Encoding UTF8
