param([string]$Path,[string]$Pattern='^word/document.xml$',[int]$Start=0,[int]$Count=100000)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression
$stream=[System.IO.File]::Open($Path,'Open','Read','ReadWrite')
$archive=[System.IO.Compression.ZipArchive]::new($stream,[System.IO.Compression.ZipArchiveMode]::Read)
try {
  $entries=@($archive.Entries | Where-Object FullName -Match $Pattern | Sort-Object { $m=[regex]::Match($_.FullName,'\d+'); if($m.Success){[int]$m.Value}else{0} })
  foreach($entry in ($entries | Select-Object -Skip $Start -First $Count)) {
    $reader=[System.IO.StreamReader]::new($entry.Open()); [xml]$xml=$reader.ReadToEnd(); $reader.Dispose()
    Write-Output $entry.FullName
    $xml.SelectNodes('//*[local-name()="p"]') | ForEach-Object {
      ($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object InnerText) -join ''
    }
  }
} finally {$archive.Dispose();$stream.Dispose()}
