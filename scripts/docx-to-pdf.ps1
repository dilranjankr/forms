# Convert a DOCX to PDF via Microsoft Word COM
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$Dst
)
$ErrorActionPreference = "Stop"
$inAbs  = (Resolve-Path -Path $Src).Path
$outAbs = [System.IO.Path]::GetFullPath($Dst)
Write-Output "DOCX: $inAbs"
Write-Output "PDF : $outAbs"

$word = $null
$doc  = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($inAbs, $false, $true)   # ReadOnly=true
  # 17 = wdFormatPDF
  $doc.SaveAs([ref]$outAbs, [ref]17)
  Write-Output "OK"
}
finally {
  if ($doc -ne $null)  { $doc.Close($false)  | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc)  | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
