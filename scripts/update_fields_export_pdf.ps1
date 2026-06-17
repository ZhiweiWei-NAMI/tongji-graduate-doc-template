param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [Parameter(Mandatory=$true)][string]$OutputPdf
)

$ErrorActionPreference = 'Stop'

$inputPath = [System.IO.Path]::GetFullPath($InputDocx)
$outputDocxPath = [System.IO.Path]::GetFullPath($OutputDocx)
$outputPdfPath = [System.IO.Path]::GetFullPath($OutputPdf)

New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputDocxPath)) | Out-Null
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputPdfPath)) | Out-Null

$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0

  $doc = $word.Documents.Open($inputPath, $false, $false)

  foreach ($toc in $doc.TablesOfContents) {
    $toc.Update()
  }

  $doc.Fields.Update() | Out-Null

  foreach ($story in $doc.StoryRanges) {
    $range = $story
    while ($null -ne $range) {
      $range.Fields.Update() | Out-Null
      $range = $range.NextStoryRange
    }
  }

  $doc.Repaginate()
  $doc.SaveAs2($outputDocxPath, 16)
  $doc.ExportAsFixedFormat($outputPdfPath, 17)

  $pageCount = $doc.ComputeStatistics(2)
  Write-Output "DOCX=$outputDocxPath"
  Write-Output "PDF=$outputPdfPath"
  Write-Output "PAGES=$pageCount"
}
finally {
  if ($null -ne $doc) {
    $doc.Close($false)
  }
  if ($null -ne $word) {
    $word.Quit()
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
