param(
    [Parameter(Mandatory=$true)][string]$DocxPath,
    [Parameter(Mandatory=$true)][string]$PdfPath
)

$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($DocxPath, $false, $true)
    $outDir = Split-Path -Path $PdfPath -Parent
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    $doc.ExportAsFixedFormat($PdfPath, 17)
    $pages = $doc.ComputeStatistics(2)
    Write-Output "PDF=$PdfPath"
    Write-Output "PAGES=$pages"
}
finally {
    if ($doc -ne $null) {
        $doc.Close([ref]0)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
    if ($doc -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
