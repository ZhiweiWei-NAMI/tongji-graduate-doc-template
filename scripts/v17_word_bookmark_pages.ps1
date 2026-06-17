param(
    [string]$InputDocx = "",
    [string]$OutputCsv = "",
    [int]$TimeoutSeconds = 240
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($InputDocx)) {
    throw "InputDocx must be provided."
}
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
    throw "OutputCsv must be provided."
}

$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputCsv)
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputPath)) | Out-Null

$started = Get-Date
$job = Start-Job -ScriptBlock {
    param($inputPath, $outputPath)
    $ErrorActionPreference = "Stop"
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($inputPath, $false, $true)
        $doc.Repaginate()
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($bookmark in $doc.Bookmarks) {
            $name = [string]$bookmark.Name
            if (-not $name.StartsWith("eq_")) {
                continue
            }
            $range = $bookmark.Range
            $text = ([string]$range.Text).Replace("`r", "").Replace("`a", "").Trim()
            $rows.Add([pscustomobject]@{
                bookmark = $name
                label = $text
                page = $range.Information(3)
                x_page = [math]::Round([double]$range.Information(5), 2)
                y_page = [math]::Round([double]$range.Information(6), 2)
            })
        }
        $rows | Sort-Object bookmark | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        $doc.Close(0)
        $word.Quit()
        "WORD_BOOKMARK_PAGE_EXPORT_OK"
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch {}
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
        }
    }
} -ArgumentList $inputPath, $outputPath

$done = Wait-Job $job -Timeout $TimeoutSeconds
if ($done) {
    Receive-Job $job
    Remove-Job $job
}
else {
    Stop-Job $job
    Remove-Job $job
    "WORD_BOOKMARK_PAGE_EXPORT_TIMEOUT"
}

Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.StartTime -gt $started } |
    Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path -LiteralPath $outputPath) {
    Get-Item -LiteralPath $outputPath
}
