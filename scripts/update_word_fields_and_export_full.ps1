param(
    [string]$InputDocx = "",
    [string]$OutputDocx = "",
    [string]$OutputPdf = "",
    [int]$TimeoutSeconds = 240
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($InputDocx)) {
    throw "InputDocx must be provided explicitly. Non-ASCII default paths are avoided for Word COM reliability."
}
if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    throw "OutputDocx must be provided explicitly. Non-ASCII default paths are avoided for Word COM reliability."
}
if ([string]::IsNullOrWhiteSpace($OutputPdf)) {
    throw "OutputPdf must be provided explicitly. Non-ASCII default paths are avoided for Word COM reliability."
}
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputDocxPath = [System.IO.Path]::GetFullPath($OutputDocx)
$outputPdfPath = [System.IO.Path]::GetFullPath($OutputPdf)
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputDocxPath)) | Out-Null
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputPdfPath)) | Out-Null

$started = Get-Date
$job = Start-Job -ScriptBlock {
    param($inputPath, $outputDocxPath, $outputPdfPath)
    $ErrorActionPreference = "Stop"

    function Update-StoryFields {
        param($Document)
        foreach ($story in $Document.StoryRanges) {
            $range = $story
            while ($null -ne $range) {
                $range.Fields.Update() | Out-Null
                $range = $range.NextStoryRange
            }
        }
    }

    function Apply-StaticCleanups {
        param($Document)
        $coverTailFragment = -join ([char[]](0x57DF, 0x670D, 0x52A1, 0x7F16, 0x6392, 0x7814, 0x7A76))
        foreach ($shape in $Document.Shapes) {
            try {
                if ($shape.TextFrame.HasText -ne 0) {
                    $paragraphs = $shape.TextFrame.TextRange.Paragraphs
                    for ($idx = $paragraphs.Count; $idx -ge 1; $idx--) {
                        $paragraph = $paragraphs.Item($idx)
                        if ($paragraph.Range.Text.Contains($coverTailFragment)) {
                            $paragraph.Range.Delete() | Out-Null
                        }
                    }
                }
            } catch {}
        }
    }

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($inputPath)

        Apply-StaticCleanups -Document $doc
        Update-StoryFields -Document $doc
        foreach ($toc in $doc.TablesOfContents) {
            $toc.Update()
        }
        foreach ($tof in $doc.TablesOfFigures) {
            $tof.Update()
        }
        $doc.Repaginate()
        Update-StoryFields -Document $doc

        $doc.SaveAs2($outputDocxPath)
        $doc.ExportAsFixedFormat($outputPdfPath, 17)
        $doc.Close(0)
        $word.Quit()
        "WORD_UPDATE_EXPORT_OK"
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch {}
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
        }
    }
} -ArgumentList $inputPath, $outputDocxPath, $outputPdfPath

$done = Wait-Job $job -Timeout $TimeoutSeconds
if ($done) {
    Receive-Job $job
    Remove-Job $job
}
else {
    Stop-Job $job
    Remove-Job $job
    "WORD_UPDATE_EXPORT_TIMEOUT"
}

Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.StartTime -gt $started } |
    Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path -LiteralPath $outputDocxPath) {
    Get-Item -LiteralPath $outputDocxPath
}
if (Test-Path -LiteralPath $outputPdfPath) {
    Get-Item -LiteralPath $outputPdfPath
}
