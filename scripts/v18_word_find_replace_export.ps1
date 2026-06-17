param(
    [string]$InputDocx = "",
    [string]$OutputDocx = "",
    [string]$OutputPdf = "",
    [string]$ReplacementsJson = "",
    [string]$AuditJson = "",
    [int]$TimeoutSeconds = 420
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($InputDocx)) { throw "InputDocx is required." }
if ([string]::IsNullOrWhiteSpace($OutputDocx)) { throw "OutputDocx is required." }
if ([string]::IsNullOrWhiteSpace($OutputPdf)) { throw "OutputPdf is required." }
if ([string]::IsNullOrWhiteSpace($ReplacementsJson)) { throw "ReplacementsJson is required." }
if ([string]::IsNullOrWhiteSpace($AuditJson)) { throw "AuditJson is required." }

$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputDocxPath = [System.IO.Path]::GetFullPath($OutputDocx)
$outputPdfPath = [System.IO.Path]::GetFullPath($OutputPdf)
$replacementPath = (Resolve-Path -LiteralPath $ReplacementsJson).Path
$auditPath = [System.IO.Path]::GetFullPath($AuditJson)
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputDocxPath)) | Out-Null
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($outputPdfPath)) | Out-Null
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($auditPath)) | Out-Null

$started = Get-Date
$job = Start-Job -ScriptBlock {
    param($inputPath, $outputDocxPath, $outputPdfPath, $replacementPath, $auditPath)
    $ErrorActionPreference = "Stop"

    function Update-StoryFields {
        param($Document)
        foreach ($story in $Document.StoryRanges) {
            $range = $story
            while ($null -ne $range) {
                try { $range.Fields.Update() | Out-Null } catch {}
                $range = $range.NextStoryRange
            }
        }
    }

    function Replace-InRange {
        param($Range, [string]$OldText, [string]$NewText)
        $find = $Range.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $OldText
        $find.Replacement.Text = $NewText
        $find.Forward = $true
        $find.Wrap = 1
        $find.Format = $false
        $find.MatchCase = $false
        $find.MatchWholeWord = $false
        $find.MatchWildcards = $false
        $find.MatchSoundsLike = $false
        $find.MatchAllWordForms = $false
        return $find.Execute($OldText, $false, $false, $false, $false, $false, $true, 1, $false, $NewText, 2)
    }

    $rows = Get-Content -LiteralPath $replacementPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $audit = New-Object System.Collections.Generic.List[object]
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($inputPath, $false, $false, $false)

        foreach ($row in $rows) {
            $hit = $false
            $errorText = ""
            try {
                foreach ($story in $doc.StoryRanges) {
                    $range = $story
                    while ($null -ne $range) {
                        $result = Replace-InRange -Range $range -OldText $row.old -NewText $row.new
                        if ($result) { $hit = $true }
                        $range = $range.NextStoryRange
                    }
                }
            } catch {
                $errorText = $_.Exception.Message
            }
            $audit.Add([PSCustomObject]@{
                unit = $row.unit
                note = $row.note
                old = $row.old
                new = $row.new
                hit = $hit
                error = $errorText
            }) | Out-Null
        }

        Update-StoryFields -Document $doc
        foreach ($toc in $doc.TablesOfContents) {
            try { $toc.Update() } catch {}
        }
        foreach ($tof in $doc.TablesOfFigures) {
            try { $tof.Update() } catch {}
        }
        $doc.Repaginate()
        Update-StoryFields -Document $doc

        $doc.SaveAs2($outputDocxPath)
        $doc.ExportAsFixedFormat($outputPdfPath, 17)
        $audit | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $auditPath -Encoding UTF8
        $doc.Close(0)
        $word.Quit()
        "WORD_FIND_REPLACE_EXPORT_OK"
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch {}
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
        }
    }
} -ArgumentList $inputPath, $outputDocxPath, $outputPdfPath, $replacementPath, $auditPath

$done = Wait-Job $job -Timeout $TimeoutSeconds
if ($done) {
    Receive-Job $job
    Remove-Job $job
}
else {
    Stop-Job $job
    Remove-Job $job
    "WORD_FIND_REPLACE_EXPORT_TIMEOUT"
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
