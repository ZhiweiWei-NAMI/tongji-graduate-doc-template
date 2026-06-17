param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [int]$TimeoutSeconds = 20
)

$ErrorActionPreference = "Stop"

$macroNames = @(
    "MTCommand_UpdateEqns",
    "MTCommand_FormatEqns",
    "MTCommand_FormatEqnNum",
    "MTCommand_InsertEqnNum",
    "MTCommand_InsertRightNumberedDispEqn",
    "MTCommand_ConvertEqns",
    "MTConvertEquations",
    "MTCommand_OnConvertEquations"
)

$beforePids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$report = [ordered]@{
    generated_at = (Get-Date).ToString("s")
    timeout_seconds = $TimeoutSeconds
    templates = @()
    macros = @()
}

function Stop-NewWordProcesses {
    param([int[]]$ExistingPids)
    $current = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    foreach ($proc in $current) {
        if ($ExistingPids -notcontains $proc.Id) {
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

$templateJob = Start-Job -ScriptBlock {
    $word = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $items = @()
        foreach ($template in $word.Templates) {
            if ($template.Name -like "*MathType*" -or $template.FullName -like "*MathType*") {
                $items += [ordered]@{
                    name = $template.Name
                    full_name = $template.FullName
                    saved = $template.Saved
                }
            }
        }
        return $items
    }
    finally {
        if ($null -ne $word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

if (Wait-Job -Job $templateJob -Timeout $TimeoutSeconds) {
    $report.templates = @(Receive-Job -Job $templateJob)
}
else {
    Stop-Job -Job $templateJob -ErrorAction SilentlyContinue
    $report.templates = @([ordered]@{ error = "timeout" })
    Stop-NewWordProcesses -ExistingPids $beforePids
}
Remove-Job -Job $templateJob -Force -ErrorAction SilentlyContinue

foreach ($macroName in $macroNames) {
    $job = Start-Job -ArgumentList $macroName -ScriptBlock {
        param($Name)
        $word = $null
        $doc = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $doc = $word.Documents.Add()
            $start = Get-Date
            $null = $word.Run($Name)
            $elapsed = ((Get-Date) - $start).TotalSeconds
            return [ordered]@{
                macro = $Name
                status = "ok"
                elapsed_seconds = [Math]::Round($elapsed, 3)
            }
        }
        catch {
            return [ordered]@{
                macro = $Name
                status = "error"
                message = $_.Exception.Message
            }
        }
        finally {
            if ($null -ne $doc) {
                $doc.Close(0)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            if ($null -ne $word) {
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        }
    }

    if (Wait-Job -Job $job -Timeout $TimeoutSeconds) {
        $report.macros += @(Receive-Job -Job $job)
    }
    else {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        $report.macros += @([ordered]@{
            macro = $macroName
            status = "timeout"
            message = "No result before timeout; likely modal or interactive."
        })
        Stop-NewWordProcesses -ExistingPids $beforePids
    }
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
}

$reportDir = Split-Path -Parent $ReportPath
if ($reportDir) {
    New-Item -ItemType Directory -Force -Path $reportDir | Out-Null
}
$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$report | ConvertTo-Json -Depth 6
