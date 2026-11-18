param(
    [Parameter(Mandatory = $true)][string]$OriginalPath
)

$ErrorActionPreference = 'Stop'
$taskOriginal = Get-Content -Raw -Encoding UTF8 -LiteralPath $OriginalPath
$taskWrapper = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $PSScriptRoot 'auto_refresh_integration.gs')
$taskReceiver = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $PSScriptRoot 'call_statistics.gs')
$taskPattern = '(?ms)^function autoRefresh\(\) \{\r?\n.*?^\}'
$taskMatches = [regex]::Matches($taskOriginal, $taskPattern)
if ($taskMatches.Count -ne 1) {
    throw 'Expected exactly one original autoRefresh function. No output was written.'
}
$taskMatch = $taskMatches[0]
$taskDashboard = $taskOriginal.Substring(0, $taskMatch.Index) + $taskWrapper.TrimEnd() + $taskOriginal.Substring($taskMatch.Index + $taskMatch.Length)
$taskCombined = $taskDashboard.TrimEnd() + "`r`n"
$taskOutput = Join-Path $PSScriptRoot 'dashboard_with_call_statistics.gs.txt'
[System.IO.File]::WriteAllText($taskOutput, $taskCombined, [System.Text.UTF8Encoding]::new($false))
$taskExportDir = Join-Path $PSScriptRoot 'exports'
New-Item -ItemType Directory -Path $taskExportDir -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path $taskExportDir '01_receiver_save.gs.txt'), $taskReceiver.TrimEnd() + "`r`n", [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $taskExportDir '02_dashboard.gs.txt'), $taskCombined, [System.Text.UTF8Encoding]::new($false))
Write-Output ('Generated receiver and dashboard GAS exports: ' + $taskExportDir)
