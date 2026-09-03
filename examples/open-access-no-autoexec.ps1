param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [switch]$AcknowledgeStartupMayRun
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database not found: $DatabasePath"
}

if (-not $AcknowledgeStartupMayRun) {
    throw 'AutomationSecurity=3 alone does not guarantee that AutoExec, startup forms, or trusted-database initialization are bypassed. Use the standard workflow, or rerun with -AcknowledgeStartupMayRun for this diagnostic example.'
}

$access = New-Object -ComObject Access.Application

# 3 = msoAutomationSecurityForceDisable
# Set this before OpenCurrentDatabase.
$access.AutomationSecurity = 3
$access.Visible = $true
$access.OpenCurrentDatabase($DatabasePath)

Write-Warning 'Startup processing may still run. This example demonstrates AutomationSecurity=3 only; it is not a complete startup bypass.'
Write-Host "Access was opened with AutomationSecurity=3."
Write-Host "Keep this PowerShell session open while using Access."
Read-Host "Press Enter after closing Access"

try {
    $access.CloseCurrentDatabase()
}
catch {
}

try {
    $access.Quit()
}
catch {
}

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
