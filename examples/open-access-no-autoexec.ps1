param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database not found: $DatabasePath"
}

$access = New-Object -ComObject Access.Application

# 3 = msoAutomationSecurityForceDisable
# Set this before OpenCurrentDatabase.
$access.AutomationSecurity = 3
$access.Visible = $true
$access.OpenCurrentDatabase($DatabasePath)

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
