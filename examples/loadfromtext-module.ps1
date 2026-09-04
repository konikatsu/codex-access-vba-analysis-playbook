param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [Parameter(Mandatory = $true)]
    [string]$ModuleName,

    [Parameter(Mandatory = $true)]
    [string]$ModuleTextPath,

    [switch]$AcknowledgeStartupMayRun
)

$ErrorActionPreference = 'Stop'

if (-not $AcknowledgeStartupMayRun) {
    throw 'This basic sample does not implement virtual Shift. Use the standard development workflow, or rerun with -AcknowledgeStartupMayRun only when startup inspection proved no automatic startup, or normal startup is deliberately permitted in an isolated test environment.'
}

Write-Warning 'AutomationSecurity=1 enables VBA and does not bypass AutoExec or startup forms.'
$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.AutomationSecurity = 1
    $access.OpenCurrentDatabase($DatabasePath)

    try {
        $access.DoCmd.DeleteObject(5, $ModuleName)
    }
    catch {
        # Module did not exist. Continue.
    }

    $access.LoadFromText(5, $ModuleName, $ModuleTextPath)
    Write-Host "LoadFromText succeeded: $ModuleName"
}
finally {
    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
