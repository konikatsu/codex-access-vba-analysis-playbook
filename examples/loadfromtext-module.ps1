param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [Parameter(Mandatory = $true)]
    [string]$ModuleName,

    [Parameter(Mandatory = $true)]
    [string]$ModuleTextPath
)

$ErrorActionPreference = 'Stop'
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
