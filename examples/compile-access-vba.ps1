param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath
)

$ErrorActionPreference = 'Stop'
$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.AutomationSecurity = 1
    $access.OpenCurrentDatabase($DatabasePath)

    # 126 = acCmdCompileAndSaveAllModules
    $access.RunCommand(126)
    Write-Host "Compile and save succeeded."
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
