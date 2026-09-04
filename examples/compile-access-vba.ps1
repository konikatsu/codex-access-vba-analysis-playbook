param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    # Optional, zero-argument Public function changed by this work.
    [string]$PublicProcedure,

    [switch]$AcknowledgeStartupMayRun
)

$ErrorActionPreference = 'Stop'

if (-not $AcknowledgeStartupMayRun) {
    throw 'This basic sample does not implement virtual Shift. Use the standard development workflow, or rerun with -AcknowledgeStartupMayRun only when startup inspection proved no automatic startup, or normal startup is deliberately permitted in an isolated test environment.'
}

Write-Warning 'AutomationSecurity=1 enables VBA and does not bypass AutoExec or startup forms.'

function Invoke-CompilePass {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Stage
    )

    $access = $null

    try {
        $access = New-Object -ComObject Access.Application
        $access.Visible = $false
        $access.AutomationSecurity = 1
        $access.OpenCurrentDatabase($DatabasePath)

        # 126 = acCmdCompileAndSaveAllModules
        $access.RunCommand(126)

        if ($PublicProcedure) {
            $access.Run($PublicProcedure)
            Write-Host "${Stage}: compiled and ran $PublicProcedure."
        }
        else {
            Write-Host "${Stage}: compiled and saved."
        }
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
}

# Reopening catches issues that a single in-memory compile can miss.
Invoke-CompilePass -Stage 'first pass'
Invoke-CompilePass -Stage 'reopen pass'
