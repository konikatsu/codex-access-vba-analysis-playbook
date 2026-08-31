param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    # Optional, zero-argument Public function changed by this work.
    [string]$PublicProcedure
)

$ErrorActionPreference = 'Stop'

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
