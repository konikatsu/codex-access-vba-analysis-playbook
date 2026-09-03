param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [string]$AccessExe = 'msaccess.exe',

    [ValidateSet('SKIP_AUTOEXEC', 'RUN_SELFTEST_READONLY', 'RUN_SELFTEST_DML')]
    [string]$CommandText = 'SKIP_AUTOEXEC'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database not found: $DatabasePath"
}

$arguments = "`"$DatabasePath`" /cmd $CommandText"

Write-Warning 'This command is effective only when the target database startup dispatcher checks Command() against the same fixed value.'
Start-Process -FilePath $AccessExe -ArgumentList $arguments
