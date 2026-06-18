param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [string]$AccessExe = 'msaccess.exe',

    [string]$CommandText = 'SKIP_AUTOEXEC'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database not found: $DatabasePath"
}

$arguments = "`"$DatabasePath`" /cmd $CommandText"

Start-Process -FilePath $AccessExe -ArgumentList $arguments
