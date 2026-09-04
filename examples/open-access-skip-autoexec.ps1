[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [string]$AccessExe = 'msaccess.exe',

    [ValidateSet('SKIP_AUTOEXEC', 'RUN_SELFTEST_READONLY', 'RUN_SELFTEST_DML')]
    [string]$CommandText = 'SKIP_AUTOEXEC',

    [string]$RunId,

    [string]$ResultPath,

    [string]$AllowlistPath,

    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,

    [switch]$AcknowledgeTrustedLocation
)

$ErrorActionPreference = 'Stop'

function Resolve-AccessExecutable {
    param([Parameter(Mandatory = $true)][string]$RequestedPath)

    if (Test-Path -LiteralPath $RequestedPath -PathType Leaf) {
        return (Resolve-Path -LiteralPath $RequestedPath).Path
    }

    $command = Get-Command -Name $RequestedPath -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $command) {
        return [IO.Path]::GetFullPath($command.Source)
    }

    foreach ($registryPath in @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE'
    )) {
        if (Test-Path -LiteralPath $registryPath) {
            $candidate = (Get-Item -LiteralPath $registryPath).GetValue('')
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and
                (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
        }
    }

    throw "Microsoft Access executable was not found. Pass its absolute path with -AccessExe: $RequestedPath"
}

if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) {
    throw "Database not found: $DatabasePath"
}

$databaseFullPath = (Resolve-Path -LiteralPath $DatabasePath).Path
$arguments = @(('"{0}"' -f $databaseFullPath), '/cmd', $CommandText)
$isSelfTest = $CommandText -in @('RUN_SELFTEST_READONLY', 'RUN_SELFTEST_DML')
$accessExeFullPath = Resolve-AccessExecutable -RequestedPath $AccessExe

Write-Warning 'This command is effective only when the target database startup dispatcher checks Command() against the same fixed value.'

if (-not $isSelfTest) {
    $process = Start-Process -FilePath $accessExeFullPath -ArgumentList $arguments -PassThru
    Write-Host "Started Access with /cmd $CommandText. PID: $($process.Id)"
    return
}

if (-not $AcknowledgeTrustedLocation) {
    throw 'Self-test execution requires a trusted location or explicitly enabled content. Rerun with -AcknowledgeTrustedLocation only after verifying that prerequisite.'
}
if ([string]::IsNullOrWhiteSpace($RunId)) {
    throw 'RunId is required for a self-test.'
}
if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    throw 'ResultPath is required for a self-test.'
}
if ($CommandText -eq 'RUN_SELFTEST_DML' -and [string]::IsNullOrWhiteSpace($AllowlistPath)) {
    throw 'RUN_SELFTEST_DML requires an allowlist file that is independent of the application INI.'
}

$existingAccess = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
if ($existingAccess.Count -gt 0) {
    throw 'Close existing Access processes before a self-test so the launched PID is unambiguous.'
}

if (-not ('SelfTestKeyboardState' -as [type])) {
    Add-Type @'
using System.Runtime.InteropServices;
public static class SelfTestKeyboardState
{
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int virtualKey);
}
'@
}

if (([SelfTestKeyboardState]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0) {
    throw 'Shift is currently pressed. Release physical and virtual Shift before running a /cmd self-test.'
}

$resultFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultPath)
$resultParent = Split-Path -Parent $resultFullPath
if (-not (Test-Path -LiteralPath $resultParent -PathType Container)) {
    throw "Result directory does not exist: $resultParent"
}
if (Test-Path -LiteralPath $resultFullPath) {
    throw "ResultPath already exists; use a new path for each run: $resultFullPath"
}

$allowlistFullPath = ''
if ($AllowlistPath) {
    if (-not (Test-Path -LiteralPath $AllowlistPath -PathType Leaf)) {
        throw "Allowlist file not found: $AllowlistPath"
    }
    $allowlistFullPath = (Resolve-Path -LiteralPath $AllowlistPath).Path
}

$environmentNames = @(
    'ACCESS_SELFTEST_RUN_ID',
    'ACCESS_SELFTEST_RESULT_PATH',
    'ACCESS_SELFTEST_ALLOWLIST_PATH'
)
$previousEnvironment = @{}
foreach ($name in $environmentNames) {
    $previousEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
}

$process = $null
$accessCreationTimeUtcTicks = 0L
try {
    [Environment]::SetEnvironmentVariable('ACCESS_SELFTEST_RUN_ID', $RunId, 'Process')
    [Environment]::SetEnvironmentVariable('ACCESS_SELFTEST_RESULT_PATH', $resultFullPath, 'Process')
    [Environment]::SetEnvironmentVariable('ACCESS_SELFTEST_ALLOWLIST_PATH', $allowlistFullPath, 'Process')

    $process = Start-Process -FilePath $accessExeFullPath -ArgumentList $arguments -PassThru
    try {
        $accessIdentity = Get-CimInstance Win32_Process -Filter "ProcessId=$($process.Id)" -ErrorAction Stop
    }
    catch {
        try {
            if (-not $process.HasExited) {
                $process.Kill()
                [void]$process.WaitForExit(15000)
            }
        }
        catch {}
        throw
    }
    if ($accessIdentity.Name -ine 'MSACCESS.EXE' -or
        $accessIdentity.ExecutablePath -ine $accessExeFullPath -or
        $null -eq $accessIdentity.CreationDate) {
        try {
            if (-not $process.HasExited) {
                $process.Kill()
                [void]$process.WaitForExit(15000)
            }
        }
        catch {}
        throw 'The launched process identity could not be verified as the expected Access executable.'
    }
    $accessCreationTimeUtcTicks = ([DateTime]$accessIdentity.CreationDate).ToUniversalTime().Ticks
}
finally {
    foreach ($name in $environmentNames) {
        [Environment]::SetEnvironmentVariable($name, $previousEnvironment[$name], 'Process')
    }
}

if ($null -eq $process) {
    throw 'Access did not start.'
}

$finished = $process.WaitForExit($TimeoutSeconds * 1000)
if (-not $finished) {
    $candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$($process.Id)" -ErrorAction SilentlyContinue
    $stopped = $false
    if ($null -ne $candidate -and
        $candidate.Name -ieq 'MSACCESS.EXE' -and
        $candidate.ExecutablePath -ieq $accessExeFullPath -and
        $null -ne $candidate.CreationDate -and
        ([DateTime]$candidate.CreationDate).ToUniversalTime().Ticks -eq $accessCreationTimeUtcTicks) {
        Stop-Process -Id $process.Id -Force
        Wait-Process -Id $process.Id -Timeout 15 -ErrorAction SilentlyContinue
        $stopped = $true
    }
    throw "Self-test timed out after $TimeoutSeconds seconds. Recorded PID: $($process.Id); exact executable match stopped: $stopped"
}

if (-not (Test-Path -LiteralPath $resultFullPath -PathType Leaf)) {
    throw 'Self-test result is missing. A zero-row cleanup check alone is not proof that the test ran.'
}

try {
    $result = Get-Content -LiteralPath $resultFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
catch {
    throw "Self-test result is not valid UTF-8 JSON: $($_.Exception.Message)"
}

$requiredProperties = @(
    'run_id',
    'command',
    'status',
    'started_at',
    'finished_at',
    'assertions_planned',
    'assertions_executed',
    'failed_assertions',
    'cleanup_remaining'
)
foreach ($propertyName in $requiredProperties) {
    if ($result.PSObject.Properties.Name -notcontains $propertyName) {
        throw "Self-test result is missing required property: $propertyName"
    }
}

if ([string]$result.run_id -cne $RunId) {
    throw 'Self-test result RunId does not match the requested run.'
}
if ([string]$result.command -cne $CommandText) {
    throw 'Self-test result command does not match the requested fixed command.'
}
if ([string]$result.status -cne 'PASS') {
    throw "Self-test reported a non-PASS status: $($result.status)"
}
if ([int]$result.assertions_planned -le 0) {
    throw 'Self-test planned assertion count must be greater than zero.'
}
if ([int]$result.assertions_executed -ne [int]$result.assertions_planned) {
    throw 'Self-test did not execute every planned assertion.'
}
if ([int]$result.failed_assertions -ne 0) {
    throw 'Self-test reported failed assertions.'
}
if ([int]$result.cleanup_remaining -ne 0) {
    throw 'Self-test cleanup verification found remaining test data.'
}

$startedAt = [DateTimeOffset]::Parse([string]$result.started_at)
$finishedAt = [DateTimeOffset]::Parse([string]$result.finished_at)
if ($finishedAt -lt $startedAt) {
    throw 'Self-test finished_at precedes started_at.'
}

if ($CommandText -eq 'RUN_SELFTEST_DML') {
    if ($result.PSObject.Properties.Name -notcontains 'target_allowlist_match' -or
        $result.target_allowlist_match -isnot [bool] -or
        -not $result.target_allowlist_match) {
        throw 'DML self-test did not prove an independent allowlist match before writing.'
    }
}

$remainingAccess = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
if ($remainingAccess.Count -gt 0) {
    throw 'An MSACCESS process remains after the self-test. Do not stop it without exact ownership evidence.'
}
$databaseDirectory = Split-Path -Parent $databaseFullPath
$databaseBaseName = [IO.Path]::GetFileNameWithoutExtension($databaseFullPath)
$possibleLockFiles = @(
    Join-Path $databaseDirectory ($databaseBaseName + '.laccdb')
    Join-Path $databaseDirectory ($databaseBaseName + '.ldb')
)
if (@($possibleLockFiles | Where-Object { Test-Path -LiteralPath $_ }).Count -gt 0) {
    throw 'An Access lock file remains beside the self-test database.'
}

Write-Host "PASS: $CommandText completed with $($result.assertions_executed) assertions."
Write-Host "RunId: $RunId"
Write-Host "Result: $resultFullPath"
