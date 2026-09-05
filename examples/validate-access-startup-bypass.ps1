[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9A-Fa-f]{64}$')]
    [string]$ExpectedInputSha256,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [string]$AccessExe = 'msaccess.exe',

    [ValidatePattern('^[0-9A-Fa-f]{32}$')]
    [string]$RunId = ([guid]::NewGuid().ToString('N')),

    [ValidateRange(15, 120)]
    [int]$TimeoutSeconds = 30,

    [switch]$AcknowledgeTrustedLocation,

    [switch]$AcknowledgeFailureIsolation
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

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

function Assert-RestrictedLocalPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\', [StringComparison]::Ordinal)) {
        throw 'OutputDirectory cannot use a UNC path.'
    }

    $root = [IO.Path]::GetPathRoot($fullPath)
    $drive = New-Object IO.DriveInfo($root)
    if ($drive.DriveType -ne [IO.DriveType]::Fixed) {
        throw "OutputDirectory requires a fixed local drive: $root"
    }

    $current = $fullPath
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "OutputDirectory cannot traverse a reparse point: $current"
            }
        }
        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }
}

function Write-Utf8Json {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $json = ConvertTo-Json -InputObject $Value -Depth 12
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, $utf8)
}

function Read-Utf8JsonWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 20)][int]$Attempts = 5,
        [ValidateRange(10, 5000)][int]$DelayMilliseconds = 200
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        }
        catch {
            if ($attempt -eq $Attempts) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Test-ExpectedAccessProcess {
    param(
        [Parameter(Mandatory = $true)][int]$ProcessId,
        [Parameter(Mandatory = $true)][string]$ExpectedPath,
        [Parameter(Mandatory = $true)][long]$ExpectedCreationTimeUtcTicks
    )

    $candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction SilentlyContinue
    return ($null -ne $candidate -and
        $candidate.Name -ieq 'MSACCESS.EXE' -and
        $candidate.ExecutablePath -ieq $ExpectedPath -and
        $null -ne $candidate.CreationDate -and
        ([DateTime]$candidate.CreationDate).ToUniversalTime().Ticks -eq $ExpectedCreationTimeUtcTicks)
}

function Test-ExpectedStartedProcessObject {
    param(
        [Parameter(Mandatory = $true)][object]$Process,
        [Parameter(Mandatory = $true)][string]$ExpectedPath,
        [Parameter(Mandatory = $true)][long]$ExpectedStartTimeUtcTicks
    )

    try {
        $Process.Refresh()
        return (-not $Process.HasExited -and
            $Process.MainModule.FileName -ieq $ExpectedPath -and
            $Process.StartTime.ToUniversalTime().Ticks -eq $ExpectedStartTimeUtcTicks)
    }
    catch {
        return $false
    }
}

if (-not $AcknowledgeTrustedLocation) {
    throw 'Startup bypass validation requires a trusted GUI stage. Verify it, then pass -AcknowledgeTrustedLocation.'
}
if (-not $AcknowledgeFailureIsolation) {
    throw 'Verify that failed startup cannot reach production or unknown external dependencies, then pass -AcknowledgeFailureIsolation.'
}
if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) {
    throw "Database not found: $DatabasePath"
}
if (@(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Close existing Access processes before validation so the launched PID is unambiguous.'
}

if (-not ('AccessStartupBypassNative' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class AccessStartupBypassNative
{
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int virtualKey);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}
if (([AccessStartupBypassNative]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0) {
    throw 'Shift is currently pressed. Release it before validation.'
}

$databaseFullPath = (Resolve-Path -LiteralPath $DatabasePath).Path
$inputHash = (Get-FileHash -LiteralPath $databaseFullPath -Algorithm SHA256).Hash
if ($inputHash -cne $ExpectedInputSha256.ToUpperInvariant()) {
    throw 'Database SHA-256 does not match ExpectedInputSha256.'
}

$accessExeFullPath = Resolve-AccessExecutable -RequestedPath $AccessExe
$outputFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
Assert-RestrictedLocalPath -Path $outputFullPath
if (Test-Path -LiteralPath $outputFullPath) {
    throw "OutputDirectory already exists: $outputFullPath"
}
[void](New-Item -ItemType Directory -Path $outputFullPath)

$attestationPath = Join-Path $outputFullPath 'startup-bypass-attestation.json'
$acknowledgePath = Join-Path $outputFullPath 'startup-bypass-attestation.ack'
$windowSnapshotPath = Join-Path $outputFullPath 'window-snapshot.json'
$watchdogMarkerPath = Join-Path $outputFullPath 'watchdog-result.json'
$validationPath = Join-Path $outputFullPath 'validation-summary.json'
$windowSnapshotTool = Join-Path $PSScriptRoot 'get-access-window-snapshot.ps1'

$result = [ordered]@{
    schema_version = 1
    status = 'FAIL'
    run_id = $RunId.ToLowerInvariant()
    requested_command = 'SKIP_AUTOEXEC'
    failure_isolation_acknowledged = [bool]$AcknowledgeFailureIsolation
    database_sha256_before = $inputHash
    database_sha256_after = $null
    access_executable_sha256 = (Get-FileHash -LiteralPath $accessExeFullPath -Algorithm SHA256).Hash
    process_id = $null
    process_identity_match = $false
    command_line_match = $false
    attestation_received = $false
    attestation_status = $null
    database_path_match = $false
    forms_count = $null
    attested_process_id_match = $false
    hwnd_process_id_match = $false
    watchdog_fired = $false
    window_enum = 'not-needed'
    access_pid_gone = $false
    lock_files_remaining = @()
    shift_released = $false
    started_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    finished_at_utc = $null
    error = $null
}

$environmentNames = @(
    'ACCESS_STARTUP_BYPASS_RUN_ID',
    'ACCESS_STARTUP_BYPASS_RESULT_PATH',
    'ACCESS_STARTUP_BYPASS_ACK_PATH',
    'ACCESS_STARTUP_BYPASS_EXPECTED_DB_PATH'
)
$previousEnvironment = @{}
foreach ($name in $environmentNames) {
    $previousEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
}

$process = $null
$watchdog = $null
$creationTimeUtcTicks = 0L
$processStartTimeUtcTicks = 0L
$validationError = $null

try {
    [Environment]::SetEnvironmentVariable('ACCESS_STARTUP_BYPASS_RUN_ID', $RunId.ToLowerInvariant(), 'Process')
    [Environment]::SetEnvironmentVariable('ACCESS_STARTUP_BYPASS_RESULT_PATH', $attestationPath, 'Process')
    [Environment]::SetEnvironmentVariable('ACCESS_STARTUP_BYPASS_ACK_PATH', $acknowledgePath, 'Process')
    [Environment]::SetEnvironmentVariable('ACCESS_STARTUP_BYPASS_EXPECTED_DB_PATH', $databaseFullPath, 'Process')

    $arguments = @(('"{0}"' -f $databaseFullPath), '/cmd', 'SKIP_AUTOEXEC')
    $process = Start-Process -FilePath $accessExeFullPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $result.process_id = $process.Id
    $processStartTimeUtcTicks = $process.StartTime.ToUniversalTime().Ticks

    $identity = Get-CimInstance Win32_Process -Filter "ProcessId=$($process.Id)" -ErrorAction Stop
    if ($null -eq $identity.CreationDate) {
        throw 'The launched Access process has no creation timestamp.'
    }
    $creationTimeUtcTicks = ([DateTime]$identity.CreationDate).ToUniversalTime().Ticks
    if ($identity.Name -ine 'MSACCESS.EXE' -or $identity.ExecutablePath -ine $accessExeFullPath) {
        throw 'The launched process identity does not match the expected Access executable.'
    }
    $result.process_identity_match = $true

    $commandLine = [string]$identity.CommandLine
    $databaseArgumentMatch = $commandLine.IndexOf($databaseFullPath, [StringComparison]::OrdinalIgnoreCase) -ge 0
    $commandArgumentMatch = [bool]($commandLine -cmatch '(?:^|\s)/cmd\s+SKIP_AUTOEXEC(?:\s|$)')
    $result.command_line_match = $databaseArgumentMatch -and $commandArgumentMatch
    if (-not $result.command_line_match) {
        throw 'The Access command line does not contain the expected database and exact command token.'
    }

    $watchdogSource = @'
param(
    [int]$AccessPid,
    [string]$ExpectedAccessExe,
    [long]$ExpectedCreationTimeUtcTicks,
    [int]$TimeoutSeconds,
    [string]$WindowSnapshotTool,
    [string]$WindowSnapshotPath,
    [string]$MarkerPath
)

$ErrorActionPreference = 'SilentlyContinue'
Start-Sleep -Seconds $TimeoutSeconds
$windowEnum = 'process-not-running'
$processStopped = $false
$candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$AccessPid"
if ($null -ne $candidate -and
    $candidate.Name -ieq 'MSACCESS.EXE' -and
    $candidate.ExecutablePath -ieq $ExpectedAccessExe -and
    $null -ne $candidate.CreationDate -and
    ([DateTime]$candidate.CreationDate).ToUniversalTime().Ticks -eq $ExpectedCreationTimeUtcTicks) {
    try {
        & $WindowSnapshotTool -ProcessId $AccessPid -ExpectedAccessExe $ExpectedAccessExe -ExpectedCreationTimeUtcTicks $ExpectedCreationTimeUtcTicks -OutputPath $WindowSnapshotPath
        $windowEnum = 'captured'
    }
    catch {
        $windowEnum = 'failed'
    }
    $candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$AccessPid"
    if ($null -ne $candidate -and
        $candidate.Name -ieq 'MSACCESS.EXE' -and
        $candidate.ExecutablePath -ieq $ExpectedAccessExe -and
        $null -ne $candidate.CreationDate -and
        ([DateTime]$candidate.CreationDate).ToUniversalTime().Ticks -eq $ExpectedCreationTimeUtcTicks) {
        Stop-Process -Id $AccessPid -Force
        Wait-Process -Id $AccessPid -Timeout 15 -ErrorAction SilentlyContinue
        $processStopped = $true
    }
}

$payload = [ordered]@{
    watchdog_fired = $true
    access_pid = $AccessPid
    process_stopped = $processStopped
    window_enum = $windowEnum
}
$json = ConvertTo-Json -InputObject $payload -Depth 5
[IO.File]::WriteAllText($MarkerPath, $json + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
'@

    $watchdogPath = Join-Path $outputFullPath 'startup-bypass-watchdog.ps1'
    [IO.File]::WriteAllText($watchdogPath, $watchdogSource, (New-Object Text.UTF8Encoding($false)))
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $watchdogArguments = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $watchdogPath),
        '-AccessPid', $process.Id,
        '-ExpectedAccessExe', ('"{0}"' -f $accessExeFullPath),
        '-ExpectedCreationTimeUtcTicks', $creationTimeUtcTicks,
        '-TimeoutSeconds', $TimeoutSeconds,
        '-WindowSnapshotTool', ('"{0}"' -f $windowSnapshotTool),
        '-WindowSnapshotPath', ('"{0}"' -f $windowSnapshotPath),
        '-MarkerPath', ('"{0}"' -f $watchdogMarkerPath)
    )
    $watchdog = Start-Process -FilePath $windowsPowerShell -ArgumentList $watchdogArguments -WindowStyle Hidden -PassThru

    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTimeOffset]::UtcNow -lt $deadline) {
        if (Test-Path -LiteralPath $attestationPath -PathType Leaf) {
            break
        }
        $process.Refresh()
        if ($process.HasExited) {
            break
        }
        Start-Sleep -Milliseconds 100
    }

    if (-not (Test-Path -LiteralPath $attestationPath -PathType Leaf)) {
        throw 'The in-process startup bypass attestation was not produced before exit or timeout.'
    }
    $result.attestation_received = $true

    try {
        $attestation = Read-Utf8JsonWithRetry -Path $attestationPath
        foreach ($propertyName in @(
            'schema_version', 'run_id', 'command', 'status', 'database_path_match',
            'forms_count', 'hwnd_access_app', 'process_id'
        )) {
            if ($attestation.PSObject.Properties.Name -notcontains $propertyName) {
                throw "Attestation is missing required property: $propertyName"
            }
        }

        $result.attestation_status = [string]$attestation.status
        $result.database_path_match = $attestation.database_path_match -is [bool] -and $attestation.database_path_match
        $result.forms_count = [int]$attestation.forms_count
        $result.attested_process_id_match = [int]$attestation.process_id -eq $process.Id

        $windowProcessId = [uint32]0
        $windowHandleValue = [long]$attestation.hwnd_access_app
        if ($windowHandleValue -lt 0) {
            $windowHandleValue += 4294967296L
        }
        $windowHandle = [IntPtr]$windowHandleValue
        [void][AccessStartupBypassNative]::GetWindowThreadProcessId($windowHandle, [ref]$windowProcessId)
        $result.hwnd_process_id_match = [int]$windowProcessId -eq $process.Id

        if ([int]$attestation.schema_version -ne 1 -or
            [string]$attestation.run_id -cne $RunId.ToLowerInvariant() -or
            [string]$attestation.command -cne 'SKIP_AUTOEXEC' -or
            $result.attestation_status -cne 'PASS' -or
            -not $result.database_path_match -or
            $result.forms_count -ne 0 -or
            -not $result.attested_process_id_match -or
            -not $result.hwnd_process_id_match) {
            throw 'The startup bypass attestation did not satisfy every required assertion.'
        }
    }
    finally {
        [IO.File]::WriteAllText($acknowledgePath, 'ack', (New-Object Text.UTF8Encoding($false)))
    }

    $remainingMilliseconds = [Math]::Max(0, [int]($deadline - [DateTimeOffset]::UtcNow).TotalMilliseconds)
    if (-not $process.WaitForExit($remainingMilliseconds)) {
        throw 'Access did not exit after the attestation was acknowledged.'
    }
}
catch {
    $validationError = $_.Exception.Message
}
finally {
    foreach ($name in $environmentNames) {
        [Environment]::SetEnvironmentVariable($name, $previousEnvironment[$name], 'Process')
    }

    if ($null -ne $process) {
        $process.Refresh()
        $safeToStop = -not $process.HasExited -and (
            ($creationTimeUtcTicks -ne 0L -and
                (Test-ExpectedAccessProcess -ProcessId $process.Id -ExpectedPath $accessExeFullPath -ExpectedCreationTimeUtcTicks $creationTimeUtcTicks)) -or
            ($processStartTimeUtcTicks -ne 0L -and
                (Test-ExpectedStartedProcessObject -Process $process -ExpectedPath $accessExeFullPath -ExpectedStartTimeUtcTicks $processStartTimeUtcTicks))
        )
        if ($safeToStop) {
            if (-not (Test-Path -LiteralPath $windowSnapshotPath)) {
                try {
                    & $windowSnapshotTool -ProcessId $process.Id -ExpectedAccessExe $accessExeFullPath -ExpectedCreationTimeUtcTicks $creationTimeUtcTicks -OutputPath $windowSnapshotPath
                    $result.window_enum = 'captured'
                }
                catch {
                    $result.window_enum = 'failed'
                }
            }
            $safeToStopAfterSnapshot =
                ($creationTimeUtcTicks -ne 0L -and
                    (Test-ExpectedAccessProcess -ProcessId $process.Id -ExpectedPath $accessExeFullPath -ExpectedCreationTimeUtcTicks $creationTimeUtcTicks)) -or
                ($processStartTimeUtcTicks -ne 0L -and
                    (Test-ExpectedStartedProcessObject -Process $process -ExpectedPath $accessExeFullPath -ExpectedStartTimeUtcTicks $processStartTimeUtcTicks))
            if ($safeToStopAfterSnapshot) {
                Stop-Process -Id $process.Id -Force
                Wait-Process -Id $process.Id -Timeout 15 -ErrorAction SilentlyContinue
            }
        }
    }

    if ($null -ne $watchdog) {
        try {
            if (-not $watchdog.HasExited) {
                Stop-Process -Id $watchdog.Id -Force
                [void]$watchdog.WaitForExit(5000)
            }
        }
        catch {}
    }

    if (Test-Path -LiteralPath $watchdogMarkerPath -PathType Leaf) {
        $watchdogResult = Read-Utf8JsonWithRetry -Path $watchdogMarkerPath
        $result.watchdog_fired = [bool]$watchdogResult.watchdog_fired
        $result.window_enum = [string]$watchdogResult.window_enum
    }
    if (Test-Path -LiteralPath $windowSnapshotPath -PathType Leaf) {
        $result.window_enum = 'captured'
    }

    $result.access_pid_gone = if ($null -eq $process) {
        $true
    }
    else {
        $null -eq (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)
    }

    $databaseDirectory = Split-Path -Parent $databaseFullPath
    $databaseBaseName = [IO.Path]::GetFileNameWithoutExtension($databaseFullPath)
    $remainingLocks = @(
        Join-Path $databaseDirectory ($databaseBaseName + '.laccdb')
        Join-Path $databaseDirectory ($databaseBaseName + '.ldb')
    ) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | ForEach-Object {
        $item = Get-Item -LiteralPath $_
        [pscustomobject][ordered]@{
            file_name = $item.Name
            byte_length = $item.Length
            last_write_time_utc = $item.LastWriteTimeUtc.ToString('o')
            sha256 = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash
        }
    }
    $result.lock_files_remaining = @($remainingLocks)
    $result.shift_released = (([AccessStartupBypassNative]::GetAsyncKeyState(0x10) -band 0x8000) -eq 0)
    $result.database_sha256_after = (Get-FileHash -LiteralPath $databaseFullPath -Algorithm SHA256).Hash
    $result.finished_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    $result.error = $validationError

    if ($null -eq $validationError -and
        -not $result.watchdog_fired -and
        $result.access_pid_gone -and
        $result.lock_files_remaining.Count -eq 0 -and
        $result.shift_released) {
        $result.status = 'PASS'
    }

    Write-Utf8Json -Path $validationPath -Value $result
}

if ($result.status -ne 'PASS') {
    throw "Startup bypass validation failed. Review: $validationPath"
}

Write-Host 'PASS: SKIP_AUTOEXEC was attested by the target Access process.'
Write-Host "Run ID: $($result.run_id)"
Write-Host "Output: $outputFullPath"
