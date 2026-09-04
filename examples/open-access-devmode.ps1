[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [string]$AccessExe = 'msaccess.exe',

    [ValidateRange(10, 120)]
    [int]$StartupTimeoutSeconds = 30,

    [switch]$AcknowledgeCurrentDesktopInput
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) } catch {}
    }
}

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

if (-not $AcknowledgeCurrentDesktopInput) {
    throw 'This helper injects Shift on the current input desktop. Close or pause other interactive work, then rerun with -AcknowledgeCurrentDesktopInput.'
}
if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) {
    throw "Database not found: $DatabasePath"
}

$databaseFullPath = (Resolve-Path -LiteralPath $DatabasePath).Path
$accessExeFullPath = Resolve-AccessExecutable -RequestedPath $AccessExe
$existingAccess = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
if ($existingAccess.Count -gt 0) {
    throw 'Close existing Access processes before using this helper so the launched process is unambiguous.'
}

$engine = $null
$database = $null
$allowBypassProperty = $null
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $database = $engine.OpenDatabase($databaseFullPath, $false, $true)
    try {
        $allowBypassProperty = $database.Properties.Item('AllowBypassKey')
        if (-not [Convert]::ToBoolean($allowBypassProperty.Value, [Globalization.CultureInfo]::InvariantCulture)) {
            throw 'AllowBypassKey is explicitly False; Shift-bypass is not available for this database.'
        }
    }
    catch {
        $daoErrorNumber = $_.Exception.HResult -band 0xFFFF
        if ($daoErrorNumber -notin @(3265, 3270)) {
            throw
        }
    }
}
finally {
    if ($null -ne $database) {
        try { $database.Close() } catch {}
    }
    Release-ComObject $allowBypassProperty
    Release-ComObject $database
    Release-ComObject $engine
}

if (-not ('AccessDevModeKeyboard' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class AccessDevModeKeyboard
{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int virtualKey);

    public static void ShiftDown() { keybd_event(0x10, 0, 0, UIntPtr.Zero); }
    public static void ShiftUp() { keybd_event(0x10, 0, 0x0002, UIntPtr.Zero); }
}
'@
}

if (([AccessDevModeKeyboard]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0) {
    throw 'Shift is already pressed. Release physical and virtual Shift before continuing.'
}

$watchdogId = [guid]::NewGuid().ToString('N')
$watchdogPath = Join-Path $env:TEMP "access-shift-release-$watchdogId.ps1"
$watchdogMarker = Join-Path $env:TEMP "access-shift-release-$watchdogId.fired"
$watchdogSource = @'
param([int]$TimeoutSeconds, [string]$MarkerPath)
$ErrorActionPreference = 'SilentlyContinue'
Start-Sleep -Seconds $TimeoutSeconds
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class AccessShiftReleaseWatchdog
{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@
[AccessShiftReleaseWatchdog]::keybd_event(0x10, 0, 0x0002, [UIntPtr]::Zero)
[IO.File]::WriteAllText($MarkerPath, 'watchdog_fired=true', (New-Object Text.UTF8Encoding($false)))
'@
[IO.File]::WriteAllText($watchdogPath, $watchdogSource, (New-Object Text.UTF8Encoding($false)))

$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$watchdogArguments = @(
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', ('"{0}"' -f $watchdogPath),
    '-TimeoutSeconds', $StartupTimeoutSeconds,
    '-MarkerPath', ('"{0}"' -f $watchdogMarker)
)
$watchdog = Start-Process -FilePath $windowsPowerShell -ArgumentList $watchdogArguments -WindowStyle Hidden -PassThru
$process = $null
$shiftIsDown = $false
$launchIdentity = $null

try {
    [AccessDevModeKeyboard]::ShiftDown()
    $shiftIsDown = $true
    $arguments = @(('"{0}"' -f $databaseFullPath))
    $process = Start-Process -FilePath $accessExeFullPath -ArgumentList $arguments -PassThru

    $launchIdentity = Get-CimInstance Win32_Process -Filter "ProcessId=$($process.Id)" -ErrorAction Stop
    if ($launchIdentity.Name -ine 'MSACCESS.EXE' -or
        $launchIdentity.ExecutablePath -ine $accessExeFullPath -or
        $null -eq $launchIdentity.CreationDate) {
        throw 'The launched process identity could not be verified as the expected Access executable.'
    }

    if (-not $process.WaitForInputIdle($StartupTimeoutSeconds * 1000)) {
        throw "Access did not reach input-idle state within $StartupTimeoutSeconds seconds."
    }
    Start-Sleep -Milliseconds 500
    if (Test-Path -LiteralPath $watchdogMarker) {
        throw 'The external Shift-release watchdog fired before Access reached a stable GUI state.'
    }
}
catch {
    if ($null -ne $process) {
        if ($null -ne $launchIdentity) {
            $candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$($process.Id)" -ErrorAction SilentlyContinue
            $expectedCreationTicks = ([DateTime]$launchIdentity.CreationDate).ToUniversalTime().Ticks
            if ($null -ne $candidate -and
                $candidate.Name -ieq 'MSACCESS.EXE' -and
                $candidate.ExecutablePath -ieq $accessExeFullPath -and
                $null -ne $candidate.CreationDate -and
                ([DateTime]$candidate.CreationDate).ToUniversalTime().Ticks -eq $expectedCreationTicks) {
                Stop-Process -Id $process.Id -Force
                Wait-Process -Id $process.Id -Timeout 15 -ErrorAction SilentlyContinue
            }
        }
        else {
            try {
                if (-not $process.HasExited) {
                    $process.Kill()
                    [void]$process.WaitForExit(15000)
                }
            }
            catch {}
        }
    }
    throw
}
finally {
    if ($shiftIsDown) {
        [AccessDevModeKeyboard]::ShiftUp()
    }
    if ($null -ne $watchdog) {
        try {
            if (-not $watchdog.HasExited) {
                Stop-Process -Id $watchdog.Id -Force
                Wait-Process -Id $watchdog.Id -Timeout 10 -ErrorAction SilentlyContinue
            }
        }
        catch {}
    }
    Remove-Item -LiteralPath $watchdogPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $watchdogMarker -Force -ErrorAction SilentlyContinue
}

Write-Warning 'Shift-bypass was requested and bounded, but the target database must still provide its own evidence that startup side effects did not run.'
Write-Host "Access opened for the one-time startup bypass edit. PID: $($process.Id)"
