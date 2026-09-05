[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [int]$ProcessId,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedAccessExe,

    [Parameter(Mandatory = $true)]
    [long]$ExpectedCreationTimeUtcTicks,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Write-Utf8Json {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $json = ConvertTo-Json -InputObject $Value -Depth 10
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, $utf8)
}

function Assert-RestrictedLocalPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\', [StringComparison]::Ordinal)) {
        throw 'OutputPath cannot use a UNC path.'
    }

    $root = [IO.Path]::GetPathRoot($fullPath)
    $drive = New-Object IO.DriveInfo($root)
    if ($drive.DriveType -ne [IO.DriveType]::Fixed) {
        throw "OutputPath requires a fixed local drive: $root"
    }

    $current = $fullPath
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "OutputPath cannot traverse a reparse point: $current"
            }
        }
        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }
}

if (-not (Test-Path -LiteralPath $ExpectedAccessExe -PathType Leaf)) {
    throw "Expected Access executable was not found: $ExpectedAccessExe"
}
$expectedExeFullPath = (Resolve-Path -LiteralPath $ExpectedAccessExe).Path

$outputFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
Assert-RestrictedLocalPath -Path $outputFullPath
$outputParent = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputParent -PathType Container)) {
    throw "Output directory does not exist: $outputParent"
}
if (Test-Path -LiteralPath $outputFullPath) {
    throw "OutputPath already exists: $outputFullPath"
}

$identity = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction SilentlyContinue
if ($null -eq $identity -or
    $identity.Name -ine 'MSACCESS.EXE' -or
    $identity.ExecutablePath -ine $expectedExeFullPath -or
    $null -eq $identity.CreationDate -or
    ([DateTime]$identity.CreationDate).ToUniversalTime().Ticks -ne $ExpectedCreationTimeUtcTicks) {
    throw 'The process identity does not match the recorded Access process.'
}

if (-not ('AccessWindowSnapshotNative' -as [type])) {
    Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class AccessWindowSnapshotNative
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder className, int count);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowEnabled(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint command);
}
'@
}

$windows = New-Object System.Collections.Generic.List[object]
$callback = [AccessWindowSnapshotNative+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)

    $ownerProcessId = 0
    [void][AccessWindowSnapshotNative]::GetWindowThreadProcessId($hWnd, [ref]$ownerProcessId)
    if ([int]$ownerProcessId -ne $ProcessId) {
        return $true
    }

    $captionLength = [AccessWindowSnapshotNative]::GetWindowTextLength($hWnd)
    $captionBuffer = New-Object Text.StringBuilder ([Math]::Max(1, $captionLength + 1))
    [void][AccessWindowSnapshotNative]::GetWindowText($hWnd, $captionBuffer, $captionBuffer.Capacity)

    $classBuffer = New-Object Text.StringBuilder 512
    [void][AccessWindowSnapshotNative]::GetClassName($hWnd, $classBuffer, $classBuffer.Capacity)

    $ownerHwnd = [AccessWindowSnapshotNative]::GetWindow($hWnd, 4)
    [void]$windows.Add([pscustomobject][ordered]@{
        hwnd = ('0x{0:X}' -f $hWnd.ToInt64())
        class_name = $classBuffer.ToString()
        caption = $captionBuffer.ToString()
        visible = [AccessWindowSnapshotNative]::IsWindowVisible($hWnd)
        enabled = [AccessWindowSnapshotNative]::IsWindowEnabled($hWnd)
        owner_hwnd = if ($ownerHwnd -eq [IntPtr]::Zero) { $null } else { '0x{0:X}' -f $ownerHwnd.ToInt64() }
    })
    return $true
}

if (-not [AccessWindowSnapshotNative]::EnumWindows($callback, [IntPtr]::Zero)) {
    throw 'EnumWindows failed.'
}

$payload = [ordered]@{
    schema_version = 1
    captured_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    process_id = $ProcessId
    process_creation_time_utc_ticks = $ExpectedCreationTimeUtcTicks
    access_executable_sha256 = (Get-FileHash -LiteralPath $expectedExeFullPath -Algorithm SHA256).Hash
    window_count = $windows.Count
    windows = @($windows.ToArray())
    interpretation = 'diagnostic-only'
}
Write-Utf8Json -Path $outputFullPath -Value $payload

Write-Host "Captured $($windows.Count) top-level window(s) for Access PID $ProcessId."
Write-Host "Output: $outputFullPath"
