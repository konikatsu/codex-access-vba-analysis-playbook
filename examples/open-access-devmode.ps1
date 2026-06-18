param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [int]$HoldSeconds = 5
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database not found: $DatabasePath"
}

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class KeyboardState {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    public const byte VK_SHIFT = 0x10;
    public const uint KEYEVENTF_KEYUP = 0x0002;

    public static void ShiftDown() {
        keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero);
    }

    public static void ShiftUp() {
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
"@

try {
    [KeyboardState]::ShiftDown()
    Start-Process -FilePath $DatabasePath
    Start-Sleep -Seconds $HoldSeconds
}
finally {
    [KeyboardState]::ShiftUp()
}
