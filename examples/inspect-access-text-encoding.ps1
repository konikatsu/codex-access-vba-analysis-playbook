param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [int]$SampleBytes = 200
)

$ErrorActionPreference = "Stop"

$resolved = (Resolve-Path -LiteralPath $Path).Path
$bytes = Get-Content -LiteralPath $resolved -Encoding Byte -TotalCount $SampleBytes

if ($null -eq $bytes) {
    throw "No bytes were read: $resolved"
}

$first16 = ($bytes | Select-Object -First 16 | ForEach-Object { $_.ToString("X2") }) -join " "
$nulCount = ($bytes | Where-Object { $_ -eq 0 }).Count

$guess = "Unknown"
if ($bytes.Count -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    $guess = "UTF-16 LE with BOM"
} elseif ($bytes.Count -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    $guess = "UTF-16 BE with BOM"
} elseif ($bytes.Count -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $guess = "UTF-8 with BOM"
} elseif ($nulCount -gt 0) {
    $guess = "Likely UTF-16 without decoded-as-text handling"
} else {
    $guess = "No BOM; CP932/SJIS or UTF-8 candidate"
}

[pscustomobject]@{
    Path = $resolved
    BytesRead = $bytes.Count
    First16 = $first16
    NulCountInSample = $nulCount
    Guess = $guess
} | Format-List
