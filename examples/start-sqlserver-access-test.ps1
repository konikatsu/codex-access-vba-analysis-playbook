param(
    [string]$ContainerName = "sqlserver-access-test01",
    [string]$VolumeName = "sqlserver-access-test01-data",
    [int]$Port = 14333,
    [string]$DatabaseName = "access_test",
    [string]$Image = "mcr.microsoft.com/mssql/server:2022-latest"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$localDir = Join-Path $root ".local"
$credentialPath = Join-Path $localDir "$ContainerName.credentials.txt"
$sqlcmd = "C:\Program Files\SqlCmd\sqlcmd.exe"

if (-not (Test-Path $sqlcmd)) {
    throw "sqlcmd not found: $sqlcmd"
}

New-Item -ItemType Directory -Force -Path $localDir | Out-Null

function New-SaPassword {
    $guidPart = ([guid]::NewGuid().ToString("N") + [guid]::NewGuid().ToString("N")).Substring(0, 24)
    return "Cdx!A1$guidPart"
}

if (Test-Path $credentialPath) {
    $saPassword = Select-String -Path $credentialPath -Pattern "^SA_PASSWORD=(.*)$" |
        ForEach-Object { $_.Matches[0].Groups[1].Value } |
        Select-Object -First 1

    if (-not $saPassword) {
        throw "Credential file exists but SA_PASSWORD is missing: $credentialPath"
    }
} else {
    $saPassword = New-SaPassword
    @(
        "# Local test SQL Server credentials. Do not commit this file."
        "SERVER=localhost,$Port"
        "USER=sa"
        "SA_PASSWORD=$saPassword"
        "DATABASE=$DatabaseName"
        "CONTAINER=$ContainerName"
        "VOLUME=$VolumeName"
    ) | Set-Content -Path $credentialPath -Encoding ASCII
}

$existingContainer = docker ps -a --filter "name=^/$ContainerName$" --format "{{.Names}}"

if ($existingContainer -eq $ContainerName) {
    docker start $ContainerName | Out-Null
} else {
    docker run `
        -e "ACCEPT_EULA=Y" `
        -e "MSSQL_SA_PASSWORD=$saPassword" `
        -e "MSSQL_PID=Developer" `
        -p "${Port}:1433" `
        --name $ContainerName `
        --hostname $ContainerName `
        -v "${VolumeName}:/var/opt/mssql" `
        -d $Image | Out-Null
}

$env:SQLCMDPASSWORD = $saPassword
try {
    $ready = $false
    for ($i = 1; $i -le 90; $i++) {
        & $sqlcmd -S "localhost,$Port" -U sa -C -Q "SELECT 1" -b | Out-Null
        if ($LASTEXITCODE -eq 0) {
            $ready = $true
            break
        }
        Start-Sleep -Seconds 2
    }

    if (-not $ready) {
        docker logs --tail 80 $ContainerName
        throw "SQL Server did not become ready in time."
    }

    $createDbSql = @"
IF DB_ID(N'$DatabaseName') IS NULL
BEGIN
    EXEC(N'CREATE DATABASE [$DatabaseName]');
END
SELECT @@VERSION AS version;
SELECT name FROM sys.databases WHERE name = N'$DatabaseName';
"@

    & $sqlcmd -S "localhost,$Port" -U sa -C -Q $createDbSql -b
    if ($LASTEXITCODE -ne 0) {
        throw "sqlcmd verification failed."
    }
} finally {
    Remove-Item Env:\SQLCMDPASSWORD -ErrorAction SilentlyContinue
}

Write-Host "SQL Server container is ready."
Write-Host "Server: localhost,$Port"
Write-Host "Database: $DatabaseName"
Write-Host "Credential file: $credentialPath"

