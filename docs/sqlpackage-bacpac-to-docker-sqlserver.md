# sqlpackageで既存SQL Server DBをDocker SQL Serverへ移行する

既存のSQL Serverデータベースを、ローカルのDocker SQL Serverへコピーする手順です。
Accessのリンクテーブルや既存アプリのテスト環境を、本物のSQL Serverから切り離したい場合に使います。

## 使う方式

`sqlpackage` で既存DBを `.bacpac` にエクスポートし、Docker SQL Serverへインポートします。

```text
既存SQL Server DB
  -> sqlpackage Export
  -> .bacpac
  -> sqlpackage Import
  -> Docker SQL Server DB
```

この方式は、テーブル定義・インデックス・データをまとめて移せます。
元DBには読み取り中心でアクセスでき、Docker側に独立したテストDBを作れます。

## 前提

- Docker SQL Serverコンテナが起動していること
- `sqlcmd` が使えること
- `sqlpackage` が使えること
- 既存SQL Serverへ接続できる認証情報があること
- Docker SQL Server側の `sa` パスワードはチャットやログに出さず、ローカルファイルか環境変数で扱うこと

`sqlpackage` がない場合は、Windowsではwingetで入れられます。

```powershell
winget install Microsoft.SqlPackage --accept-package-agreements --accept-source-agreements
```

既存PowerShellセッションではPATHが更新されていない場合があります。

```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
sqlpackage /Version
```

## 1. Docker側の接続確認

```powershell
$env:SQLCMDPASSWORD = '<docker-sa-password>'
& 'C:\Program Files\SqlCmd\sqlcmd.exe' `
  -S localhost,14333 `
  -U sa `
  -C `
  -Q "SELECT @@VERSION"
Remove-Item Env:\SQLCMDPASSWORD -ErrorAction SilentlyContinue
```

## 2. 移行元DBのテーブル数を確認

```powershell
$sourceConnection = "Server=<source-server>;Database=<source-db>;User ID=<user>;Password=<password>;Encrypt=False;TrustServerCertificate=True;Connection Timeout=30;"

$query = @"
SELECT COUNT(*) AS user_tables
FROM sys.tables
WHERE is_ms_shipped = 0;
"@

Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection($sourceConnection)
$conn.Open()
try {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query
    $cmd.ExecuteScalar()
} finally {
    $conn.Close()
}
```

## 3. .bacpacへエクスポート

```powershell
New-Item -ItemType Directory -Force -Path .\sql_exports | Out-Null

$bacpac = ".\sql_exports\<source-db>_$(Get-Date -Format 'yyyyMMdd_HHmmss').bacpac"
$sourceConnection = "Server=<source-server>;Database=<source-db>;User ID=<user>;Password=<password>;Encrypt=False;TrustServerCertificate=True;Connection Timeout=30;"

sqlpackage `
  /Action:Export `
  /SourceConnectionString:"$sourceConnection" `
  /TargetFile:"$bacpac" `
  /p:VerifyExtraction=True
```

## 4. Docker SQL Serverへインポート

Docker側へインポートするときは、`Encrypt=False;TrustServerCertificate=True` を明示します。
環境によっては `Encrypt=True` のままだと、暗号化条件の不一致で長時間リトライしてから失敗することがあります。

```powershell
$targetConnection = "Server=localhost,14333;Database=<target-db>;User ID=sa;Password=<docker-sa-password>;Encrypt=False;TrustServerCertificate=True;Connection Timeout=30;"

sqlpackage `
  /Action:Import `
  /TargetConnectionString:"$targetConnection" `
  /SourceFile:"$bacpac"
```

## 5. 件数確認

インポート後、移行元と移行先でテーブル数・行数を比較します。

```sql
SELECT
    s.name + N'.' + t.name AS table_name,
    SUM(p.rows) AS row_count
FROM sys.tables t
JOIN sys.schemas s
    ON s.schema_id = t.schema_id
JOIN sys.partitions p
    ON p.object_id = t.object_id
   AND p.index_id IN (0, 1)
WHERE t.is_ms_shipped = 0
GROUP BY s.name, t.name
ORDER BY s.name, t.name;
```

Docker側で確認する例:

```powershell
$env:SQLCMDPASSWORD = '<docker-sa-password>'
& 'C:\Program Files\SqlCmd\sqlcmd.exe' `
  -S localhost,14333 `
  -U sa `
  -C `
  -d <target-db> `
  -Q "SELECT COUNT(*) AS user_tables FROM sys.tables WHERE is_ms_shipped = 0"
Remove-Item Env:\SQLCMDPASSWORD -ErrorAction SilentlyContinue
```

## 注意点

- 本物のSQL Serverを止める予定がある場合は、作業対象をDocker DBへ切り替えたことを関係者へ明確に共有する。
- Accessやアプリの接続先は、名前付きインスタンス形式ではなく `localhost,14333` のような `server,port` 形式にする。
- `.bacpac` はデータを含むため、公開リポジトリへコミットしない。
- パスワードはチャット、コミット、ログに残さない。
- `sqlpackage Import` で暗号化エラーが出る場合は、まず `Encrypt=False;TrustServerCertificate=True` を試す。

