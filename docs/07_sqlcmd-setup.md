# sqlcmdを使えるようにする

SQL ServerのDDL確認やクエリ実行では、`sqlcmd` があると便利です。

## インストール

Windowsでは、wingetで新しいGo版の `sqlcmd` を入れられます。

```powershell
winget install sqlcmd --accept-package-agreements --accept-source-agreements
```

インストール確認:

```powershell
sqlcmd --version
```

## 既存PowerShellセッションで見つからない場合

インストール直後、既に開いているPowerShellやCodexのシェルではPATHがまだ反映されず、次のようになることがあります。

```text
sqlcmd : The term 'sqlcmd' is not recognized
```

その場合は、セッション内でPATHを読み直します。

```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
sqlcmd --version
```

または、インストール先をフルパスで呼びます。

```powershell
& 'C:\Program Files\SqlCmd\sqlcmd.exe' --version
```

## Go版とODBC版

wingetで入る `Microsoft.Sqlcmd` はGo版のsqlcmdです。

通常のDDL確認やクエリ実行には使えますが、古いODBC版sqlcmd前提のスクリプトでは、細かいオプション互換に注意してください。

確認ポイント:

- `sqlcmd --version` が通るか
- 接続先SQL Serverへログインできるか
- 既存スクリプトで使っているオプションがGo版でも使えるか

## sqlcmdがTLSで失敗する場合

`sqlcmd` が次のようなエラーで失敗しても、すぐに「SQL Serverへ接続不能」と判断しないでください。

```text
TLS Handshake failed: cannot read handshake packet: EOF
```

Microsoft ODBC Driver 18以降では接続暗号化が既定で有効になり、証明書検証や暗号化設定の影響で接続に失敗する場合があります。
ただし、`TLS Handshake failed: EOF` が出た原因を証明書検証だけと断定しないでください。
sqlcmdの実装、バージョン、名前付きインスタンス、プロトコル解決、TLS設定なども切り分け対象です。

まず分けて記録します。

```text
確認済み:
推定:
暫定回避:
未確認:
```

WebアプリやPHPがODBC/PDOで接続できている環境では、先に同じ接続経路をPowerShellで再現すると早いです。
ODBC Driver 18を使っている場合の例:

```powershell
$conn = New-Object System.Data.Odbc.OdbcConnection
$conn.ConnectionString = "Driver={ODBC Driver 18 for SQL Server};Server=SERVER\INSTANCE;Database=DBNAME;Uid=USER;Pwd=PASSWORD;Encrypt=yes;TrustServerCertificate=yes;"
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT TOP (1) name FROM sys.tables"
$reader = $cmd.ExecuteReader()
while ($reader.Read()) { $reader.GetString(0) }
$reader.Close()
$conn.Close()
```

注意:

- パスワードや資格情報をチャットや公開ドキュメントに残さない。
- `sqlcmd` が失敗した事実と、ODBCで成功した事実を分ける。
- `-C` や暗号化オプションを試しても失敗する場合は、同じ操作を繰り返さず、接続経路を変えて実データ確認を先に進める。
- `sqlcmd` 自体を検証したい場合だけ、`sqlcmd --version`、`-N` / `-C`、サーバ名、ポート、インスタンス指定、ODBC Driverとの差を切り分ける。

参考:

- [ODBC Driver for SQL Server connection encryption troubleshooting](https://learn.microsoft.com/en-us/sql/connect/odbc/connection-troubleshooting?view=sql-server-ver17)
- [Certificate chain not trusted after driver upgrade](https://learn.microsoft.com/en-us/troubleshoot/sql/database-engine/connect/certificate-chain-not-trusted)

## Codex間で共有するときのメモ

別のCodexへ伝えるときは、次を渡すと十分です。

```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
sqlcmd --version
```

フルパス確認:

```powershell
& 'C:\Program Files\SqlCmd\sqlcmd.exe' --version
```
