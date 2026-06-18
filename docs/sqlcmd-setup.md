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
