# Accessリンクテーブル向けSQL Serverテスト環境の選び方

AccessアプリケーションがSQL Serverのリンクテーブルを使っている場合に、
ローカルのテスト用SQL Serverをどう用意するかの判断メモです。

## 結論

Access側のリンクテーブルを張り直せるなら、DockerでSQL Serverを動かす方法は管理しやすいです。

Docker方式では、接続先は次のような `サーバ名,ポート番号` 形式にします。

```text
localhost,14333
```

または、ローカルのhostsやDNSで別名を付けるなら次のようにできます。

```text
app-sql,14333
```

一方で、既存アプリケーションが次のようなWindows SQL Serverの名前付きインスタンス形式を
どうしても維持する必要がある場合は、DockerではなくSQL Server Developer Editionや
SQL Server Expressをローカルに名前付きインスタンスとしてインストールする方が素直です。

```text
server-name\instance-name
```

## Dockerで名前付きインスタンス形式が難しい理由

SQL Serverコンテナは、基本的に「1コンテナ = 1つの既定インスタンス」として扱います。

ホスト側のポートを、コンテナ内のSQL Serverポートへ割り当てます。

```text
ホスト側 14333 -> コンテナ側 1433
```

そのため、クライアントからは次の形式で接続します。

```text
server,port
```

名前付きインスタンスの接続形式である次の形ではありません。

```text
server\instance
```

`server\instance` 形式は、SQL Server Browserによるインスタンス名からポート番号への解決に依存します。
Linux版SQL Serverコンテナの周辺でこの挙動を再現することも不可能ではありませんが、
開発・テスト環境としては手間が大きく、割に合わないことが多いです。

## Dockerで作る場合の基本形

テストDBやプロジェクトごとに、コンテナ名・ポート・volumeを分けます。

```powershell
docker run `
  -e "ACCEPT_EULA=Y" `
  -e "MSSQL_SA_PASSWORD=<strong-password>" `
  -e "MSSQL_PID=Developer" `
  -p 14333:1433 `
  --name sqlserver-test01 `
  --hostname sqlserver-test01 `
  -v sqlserver-test01-data:/var/opt/mssql `
  -d mcr.microsoft.com/mssql/server:2022-latest
```

接続確認は `sqlcmd` で行えます。

```powershell
& 'C:\Program Files\SqlCmd\sqlcmd.exe' `
  -S localhost,14333 `
  -U sa `
  -P '<strong-password>' `
  -C `
  -Q "SELECT @@VERSION"
```

volumeを使うことが重要です。
volumeなしでコンテナを削除すると、コンテナ内のデータベースファイルも失われます。

## Accessリンクテーブルの張り直し方針

Docker方式にする場合は、Access側のリンクテーブルをコンテナの接続先へ張り直します。

```text
変更前: server-name\instance-name
変更後: localhost,14333
```

または、ホスト名を寄せたい場合は次のようにします。

```text
変更後: app-sql,14333
```

ホスト名の別名を使う場合は、ローカルのhostsまたはDNSでDockerホストを指すようにします。
ただし、接続形式はあくまでカンマ付きの `server,port` です。
バックスラッシュ付きの `server\instance` ではありません。

## 判断基準

- Accessのリンクテーブルを `server,port` 形式へ張り直せるなら、Docker方式が扱いやすい。
- 名前付きインスタンス形式 `server\instance` を維持したいなら、通常インストールのSQL Serverを使う。
- 多数のテストDBを作るなら、SQL Server Developer EditionのDockerコンテナが便利。
- Expressの制限に近い環境を再現したいなら、SQL Server Expressの通常インストールも候補になる。
- LocalDBは単独開発者の軽い検証向けで、Accessリンクテーブルの本格的な接続先としては優先度が下がる。
- Docker版SQL ServerにWindowsの名前付きインスタンスの挙動を無理に再現させるのは避ける。

## 実務上のおすすめ手順

Accessシステムのテスト環境では、次の流れが単純です。

1. DockerでSQL Serverを固定ポート付きで起動する。
2. そのコンテナ内にテストDBを作る、またはバックアップを復元する。
3. Accessのリンクテーブルを `localhost,<port>` または別名ホストの `server,<port>` に張り直す。
4. プロジェクトごとに、コンテナ名・ポート番号・volume名を分ける。

## 今回の構築例

実際に作った構成例です。

サンプルスクリプト:

```text
examples/start-sqlserver-access-test.ps1
```

実行例:

```powershell
powershell -ExecutionPolicy Bypass -File ".\examples\start-sqlserver-access-test.ps1"
```

```text
コンテナ名: sqlserver-access-test01
volume名: sqlserver-access-test01-data
ホスト側ポート: 14333
SQL Server: 2022 Developer
接続先: localhost,14333
DB名: access_test
```

起動確認:

```powershell
docker ps
```

表示例:

```text
sqlserver-access-test01   mcr.microsoft.com/mssql/server:2022-latest   Up ...   0.0.0.0:14333->1433/tcp
```

接続確認:

```powershell
& 'C:\Program Files\SqlCmd\sqlcmd.exe' `
  -S localhost,14333 `
  -U sa `
  -C `
  -Q "SELECT @@VERSION"
```

パスワードをコマンドラインに直接残したくない場合は、`SQLCMDPASSWORD` 環境変数を使います。

```powershell
$env:SQLCMDPASSWORD = '<strong-password>'
& 'C:\Program Files\SqlCmd\sqlcmd.exe' -S localhost,14333 -U sa -C -Q "SELECT name FROM sys.databases"
Remove-Item Env:\SQLCMDPASSWORD
```

## Docker Desktopの注意

Docker Desktopの初回起動時にサインイン画面が出ることがあります。
ローカルのテスト用コンテナを動かすだけなら、サインインをスキップして進められる場合があります。

Codexなどの制限付き環境からDockerを操作する場合、通常権限ではDocker Desktopの名前付きパイプに
アクセスできず、次のようなエラーになることがあります。

```text
open //./pipe/dockerDesktopLinuxEngine: Access is denied.
```

その場合は、Docker操作を権限付きで実行します。

## 参考

- Microsoft Learn: [Quickstart: Run SQL Server Linux container images with Docker](https://learn.microsoft.com/en-us/sql/linux/install-upgrade/quickstart-install-docker)
