# Codex Access/VBA Analysis Playbook

Access/VBA の既存システムを Codex などのAIエージェントで解析・修正するときの実践メモです。

特に、以下のような作業で詰まりやすい点をまとめています。

- Access DB のフォーム、標準モジュール、クエリ、テーブル定義を静的解析する
- `Access.Application` の外部COMから `LoadFromText` / `SaveAsText` を使う
- 解析用モジュールを一時的に取り込んで実行する
- 64bit Access 対応で `Declare PtrSafe` / `LongPtr` を確認する
- 外部COMからVBAをコンパイルする

## まず読む

- [Access解析プレイブック](docs/access-analysis-playbook.md)
- [Access COM自動化の基本](docs/access-com-automation.md)
- [LoadFromTextトラブルシュート](docs/loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイルする](docs/compile-with-runcommand.md)
- [解析情報をエクスポートする考え方](docs/export-analysis-info.md)
- [sqlcmdを使えるようにする](docs/sqlcmd-setup.md)

## 重要な教訓

Access外部COMでは、次を別々に切り分けます。

- `LoadFromText` が成功するか
- `SaveAsText` が成功するか
- `Application.Run` が成功するか
- VBEオブジェクトモデルでコードを読めるか
- `RunCommand(126)` でコンパイルできるか

また、`OpenCurrentDatabase` の前に `AutomationSecurity = 1` を設定します。

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)
```

サンドボックス環境や制限付き実行では、`LoadFromText` が `予約済みエラー` になることがあります。  
この場合、DB破損やファイル形式の問題と決めつけず、権限付き実行で再試行します。

## サンプル

- [LoadFromTextで標準モジュールを取り込む](examples/loadfromtext-module.ps1)
- [Access VBAをコンパイルする](examples/compile-access-vba.ps1)
- [Shift-bypassでAccess DBを開く](examples/open-access-devmode.ps1)
- [AutomationSecurity=3でAccess DBを開く](examples/open-access-no-autoexec.ps1)
- [/cmd SKIP_AUTOEXEC でAccess DBを開く](examples/open-access-skip-autoexec.ps1)

## 注意

このリポジトリは実案件のDBや業務ロジックを含みません。  
公開用に汎用化した手順メモです。

実DBを扱う場合は、必ずバックアップまたはコピーで検証してください。
