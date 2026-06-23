# Codex Access/VBA Analysis Playbook

Microsoft Access / VBA の既存システムを、Codex などのAIエージェントで解析・修正・検証するときの実践メモです。

このリポジトリは、特定の顧客名、DB名、サーバ名、実データを含めない公開用のノウハウ集です。

## まず読む

- [Access作業共通ルール](docs/access-work-common-rules.md)
- [Access資産をAI(Codex)にエクスポートさせる手順](docs/export-analysis-info.md)

## 要件別に読む

- [Access資産をAI向けにエクスポートする](docs/requirements/export-access-assets-for-ai.md)
- [ExportAnalysisInfoを使える状態にする手順](docs/requirements/import-analysis-module.md)
- [起動処理を止めて開発モードで開く](docs/requirements/startup-bypass.md)
- [フォーム/レポート/モジュールをLoadFromTextで差し替える](docs/requirements/loadfromtext-replace-object.md)
- [別PC・別Codexへ申し送りする](docs/requirements/handoff-to-another-ai-agent.md)
- [作業コピーと進捗メモで安全に進める](docs/requirements/work-copy-and-progress.md)

## 関連メモ

- [Access/VBA作業プレイブック](docs/access-ai-agent-workflow.md)
- [Access作業共通ルール](docs/access-work-common-rules.md)
- [Access解析の基本](docs/access-analysis-playbook.md)
- [Access COM自動化の基本](docs/access-com-automation.md)
- [LoadFromTextトラブルシュート](docs/loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイルする](docs/compile-with-runcommand.md)
- [sqlcmdを使えるようにする](docs/sqlcmd-setup.md)
- [Accessリンクテーブル向けSQL Serverテスト環境](docs/sql-server-test-environment-for-access.md)
- [sqlpackageで既存SQL Server DBをDocker SQL Serverへ移行する](docs/sqlpackage-bacpac-to-docker-sqlserver.md)

## 重要な方針

- 本体DBを直接触らず、必ず作業コピーで検証する。
- Access資産エクスポートは、対象DBで `ExportAnalysisInfo` を実行して `Latest` フォルダを作る。
- 成功した作業コピーを次の土台にする。
- 失敗した作業コピーは修復しながら続けず、破棄する。
- フォームやレポートの差し替えは、作業コピー上なら `DeleteObject -> LoadFromText -> Compile` でよい。
- 差し替え前の `SaveAsText` は必須ではない。DBコピー単位で戻れる運用を優先する。
- GUIテストが必要な画面は、AIエージェントでも押しやすい固定ボタン、十分な幅、ツールチップを用意する。
- 進捗は Markdown に残し、分母・分子が分かる形にする。

## Access COMで最初に確認すること

Access 外部COMでは、次を別々に切り分けます。

- `LoadFromText` が成功するか
- `SaveAsText` が成功するか
- `Application.Run` が成功するか
- VBEオブジェクトモデルでコードを読めるか
- `RunCommand(126)` でコンパイルできるか

`OpenCurrentDatabase` の前に、用途に応じて `AutomationSecurity` を設定します。

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)
```

使い分けの目安:

- `AutomationSecurity = 1`: モジュール取り込み、`Application.Run`、コンパイル向け
- `AutomationSecurity = 3`: GUI確認時にマクロを止めたい場合の候補
- `/cmd SKIP_AUTOEXEC`: 開発モード起動の標準手段としておすすめ

## サンプル

- [標準モジュールをLoadFromTextで取り込む](examples/loadfromtext-module.ps1)
- [Access VBAをコンパイルする](examples/compile-access-vba.ps1)
- [Shift-bypassでAccess DBを開く](examples/open-access-devmode.ps1)
- [AutomationSecurity=3でAccess DBを開く](examples/open-access-no-autoexec.ps1)
- [/cmd SKIP_AUTOEXEC でAccess DBを開く](examples/open-access-skip-autoexec.ps1)
- [Docker SQL Serverテスト環境を起動する](examples/start-sqlserver-access-test.ps1)

## 注意

このリポジトリは公開用に一般化したメモです。実案件のDB名、サーバ名、ユーザー名、パスワード、業務ロジック固有名は入れないでください。

実DBを扱う場合は、必ずバックアップまたは作業コピーで検証してください。
