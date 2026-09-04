# Codex Access/VBA Analysis Playbook

Microsoft Access / VBA の既存システムを、Codex などのAIエージェントで解析・修正・検証するときの実践メモです。

このリポジトリは、特定の顧客名、DB名、サーバ名、実データを含めない公開用のノウハウ集です。

## まず読む

- [Access作業共通ルール](docs/00_access-work-common-rules.md)
- [Access修正の標準開発手順](docs/15_access-development-workflow.md)
- [Access外部Exportツール](docs/16_access-external-export.md)
- [DB内ExportAnalysisInfoによる初見解析](docs/01_export-analysis-info.md)

## 要件別に読む

- [コンテンツの有効化を確認する](docs/requirements/00_enable-active-content.md)
- [起動処理を止めて開発モードで開く](docs/requirements/01_startup-bypass.md)
- [ExportAnalysisInfoを使える状態にする手順](docs/requirements/02_import-analysis-module.md)
- [Access資産をAI向けにエクスポートする](docs/requirements/03_export-access-assets-for-ai.md)
- [作業コピーと進捗メモで安全に進める](docs/requirements/04_work-copy-and-progress.md)
- [フォーム/レポート/モジュールをLoadFromTextで差し替える](docs/requirements/05_loadfromtext-replace-object.md)
- [別PC・別Codexへ申し送りする](docs/requirements/06_handoff-to-another-ai-agent.md)
- [空DBインポートによる救出解析](docs/requirements/07_empty-database-recovery-import.md)

## 関連メモ

- [Access/VBA作業プレイブック](docs/02_access-ai-agent-workflow.md)
- [Access作業共通ルール](docs/00_access-work-common-rules.md)
- [Access解析の基本](docs/03_access-analysis-playbook.md)
- [Access COM自動化の基本](docs/04_access-com-automation.md)
- [LoadFromTextトラブルシュート](docs/05_loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイルする](docs/06_compile-with-runcommand.md)
- [sqlcmdを使えるようにする](docs/07_sqlcmd-setup.md)
- [Accessリンクテーブル向けSQL Serverテスト環境](docs/08_sql-server-test-environment-for-access.md)
- [sqlpackageで既存SQL Server DBをDocker SQL Serverへ移行する](docs/09_sqlpackage-bacpac-to-docker-sqlserver.md)
- [Accessテキスト資産の文字コード](docs/10_access-text-encoding.md)
- [Accessフォームデザイン確認](docs/11_access-form-design.md)
- [Access VBA 64bit移行](docs/11_access-vba-64bit-migration.md)
- [DOCX成果物のレンダリングQA](docs/12_docx-rendering-qa.md)
- [Access Web化UI表示崩れの切り分け](docs/13_access-web-ui-troubleshooting.md)
- [Gemini CLI / Antigravity CLI連携](docs/14_gemini-cli-collaboration.md)

## 重要な方針

- 本体DBを直接触らず、必ず作業コピーで検証する。
- 依頼時に既知の自動起動情報を渡し、作業担当は実装前に自動起動の有無、経路、無効化方法、検証証跡、採用baselineを宣言する。`未確認`のまま実装へ進まない。
- 新しい原本を初めて扱うときは、DAOとStartupProbeで自動起動を最初に調べる。呼出先関数に完全一致の`SKIP_AUTOEXEC`分岐がなければ作業コピーへ一度だけ追加し、既にあれば重複追加しない。
- 通常起動とスキップ起動を確認した自動起動無効化対応版を開発baselineにし、以後の候補はそこから作る。
- 元ACCDBのSHA-256と検証記録が一致する再作業では、StartupProbeやIF追加を繰り返さず、そのbaselineを再利用する。
- 正式なbaselineと候補の全資産Exportは、DBを変更しない外部ツールを同じバージョン、同じ条件で使う。
- DB内の`ExportAnalysisInfo`は、解析専用コピーでの初見解析や救出調査に限定する。正式差分のbaselineへ解析モジュールを取り込まない。
- 成功した作業コピーを次の土台にする。
- 失敗した作業コピーは修復しながら続けず、破棄する。
- フォームやレポートの差し替えは、作業コピー上なら `DeleteObject -> LoadFromText -> Compile` でよい。
- 標準開発手順では、変更対象の`SaveAsText`と予定diffの一致確認を必須にする。使い捨て調査でDBコピー単位の復旧だけを目的とする場合は省略できる。
- `SaveAsText` 出力は文字コードを決め打ちしない。先頭バイトを確認し、UTF-16 LE / UTF-8 / CP932を切り分ける。
- 文字コード確認は、長い `powershell.exe -Command ... ReadAllBytes ... ToString('X2')` ではなく、名前付きスクリプトや読みやすい短いコマンドで行う。
- GUIテストが必要な画面は、AIエージェントでも押しやすい固定ボタン、十分な幅、ツールチップを用意する。
- Accessフォームのデザイン変更は、基準フォームの行高、色、枠線、フォント、列幅を採寸し、キャプチャ比較で確認する。
- 明細グリッドに属する合計行は、親フォームではなくサブフォームのフッターへ置く。
- Accessフォームの自動テストでは、物理クリックや `SendKeys` に依存せず、Clickイベントから呼ぶ実処理をPublicメソッド化して直接テストする。
- 最小修正は成功baselineから名前付き作業コピーを作り、コンパイルは `AutomationSecurity = 1` で再オープン後まで確認する。`AutomationSecurity = 3` の例外なしだけをコンパイルPASSとしない。
- スキーマ限定変更は、対象制約だけの冪等DDL、変更前後の行数・制約数比較、全`SaveAsText`差分を基本ゲートにする。新機能が既存機能の更新・削除を拘束する変更は原則禁止する。
- DOCX成果物は、PDF/PNGへレンダリングして表、図、改ページの崩れを確認する。WindowsではLibreOfficeのCLIに `soffice.com` を優先する。
- Web化画面の表示崩れは、見た目だけでデータ起因と断定せず、DB実値、HTML構造、CSSの順で切り分ける。
- Gemini CLI / Antigravity CLIは調査・要約・一次レビューの補助として使い、最終判断、編集、テスト、pushはCodexが行う。個人Google OAuthのGemini CLI利用は対象外のため、Antigravity CLIへ移行する。
- 修正案は、セルフレビュー、独立レビュー、指摘の採用・棄却・要検証、修正後再検証を通す。レビュアーの出力を根拠確認なしで反映しない。
- 進捗は Markdown に残し、分母・分子が分かる形にする。

## Access COMで最初に確認すること

Access 外部COMでは、次を別々に切り分けます。

- `LoadFromText` が成功するか
- `SaveAsText` が成功するか
- `Application.Run` が成功するか
- VBEオブジェクトモデルでコードを読めるか
- `RunCommand(126)` でコンパイルできるか

`OpenCurrentDatabase` の前に、用途に応じて `AutomationSecurity` を設定します。次のコードは値の説明用で、起動バイパス、PID記録、タイムアウトを省略しています。実作業では[標準開発手順](docs/15_access-development-workflow.md)を優先します。

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)
```

使い分けの目安:

- `AutomationSecurity = 1`: モジュール取り込み、`Application.Run`、コンパイル向け
- `AutomationSecurity = 3`: COM静的操作でマクロを無効化する設定。起動バイパスは仮想Shiftなどを別途使う
- DAO `OpenDatabase(copy, False, True)`: 構造メタデータだけをAccess非起動で読む
- `/cmd SKIP_AUTOEXEC`: DB側が`Command()`の完全一致分岐を実装済みの場合の開発モード起動

## サンプル

- [標準モジュールをLoadFromTextで取り込む](examples/loadfromtext-module.ps1)
- [Access VBAをコンパイルする](examples/compile-access-vba.ps1)
- [Accessテキスト資産の文字コードを確認する](examples/inspect-access-text-encoding.ps1)
- [Shift-bypassでAccess DBを開く](examples/open-access-devmode.ps1)
- [AutomationSecurity=3の制約を確認してAccess DBを開く](examples/open-access-no-autoexec.ps1)
- [/cmd SKIP_AUTOEXEC でAccess DBを開く](examples/open-access-skip-autoexec.ps1)
- [Docker SQL Serverテスト環境を起動する](examples/start-sqlserver-access-test.ps1)
- [Access資産を外部ツールでExportする](examples/export-access-assets.ps1)

## 注意

このリポジトリは公開用に一般化したメモです。実案件のDB名、サーバ名、ユーザー名、パスワード、業務ロジック固有名は入れないでください。

実DBを扱う場合は、必ずバックアップまたは作業コピーで検証してください。
