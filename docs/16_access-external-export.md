# Access外部Exportツール

`examples/export-access-assets.ps1`は、正式なbaselineと実装候補を同じ条件で比較するための標準Exportツールです。

DB内へ解析モジュールを取り込まず、Access外のPowerShellから`SaveAsText`とDAOメタデータ取得を直接実行します。元ACCDBはハッシュ取得と二次コピー作成にしか使わず、Accessで開くのは二次コピーだけです。

## 用途

- 最初の自動起動調査
- baselineの初回全資産Export
- 実装候補の最終全資産Export
- 変更対象を含む資産の機械可読な棚卸し

初見調査のためにDB内へ`ExportAnalysisInfo`を取り込む方法とは用途が違います。正式差分では、この外部ツールをbaselineと候補の両方に同じバージョン、同じ条件で使います。

## 最初の自動起動調査

ほかの修正より先に`StartupProbe`を実行します。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\examples\export-access-assets.ps1" `
  -DatabasePath "C:\work\project_work\stepNNN_baseline\01_source_copy\app.source-copy.accdb" `
  -OutputDirectory "C:\work\project_work\stepNNN_baseline\02_startup_probe" `
  -Mode StartupProbe
```

`StartupProbe`は全マクロ、全フォーム、全標準/クラスモジュール、全保存クエリ、ローカルテーブルのデータマクロ、起動設定、オブジェクトカタログを出力し、レポートとテーブル等の詳細メタデータは省略します。`catalog.json`、AutoExecの定義、起動フォームやAutoExecから開かれるフォームの`Open`/`Load`イベントから呼出先関数を追い、その先頭に完全一致の`SKIP_AUTOEXEC`分岐があるか確認します。AutoExecがレポートを直接開くなど省略対象へ続く場合は、判定前に`Full`で追加確認します。

自動起動がなければ追加作業は不要です。既に同等分岐があれば重複追加しません。分岐がない場合だけ作業コピーへ一度追加し、通常起動とスキップ起動を別々に確認してから開発baselineにします。

## 実行例

Windows PowerShellとOffice/ACEのbitnessを合わせます。
`-OutputDirectory`には、まだ存在しない新しいディレクトリを指定します。ツールが出力先と`_working`を作成します。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\examples\export-access-assets.ps1" `
  -DatabasePath "C:\work\project_work\stepNNN_baseline\04_baseline\app.baseline.accdb" `
  -OutputDirectory "C:\work\project_work\stepNNN_baseline\05_before_export"
```

パスワード付きDBでは、平文をコマンド履歴へ残さず`SecureString`として渡します。

```powershell
$password = Read-Host 'Database password' -AsSecureString
& '.\examples\export-access-assets.ps1' `
  -DatabasePath 'C:\work\project_work\stepNNN_baseline\04_baseline\app.baseline.accdb' `
  -OutputDirectory 'C:\work\project_work\stepNNN_baseline\05_before_export' `
  -DatabasePassword $password
```

リンクテーブルのフィールドやインデックスは、ドライバーが接続先へ到達する場合があります。既定ではリンク先詳細を読みません。隔離した検証環境で必要性を確認した場合だけ`-IncludeLinkedTableDetails`を使います。

## 起動と後片付け

ツールは次の順序で動きます。

1. 元ACCDBのSHA-256を記録する。
2. 出力先の`_working`へ二次コピーを作り、コピー直後のSHA-256を照合する。
3. DAOで二次コピーの`AllowBypassKey`を読み、明示的な`False`ならAccess起動前に失敗する。
4. 専用のAccess COMインスタンスを作り、`hWndAccessApp`からPID、実行ファイル、生成時刻を記録する。
5. `AutomationSecurity=3`を設定する。
6. 外部ウォッチドッグを起動し、仮想Shiftを押して二次コピーを開く。
7. `OpenCurrentDatabase`復帰直後、例外時、時間上限到達時のすべてでShiftを解放する。
8. 起動直後にフォームが1件でも開いていればFAILとして後続Exportへ進まない。
9. `SaveAsText`とメタデータ取得を行い、Accessを閉じる。
10. 専用PID、ロック、元ACCDBのSHA-256、manifestを検証する。

ウォッチドッグは時間上限でShiftを解放し、`hWndAccessApp`から記録したPIDと実行ファイルが一致するAccessだけを停止します。既存の`MSACCESS.EXE`を一括停止しません。

現行ツールは、タイムアウト直前のトップレベルウィンドウ列挙を実装していません。現行ツールだけでタイムアウトした場合は`window_enum=not-implemented`相当として扱い、モーダルダイアログが原因と断定しません。ウィンドウ証跡が必要な案件では、外部ラッパーで記録PIDを時間制限付きで列挙します。

Shiftが既に押されている場合、ツールは入力状態を変更せず失敗します。環境変数の有無と値は、ユーザープロファイル等をマスクして`environment.json`へ記録します。`PATHEXT`や`CommonProgramFiles`が欠けた制限環境では、DLL消失と誤診断しないようCOM起動前に失敗させます。

`open_forms_after_open=0`は補助証拠にすぎません。通常起動処理が走らなかったことは、対象DB固有の起動ログ、カナリア、副作用の不在などをstage記録でも確認します。

## 出力構成

```text
05_before_export/
  native/
    saveastext/
      forms/
      reports/
      modules/
      macros/
      queries/
      data_macros/
    metadata/
      environment.json
      startup-preflight.json
      tables.json
      relations.json
      references.json
      database-settings.json
      catalog.json
  review_utf8/
    saveastext/
      forms/
      reports/
      modules/
      macros/
      queries/
      data_macros/
  manifest.json
  manifest.csv
  export-summary.json
  export-errors.json
  export.log
```

`native/saveastext`はAccessの出力バイトを変更しません。`review_utf8`はBOMと実バイトから文字コードを判定して作るUTF-8派生物です。差分の正本は`native`です。SaveAsText本文は完全性を守るためマスクしないので、保存クエリやVBAに埋め込まれた接続文字列もそのまま含み得ます。

`tables.json`にはローカルテーブルのフィールドとインデックス、リンクテーブルの安全化した接続情報を記録します。接続文字列の資格情報、サーバー名、DB名、ファイルパスは既定でマスクします。

`environment.json`は実行環境の診断証跡です。ユーザープロファイル配下は`<user-profile>`へ置換されますが、公開前には顧客固有のパスや値が残っていないことを別途確認します。

ツールは`review_utf8`を走査し、接続資格情報・接続先に使われるキーの候補を見つけると、値を記録せず`secret-scan`エラーにして全体をFAILにします。これは候補検出であり、機密情報がないことの完全証明ではありません。検出された`native`と`review_utf8`を公開せず、隔離したstageで内容を確認します。

`-Mode StartupProbe`ではレポート、`tables.json`、`relations.json`、`references.json`を省略し、起動経路の確認に必要な資産を優先して出します。`export-summary.json`の`skipped_by_mode`に`reports`を記録し、実在数は`catalog.json`で保持します。AutoExecから省略対象へ続く場合と、正式なbaseline/候補比較では既定の`Full`を使います。

データマクロは、各ローカルテーブルへ`SaveAsText`の`acTableDataMacro`（値`12`）を試みます。この環境で「データマクロなし」を示したAccessエラー2950だけを不在として扱い、それ以外はFAILにします。probe件数、不在件数、未解決エラー件数はsummaryへ記録します。

## PASS条件

`export-summary.json`の`status`が`PASS`であり、次が成立することを確認します。

- `source_sha256_before`と`source_sha256_after`が一致する。
- `skipped_by_mode`以外の各カテゴリで`discovered`と`exported`が一致する。
- `export-errors.json`が空配列である。
- manifest記載ファイルが存在し、非空で、SHA-256が一致する。
- `manifest_sha256`と実際の`manifest.json`が一致する。
- 専用Access PIDと`.laccdb`または`.ldb`が残っていない。
- `sensitive_token_hit_count`が0である。
- 対象DB固有の証拠でも通常起動処理の不実行を確認できる。

失敗時は`_working`を残します。強制停止された二次コピーやロックが残ったコピーは失敗物として扱い、次の作業へ再利用しません。

## 記録する識別子

baseline記録には次を転記します。

- 元ACCDBのSHA-256
- 対応INIのSHA-256
- `exporter_sha256`
- `access_version`
- `manifest_sha256`
- Export日時と結果

元ACCDBのパス、接続先、資格情報、顧客固有情報を公開ログや申し送りへ転載しません。

関連する公式定数は[AcObjectType列挙](https://learn.microsoft.com/en-us/office/vba/api/access.acobjecttype)を参照してください。
