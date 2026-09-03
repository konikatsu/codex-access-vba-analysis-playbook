# Access修正の標準開発手順

この文書は、Access資産を継続的に解析、修正、検証するときの標準手順です。
安全性を落とさず、同じコンパイルや全資産エクスポートを繰り返さないことを目的にします。

特定の案件名、DB名、サーバ名、資格情報、実データは前提にしません。

## 1. 基本方針

- 本体ACCDBを直接編集しない。
- ACCDBと対応INIを一組として扱う。
- 検証済みbaselineをSHA-256で識別し、条件が同じなら再利用する。
- AutoExecは資産として保持し、目的に応じて起動経路だけを切り替える。
- baselineの全資産エクスポートは1回、実装中は対象資産だけ、最終監査で全資産を1回出力する。
- 空DBへのインポート結果は解析専用とし、正式な差分基準には使わない。
- 失敗した作業コピーを継ぎ足し修復せず、直前の成功baselineから作り直す。
- 本番反映は開発・検証とは別工程とし、明示的な承認を必要とする。

## 2. AutoExecと起動経路

「AutoExecを無効化する」を単一の操作として扱いません。AutoExecを削除、改名、書換えせず、目的別に次の経路を使います。

| 目的 | AutoExecの扱い | 標準経路 |
| --- | --- | --- |
| 手動GUI/VBE開発 | AutoExecから呼ばれる起動関数が早期終了 | `MSACCESS.EXE <work-copy> /cmd SKIP_AUTOEXEC` |
| 構造メタデータだけの確認 | Accessの起動経路に入らない | DAO `OpenDatabase(copy, False, True)` |
| COMによる構造確認・Export | AutoExecを保持したまま起動をバイパス | 仮想Shift + `AutomationSecurity=3` |
| COMによる正式コンパイル | AutoExecを保持したまま起動をバイパス | 仮想Shift + `AutomationSecurity=1` |
| 読み取り自己テスト | AutoExecを固定ディスパッチャーとして使う | `/cmd RUN_SELFTEST_READONLY` |
| 更新系自己テスト | AutoExecを固定ディスパッチャーとして使う | `/cmd RUN_SELFTEST_DML` |
| フォーム・VBAを含む救出解析 | AutoExecを持たない空DBを開く | 空DBへ対象資産をインポート |

### 2.1 `/cmd`の意味

`/cmd`はAutoExecへ直接引数を渡す機能ではありません。Access起動時の値を`Command()`へ渡し、AutoExecから呼ばれた起動関数が完全一致で判定します。

```text
MSACCESS.EXE /cmd <fixed-command>
-> AutoExec
-> startup function
-> Command()を完全一致判定
-> 許可された処理だけを実行
```

`CreateObject`または`New-Object -ComObject Access.Application`から`OpenCurrentDatabase`する経路には、`/cmd`を直接渡せません。COM処理に`SKIP_AUTOEXEC`が届くと仮定しないでください。

`SKIP_AUTOEXEC`と自己テスト用コマンドは、対象DBの起動関数が`Command()`を完全一致で判定する実装を持つ場合だけ有効です。その実装がないDBへコマンド名を渡しても、起動処理は抑止されません。

固定コマンドから任意の関数名、SQL、ファイル名を実行できる設計にはしません。

### 2.2 DAO読み取り専用経路

テーブル、クエリ、リレーションなどの構造メタデータだけが必要なら、Access Applicationを起動する前にDAOを使います。

```powershell
$engine = $null
$db = $null

try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $db = $engine.OpenDatabase($copyPath, $false, $true)

    # TableDefs / QueryDefs / Relations / Database.Propertiesを読み取る
}
finally {
    if ($null -ne $db) {
        try { $db.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)
    }
    if ($null -ne $engine) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
    }
}
```

第2引数の`False`は共有モード、第3引数の`True`は読み取り専用です。DAOはAccess Applicationの起動処理へ入らないため、AutoExecや起動フォームを実行せず、`MSACCESS.EXE`も起動しません。

2026-09-03にAccess 16.0.20326.20112で行った検証では、起動フォームに副作用を設定した作業用ACCDBをDAOで開いても副作用は0件で、新しい`MSACCESS.EXE`も0件でした。`TableDefs`、`QueryDefs`に加え、`Containers`からフォーム、レポート、モジュールの名前と更新日時を列挙できました。

ただし、アプリケーション固有のContainerはDAOで常にサポートされる保証がありません。取得できても名前などのカタログ情報に限られ、フォーム・レポートの完全な定義やVBAソースは取得できません。DAOの`Database`には`SaveAsText`もありません。完全な資産Exportが必要ならCOM経路を使い、直接開けない場合だけ空DBインポートを救出解析として使います。

読み取り専用はACCDBへの書込みを防ぐ設定であり、リンクテーブルのメタデータ参照が接続先へ到達しない保証ではありません。

DAOを閉じた後は、作業コピー横の`.laccdb`または`.ldb`が消えたことも確認します。`DAO.DBEngine.120`の生成が`0x80040154`で失敗する場合は、ACE未導入だけでなく、実行したPowerShellとOffice/ACEの32bit・64bit不一致も切り分けます。

### 2.3 COM起動ゲート

仮想Shiftを使うCOM処理は、次のゲートを一組として実行します。

1. 環境変数の総数と、`PATHEXT`、`COMSPEC`、`CommonProgramFiles`、`CommonProgramFiles(x86)`、`TEMP`、`USERPROFILE`、`APPDATA`の実測値を環境フィンガープリントとして記録する。公開ログではユーザー固有部分をマスクする。
2. 本体ではなく作業コピーを対象にする。
3. DAOで`AllowBypassKey`と起動設定を記録する。
4. `CreateObject`経路で専用のAccess COMインスタンスを起動する。GUI試験では、先に代替デスクトップ上で起動したラッパープロセスの中から`CreateObject`する。COMは`DcomLaunch`経由になり得るため、呼出元との親子関係は識別根拠にしない。
5. `hWndAccessApp`から、その処理が起動したAccess PIDを直ちに記録する。
6. 用途に合う`AutomationSecurity`を`OpenCurrentDatabase`より前に設定する。
7. 外部ウォッチドッグの時間上限を開始し、`OpenCurrentDatabase`の直前に仮想Shiftを押す。
8. Shiftを押したまま`OpenCurrentDatabase`を呼び、正常復帰直後にShiftを解放する。例外、ハング、タイムアウトでもウォッチドッグが必ず解放する。
9. 起動直後に、フォーム、起動ログ、段階ログなどから通常起動処理が走っていないことを確認する。
10. タイムアウト時はShiftを解放してから記録したPIDだけを停止する。既存のAccessプロセスを一括停止しない。
11. 終了後にPID消滅、`.laccdb`消失、仮想Shift解放を確認する。

親プロセスの環境変数が制限されていると、`PATHEXT`欠落によるexe解決失敗や、`CommonProgramFiles`欠落によるCOM DLL解決失敗が起きます。後者は`0x8007007E`となり、DLLやプロバイダー自体が消えたように見える場合があります。環境フィンガープリントのない観測ログだけで、アプリ、DLL、プロバイダーの障害と結論しません。完全なログオン環境でも再現するかを比較します。

レジストリの`REG_EXPAND_SZ`を確認するときは、値に含まれる環境変数を`[Environment]::ExpandEnvironmentVariables()`で展開し、`%...%`が残っていないことを確認してから`Test-Path`します。`0x8007007E`はモジュールを解決できなかった証拠であり、対象ファイルが存在しない証拠ではありません。

`Application.Forms.Count=0`は確認項目の一つですが、それだけで起動処理が何も実行されなかったとは断定しません。フォームを開かない初期処理もあるため、段階ログや起動処理固有の証跡も確認します。

`hWndAccessApp`からPIDを記録できなかった場合は、起動前の`MSACCESS.EXE`スナップショットと`Win32_Process`を使います。停止候補は「起動前に存在しない」「起動操作の時間窓内に生成」「実行ファイルが期待する`MSACCESS.EXE`」「コマンドラインに独立した引数`-Embedding`がある」をすべて満たすものに限定します。候補が一意でなければ停止しません。`MainWindowHandle=0`、プロセス名だけ、親子関係だけでは識別しません。

`AllowBypassKey=False`の場合はShift-bypassを強行しません。解析だけなら空DB方式へ切り替えます。正式コンパイルが必要なら、承認済みの方法で作成したコンパイル専用コピーを使用し、本体や製品候補のAutoExecを変更しません。

## 3. baselineの再利用判定

新しい作業を始めるたびに、baseline作成を繰り返してはいけません。次がすべて一致し、前回結果が`PASS`なら、既存baselineとそのExportを再利用します。

- コンパイル済みbaseline ACCDBのSHA-256
- 対応INIのSHA-256
- Microsoft Accessのバージョンとビルド
- エクスポータのバージョンまたはSHA-256
- 参照設定
- Export manifestのSHA-256
- 前回のコンパイル、再オープン、Export検証結果

ACCDB、INI、Access、参照設定、エクスポータのいずれかが変わった場合はbaselineを作り直します。由来が説明できない既存コピーは再利用しません。

baseline記録には、最低限次を残します。

```text
source_path
source_sha256
baseline_path
baseline_sha256
ini_sha256
access_version
exporter_sha256
manifest_sha256
created_at
compile_status
reopen_compile_status
export_status
```

## 4. baselineを初めて作る

### 4.1 作業コピー

1. 最新成功ACCDBを選ぶ。
2. 新しいstageへACCDBと対応INIをコピーする。
3. コピー元ACCDBの操作前SHA-256を記録する。
4. コピー直後に、コピー元と作業コピーのSHA-256が一致することを確認する。
5. INIの接続先が、作業記録に記載した検証環境と一致することを確認する。
6. 無関係な`MSACCESS.EXE`と残留`.laccdb`がないことを確認する。

例:

```text
stepNNN_baseline/
  00_stage.md
  01_source_copy/
    app.baseline.accdb
    app.ini
  02_before_export/
```

### 4.2 コンパイル済みbaseline

差分基準は、コンパイル条件をそろえたACCDBから作ります。

1. 仮想Shift + `AutomationSecurity=1`で作業コピーを開く。
2. `RunCommand(126)`でコンパイルする。
3. `Application.Quit(1)`で保存して終了する。
4. 同じ作業コピーを同じ条件で再オープンする。
5. もう一度コンパイルし、エラーがないことを確認する。
6. このコンパイル済みACCDBをbaselineとして凍結する。

`AutomationSecurity=3`で`RunCommand(126)`が例外を返さなかったことだけを、正式なコンパイル合格にしません。ユーザーVBAを有効にした`AutomationSecurity=1`の検証結果を使います。

AutoExecが正式コンパイルを妨げる場合、baselineと同じ内容からAutoExecだけを除いたコンパイル専用コピーを使えます。ただし、そのコピーを全資産Exportや製品候補の正本にはしません。

### 4.3 全資産Export

コンパイル済みbaselineから、全資産を直接出力します。この1回を、初期解析と正式な`before`差分基準の両方に使います。

最低出力対象:

- フォームとコードビハインド
- レポートとコードビハインド
- 標準モジュールとクラスモジュール
- マクロ
- 保存クエリ定義
- ローカル・リンクテーブル定義
- フィールド、インデックス、リレーション
- VBA参照設定
- データベースと起動設定

Export処理では、保存クエリ、VBA、マクロ、DDL、DML、リンク先の実データ取得を実行しません。リンク情報を出力する場合は資格情報をマスクします。

リンクテーブルの`TableDef.Fields`や`Indexes`は、AccessやODBCドライバーによって接続先のスキーマ確認を行う場合があります。接続自体を禁止する検証では、ネットワークを観測または遮断できる環境で実行し、取得できなかった項目を未検証として記録します。接続待ちやタイムアウト時に、同じプロパティ参照を繰り返しません。

出力構成:

```text
02_before_export/
  native/
  review_utf8/
  manifest.json
  manifest.csv
  export-summary.json
  export-errors.json
  export.log
```

`native`はAccessが出力した原本として上書きしません。`review_utf8`は検索・レビュー用の派生物です。文字コードは拡張子で決めず、BOMと実バイトで判定します。

### 4.4 Exportの合格条件

次がすべて成立したときだけ`PASS`とします。

- コピー元ACCDBの操作前後SHA-256が一致する。
- 発見したオブジェクト件数と出力件数が一致する。
- 必須カテゴリが省略されていない。
- `export-errors.json`が空である。
- manifest記載ファイルが存在し、空ではない。
- manifestのSHA-256が実ファイルと一致する。
- UTF-8派生物に変換失敗、文字化け兆候、想定外NULがない。
- 出力、ログ、申し送りに接続文字列、資格情報、実サーバ名が残っていない。
- 通常の起動処理が走っていない。
- 新しいAccess PIDと`.laccdb`が残っていない。

失敗時は同じ操作を繰り返しません。段階ログから停止位置を特定し、原因、確認済み事実、推定、次の安全な手を分けて記録します。

## 5. 修正案をAccess外で作る

1. `native`を保存したまま、編集用コピーを作る。
2. 対象オブジェクトと依存先を特定する。
3. before、after、unified diff、変更理由を保存する。
4. 文字コード、CRCRLF、過剰空行、想定外NUL、孤立した`()`を検査する。
5. 仕様、変更ソース、diff、テスト観点を独立レビューへ渡す。
6. 重大な未解決指摘がなくなるまで、Accessへ反映せず修正案を直す。

レビュー前の不要なAccess起動と、途中候補の全資産Exportを避けます。

## 6. レビュー済み修正を実装する

1. 凍結baselineから新しい実装候補ACCDBを作る。
2. 対応INIを同じフォルダへコピーする。
3. 対象、操作、成功条件、ロールバック元をstage記録へ書く。
4. 仮想Shift + `AutomationSecurity=1`の実装セッションとして候補を開く。
5. レビュー済み変更だけを反映する。
6. 反映直後にコンパイルする。
7. 変更対象だけを`SaveAsText`し、予定した差分と一致することを確認する。

変更方法の優先順位:

- 既存フォームや既存クラスのコードだけ: VBE `CodeModule`へのアンカー付き最小変更
- フォーム・レポートの定義またはレイアウト: 作業コピー上の`SaveAsText -> DeleteObject -> LoadFromText`
- 新規標準モジュール: 文字コードと改行を検証して`LoadFromText`
- クラスモジュール: `acModule`として`LoadFromText`しない

置換件数が期待値と違う、文字コードが不明、または`LoadFromText`後にコンパイルできない場合は、その候補を失敗扱いにします。

## 7. 再オープン監査

実装候補を一度閉じ、次の1回で最終静的監査をまとめます。

1. 実装候補を再オープンする。
2. `AutomationSecurity=1`で正式コンパイルする。
3. 変更フォームのイベント手続きと経路を監査する。
4. 全資産を直接Exportする。
5. baseline manifestと候補manifestを比較する。
6. `native`同士を比較し、意図した変更以外がないことを確認する。
7. Accessを閉じ、PIDと`.laccdb`の消失を確認する。

Accessによる既知の再シリアライズ差分と、業務上の変更を分けます。許容差分は案件ごとに明示し、見慣れた差分だからという理由だけで無条件に除外しません。

## 8. 自己テスト

静的監査に合格した同じ候補の複製で、必要なテストだけを実行します。

読み取りテスト:

```text
MSACCESS.EXE <verification-copy> /cmd RUN_SELFTEST_READONLY
```

更新系テスト:

```text
MSACCESS.EXE <verification-copy> /cmd RUN_SELFTEST_DML
```

`RUN_SELFTEST_DML`は、書込み前にAccess自身の接続でサーバ名とDB名を取得し、許可された使い捨て検証DBとの完全一致を要求します。本番、共有DB、接続先不明では、最初の書込みより前に拒否します。

試験データには一意なマーカーを付け、終了後に別の読み取り確認で0件になったことを検証します。直接SQL接続の成功だけで、Accessが正しいINIを読んだとは判断しません。

## 9. GUI・帳票・PDF確認

- 無人GUI試験は`acWindowNormal`で開き、代替デスクトップ上で実行する。
- `Application.Visible=False`だけでは画面分離を保証しない。ユーザーが起動したAccessは`UserControl=True`となり、`Visible=False`へ変更できない。Automation起動でもフォームやダイアログを含む一連の処理を不可視に保てるとは限らない。
- 確実な画面分離は、Win32のDesktopオブジェクト（`HDESK`）を作り、その上でAccessを起動して行う。これは窓と入力の分離であり、権限、ファイル、ネットワークを隔離するセキュリティサンドボックスではない。
- 実装順は、`CreateDesktop`で`HDESK`を作成し、`STARTUPINFO.lpDesktop`を指定した`CreateProcess`でPowerShellなどのラッパーをそのデスクトップ上に起動し、ラッパー内から`CreateObject("Access.Application")`を呼ぶ。通常デスクトップでAccessを起動してから移動しようとしない。
- `acHidden`を無人GUI試験の標準にしない。
- 試験専用コピーに限り、修飾なし`MsgBox`をログ化する無人モジュールを使用できる。
- `VBA.Interaction.MsgBox`、Access本体、外部COMのダイアログは別途検出する。
- 無人ログに記録があれば、その試験を無条件に`PASS`としない。
- 帳票とPDFは、複数ページ、改ページ、見出し継続、集計値、文字切れを確認する。

## 10. 成功版の保存と本番反映

成功と判断する前に、次を保存します。

- 成功ACCDBと対応INI
- baselineと候補のSHA-256
- before/after/diff
- compileとreopen compileの結果
- baselineと候補のmanifest
- 自己テスト、GUI、帳票、PDFの証跡
- 既知の未検証項目

判定は`PASS`、`FAIL`、`静的レビュー合格（実機未確認）`、`要実機確認`、`未検証`のいずれかで明示します。

本番反映は、成功候補の作成とは別の操作です。対象、バックアップ、反映方法、確認方法、戻し方について明示的な承認を得てから実行します。

## 11. 最小実行回数

検証済みbaselineがある通常のコード修正は、Access作業を次に集約します。

```text
Access外: 解析 -> 修正案 -> diff -> 独立レビュー

Access 1回目:
  baselineから候補作成 -> 変更反映 -> 即時コンパイル -> 対象だけExport

Access 2回目:
  再オープン -> 正式コンパイル -> 全資産Export -> 差分監査

必要な場合だけ:
  固定/cmd自己テスト -> GUI・帳票・PDF確認
```

baselineがない場合だけ、baselineのコンパイル、再オープン、全資産Exportを追加します。初見調査以外で、未コンパイル状態の全資産Exportを先に重ねません。

## 12. 例外時の切替

| 状況 | 次の手 |
| --- | --- |
| ACCDBハッシュと環境が既存baselineと一致 | baseline作成とbefore Exportを省略 |
| テーブル・クエリなどの構造情報だけが必要 | DAO読み取り専用経路を使い、Accessを起動しない |
| COM openがタイムアウト | 同じ呼出しを繰り返さず、PIDと段階ログを確認 |
| exeやCOM DLLが見つからないように見える | 環境フィンガープリントと展開済みレジストリ値を確認し、完全環境と比較 |
| `AllowBypassKey=False` | Shiftを強行せず、別の承認済み経路へ切替 |
| 解析だけ必要で直接開けない | 空DBインポート方式を使用 |
| 空DB方式で解析できた | 構造理解には使用可、正式差分基準には使用不可 |
| AutoExecがコンパイルを妨げる | AutoExecを除いたコンパイル専用コピーを使用 |
| `AutomationSecurity=3`でコンパイルが通った | 正式合格にせず、`1`で再確認 |
| 失敗コピーができた | 成功baselineから新しい候補を作成 |

## 関連文書

- [Access作業共通ルール](00_access-work-common-rules.md)
- [Access資産のエクスポート](01_export-analysis-info.md)
- [Access/VBA作業プレイブック](02_access-ai-agent-workflow.md)
- [Access COM自動化](04_access-com-automation.md)
- [LoadFromTextトラブルシュート](05_loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイル](06_compile-with-runcommand.md)
- [Accessテキスト資産の文字コード](10_access-text-encoding.md)
- [Workspace.OpenDatabase method (DAO)](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-opendatabase-method-dao)
- [Application.SaveAsText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text)
- [Application.Visible property](https://learn.microsoft.com/en-us/office/vba/api/access.application.visible)
- [Win32_Process class](https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process)
- [Registry value types](https://learn.microsoft.com/en-us/windows/win32/sysinfo/registry-value-types)
