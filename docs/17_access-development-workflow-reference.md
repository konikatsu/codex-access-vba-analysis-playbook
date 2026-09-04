# Access修正の標準開発手順 詳細リファレンス

この文書は、[短い標準開発手順](15_access-development-workflow.md)の判断根拠、実装例、検証条件をまとめた詳細版です。
通常作業では短い手順を使い、例外処理や根拠の確認が必要な場合に参照します。

特定の案件名、DB名、サーバ名、資格情報、実データは前提にしません。

## 0. 作業着手の前提条件

Access DBには、AutoExec、起動フォーム、起動イベントから直ちに外部DBへ接続するものがあります。接続先が停止している、ネットワークやINIが合わない、資格情報に問題がある場合、接続エラーやモーダルダイアログで通常起動とCOM保守処理が入口から停止します。その状態ではDBを開いて原因調査や修正を行うこと自体が難しくなるため、自動起動の把握と安全な無効化方法の確保を作業着手の前提条件にします。

作業担当は、対象Access資産へ触る前に、[Access作業共通ルールの必須ヒアリング](00_access-work-common-rules.md#前提条件-依頼元へ自動起動情報を必ずヒアリングする)を完了します。質問、復唱、記録、完了条件は同節を正典とし、原本のコピー、ハッシュ取得、DAO確認、StartupProbe、Access起動も完了前には開始しません。

`不明`はヒアリング回答としては有効ですが、作業開始宣言の最終判定には使えません。ヒアリング完了後は安全な事前調査だけに進み、[2.1](#21-最初に自動起動を調べる)で確定します。

## 1. 基本方針

- 本体ACCDBを直接編集しない。
- 依頼元への自動起動ヒアリング完了を、すべてのAccess作業の前提条件とする。
- ACCDBと対応INIを一組として扱う。
- 検証済みbaselineをSHA-256で識別し、条件が同じなら再利用する。
- 実装を始める前に、自動起動の有無、起動経路、無効化方法、検証証跡、採用baselineを依頼元へ宣言する。
- 最初に自動起動の有無と経路を調べる。AutoExecマクロは保持し、呼出先の起動関数へ完全一致の`SKIP_AUTOEXEC`分岐を一度だけ追加する。
- 既に同等の分岐がある場合は追加しない。以後は自動起動無効化対応版を開発baselineにする。
- baselineと候補の全資産Exportは、閉じたACCDBから二次コピーを作る同じ外部ツールで行う。
- baselineの全資産Exportは1回、実装中は対象資産だけ、最終監査で全資産を1回出力する。
- 空DBへのインポート結果は解析専用とし、正式な差分基準には使わない。
- 修正案は、セルフレビュー、独立レビュー、指摘のトリアージ、修正後再検証を通す。
- 失敗した作業コピーを継ぎ足し修復せず、直前の成功baselineから作り直す。
- 本番反映は開発・検証とは別工程とし、明示的な承認を必要とする。

### 1.1 事前調査後の作業開始宣言ゲート

前提条件のヒアリング後、依頼元が把握していない情報は、作業担当が原本のハッシュ取得、コピー、DAO確認、`StartupProbe`までを安全な事前調査として確認します。依頼元がAccess内部を知っていることは前提にしません。

作業担当は、Access資産への実装、VBE編集、`LoadFromText`、DDL、データ更新、コンパイルを始める前に、[共通ルールの作業前ゲート](00_access-work-common-rules.md#05-作業前ゲートを必ず書く)を記入して依頼元へ短く宣言します。同ゲートの項目12から17を、自動起動に関する正典の宣言項目とします。

`自動起動: 未確認`のまま実装へ進みません。事前調査で判定できない場合は、分からない箇所と安全な次手を報告して作業を止めます。既存baselineを再利用する場合は、保存済みの調査・検証証跡を示せばよく、StartupProbeやIF追加を繰り返しません。

## 2. AutoExecと起動経路

最初に自動起動経路を調べ、必要ならAutoExecの呼出先関数へ保守用の早期終了分岐を追加します。AutoExecマクロ自体は削除、改名、書換えせず、目的別に次の経路を使います。

| 目的 | AutoExecの扱い | 標準経路 |
| --- | --- | --- |
| 手動GUI/VBE開発 | AutoExecから呼ばれる起動関数が早期終了 | 対応済みDBは`/cmd SKIP_AUTOEXEC`、未対応DBはShift-bypass |
| 構造メタデータだけの確認 | Accessの起動経路に入らない | DAO `OpenDatabase(copy, False, True)` |
| baseline/候補の全資産Export | 元ACCDBを開かず二次コピーでバイパス | [外部Exportツール](16_access-external-export.md) |
| COMによる個別の構造確認 | AutoExecを保持したまま起動をバイパス | 仮想Shift + `AutomationSecurity=3` |
| COMによる正式コンパイル | AutoExecを保持したまま起動をバイパス | 仮想Shift + `AutomationSecurity=1` |
| 読み取り自己テスト | AutoExecを固定ディスパッチャーとして使う | `/cmd RUN_SELFTEST_READONLY` |
| 更新系自己テスト | AutoExecを固定ディスパッチャーとして使う | `/cmd RUN_SELFTEST_DML` |
| フォーム・VBAを含む救出解析 | AutoExecを持たない空DBを開く | [空DBへ対象資産を選択インポート](requirements/07_empty-database-recovery-import.md) |

### 2.1 最初に自動起動を調べる

必須ヒアリングの完了後、新しい原本を初めて扱うとき、または[baselineの再利用条件](#3-baselineの再利用判定)が崩れたときは、原本を複製した直後、ほかの修正より先に次を確認します。この調査で本体や作業コピーの通常起動を行いません。有効な検証記録を持つbaselineを再利用できる場合は、この調査とIF追加を繰り返しません。

1. DAOで`StartupForm`、`AllowBypassKey`、マクロとフォームのカタログを読む。
2. `AutoExec`マクロ、全マクロ、起動フォーム、AutoExecから開かれるフォーム、それらの`Open`/`Load`イベント、保存クエリ、テーブルデータマクロ、そこから呼ばれる関数を列挙する。
3. 定義が必要で`AllowBypassKey`が明示的な`False`でなければ、外部Exportツールを`-Mode StartupProbe`で実行する。ツールが仮想Shiftで開くのは二次コピーだけである。`False`ならツールを繰り返さず、[空DB救出解析](requirements/07_empty-database-recovery-import.md)へ切り替える。
4. `AutoExec`の`RunCode`等が呼ぶ最初の起動関数を特定する。
5. その関数の先頭に、完全一致の`SKIP_AUTOEXEC`分岐が既にあるか確認する。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\examples\export-access-assets.ps1" `
  -DatabasePath "C:\work\project_work\stepNNN_baseline\01_source_copy\app.source-copy.accdb" `
  -OutputDirectory "C:\work\project_work\stepNNN_baseline\02_startup_probe" `
  -Mode StartupProbe
```

自動起動があり、同等の分岐がなければ、初回だけShift-bypassで作業コピーを開き、呼出先関数の最初の実行文として次を追加します。

```vb
If StrComp(Nz(Command(), vbNullString), "SKIP_AUTOEXEC", vbBinaryCompare) = 0 Then
    Debug.Print "Startup skipped by SKIP_AUTOEXEC."
    Exit Function
End If
```

既に`Command()`を完全一致で比較して早期終了する同等実装がある場合は、表現が違っても重複追加しません。部分一致、任意関数名の実行、SQLやファイル名を受け取る実装は同等とみなしません。

AutoExecマクロ自体の削除、改名、コメントアウトは行いません。IFを置く関数がAutoExecまたは起動フォームから副作用より前に呼ばれることを確認します。AutoExecがその関数より先にクエリ、フォーム、外部接続などを実行する構成では、このIFだけで無効化できたと判定しません。通常起動では`Command()`が空なので既存処理へ進み、`/cmd SKIP_AUTOEXEC`のときだけ最初の副作用より前に終了する形にします。

追加後は、変更した起動関数だけを`SaveAsText`して予定diffと照合し、コンパイルします。次に、検証環境で通常起動が従来どおり進むことと、`/cmd SKIP_AUTOEXEC`で外部接続や画面初期化より前に終了することを別々の証跡で確認します。この確認に合格したコピーを自動起動無効化対応版とし、以後の開発baselineにします。

起動フォームだけで自動処理が始まり、共通の起動関数がないDBでは、フォームイベント先頭の`Exit Sub`だけで「無効化完了」としません。フォーム自体は開くため、最初の共通入口を作るか、対象DB固有の起動経路を設計してから確認します。

### 2.2 `/cmd`の意味

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

### 2.3 DAO読み取り専用経路

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

`AllowBypassKey`は未作成の場合があります。未定義をエラーや`False`と同一視せず、`defined=false`と既定動作を分けて記録します。パスワード付きDBでは`OpenDatabase`のConnect引数が必要です。平文パスワードをコマンド履歴、ログ、申し送りへ残しません。`TableDef.Connect`や起動設定を記録するときも、資格情報、サーバー名、DB名、ユーザー固有パスをマスクします。

DAOを閉じた後は、作業コピー横の`.laccdb`または`.ldb`が消えたことも確認します。`DAO.DBEngine.120`はACE 12.0のProgIDで、`.accdb`を扱う標準経路です。Access 2010専用ではありません。Jetのみの旧環境を対象にする場合は別経路として明示し、`.accdb`を開けない`DAO.DBEngine.36`へ自動でフォールバックしません。`DAO.DBEngine.120`の生成が`0x80040154`で失敗する場合は、ACE未導入だけでなく、実行したPowerShellとOffice/ACEの32bit・64bit不一致も切り分けます。

### 2.3.1 通常起動前の外部接続棚卸し

作業コピーを通常起動する前、または自己テスト・GUIで業務コードを実行する前に、無効化方法の有無にかかわらず、対象操作から到達し得る外部接続元を静的に棚卸しします。`/cmd SKIP_AUTOEXEC`は起動時の分岐を示すだけで、自己テストやその後の画面操作が使う接続先まで証明しません。

1. DAOでリンクテーブルの`TableDef.Connect`と`SourceTableName`、パススルークエリの`QueryDef.Connect`と全保存クエリの`QueryDef.SQL`を読む。`Fields`、`Indexes`、`RefreshLink`、クエリ実行、レコード参照は行わない。
2. `StartupProbe`または空DB救出解析のSaveAsTextから、マクロ、保存クエリ、フォーム・レポートの`RecordSource`、コントロールの`RowSource`、VBAの動的SQLを追う。INI読込み、DSN、`ConnectionString`、`OpenDatabase`、`CurrentProject.Connection`、`ADODB.Connection.Open`に加え、SQLの`IN`句、`[ODBC;`、`DATABASE=`、`DSN=`、`DBQ=`も接続元候補にする。コードは実行しない。
3. `AllowBypassKey=False`では現行ツールの`StartupProbe`を使えないため、DAOと空DB救出解析を使う。接続先キーを含むDBで`StartupProbe`が`secret-scan`によりFAILになった場合は安全側の停止として扱い、隔離stageの出力をローカルだけで確認する。PASSやbaselineにせず、外部へ送らない。
4. 対象操作から到達する接続ごとに、設定元、マスクした接続先識別子、検証用接続先、到達性の証跡をstageへ残す。資格情報や実サーバー名は残さない。

原本ACCDB、対応INI、起動経路、対象操作が前回記録と一致し、保存済み棚卸しのSHA-256も一致する場合は、baselineの棚卸しを再利用できます。変更後の候補で再利用するには、最終unified diffと依存先の再列挙により、この節で列挙した接続生成箇所と、それらへ到達する制御条件に追加・変更がないことも必要です。接続元、接続先、到達条件へ影響する変更があれば、候補を対象に棚卸しをやり直します。再利用時も、資格情報を含まない接続先allowlistのSHA-256と検証用接続先の到達性は実行前に再確認します。

INI差替えだけで安全と判断できるのは、対象操作から到達する全接続がそのINIから生成されると確認できた場合だけです。既存のUser DSN、System DSN、File DSNは、ほかのアプリや利用者へ影響し得るため書き換えません。必要なら明示承認した使い捨て保守コピーだけで、専用の検証用DSNまたは対応ドライバーが許すDSN-less接続へ置き換えます。DSNを使う場合は、その名前だけをallowlistと比べず、実行プロセスと同じ32bit・64bit側のDSN定義から接続先識別子を読み取って照合します。

リンクテーブルの`TableDef.Connect`を変更する場合は、接続文字列から資格情報を除いて解析した接続先識別子が独立したallowlistと完全一致することを確認し、専用プロセスと時間制限の下で`RefreshLink`を呼びます。`RefreshLink`は外部通信を伴い得るため、意図しない接続先への通信を遮断または監視できない環境では実行しません。処理後はDAOを閉じ、再オープンして永続化された`Connect`を再読します。再読した値も事前確認と同じ方法で接続先識別子へ正規化し、allowlistと完全一致することを確認します。`QueryDef.Connect`、保存SQL、VBAなどを変更した場合も、閉じて再オープンしたコピーから値を再取得します。変更、再接続、再読、再照合のいずれかが失敗したコピーは`failed`とし、通常起動しません。現行の共通ツールは、この書換えと`RefreshLink`を実装していません。

接続元または接続先を確定できない場合は、通常起動、自己テスト、GUIを行いません。停止中または不明な本番依存先を、この作業のために起動しません。

### 2.4 COM起動ゲート

仮想Shiftを使うCOM処理は、次のゲートを一組として実行します。

1. 環境変数の総数と、`PATH`、`PATHEXT`、`COMSPEC`、`SystemRoot`、`ProgramFiles`、`ProgramFiles(x86)`、`CommonProgramFiles`、`CommonProgramFiles(x86)`、`PROCESSOR_ARCHITECTURE`、`PROCESSOR_ARCHITEW6432`、`LOCALAPPDATA`、`ProgramData`、`TEMP`、`USERPROFILE`、`APPDATA`の実測値を環境フィンガープリントとして記録する。公開ログではユーザー固有部分をマスクする。
2. 本体ではなく作業コピーを対象にする。
3. DAOで`AllowBypassKey`と起動設定を記録する。
4. `CreateObject`経路で専用のAccess COMインスタンスを起動する。GUI試験では、先に代替デスクトップ上で起動したラッパープロセスの中から`CreateObject`する。COMは`DcomLaunch`経由になり得るため、呼出元との親子関係は識別根拠にしない。
5. `hWndAccessApp`から、その処理が起動したAccess PIDを直ちに記録する。
6. 用途に合う`AutomationSecurity`を`OpenCurrentDatabase`より前に設定する。`3`は`msoAutomationSecurityForceDisable`（プログラムから開くファイルのマクロを無効化）、`1`は`msoAutomationSecurityLow`（マクロを有効化）である。
7. 外部ウォッチドッグの時間上限を開始し、`OpenCurrentDatabase`の直前に仮想Shiftを押す。
8. Shiftを押したまま`OpenCurrentDatabase`を呼び、正常復帰直後にShiftを解放する。例外、ハング、タイムアウトでもウォッチドッグが必ず解放する。
9. 起動直後に、フォーム、起動ログ、段階ログなどから通常起動処理が走っていないことを確認する。
10. タイムアウト時はウォッチドッグでShiftを解放する。列挙機能がある実行系では、停止前に記録PIDに属するトップレベルウィンドウを時間制限付きで1回だけ列挙し、クラス名、キャプション、可視・有効状態、所有関係をstage記録へ残す。列挙機能がなければ`window_enum=not-implemented`、失敗したら`window_enum=failed`とstage記録へ手作業で残す。現行ツールのsummaryには出ない。その後、記録したPIDだけを停止し、既存のAccessプロセスを一括停止しない。
11. 終了後にPID消滅、`.laccdb`消失、仮想Shift解放を確認する。

環境フィンガープリントは、同一ホストの直近成功記録または新しい対話型PowerShellの値と比較します。必須変数の欠落や想定外の差があればCOM診断を中断し、既知の正常な実行環境を復元してから再確認します。

初めて代替デスクトップを使う環境では、カナリアDBで「Accessウィンドウが作成した`HDESK`に属する」「そのデスクトップ上で仮想Shiftによる起動バイパスが成立する」を確認してから標準経路にします。確認できない環境では無人GUI試験を続行せず、手動確認または承認済みの別経路へ切り替えます。

親プロセスの環境変数が制限されていると、`PATHEXT`欠落によるexe解決失敗や、`CommonProgramFiles`欠落によるCOM DLL解決失敗が起きます。後者は`0x8007007E`となり、DLLやプロバイダー自体が消えたように見える場合があります。環境フィンガープリントのない観測ログだけで、アプリ、DLL、プロバイダーの障害と結論しません。完全なログオン環境でも再現するかを比較します。

レジストリの`REG_EXPAND_SZ`を確認するときは、値に含まれる環境変数を`[Environment]::ExpandEnvironmentVariables()`で展開し、`%...%`が残っていないことを確認してから`Test-Path`します。`0x8007007E`はモジュールを解決できなかった証拠であり、対象ファイルが存在しない証拠ではありません。展開済みファイルの`Test-Path`成功も、依存DLL、bitness、ロード条件まで満たす証拠にはなりません。DAOのin-process生成、Access.Applicationのout-of-process生成、exe直接起動を分けて記録します。

`Application.Forms.Count=0`は確認項目の一つですが、それだけで起動処理が何も実行されなかったとは断定しません。フォームを開かない初期処理もあるため、段階ログや起動処理固有の証跡も確認します。

ウィンドウ列挙でダイアログを観測できた場合だけ`modal-observed`と記録します。ウィンドウが見つからない、列挙未実装、列挙失敗のいずれもモーダル不在の証明にはせず、タイムアウト原因は`unknown`のまま段階ログや接続先を切り分けます。

ログや副作用が無いという否定的証拠だけで抑止成功と判定しません。外部接続より前に必ずマーカーを残す安全なカナリアDB、または承認済みの使い捨て検証コピーで肯定対照を取り、その証跡が確認対象では存在しないことを確認します。対照取得だけを目的に、本体や到達先不明の起動処理を実行しません。

`hWndAccessApp`からPIDを記録できなかった場合は、起動前の`MSACCESS.EXE`スナップショットと`Win32_Process`を使います。停止候補は「起動前に存在しない」「起動操作の時間窓内に生成」「実行ファイルが期待する`MSACCESS.EXE`」「コマンドラインに独立した引数`-Embedding`がある」をすべて満たすものに限定します。候補が一意でなければ停止しません。`MainWindowHandle=0`、プロセス名だけ、親子関係だけでは識別しません。

タイムアウトや強制停止を経験した作業コピーは`failed`扱いにし、以後の候補へ昇格させません。`.laccdb`または`.ldb`は対象PIDの消滅を確認した後だけ片付けます。停止候補を一意に識別できない場合は停止せず、前提が回復するまでそのstageを中断します。

`AllowBypassKey=False`の場合はShift-bypassを強行しません。解析だけならDAO、次節のdisabled mode、または空DB方式へ切り替えます。修正が必要なら、まず[前節の棚卸し](#231-通常起動前の外部接続棚卸し)で起動時接続元を確定します。承認済みの使い捨て検証サービスと、必要なINI、`TableDef.Connect`、`QueryDef.Connect`、DSN等がすべて同じ検証環境を指すことを再読して証明できた場合だけ、通常起動経路を候補にします。

依存先を用意できず、起動関数へ保守分岐を追加する必要がある場合に限り、明示承認を得た使い捨てコピーをDAOで排他・読み書きオープンし、`AllowBypassKey=True`へ一時変更する救出経路を候補にできます。変更前の値、承認者、承認日時、作業前SHA-256、操作ログを記録し、次回オープンから有効になることを前提に、保守分岐を追加するためShift-bypassで開きます。原本では行わず、排他オープンまたはプロパティ変更を確認できない場合は中断します。

一時的な`True`の間に、保守分岐の追加、再オープン正式コンパイル、全資産Exportまでの証跡を取得します。監査後に`AllowBypassKey`を元の`False`へ戻し、DAOを閉じて再オープンした後の復元値を確認します。復元値が`False`でなければ、その候補は`failed`とし、最終候補にしません。このプロパティ差分はstageに明記しますが、復元前のExportを復元後の最終候補そのものの証明にはしません。

この救出経路は対象のAccess/ACE・ファイル形式で実機確認が必要です。現行の外部Exportツールは一時変更も、`AllowBypassKey=False`へ戻した最終候補のExportも実装していません。最終候補の起動抑止、正式コンパイル、全資産diffを承認済み経路で証明できるまでは、開発baselineへ昇格せず`要実機確認`とします。

### 2.5 disabled modeによる静的救出

Trusted Locationでは、マクロ、外部データ接続などのactive contentが有効になります。そのため、Trusted Location上のDBを`AutomationSecurity=3`だけで安全に開けるとは判断しません。

Shift-bypassが使えずDAOだけでは不足する場合、検証済みの非信頼フォルダへ使い捨てコピーを置き、disabled modeで静的に調査する方法を例外候補にできます。ただし、場所を移しただけで全AutoExecアクションや起動設定が止まる保証はありません。事前に同じ環境のカナリアでactive contentが無効になることを確認し、`AutomationSecurity=3`、専用PID、ウォッチドッグを併用して、コンテンツを有効化せず、対象固有の起動証跡と副作用も確認します。

この経路は読み取りと静的Export専用です。開発baseline、正式コンパイル、編集、自己テストには使いません。抑止を証明できない場合は[空DBインポート方式](requirements/07_empty-database-recovery-import.md)へ切り替えます。

現行の[外部Exportツール](16_access-external-export.md)はShift-bypass前提で、`AllowBypassKey=False`を起動前に拒否します。disabled modeへの自動切替は実装していないため、同ツールで代替できると仮定しません。

## 3. baselineの再利用判定

必須ヒアリングの完了後、まずbaseline記録の自動起動調査結果を確認します。元ファイルのSHA-256と起動経路が変わっておらず、自動起動無効化対応版が`PASS`ならStartupProbeを再実行せず、そのbaselineを再利用できます。記録がない、元ファイルが変わった、起動経路が変わった場合は2.1からやり直します。

baseline作成を毎回繰り返してはいけません。次がすべて一致し、前回結果が`PASS`なら、既存baselineとそのExportを再利用します。

- コンパイル済みbaseline ACCDBのSHA-256
- 元ファイルのSHA-256と自動起動経路
- `SKIP_AUTOEXEC`分岐の状態と検証済みdiff
- 対応INIのSHA-256
- Microsoft Accessのバージョンとビルド
- エクスポータのバージョンまたはSHA-256
- 参照設定
- Export manifestのSHA-256
- 前回のコンパイル、再オープン、Export検証結果
- baselineと候補へ適用するコンパイル・Exportのセッション列

ACCDBのSHA-256は、閉じて凍結したファイルのバイト同一性を確認する識別子です。Accessで開いた作業コピーや二次コピーは内部状態が変わり得るため、凍結baselineとの同一性判定には使いません。Export manifestと各資産のSHA-256は内容差分の証跡として併用し、ACCDBのSHA-256を置き換えません。

ACCDB、INI、Access、参照設定、エクスポータのいずれかが変わった場合はbaselineを作り直します。由来が説明できない既存コピーは再利用しません。

baseline記録には、最低限次を残します。

```text
source_path
source_sha256_before
source_sha256_after
source_sha256_match
startup_interview_status
startup_interview_at
startup_interview_respondent
startup_interview_answers_sha256
baseline_path
baseline_sha256
ini_sha256
access_version
exporter_sha256
manifest_sha256
startup_probe_sha256
startup_connection_inventory_sha256
test_target_allowlist_sha256
startup_bypass_status
startup_bypass_diff_sha256
created_at
compile_status
reopen_compile_status
export_status
session_sequence
```

## 4. baselineを初めて作る

### 4.1 作業コピー

1. 最新成功ACCDBを選ぶ。
2. 新しいstageへACCDBと対応INIをコピーする。
3. 原本の作業前SHA-256を記録し、作業前ゲートの項目12と一致することを確認する。
4. コピー直後に、コピー元と作業コピーのSHA-256が一致することを確認する。
5. 通常起動、自己テスト、GUI試験を許す場合は、[通常起動前の外部接続棚卸し](#231-通常起動前の外部接続棚卸し)を行い、対象操作から到達する全接続先が承認済みの使い捨て検証環境と一致し、同じプロトコル・ドライバーによる時間制限付きの読み取り確認で到達できることを確かめる。静的作業でも、自動起動抑止を肯定対照で確認できていない間は、停止中の依存先があることを安全の根拠にしない。対象コピーが信頼済み場所にある場合は、依存先の到達性を確認するか、目的に合う非信頼stageへ移してから開く。
6. 無関係な`MSACCESS.EXE`と残留`.laccdb`がないことを確認する。

例:

```text
project_work/                    # この親フォルダは信頼済みにしない
  stepNNN_baseline/              # 原本コピー、静的調査、Export
    00_stage.md
    01_source_copy/
      app.source-copy.accdb
      app.ini
    02_startup_probe/
    03_startup_bypass_diff/
    04_baseline/
      app.baseline.accdb
      app.ini
    05_before_export/
    06_test_results/
  stage_gui_trusted/             # このフォルダだけGUI/VBE・自己テスト用に信頼
    app.verification.accdb
    app.ini
```

この図は処理後の構成です。先に作るのは`project_work`、`stepNNN_baseline`、`01_source_copy`までとし、外部Exportツールへ渡す`02_startup_probe`と`05_before_export`は作成しません。ツールは存在しない新規出力先だけを受け付け、自身でディレクトリを作ります。

### 4.2 自動起動無効化対応版を作る

1. [2.1](#21-最初に自動起動を調べる)のStartupProbeで自動起動経路を特定する。
2. 自動起動がなければ`startup_bypass_status=not-required`と記録する。
3. 起動関数に同等の完全一致分岐があれば`already-supported`と記録し、追加しない。
4. 分岐がなければ初回だけShift-bypassで作業コピーを開き、起動関数の先頭へ最小のIF分を追加する。
5. 変更対象のbefore、after、diff、SHA-256を保存し、起動無効化以外の差分がないことをセルフレビューする。
6. [外部接続棚卸し](#231-通常起動前の外部接続棚卸し)に合格した後、通常起動と`/cmd SKIP_AUTOEXEC`を検証環境で別々に確認する。

この起動無効化対応は以後の保守を安全にする恒久的な開発入口です。通常起動の処理を変えない完全一致分岐として製品候補にも残します。毎回追加・削除したり、作業終了時に戻したりしません。

### 4.3 コンパイル済みbaseline

差分基準は、コンパイル条件をそろえたACCDBから作ります。

1. 仮想Shift + `AutomationSecurity=1`で作業コピーを開く。
2. `RunCommand(126)`でコンパイルする。
3. `Application.Quit(1)`で保存して終了する。
4. 同じ作業コピーを同じ条件で再オープンする。
5. もう一度コンパイルし、エラーがないことを確認する。
6. Accessを閉じ、専用PIDとロックが消えたことを確認する。
7. 閉じたACCDBのSHA-256を測定し、この時点のファイルをbaselineとして凍結する。

`AutomationSecurity=3`で`RunCommand(126)`が例外を返さなかったことだけを、正式なコンパイル合格にしません。ユーザーVBAを有効にした`AutomationSecurity=1`の検証結果を使います。

COMの`OpenCurrentDatabase`には`/cmd SKIP_AUTOEXEC`を渡せないため、起動無効化対応後もCOMコンパイルでは仮想Shiftを使います。GUI/VBE作業では自動起動無効化対応版を`/cmd SKIP_AUTOEXEC`で開きます。

### 4.4 全資産Export

閉じたコンパイル済みbaselineを[Access外部Exportツール](16_access-external-export.md)へ渡します。ツールはbaselineのSHA-256を測定して二次コピーを作り、Accessで開くのは二次コピーだけです。凍結後のbaselineをExport目的で直接開きません。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\examples\export-access-assets.ps1" `
  -DatabasePath "C:\work\project_work\stepNNN_baseline\04_baseline\app.baseline.accdb" `
  -OutputDirectory "C:\work\project_work\stepNNN_baseline\05_before_export"
```

この1回を初期解析と正式な`before`差分基準の両方に使います。DB内へ`ExportAnalysisInfo`を取り込む旧経路は、解析専用コピーに限り、正式なbaselineには使いません。

最低出力対象:

- フォームとコードビハインド
- レポートとコードビハインド
- 標準モジュールとクラスモジュール
- マクロ
- ローカルテーブルのデータマクロ
- 保存クエリ定義
- ローカル・リンクテーブル定義
- フィールド、インデックス、リレーション
- VBA参照設定
- データベースと起動設定

Export処理では、保存クエリ、VBA、マクロ、DDL、DML、リンク先の実データ取得を実行しません。DAOメタデータ内のリンク接続情報は安全なキーだけを残してマスクします。SaveAsText原本は変更せず、接続情報候補を検出した場合は全体をFAILにします。

リンクテーブルの`TableDef.Fields`や`Indexes`は、AccessやODBCドライバーによって接続先のスキーマ確認を行う場合があります。接続自体を禁止する検証では、ネットワークを観測または遮断できる環境で実行し、取得できなかった項目を未検証として記録します。接続待ちやタイムアウト時に、同じプロパティ参照を繰り返しません。

Name AutoCorrectの設定状態は事前調査で記録しますが、標準手順が自動でOFFへ変更することはしません。設定変更は名前マップの作成・削除やオブジェクトの開閉を伴い、対象DB自体を変えるためです。変更が必要な場合は承認済み作業コピーだけで行い、その差分も監査対象にします。リンクテーブルのローカル化や`Connect`書換えもbaselineには行わず、必要なら由来と対応表を記録した検証専用コピーとして分離します。

出力構成:

```text
05_before_export/
  native/
    saveastext/
    metadata/
  review_utf8/
    saveastext/
  manifest.json
  manifest.csv
  export-summary.json
  export-errors.json
  export.log
```

`native/saveastext`はAccessが出力した原本として上書きしません。`native/metadata`は外部ツールがDAO等から作る正規JSONです。`review_utf8`は検索・レビュー用の派生物です。文字コードは拡張子で決めず、BOMと実バイトで判定します。

### 4.5 Exportの合格条件

次がすべて成立したときだけ`PASS`とします。

- 凍結baselineの操作前後SHA-256が一致する。Accessで開いた二次コピーのハッシュとは混同しない。
- 発見したオブジェクト件数と出力件数が一致する。
- 必須カテゴリが省略されていない。
- `export-summary.json`の`status`が`PASS`である。
- `export-errors.json`が空配列である。
- manifest記載ファイルが存在し、空ではない。
- manifestのSHA-256が実ファイルと一致する。
- UTF-8派生物に変換失敗、文字化け兆候、想定外NULがない。
- SaveAsText内の接続資格情報・接続先候補スキャンが0件である。検出値そのものをログへ書かない。
- 出力、ログ、申し送りに接続文字列、資格情報、実サーバ名が残っていない。
- 通常の起動処理が走っていない。
- 新しいAccess PIDと`.laccdb`が残っていない。

`open_forms_after_open=0`だけで通常起動処理の不実行を証明したことにしません。対象DB固有の起動ログ、カナリア、副作用の不在もstage記録へ残します。

失敗時は同じ操作を繰り返しません。段階ログから停止位置を特定し、原因、確認済み事実、推定、次の安全な手を分けて記録します。

## 5. 修正案をAccess外で作る

1. `native`を保存したまま、編集用コピーを作る。
2. 対象オブジェクトと依存先を特定する。
3. before、after、unified diff、変更理由を保存する。
4. 文字コード、CRCRLF、過剰空行、想定外NUL、孤立した`()`を検査する。
5. セルフレビューを完了する。
6. 次の必須材料、テスト観点、セルフレビュー結果を独立レビューへ渡す。必須材料が欠けた状態では依頼しない。
7. 指摘を採用、棄却、要検証に分類し、根拠をstage記録へ残す。
8. 採用分を直し、影響範囲を再列挙してセルフレビューをやり直す。
9. 重大な未解決指摘がなく、要検証項目の扱いが決まるまでAccessへ反映しない。

レビュー前の不要なAccess起動と、途中候補の全資産Exportを避けます。

独立レビューの必須材料:

1. 前後の文脈と行番号を含むunified diff
2. 変更箇所の呼出元と依存先
3. 既存の早期終了、型宣言、入力検証などのガード
4. 仕様、および意図的に対象外とした事項と理由

該当しない材料は`該当なし`と理由を明記します。項目そのものを省略しません。

### 5.1 セルフレビュー

独立レビューへ渡す前に、作成者が次を確認します。

- 要求、変更理由、before/after、diffが一致している。
- 対象オブジェクトと依存先を再列挙し、変更範囲の漏れがない。
- [外部接続棚卸し](#231-通常起動前の外部接続棚卸し)で列挙した接続生成箇所と、それらへ到達する制御条件について、追加・変更の有無を記録し、影響があれば候補の棚卸しをやり直している。
- 予定したファイルとオブジェクトだけが変わり、置換件数が期待値と一致する。
- 確認済み事実、推定、要実機確認を分け、断定に対応する証拠がある。
- 文字コード、改行、BOM、NUL、過剰空行を検査済みである。
- 顧客情報、資格情報、接続文字列、実サーバー名、ユーザー固有パスが成果物へ混入していない。
- 変更リスクごとのテスト観点、失敗条件、戻し方がある。

### 5.2 独立レビューと指摘トリアージ

独立レビューは、変更を作った文脈から離れた別セッション、別AI、または別担当者が行います。レビュー結果は次のいずれかに分類します。

| 判定 | 扱い |
| --- | --- |
| 採用 | 実ファイル、公式仕様、再現試験のいずれかで根拠を確認し、修正する |
| 棄却 | 根拠が反証された指摘。棄却理由と確認証拠を残す |
| 要検証 | 外部資料や実機確認が必要。検証方法、担当、実装前に必要かを決める |

レビュアーの知名度やモデル性能だけを採用理由にしません。同じ指摘を繰り返さないため、指摘ID、判定、理由、対応差分、再検証結果をstage記録へ残します。

### 5.3 修正後再検証

指摘対応後は、最初から最終diffを読み直し、依存先、置換件数、文字コード、機密情報、テスト観点を再確認します。採用した修正で対象範囲が広がった場合、重大・高の指摘を修正した場合、または安全条件を変えた場合は、独立レビューへ戻します。表現だけの限定修正は影響箇所を絞った再確認でよいです。

## 6. レビュー済み修正を実装する

1. 凍結baselineから新しい実装候補ACCDBを作る。
2. 対応INIを同じフォルダへコピーする。
3. 対象、操作、成功条件、ロールバック元をstage記録へ書く。
4. 仮想Shift + `AutomationSecurity=1`の実装セッションとして候補を開く。
5. レビュー済み変更だけを反映する。
6. 反映直後にコンパイルする。
7. 変更対象だけを`SaveAsText`し、予定した差分と一致することを確認する。

変更方法の優先順位:

- 既存フォームや既存クラスのコードだけ: VBE `CodeModule`へのアンカー付き最小変更。「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を事前確認する
- フォーム・レポートの定義またはレイアウト: 作業コピー上の`SaveAsText -> DeleteObject -> LoadFromText`
- 新規標準モジュール: 文字コードと改行を検証して`LoadFromText`
- 新規クラスモジュール: `acModule`として`LoadFromText`せず、`VBComponents.Import`または手動インポート

置換件数が期待値と違う、文字コードが不明、または`LoadFromText`後にコンパイルできない場合は、その候補を失敗扱いにします。

## 7. 再オープン監査

実装候補を一度閉じ、baselineと同じコンパイル回数とExport経路で最終静的監査を行います。

1. 仮想Shift + `AutomationSecurity=1`で実装候補を再オープンする。
2. 正式コンパイルする。
3. 変更フォームのイベント手続きと経路を監査する。
4. Accessを閉じ、PIDと`.laccdb`の消失を確認する。
5. 閉じた候補を、baselineと同じ外部Exportツール、引数、リンク詳細設定で全資産Exportする。ツールが開くのは二次コピーだけとする。
6. baseline manifestと候補manifestを比較する。
7. `native`同士を比較し、意図した変更以外がないことを確認する。
8. 最終diffをセルフレビューし、独立レビュー指摘の採用・棄却・要検証と対応が一致することを確認する。

baselineと候補はいずれも「実装/基準作成セッションでコンパイル、閉じる、再オープンしてコンパイル、閉じる、外部ツールが二次コピーをExport」の順に揃えます。Accessによる既知の再シリアライズ差分と、業務上の変更を分けます。識別子の大文字小文字だけの差分も一律許容せず、新しい宣言や命名を揃えて解消できるか確認します。許容差分は案件ごとに明示し、見慣れた差分だからという理由だけで無条件に除外しません。

## 8. 自己テスト

静的監査に合格した候補と対応INIを一組で`stage_gui_trusted`へ複製し、両方のSHA-256を記録します。[外部接続棚卸しの再利用条件](#231-通常起動前の外部接続棚卸し)を満たす場合は、保存済み棚卸しのSHA-256、候補の接続影響判定、allowlistのSHA-256、検証用接続先の到達性を確認します。条件を満たさない場合は、複製後の候補で棚卸しをやり直します。その後に[4.1の接続先確認](#41-作業コピー)を満たしてから、必要なテストだけを実行します。監査済み候補そのものは移動せず、信頼済みstageへ原本、静的Export、救出解析のコピーを置きません。

自己テストの前提:

- 検証コピーのフォルダが信頼済み場所であるか、コンテンツが明示的に有効化されている。
- 読み取り系・更新系とも、最初の外部接続を開く前に、アプリ自身が接続設定から導出した接続先をアプリケーションINIとは独立したallowlistと完全一致で照合する。
- 物理Shiftと仮想Shiftが解放されている。
- AutoExecから呼ぶ起動関数が固定`Command()`を完全一致で判定する。
- テスト処理は終了時に結果JSONを書き、Access自身を閉じる。
- 実行前に既存の`MSACCESS.EXE`がなく、起動したPIDを一意に記録できる。

読み取りテスト:

```powershell
$runId = [guid]::NewGuid().ToString('N')
& '.\examples\open-access-skip-autoexec.ps1' `
  -DatabasePath 'C:\work\project_work\stage_gui_trusted\app.verification.accdb' `
  -CommandText RUN_SELFTEST_READONLY `
  -RunId $runId `
  -ResultPath "C:\work\project_work\stepNNN_baseline\06_test_results\$runId.json" `
  -AllowlistPath 'C:\work\project_work\stepNNN_baseline\selftest-allowlist.json' `
  -AcknowledgeTrustedLocation
```

更新系テスト:

```powershell
$runId = [guid]::NewGuid().ToString('N')
& '.\examples\open-access-skip-autoexec.ps1' `
  -DatabasePath 'C:\work\project_work\stage_gui_trusted\app.verification.accdb' `
  -CommandText RUN_SELFTEST_DML `
  -RunId $runId `
  -ResultPath "C:\work\project_work\stepNNN_baseline\06_test_results\$runId.json" `
  -AllowlistPath 'C:\work\project_work\stepNNN_baseline\selftest-allowlist.json' `
  -AcknowledgeTrustedLocation
```

### 8.1 ラッパーとDB側ディスパッチャーの契約

`open-access-skip-autoexec.ps1`はallowlistの内容や実際の接続先を照合しません。次の値をプロセス環境変数でAccessへ渡し、終了後に結果JSONの`target_allowlist_match=true`を検査します。接続前の照合は、DB側の固定ディスパッチャーが実装する責務です。

| 環境変数 | DB側での用途 |
| --- | --- |
| `ACCESS_SELFTEST_RUN_ID` | 結果JSONへ同じrun IDを記録する |
| `ACCESS_SELFTEST_RESULT_PATH` | UTF-8の結果JSONを書き出す |
| `ACCESS_SELFTEST_ALLOWLIST_PATH` | アプリケーションINIとは独立したallowlistを読む |

allowlistはUTF-8 JSONとし、使用する種類だけを残します。次は公開用の形式例であり、実際の接続先は書いていません。

```json
{
  "schema_version": 1,
  "targets": [
    {
      "kind": "server-database",
      "server": "<approved-test-server>",
      "database": "<approved-test-database>"
    },
    {
      "kind": "file",
      "path": "<approved-absolute-test-file>"
    }
  ]
}
```

DB側ディスパッチャーは、環境変数やJSONが空、不正、重複、不明な種類、必須項目不足、ワイルドカードを含む場合に接続前に`FAIL`とします。INI、`Connect`、保存SQLなどから接続先を接続せずに導出し、案件で定義した正規化後の`kind`と接続先項目を完全一致で比較します。部分一致やDSN名だけの一致を認めません。不一致なら結果JSONへ`target_allowlist_match=false`を記録してAccessを閉じ、外部接続を開きません。

`target_allowlist_match=true`はDB側の自己申告なので、それだけを接続前照合の証明にしません。ディスパッチャーのコードレビューに加え、意図的に一致しないallowlistを使う否定試験で、外部通信が発生せず`FAIL`になることを確認します。実値を入れたallowlistは隔離stage内だけに保存し、公開リポジトリ、AIへの依頼、申し送りへ転記しません。

`RUN_SELFTEST_READONLY`と`RUN_SELFTEST_DML`はこの契約に従います。INIが指す値をそのままallowlistへ複製せず、本番、共有DB、接続先不明では接続を開く前に拒否します。

結果JSONには最低限、`run_id`、固定`command`、開始・終了時刻、`status`、予定/実行アサーション数、失敗数、削除後残件数、`target_allowlist_match=true`を入れます。結果ファイルがない、run IDが違う、実行アサーションが0件、予定数と実行数が違う、接続先の一致を証明できない、Accessが時間内に閉じない、終了後にAccess PIDまたはロックが残る場合は`FAIL`です。

試験データには一意なマーカーを付け、終了後に別の読み取り確認で0件になったことを検証します。マーカー0件だけではテスト自体が走った証拠になりません。直接SQL接続の成功だけで、Accessが正しいINIを読んだとも判断しません。

自己テスト用コードを作業時に注入する場合は、静的監査済み候補ではなく検証専用コピーだけへ取り込みます。テスト後に成功ACCDBとして保存するのは、自己テストコードを注入していない静的監査済み候補です。製品に恒久搭載する設計なら、そのコード自体も通常の差分レビュー対象にします。

## 9. GUI・帳票・PDF確認

- 無人GUI試験は`acWindowNormal`で開き、代替デスクトップ上で実行する。
- `Application.Visible=False`だけでは画面分離を保証しない。ユーザーが起動したAccessは`UserControl=True`となり、`Visible=False`へ変更できない。Automation起動でもフォームやダイアログを含む一連の処理を不可視に保てるとは限らない。
- 確実な画面分離は、Win32のDesktopオブジェクト（`HDESK`）を作り、その上でAccessを起動して行う。これは窓と入力の分離であり、権限、ファイル、ネットワークを隔離するセキュリティサンドボックスではない。
- 実装順は、`CreateDesktop`で`HDESK`を作成し、`STARTUPINFO.lpDesktop`を指定した`CreateProcess`でPowerShellなどのラッパーをそのデスクトップ上に起動し、ラッパー内から`CreateObject("Access.Application")`を呼ぶ。通常デスクトップでAccessを起動してから移動しようとしない。
- 環境ごとの初回はカナリアDBを使い、AccessウィンドウのDesktop所属とShift-bypass成立を確認する。未確認の環境では、この組合せを成立済みと断定しない。
- `acHidden`を無人GUI試験の標準にしない。
- 試験専用コピーに限り、修飾なし`MsgBox`をログ化する無人モジュールを使用できる。
- `VBA.Interaction.MsgBox`、Access本体、外部COMのダイアログは別途検出する。
- 無人ログに記録があれば、その試験を無条件に`PASS`としない。
- 帳票とPDFは、複数ページ、改ページ、見出し継続、集計値、文字切れを確認する。

## 10. 成功版の保存と本番反映

成功と判断する前に、次を保存します。

- 成功ACCDBと対応INI
- 原本の作業前SHA-256と作業後SHA-256、および一致結果
- baselineと候補のSHA-256
- before/after/diff
- セルフレビュー結果と実施日時
- 独立レビューの担当またはモデル、実施日時、各指摘の採用・棄却・要検証、根拠
- 指摘対応後の再セルフレビューと、必要な再独立レビューの結果
- compileとreopen compileの結果
- baselineと候補のmanifest
- 候補diffの外部接続影響判定と、棚卸しを再利用または再実施した根拠
- 自己テスト、GUI、帳票、PDFの証跡
- 既知の未検証項目
- `AllowBypassKey`を一時変更した場合の元値、承認者、承認日時、復元確認

判定は`PASS`、`FAIL`、`静的レビュー合格（実機未確認）`、`要実機確認`、`未検証`のいずれかで明示します。

本番反映は、成功候補の作成とは別の操作です。対象、バックアップ、反映方法、確認方法、戻し方について明示的な承認を得てから実行します。

## 11. 最小実行回数

検証済みbaselineがある通常のコード修正は、Access作業を次に集約します。

```text
Access外: 解析 -> 修正案 -> diff -> セルフレビュー -> 独立レビュー
          -> 指摘を採用・棄却・要検証 -> 修正後再セルフレビュー

Access 1回目:
  baselineから候補作成 -> 変更反映 -> 即時コンパイル -> 対象だけExport

Access 2回目:
  仮想Shiftで再オープン -> 正式コンパイル -> 閉じる

Access外部Export:
  閉じた候補から二次コピー作成 -> 全資産Export -> 差分監査

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
| `RPC_E_DISCONNECTED` / `0x800706BE` | Access障害と断定せず、モーダル、段階ログ、INI接続先の同一性・到達性を確認 |
| exeやCOM DLLが見つからないように見える | 環境フィンガープリントと展開済みレジストリ値を確認し、完全環境と比較 |
| `AllowBypassKey=False` | Shiftを強行しない。解析はDAO、検証済みdisabled mode、または空DB方式、修正は依頼元確認後の承認済み経路へ切替 |
| `AllowBypassKey=False`かつ起動依存先が停止 | [起動経路上の外部接続を棚卸し](#231-通常起動前の外部接続棚卸し)し、すべて検証環境を指すと証明できた場合だけ通常起動を検討する。確定できなければ明示承認した使い捨てコピーで`AllowBypassKey`の一時変更を検討し、最終候補の検証経路がなければ`要実機確認`で中断する |
| 固定`/cmd`未対応DBをGUI/VBEで開く | Shift-bypassを使う。`AllowBypassKey=False`なら空DB方式または承認済みコピーへ切替 |
| 解析だけ必要で直接開けない | [空DBインポート方式](requirements/07_empty-database-recovery-import.md)を使用し、作業コピーからAutoExecを除外して選択インポート |
| 空DB方式で解析できた | 構造理解には使用可、正式差分基準には使用不可 |
| 自動起動無効化分岐を安全に追加できない | 標準作業を中断し、依存先を使い捨て環境へ向けた承認済み保守コピー等を用意 |
| `AutomationSecurity=3`でコンパイルが通った | 正式合格にせず、`1`で再確認 |
| VBA破損を疑い`/decompile`を検討 | Microsoftの公開スイッチ一覧にないため定常工程へ入れない。症状がある場合だけ、承認済み使い捨てコピーで前後差分を記録して検証 |
| 失敗コピーができた | 成功baselineから新しい候補を作成 |
| 独立レビューの指摘に根拠がない | 無条件採用せず要検証とし、公式仕様または再現試験で確認 |

## 関連文書

- [Access作業共通ルール](00_access-work-common-rules.md)
- [DB内ExportAnalysisInfoによる初見解析](01_export-analysis-info.md)
- [Access外部Exportツール](16_access-external-export.md)
- [Access/VBA作業プレイブック](02_access-ai-agent-workflow.md)
- [Access COM自動化](04_access-com-automation.md)
- [LoadFromTextトラブルシュート](05_loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイル](06_compile-with-runcommand.md)
- [Accessテキスト資産の文字コード](10_access-text-encoding.md)
- [空DBインポートによる救出解析](requirements/07_empty-database-recovery-import.md)
- [Workspace.OpenDatabase method (DAO)](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-opendatabase-method-dao)
- [TableDef object (DAO)](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-object-dao)
- [TableDef.RefreshLink method (DAO)](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-refreshlink-method-dao)
- [QueryDef.Connect property (DAO)](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/querydef-connect-property-dao)
- [IN clause (Microsoft Access SQL)](https://learn.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/in-clause-microsoft-access-sql)
- [Managing Data Sources (ODBC)](https://learn.microsoft.com/en-us/sql/odbc/admin/managing-data-sources?view=sql-server-ver17)
- [Application.SaveAsText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text)
- [Application.Visible property](https://learn.microsoft.com/en-us/office/vba/api/access.application.visible)
- [Decide whether to trust a database](https://support.microsoft.com/en-us/access/decide-whether-to-trust-a-database)
- [Bypass startup options when you open a database](https://support.microsoft.com/en-us/access/bypass-startup-options-when-you-open-a-database)
- [Trusted Locations for Office files](https://learn.microsoft.com/en-us/microsoft-365-apps/security/trusted-locations)
- [Set name AutoCorrect options](https://support.microsoft.com/en-us/access/set-name-autocorrect-options)
- [Command-line switches for Microsoft Office products](https://support.microsoft.com/en-us/office/lifecycle/command-line-switches-for-microsoft-office-products)
- [Win32_Process class](https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process)
- [EnumWindows function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumwindows)
- [GetWindowThreadProcessId function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowthreadprocessid)
- [Registry value types](https://learn.microsoft.com/en-us/windows/win32/sysinfo/registry-value-types)
