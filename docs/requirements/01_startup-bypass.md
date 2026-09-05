# 起動処理を止めて開発モードで開く

## 目的

Access DBを開いた瞬間にAutoExecマクロ、起動フォーム、初期処理が動くと、解析やVBE作業が進めにくくなります。

特に、起動処理が最初に外部DBへ接続する構成では、接続先停止、ネットワーク断、INI不一致、資格情報の問題により、接続エラーやモーダルダイアログで通常起動とCOM保守処理が入口から停止します。DBを開いて調査・修正する保守経路を失わないため、自動起動の把握と安全なスキップ方法の確保を作業着手の前提条件とします。

解析・修正作業では、まず起動処理の有無と経路を安全に調べます。その後、必要なDBだけに保守用の早期終了分岐を一度だけ追加します。

起動処理を止める前に、Access画面上部のセキュリティ警告も確認します。
開発baselineや動作試験では、「コンテンツの有効化」をクリックしていないとVBAやフォームイベントが止まり、DB破損のように見えるエラーが出ることがあります。一方、非信頼場所での静的救出ではコンテンツを有効化しません。目的を混同しないでください。

詳しくは [コンテンツの有効化を確認する](00_enable-active-content.md) を参照してください。

## 前提条件: 依頼元への必須ヒアリング

対象Access資産へ触る前に、[Access作業共通ルールの必須ヒアリング](../00_access-work-common-rules.md#前提条件-依頼元へ自動起動情報を必ずヒアリングする)を完了します。質問、復唱、記録、完了条件、完了前に禁止する操作は同節を正典とし、この文書では再定義しません。回答が`不明`なら、ヒアリング完了後に次節の安全な調査で確定します。

## 最初に調べること

本体ではなく原本コピーを対象に、通常起動せず次を確認します。

1. DAOで`StartupForm`、`AllowBypassKey`、マクロとフォームの名前を読む。
2. [外部Exportツール](../16_access-external-export.md)の`StartupProbe`で全マクロ、全フォーム、全標準/クラスモジュール、保存クエリ、ローカルテーブルのデータマクロを二次コピーから出力する。
3. AutoExecの`RunCode`、AutoExecから開かれるフォーム、起動フォームの`Open`/`Load`イベント、呼出先関数を追う。
4. 呼出先関数の先頭に、`Command()`と`SKIP_AUTOEXEC`を完全一致で比べる早期終了分岐があるか確認する。

自動起動がない場合は追加しません。既に同等の分岐がある場合も追加しません。

通常起動、自己テスト、GUIを検討する場合は、`AllowBypassKey`の値や無効化方法にかかわらず、先に[通常起動前の外部接続棚卸し](../17_access-development-workflow-reference.md#231-通常起動前の外部接続棚卸し)を行います。起動経路から到達し得る接続元または接続先が不明なら、実行しません。

## 実装前に宣言する

必須ヒアリングの完了後、作業担当は不足分をDAOと`StartupProbe`で確認します。実装、VBE編集、`LoadFromText`、コンパイルより前に、[共通ルールの作業前ゲート](../00_access-work-common-rules.md#05-作業前ゲートを必ず書く)を記入します。同ゲートの項目12から17を、自動起動に関する正典の宣言項目とします。

`未確認`は開始宣言として認めません。判定できない場合は実装へ進まず、分からない箇所と次の安全な調査方法を報告します。検証済みbaselineを再利用できる場合は、保存済み証跡を示せば十分であり、調査やIF追加を繰り返しません。

## 推奨: /cmd SKIP_AUTOEXEC

```powershell
Start-Process msaccess.exe "`"C:\work\sample\app.accdb`" /cmd SKIP_AUTOEXEC"
```

DB側の起動関数に、次のような分岐を用意します。

```vb
Public Function AutoExecMain()
    If StrComp(Nz(Command(), vbNullString), "SKIP_AUTOEXEC", vbBinaryCompare) = 0 Then
        Debug.Print "AutoExec skipped."
        Exit Function
    End If

    Call StartUp
End Function
```

この方式なら、DB内部の起動処理を毎回コメントアウトしたり、戻し忘れたりする事故を避けられます。

追加は作業コピーに対して一度だけ行い、通常起動と`/cmd SKIP_AUTOEXEC`起動を別々に確認します。合格した自動起動無効化対応版を開発baselineにし、以後はそのコピーから作業候補を作ります。通常起動に影響しない完全一致分岐なので、作業終了時に削除しません。

AutoExecマクロ自体は削除、改名、コメントアウトしません。IFを置く関数が最初の副作用より前に呼ばれることを確認します。AutoExecがその前にクエリ、フォーム、外部接続などを実行する場合、このIFだけで無効化完了とはしません。起動フォームだけで処理が始まり共通関数がない場合も、フォームイベントへ単純な`Exit Sub`を足して完了とせず、最初の共通入口を設計してから確認します。

### 自動検証は対象Access自身に証跡を出させる

exe直接起動後の`Marshal.GetActiveObject('Access.Application')`は、`SKIP_AUTOEXEC`の合格証明に使いません。Officeはフォーカスを失うまでROTへ登録しない場合があり、複数インスタンスでは別のAccessを返す可能性もあります。

自動検証するDBでは、完全一致分岐の中から[検証用ヘルパー](../../examples/AccessPlaybookStartupBypass.bas)を呼びます。

```vb
If StrComp(Nz(Command(), vbNullString), "SKIP_AUTOEXEC", vbBinaryCompare) = 0 Then
    AccessPlaybookAttestStartupBypassIfRequested
    Exit Function
End If
```

通常の`/cmd SKIP_AUTOEXEC`では環境変数がないため、ヘルパーは何も出力せず従来どおり画面を開いたままにします。[検証ラッパー](../../examples/validate-access-startup-bypass.ps1)から起動した場合だけ、一意なrun ID、固定command、対象DB一致、開いているフォーム数、`hWndAccessApp`、PIDをローカル結果へ出し、ラッパーの確認応答後に保存せず終了します。

検証は信頼済みGUI用stageの使い捨てコピーで行います。ただし分岐が失敗すれば通常起動処理が続くため、到達し得る全接続先が承認済み検証環境を指すか、失敗時にも本番・共有・不明な外部依存先へ到達できないことを別の制御試験で確認した隔離環境が必要です。結果がない、run IDが違う、対象DB不一致、フォームが開く、hWnd由来PIDが専用PIDと違う、タイムアウト、PIDまたはlockが残る場合は`FAIL`です。通常起動は、外部接続棚卸しが別途合格するまで行いません。

## 代替: Shift-bypass

GUI作業では、Shiftキーを押しながらDBを開く方法も使えます。

自動操作では[open-access-devmode.ps1](../../examples/open-access-devmode.ps1)を使えます。この補助スクリプトは、DAOで`AllowBypassKey`を先に確認し、既存Accessと既押下Shiftを拒否し、外部ウォッチドッグでShift解放時間を制限します。現在の入力デスクトップへキー状態を注入するため、ほかの対話操作を止めたうえで`-AcknowledgeCurrentDesktopInput`を明示します。パスワード保護DBには対応しません。

入力待機まで到達したことは、起動処理が抑止された完全な証明ではありません。対象DBのフォーム数、起動ログ、副作用の不在を別途確認します。`AllowBypassKey=False`の場合は使用しません。

## 例外: disabled modeによる静的救出

Trusted Locationではactive contentが有効になるため、`AutomationSecurity=3`だけを起動抑止の証拠にしません。Shift-bypassが使えずDAOでは情報が足りない場合、検証済みの非信頼フォルダに使い捨てコピーを置き、disabled modeで静的に調べる方法を例外候補にできます。

この方法も、場所を移すだけで全起動処理が止まる保証はありません。同じ環境の安全なカナリアでactive contentが無効になることを確認し、コンテンツを有効化せず、`AutomationSecurity=3`、専用PID、ウォッチドッグ、対象固有の起動証跡を併用します。読み取りと静的Exportだけに限定し、開発baseline、編集、正式コンパイル、自己テストには使いません。証明できない場合は[空DBインポート方式](07_empty-database-recovery-import.md)へ切り替えます。

現行の[外部Exportツール](../16_access-external-export.md)はShift-bypass前提で、`AllowBypassKey=False`を拒否します。disabled modeへの自動切替は未実装です。

## COM経路: 仮想Shift + AutomationSecurity

COMの`OpenCurrentDatabase`には`/cmd`を直接渡せません。作業コピーを仮想Shiftで起動バイパスし、目的に応じた`AutomationSecurity`を`OpenCurrentDatabase`より前に設定します。

```powershell
$access = New-Object -ComObject Access.Application
$access.AutomationSecurity = 3

# OpenCurrentDatabaseより前に仮想Shiftを押し、PIDを記録する。
$access.OpenCurrentDatabase($dbPath)
# OpenCurrentDatabaseの完了または時間上限到達後、必ず仮想Shiftを解放する。
```

注意:

- `AutomationSecurity = 3`だけでは、起動フォームや信頼済みDBの全初期処理を止める保証になりません。
- `SaveAsText`などの静的操作は`3`、正式コンパイルや`Application.Run`は`1`を使います。
- 仮想Shiftを使えない場合、構造メタデータだけならDAO、静的Exportが必要なら検証済みdisabled modeまたは空DB方式へ切り替えます。
- `AllowBypassKey=False`で修正が必要な場合は、[詳細リファレンスの例外時の切替](../17_access-development-workflow-reference.md#12-例外時の切替)に従い、原本を変更しません。
- PID、仮想Shift、タイムアウト、後片付けを一組にします。

詳細は[Access修正の標準開発手順](../15_access-development-workflow.md)を参照してください。

参考:

- [Decide whether to trust a database](https://support.microsoft.com/en-us/access/decide-whether-to-trust-a-database)
- [Trusted Locations for Office files](https://learn.microsoft.com/en-us/microsoft-365-apps/security/trusted-locations)
- [How to use Visual C# to automate a running instance of an Office program](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance)
