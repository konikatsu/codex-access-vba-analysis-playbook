# 起動処理を止めて開発モードで開く

## 目的

Access DBを開いた瞬間にAutoExecマクロ、起動フォーム、初期処理が動くと、解析やVBE作業が進めにくくなります。

解析・修正作業では、まず起動処理の有無と経路を安全に調べます。その後、必要なDBだけに保守用の早期終了分岐を一度だけ追加します。

起動処理を止める前に、Access画面上部のセキュリティ警告も確認します。
「コンテンツの有効化」をクリックしていないと、VBAやフォームイベントが止まり、DB破損のように見えるエラーが出ることがあります。

詳しくは [コンテンツの有効化を確認する](00_enable-active-content.md) を参照してください。

## 最初に調べること

本体ではなく原本コピーを対象に、通常起動せず次を確認します。

1. DAOで`StartupForm`、`AllowBypassKey`、マクロとフォームの名前を読む。
2. [外部Exportツール](../16_access-external-export.md)の`StartupProbe`で全マクロ、全フォーム、全標準/クラスモジュール、保存クエリ、ローカルテーブルのデータマクロを二次コピーから出力する。
3. AutoExecの`RunCode`、AutoExecから開かれるフォーム、起動フォームの`Open`/`Load`イベント、呼出先関数を追う。
4. 呼出先関数の先頭に、`Command()`と`SKIP_AUTOEXEC`を完全一致で比べる早期終了分岐があるか確認する。

自動起動がない場合は追加しません。既に同等の分岐がある場合も追加しません。

## 実装前に宣言する

依頼元は既知の情報を依頼時に渡し、作業担当は不足分をDAOと`StartupProbe`で確認します。作業担当は実装、VBE編集、`LoadFromText`、コンパイルより前に、次を依頼元へ伝えます。

- 自動起動があるか、ないか
- AutoExec、起動フォーム、イベント、最初の呼出先関数からなる起動経路
- 無効化が不要か、既存`SKIP_AUTOEXEC`を使うか、一度だけ追加するか
- 通常起動と無効化起動の検証証跡
- 以後の作業に使うbaselineとSHA-256

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

## 代替: Shift-bypass

GUI作業では、Shiftキーを押しながらDBを開く方法も使えます。

自動操作では[open-access-devmode.ps1](../../examples/open-access-devmode.ps1)を使えます。この補助スクリプトは、DAOで`AllowBypassKey`を先に確認し、既存Accessと既押下Shiftを拒否し、外部ウォッチドッグでShift解放時間を制限します。現在の入力デスクトップへキー状態を注入するため、ほかの対話操作を止めたうえで`-AcknowledgeCurrentDesktopInput`を明示します。パスワード保護DBには対応しません。

入力待機まで到達したことは、起動処理が抑止された完全な証明ではありません。対象DBのフォーム数、起動ログ、副作用の不在を別途確認します。`AllowBypassKey=False`の場合は使用しません。

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
- 仮想Shiftを使えない場合、構造メタデータだけならDAO読み取り専用経路へ切り替えます。
- PID、仮想Shift、タイムアウト、後片付けを一組にします。

詳細は[Access修正の標準開発手順](../15_access-development-workflow.md)を参照してください。
