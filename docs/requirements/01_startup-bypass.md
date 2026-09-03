# 起動処理を止めて開発モードで開く

## 目的

Access DBを開いた瞬間にAutoExecマクロ、起動フォーム、初期処理が動くと、解析やVBE作業が進めにくくなります。

解析作業では、まず起動処理を安全に止めます。

起動処理を止める前に、Access画面上部のセキュリティ警告も確認します。
「コンテンツの有効化」をクリックしていないと、VBAやフォームイベントが止まり、DB破損のように見えるエラーが出ることがあります。

詳しくは [コンテンツの有効化を確認する](00_enable-active-content.md) を参照してください。

## 推奨: /cmd SKIP_AUTOEXEC

```powershell
Start-Process msaccess.exe "`"C:\work\sample\app.accdb`" /cmd SKIP_AUTOEXEC"
```

DB側の起動関数に、次のような分岐を用意します。

```vb
Public Function AutoExecMain()
    If StrComp(Trim$(Nz(Command(), "")), "SKIP_AUTOEXEC", vbTextCompare) = 0 Then
        Debug.Print "AutoExec skipped."
        Exit Function
    End If

    Call StartUp
End Function
```

この方式なら、DB内部の起動処理を毎回コメントアウトしたり、戻し忘れたりする事故を避けられます。

## 代替: Shift-bypass

GUI作業では、Shiftキーを押しながらDBを開く方法も使えます。

ただし、AIエージェントによる自動操作ではShift押下状態の制御が難しい場合があります。

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
