# 起動処理を止めて開発モードで開く

## 目的

Access DBを開いた瞬間にAutoExecマクロ、起動フォーム、初期処理が動くと、解析やVBE作業が進めにくくなります。

解析作業では、まず起動処理を安全に止めます。

## 推奨: /cmd SKIP_AUTOEXEC

```powershell
Start-Process msaccess.exe "`"C:\work\sample\app.accdb`" /cmd SKIP_AUTOEXEC"
```

DB側の起動関数に、次のような分岐を用意します。

```vb
Public Function AutoExecMain()
    If InStr(1, Nz(Command(), ""), "SKIP_AUTOEXEC", vbTextCompare) > 0 Then
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

## 代替: AutomationSecurity=3

COM経由でDBを開き、マクロを止めたい場合に使います。

```powershell
$access = New-Object -ComObject Access.Application
$access.AutomationSecurity = 3
$access.Visible = $true
$access.OpenCurrentDatabase($dbPath)
```

注意:

- ダブルクリックや `Start-Process app.accdb` には効かない。
- `Application.Run` には向かない場合がある。
- 起動フォームのイベントなど、マクロ以外の処理は完全には止まらない可能性がある。
