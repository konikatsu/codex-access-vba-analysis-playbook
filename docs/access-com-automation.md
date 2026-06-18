# Access COM自動化の基本

PowerShellから `Access.Application` を使ってAccess DBを操作する場合の基本形です。

## 最小形

```powershell
$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.AutomationSecurity = 1
    $access.OpenCurrentDatabase($dbPath)

    # ここで LoadFromText / Run / SaveAsText などを行う
}
finally {
    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
```

## AutomationSecurity

`Application.Run` を使う場合は、`OpenCurrentDatabase` の前に次を設定します。

```powershell
$access.AutomationSecurity = 1
```

これを忘れると、次のような失敗が起きることがあります。

- `LoadFromText` は成功するが `Application.Run` が `HRESULT 0x800A9D9F` で失敗する。
- 何もしない診断関数でも実行できない。
- コード実行やコンパイルがセキュリティ状態で止まる。

用途によって値を使い分けます。

```text
1 = msoAutomationSecurityLow
3 = msoAutomationSecurityForceDisable
```

- 解析用モジュールを取り込んで `Application.Run` したい場合: `1`
- AutoExecなどの起動マクロを止めて、GUI/VBEで安全に開きたい場合: `3`

`3` はマクロ実行を強制無効化する目的に向きます。  
その一方で、DBを開いた後に `Application.Run` でVBAを実行したい作業では、実行自体も止まる可能性があるため、目的に応じて使い分けます。

## 起動処理の扱い

自動起動を止めるために、DB側の `AutoExec`、`StartUp`、起動用関数を直接書き換える方法は、戻し忘れのリスクがあります。

COM自動化では、まず次の方針を優先します。

- DB本体の起動処理は原則触らない。
- `OpenCurrentDatabase` の前に `AutomationSecurity = 1` を設定する。
- それでもダイアログや業務処理が動く場合だけ、コピーDBで一時的な無効化を検討する。
- 一時変更した場合は、変更箇所と戻し手順を記録する。

GUI/VBE作業では、DB内の起動処理を書き換える代わりに、Shiftキーを押下した状態でDBを開く補助スクリプトを使う方法があります。

例:

```powershell
powershell -ExecutionPolicy Bypass -File ".\examples\open-access-devmode.ps1" -DatabasePath "C:\work\access-project\Sample.accdb"
```

この方法なら、DB内の `AutoExec` や起動用関数を変更せずに、開発者モード相当で起動できます。

もう一つの方法として、Access COMで起動し、`AutomationSecurity = 3` を設定してから `OpenCurrentDatabase` する方法があります。

例:

```powershell
powershell -Sta -ExecutionPolicy Bypass -File ".\examples\open-access-no-autoexec.ps1" -DatabasePath "C:\work\access-project\Sample.accdb"
```

これは、DBをダブルクリックや `Start-Process` で直接開くのではなく、`Access.Application` から開く点が重要です。  
ただし、起動処理がAutoExecマクロではなく起動フォームのLoadイベントやスタートアップ設定にある場合、これだけでは完全に止まらないことがあります。

## /cmd で解析用の起動モードを作る

長期的には、Access DB側に「解析や保守では初期処理をスキップできる公式の入口」を用意するのが扱いやすいです。

AutoExecマクロから業務初期処理を直接呼ぶのではなく、いったん起動用関数を呼びます。

AutoExecマクロ:

```text
RunCode: AutoExecMain()
```

標準モジュール:

```vb
Public Function AutoExecMain()

    If IsSkipAutoExecMode() Then
        Debug.Print "AutoExec skipped by command line."

        ' 必要ならメニュー画面だけ開く
        ' DoCmd.OpenForm "Menu"

        Exit Function
    End If

    Call InitialProcess

End Function

Private Function IsSkipAutoExecMode() As Boolean
    Dim cmd As String

    cmd = Nz(Command(), "")

    IsSkipAutoExecMode = _
        (InStr(1, cmd, "SKIP_AUTOEXEC", vbTextCompare) > 0)
End Function
```

起動例:

```powershell
Start-Process "msaccess.exe" -ArgumentList '"C:\work\access-project\Sample.accdb"', '/cmd', 'SKIP_AUTOEXEC'
```

既存のAutoExecがすでに初期処理関数を直接呼んでいる場合は、最小修正として初期処理関数の先頭に判定を入れる方法もあります。

```vb
Public Function InitialProcess()

    If InStr(1, Nz(Command(), ""), "SKIP_AUTOEXEC", vbTextCompare) > 0 Then
        Debug.Print "InitialProcess skipped."
        Exit Function
    End If

    ' 既存の初期処理

End Function
```

ただし、将来的には `AutoExecMain` を挟む方が、通常起動、保守起動、テスト起動を管理しやすくなります。

## 権限付き実行

サンドボックスや制限付き環境では、`LoadFromText` が `予約済みエラー` になることがあります。  
この場合は、DB破損や `.mdl` 形式問題と決めつけず、権限付きで同じ手順を再試行します。

## 後始末

Access COMはプロセスが残りやすいので、必ず次を行います。

- `CloseCurrentDatabase`
- `Quit`
- `ReleaseComObject`
- `[GC]::Collect()`
- `[GC]::WaitForPendingFinalizers()`

必要に応じて、対象フォルダの `.laccdb` / `.ldb` も確認します。
