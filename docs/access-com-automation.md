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

## 起動処理の扱い

自動起動を止めるために、DB側の `AutoExec`、`StartUp`、起動用関数を直接書き換える方法は、戻し忘れのリスクがあります。

COM自動化では、まず次の方針を優先します。

- DB本体の起動処理は原則触らない。
- `OpenCurrentDatabase` の前に `AutomationSecurity = 1` を設定する。
- それでもダイアログや業務処理が動く場合だけ、コピーDBで一時的な無効化を検討する。
- 一時変更した場合は、変更箇所と戻し手順を記録する。

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
