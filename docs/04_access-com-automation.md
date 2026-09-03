# Access COM自動化の基本

PowerShellから `Access.Application` を使ってAccess DBを操作する場合の基本形です。

この文書のコードはCOMの基本形です。起動バイパス、DAO経路、PID限定回収を含む実作業の標準は[Access修正の標準開発手順](15_access-development-workflow.md)を優先してください。`AutomationSecurity`の値だけでAutoExecや起動フォームを止められるとは判断しません。

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
- COM静的操作でマクロを強制無効化したい場合: `3`。起動バイパスは仮想Shiftなどを別途使う

`3` はマクロ実行を強制無効化する目的に向きます。  
その一方で、DBを開いた後に `Application.Run` でVBAを実行したい作業では、実行自体も止まる可能性があるため、目的に応じて使い分けます。

重要:

- `AutomationSecurity = 1` はVBA実行を許可しますが、AutoExecや起動処理をスキップしません。
- `AutomationSecurity = 1` で開くと、DB側の初期処理、起動フォーム、外部DB接続が通常どおり動く場合があります。
- `Application.Run` が不要なフォーム/レポートの `LoadFromText`、`SaveAsText` だけなら、まず `AutomationSecurity = 3` を検討します。
- VBA実行を伴う動作テスト、`Application.Run`、実行状態を前提にした確認では `AutomationSecurity = 3` を使いません。
- `Application.Run`が必要で`AutomationSecurity = 1`を使う場合は、作業コピー横の`.ini`や接続設定がテスト環境を向いていることと、仮想ShiftなどCOM経路用の起動抑止が効くことを確認します。`/cmd`は`OpenCurrentDatabase`へ直接渡せません。

典型的な失敗:

```text
AutomationSecurity = 1 で作業コピーを開く
-> AutoExec/初期処理が動く
-> 作業コピー横に必要な ini がない
-> 旧SQL Serverや本番接続へ行く
-> COMがタイムアウト/RPCエラーになる
```

この場合、`LoadFromText` やフォーム定義の破損ではなく、起動処理と接続設定の問題として切り分けます。

## コンテンツの有効化

AccessをGUIで開いたときにセキュリティ警告が出ている場合、「コンテンツの有効化」をクリックしていないと、VBA、マクロ、フォームイベント、COM操作が期待どおり動かないことがあります。

この状態では、次のようなエラーがDB破損のように見える場合があります。

- `Application.Run` が失敗する。
- VBEコンパイルが失敗する。
- フォームやVBAの保存が失敗する。
- `データベースの Visual Basic for Applications プロジェクトが破損しています。` と表示される。

本当に破損している可能性もありますが、まずは作業コピーをGUIで開き、「コンテンツの有効化」をクリックしたうえで、VBEの `デバッグ -> コンパイル` を確認します。

詳しくは [コンテンツの有効化を確認する](requirements/00_enable-active-content.md) を参照してください。

### NGパターン

Access COMで詰まったときに、次の順番で判断すると遠回りになります。

- コンテンツ未有効化を確認せず、VBAプロジェクト破損と判断する。
- `ActiveVBProject = null` を見て、DB破損またはCOM不安定と判断する。
- `AutomationSecurity = 3` のまま `Application.Run` やコンパイル可否を判断する。
- 非表示COMで重いフォームを開き、タイムアウト原因を推測する。

先に、GUIでコンテンツ有効化、VBE手動コンパイル、VBAプロジェクト オブジェクトモデルへのアクセス信頼を確認します。

## 起動処理の扱い

自動起動を止めるために、DB側の `AutoExec`、`StartUp`、起動用関数を直接書き換える方法は、戻し忘れのリスクがあります。

COM自動化では、まず次の方針を優先します。

- DB本体の起動処理は原則触らない。
- 作業前に環境フィンガープリントと起動前`MSACCESS.EXE`スナップショットを記録する。
- 仮想Shiftで起動をバイパスし、`OpenCurrentDatabase`の前に用途に合う`AutomationSecurity`を設定する。
- PIDは`hWndAccessApp`から直ちに取得し、終了時は記録したPIDだけを対象にする。
- 構造メタデータだけならAccessを起動せず、DAO読み取り専用経路を使う。

GUI/VBE作業では、DB内の起動処理を書き換える代わりに、Shiftキーを押下した状態でDBを開く補助スクリプトを使う方法があります。

例:

```powershell
powershell -ExecutionPolicy Bypass -File ".\examples\open-access-devmode.ps1" -DatabasePath "C:\work\access-project\Sample.accdb"
```

この方法なら、DB内の `AutoExec` や起動用関数を変更せずに、開発者モード相当で起動できます。

マクロを無効化した静的操作では、Access COMで`AutomationSecurity = 3`を設定してから`OpenCurrentDatabase`します。ただし、次の例は`AutomationSecurity`の制約確認用であり、完全な起動バイパスではありません。

例:

```powershell
powershell -Sta -ExecutionPolicy Bypass -File ".\examples\open-access-no-autoexec.ps1" -DatabasePath "C:\work\access-project\Sample.accdb" -AcknowledgeStartupMayRun
```

実作業では仮想Shift、PID記録、タイムアウト、後片付けを組み合わせます。起動フォームのLoadイベントやスタートアップ設定は、`AutomationSecurity = 3`だけでは止まらないことがあります。

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
        (StrComp(Trim$(cmd), "SKIP_AUTOEXEC", vbTextCompare) = 0)
End Function
```

起動例:

```powershell
Start-Process msaccess.exe "`"C:\work\access-project\Sample.accdb`" /cmd SKIP_AUTOEXEC"
```

既存のAutoExecがすでに初期処理関数を直接呼んでいる場合は、最小修正として初期処理関数の先頭に判定を入れる方法もあります。

```vb
Public Function InitialProcess()

    If StrComp(Trim$(Nz(Command(), "")), "SKIP_AUTOEXEC", vbTextCompare) = 0 Then
        Debug.Print "InitialProcess skipped."
        Exit Function
    End If

    ' 既存の初期処理

End Function
```

ただし、将来的には `AutoExecMain` を挟む方が、通常起動、保守起動、テスト起動を管理しやすくなります。

この方式をDB側に実装できている場合、GUI/VBE作業では `/cmd SKIP_AUTOEXEC` を第一候補にします。

推奨優先順位:

1. GUI/VBE作業: コンテンツの有効化、または信頼済み場所を確認する
2. GUI/VBE作業: `/cmd SKIP_AUTOEXEC`
3. フォールバック: Shift-bypass
4. COM静的操作: 仮想Shift + `AutomationSecurity = 3`
5. COMでVBA実行が必要: 仮想Shift + `AutomationSecurity = 1`

注意:

- COMの `OpenCurrentDatabase` では、通常 `Command()` は期待どおり渡りません。
- `/cmd SKIP_AUTOEXEC` を使う場合は、`msaccess.exe` の起動引数として渡します。
- 標準コマンドは `Start-Process msaccess.exe "`"C:\path\to\db.accdb`" /cmd SKIP_AUTOEXEC"` の形式にします。
- `/cmd SKIP_AUTOEXEC` はAccessのセキュリティ警告を解除しません。
- VBA/マクロ/フォームイベントの動作テストでは、DBを置くフォルダをAccessの信頼済み場所に追加するか、画面上で「コンテンツの有効化」を行います。
- `AutomationSecurity = 1` はVBA実行用であり、初期処理をスキップする仕組みではありません。
- `AutomationSecurity = 3` はマクロを無効化するため、VBA/フォームイベントの動作テストには使いません。

例:

```powershell
# 起動処理だけ止め、VBA/フォームイベントは有効な状態でGUIテストする。
# 前提: C:\work\access-project\ がAccessの信頼済み場所に登録済み。
Start-Process msaccess.exe "`"C:\work\access-project\Sample.accdb`" /cmd SKIP_AUTOEXEC"
```

このコマンドでセキュリティ警告が出る場合、`/cmd SKIP_AUTOEXEC` の失敗ではありません。
信頼済み場所またはコンテンツ有効化の問題として切り分けます。

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
