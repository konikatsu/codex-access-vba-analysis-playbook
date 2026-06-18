# LoadFromTextトラブルシュート

`LoadFromText` はAccessオブジェクトをテキストから取り込むための便利なAPIですが、失敗理由が分かりにくいことがあります。

## オブジェクト種別

```text
1 = Query
2 = Form
3 = Report
4 = Macro
5 = Module
```

標準モジュールを取り込む例:

```powershell
$access.LoadFromText(5, 'ToolsModule', 'C:\work\access-project\tools\ToolsModule.mdl')
```

## `.bas` と `.mdl`

VBEの手動インポートでは `.bas` を使います。  
`LoadFromText` では、Accessの `SaveAsText` 形式に近い `.mdl` を使うほうが安定します。

## よくある失敗

### 予約済みエラー

確認すること:

- サンドボックス制限で失敗していないか。
- 権限付き実行で再試行したか。
- `OpenCurrentDatabase` の前に `AutomationSecurity = 1` を設定したか。
- DBがロック中ではないか。
- 同名モジュールが壊れた状態で残っていないか。

### Sub または Function が定義されていません

確認すること:

- 対象プロシージャが標準モジュールにあるか。
- `Public Sub` / `Public Function` か。
- モジュール名とプロシージャ名を同じにしていないか。
- `Application.Run` の前にコード実行が許可されているか。
- 対象DBのVBAプロジェクトがコンパイル可能か。

## 切り分けの順番

1. `LoadFromText` が成功するか。
2. `SaveAsText` で取り込んだモジュールを再出力できるか。
3. `Application.VBE.ActiveVBProject` が対象DBを指しているか。
4. 何もしない診断関数を `Application.Run` できるか。
5. 本来の処理を `Application.Run` できるか。
6. `RunCommand(126)` でコンパイルできるか。

`LoadFromText` 成功と `Application.Run` 成功は別問題として扱います。
