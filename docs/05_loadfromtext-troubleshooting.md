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

## 文字コードの基本

`LoadFromText` に戻すファイルは、Accessが `SaveAsText` で出した形式に寄せます。

実務上よく見る組み合わせ:

```text
acModule             -> CP932/SJIS系
acForm / acReport    -> UTF-16 LE with BOM
```

フォームやレポートをCP932として編集・保存し直すと、`LoadFromText` が通ってもフォーム定義や日本語文字列が壊れることがあります。
逆に、モジュールをUTF-16 LEとして扱うと、VBAコードの文字化けやインポート失敗につながることがあります。

最終判断は、必ず対象ファイルの先頭バイトで行います。
詳しくは [Accessテキスト資産の文字コード](10_access-text-encoding.md) を参照してください。

## よくある失敗

### 文字化けしている

`SaveAsText` で出したフォームやモジュールが次のように見える場合、文字コードの誤読を疑います。

```text
V<NUL>e<NUL>r<NUL>s<NUL>i<NUL>o<NUL>n<NUL> <NUL>=<NUL>2<NUL>1<NUL>
```

確認すること:

- 対象がVBAモジュール (`acModule`) か、フォーム/レポート (`acForm` / `acReport`) か。
- VBAモジュールならCP932/SJIS系、フォーム/レポートならUTF-16 LE with BOMをまず考える。
- 先頭バイトが `FF FE` ではないか。
- 本文にNULLバイト `00` が大量に残っていないか。
- UTF-16 LEのファイルを `-Encoding UTF8` やCP932として読んでいないか。
- `_utf8` というフォルダ名だけで、UTF-8変換済みとして扱っていないか。

`LoadFromText` に戻すファイルは、できるだけAccessが `SaveAsText` で出した形式に寄せます。
AI向けUTF-8コピーは解析用として扱い、そのまま取り込み用に使う前提にしないでください。

詳しくは [Accessテキスト資産の文字コード](10_access-text-encoding.md) を参照してください。

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

### 構文エラー

`End Function` や `End Sub` の直後で構文エラーになる場合、変換処理や貼り付け処理で孤立した括弧行が混入していないか確認します。

例:

```vb
End Function
()
()
)
```

検索例:

```powershell
rg -n "^\s*\(\)\s*$|^\s*\)\s*$" "C:\work\access-project\exports\after"
```

孤立行が見つかったら、該当モジュールを `SaveAsText` で出力し、適用前との差分を見てから削除します。  
複数モジュールで同時に出る場合は、全文書き戻し型の変換ロジックを疑います。

### 型が一致しません

64bit対応後に `型が一致しません` が出た場合、`Declare PtrSafe` 側だけでなく、呼び出し側の変数も確認します。

典型例:

- API戻り値を `LongPtr` にした。
- 既存コードでは戻り値を `Long` 変数へ代入していた。
- 64bit VBAで型不一致になった。

戻り値を使っていないなら、代入せず `Call SomeApi(...)` に変える選択肢があります。

## 切り分けの順番

1. `LoadFromText` が成功するか。
2. `SaveAsText` で取り込んだモジュールを再出力できるか。
3. `Application.VBE.ActiveVBProject` が対象DBを指しているか。
4. 何もしない診断関数を `Application.Run` できるか。
5. 本来の処理を `Application.Run` できるか。
6. `RunCommand(126)` でコンパイルできるか。

`LoadFromText` 成功と `Application.Run` 成功は別問題として扱います。
