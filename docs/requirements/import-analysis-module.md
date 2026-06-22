# 解析用モジュールをAccess DBへ取り込む

## 目的

対象DB自身に解析用の標準モジュールを入れ、`ExportAnalysisInfo` を実行できるようにします。

外部解析DBから対象DBを操作する方法もありますが、Access COMやVBE操作が遅くなりやすいため、対象DB内に一時的な解析用モジュールを置く方式が扱いやすいです。

## VBE手動インポート

```text
VBE
-> ファイル
-> ファイルのインポート
-> GsTools_analysisinfo.bas を選択
```

削除する場合:

```text
VBE
-> プロジェクト エクスプローラーで対象モジュールを選択
-> ファイル
-> ファイルの解放
```

## LoadFromTextで取り込む

COMで取り込む場合は、`SaveAsText` 形式に近い `.mdl` を使います。

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)
$access.LoadFromText(5, 'GsTools_analysisinfo', 'C:\work\tools\GsTools_analysisinfo_loadfromtext.mdl')
```

Accessオブジェクト種別:

```text
1 = Query
2 = Form
3 = Report
4 = Macro
5 = Module
```

## 注意

- 標準モジュール名とPublic関数名を同じにしない。
- `ExportAnalysisInfo` は `Public Sub` または `Public Function` にする。
- 同名モジュールがある場合は、作業コピー上で削除してから取り込む。
- `AutomationSecurity=1` は `Application.Run` やコンパイル向け。
- サンドボックスや権限不足で `LoadFromText` が予約済みエラーになる場合がある。

## よくあるエラー

### Sub または Function が定義されていません

確認すること:

- 実行しているVBEプロジェクトが対象DBか。
- VBE検索時に「カレントプロジェクト」を選んでいるか。
- `ExportAnalysisInfo` がPublicか。
- 標準モジュールに入っているか。
- モジュール名と関数名が同じではないか。
