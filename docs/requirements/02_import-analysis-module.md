# ExportAnalysisInfoを使える状態にする手順

## 目的

対象のAccess DBで、次のコマンドを実行できる状態にします。

```vb
ExportAnalysisInfo
```

`ExportAnalysisInfo` は、Access DB内のフォーム、レポート、モジュール、クエリ、テーブル定義、リレーション、参照設定などを、AIが読めるテキストとして出力するための手順です。

このページは、人間が手作業でVBEを操作することを前提にした依頼書ではありません。人間はCodexへ対象DBの場所と目的を指示し、CodexがAccess COM、VBE操作、または必要な自動操作で `ExportAnalysisInfo` を使える状態にします。

## 使うファイル

このリポジトリには、3種類のファイルを用意しています。

- 通常のVBE追加用: [`tools/GsTools_analysisinfo.bas`](../../tools/GsTools_analysisinfo.bas)
- VBE貼り付け用: [`tools/GsTools_analysisinfo_for_vbe_paste.bas`](../../tools/GsTools_analysisinfo_for_vbe_paste.bas)
- COM / `LoadFromText` 用: [`tools/GsTools_analysisinfo_loadfromtext.mdl`](../../tools/GsTools_analysisinfo_loadfromtext.mdl)

まずは `tools/GsTools_analysisinfo.bas` を使う方法がおすすめです。CodexがCOMで追加できる場合は、`.mdl` を使った `LoadFromText` 方式を優先します。

## 前提

- Access DBを開けること。
- VBEを開けること。
- 起動処理が邪魔をしない状態であること。
- 対象DBは本体ではなく作業コピーが望ましい。

開発モードで開く例:

```powershell
Start-Process msaccess.exe "`"C:\work\sample\app.accdb`" /cmd SKIP_AUTOEXEC"
```

## 方法1: VBEから追加する

COMで追加できない場合に、CodexがVBE操作で追加します。

対象DBをAccessで開きます。

VBEを開きます。

```text
Alt + F11
```

VBEで次を選びます。

```text
ファイル
-> ファイルのインポート
-> tools/GsTools_analysisinfo.bas を選択
```

プロジェクトエクスプローラーの標準モジュールに、次が追加されていることを確認します。

```text
GsTools_analysisinfo
```

## 方法2: VBEに貼り付ける

VBEのファイルインポートが使えない場合は、Codexが標準モジュールを新規作成して貼り付けます。

```text
VBE
-> 挿入
-> 標準モジュール
```

作成された標準モジュールに、次のファイルの内容を貼り付けます。

```text
tools/GsTools_analysisinfo_for_vbe_paste.bas
```

貼り付け後、必要ならモジュール名を次に変更します。

```text
GsTools_analysisinfo
```

## 方法3: COMからLoadFromTextで追加する

CodexなどからPowerShell COMで追加する場合は、`.mdl` を使います。AIエージェントが扱える環境では、この方法を優先します。

```powershell
$dbPath = 'C:\work\sample\app.accdb'
$modulePath = 'C:\work\codex-access-vba-analysis-playbook\tools\GsTools_analysisinfo_loadfromtext.mdl'

$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)

$access.LoadFromText(5, 'GsTools_analysisinfo', $modulePath)
$access.RunCommand(126)

$access.CloseCurrentDatabase()
$access.Quit()
```

`5` は標準モジュールを表します。

```text
1 = Query
2 = Form
3 = Report
4 = Macro
5 = Module
```

同名モジュールが既にある場合は、作業コピー上で削除してから追加します。

```powershell
$access.DoCmd.DeleteObject(5, 'GsTools_analysisinfo')
$access.LoadFromText(5, 'GsTools_analysisinfo', $modulePath)
```

## 追加できたか確認する

VBEのイミディエイトウィンドウを開きます。

```text
Ctrl + G
```

次を実行します。

```vb
? CurrentProject.Name
```

対象DB名が表示されることを確認します。

次に、`ExportAnalysisInfo` を試します。

```vb
ExportAnalysisInfo
```

成功すると、DBと同じフォルダに次のようなフォルダができます。

```text
Defines<DB名>\
  Exports\
    20260622_113000\
  Latest\
```

## よくあるエラー

### Sub または Function が定義されていません

確認すること:

- 実行しているVBEプロジェクトが対象DBか。
- VBE検索時に「カレントプロジェクト」を選んでいるか。
- `ExportAnalysisInfo` がPublicか。
- 標準モジュールに入っているか。
- モジュール名と関数名が同じではないか。

### LoadFromTextが予約済みエラーになる

確認すること:

- Accessプロセスが残っていないか。
- `.laccdb` が残っていないか。
- `AutomationSecurity = 1` を `OpenCurrentDatabase` より前に設定しているか。
- 権限付き実行が必要ではないか。
- 対象DBがロック中ではないか。

失敗した作業コピーは深追いせず、新しいコピーでやり直します。

### 実行すると起動処理が動いてしまう

起動処理を止めてから開きます。

```powershell
Start-Process msaccess.exe "`"C:\work\sample\app.accdb`" /cmd SKIP_AUTOEXEC"
```

詳しくは [起動処理を止めて開発モードで開く](01_startup-bypass.md) を参照してください。

## 次に読む

`ExportAnalysisInfo` が実行できるようになったら、次に進みます。

- [Access資産をAI向けにエクスポートする](03_export-access-assets-for-ai.md)
