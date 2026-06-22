# Access資産をAI向けにエクスポートする

## 目的

Access DB内の資産をテキスト化し、AIエージェントが検索・解析できる状態にします。

AIが見るべき対象:

- 標準モジュール
- クラスモジュール
- フォーム定義
- フォームのコードビハインド
- レポート定義
- レポートのコードビハインド
- クエリSQL
- テーブル定義
- リレーション
- 参照設定
- DBプロパティ
- 起動設定

## 最短手順

このリポジトリには、Access資産をエクスポートするためのVBAコードを同梱しています。

ここでいう「取り込み」は、Access資産を取り込むという意味ではありません。  
エクスポート処理を実行するための一時ツール `GsTools_analysisinfo` を、対象DBへ追加するという意味です。

- エクスポート実行用モジュール: [`tools/GsTools_analysisinfo.bas`](../../tools/GsTools_analysisinfo.bas)
- VBE貼り付け用コード: [`tools/GsTools_analysisinfo_for_vbe_paste.bas`](../../tools/GsTools_analysisinfo_for_vbe_paste.bas)
- COM / `LoadFromText` 用モジュール: [`tools/GsTools_analysisinfo_loadfromtext.mdl`](../../tools/GsTools_analysisinfo_loadfromtext.mdl)

まずは次の流れで実行します。

```text
1. Access DBを開発モードで開く
2. エクスポート実行用モジュール tools/GsTools_analysisinfo.bas を対象DBに追加する
3. イミディエイトウィンドウで ExportAnalysisInfo を実行する
4. DBと同じフォルダに Defines<DB名>\Latest ができる
5. AIには Latest フォルダを読ませる
```

## エクスポート実行用モジュールをVBEで追加する

Accessで対象DBを開いたら、VBEを開きます。

```text
Alt + F11
```

VBEで次を実行します。

```text
ファイル
-> ファイルのインポート
-> tools/GsTools_analysisinfo.bas を選択
```

取り込み後、標準モジュールに `GsTools_analysisinfo` が追加されていることを確認します。

## イミディエイトウィンドウから実行する

VBEでイミディエイトウィンドウを開きます。

```text
Ctrl + G
```

ログを見やすくするため、実行前にイミディエイトウィンドウをクリアします。

次を入力してEnterを押します。

```vb
ExportAnalysisInfo
```

成功すると、イミディエイトウィンドウに次のようなログが出ます。

```text
Started: 2026-06-22 11:30:00
Database: C:\work\sample\app.accdb
OutputDir: C:\work\sample\Definesapp.accdb\Exports\20260622_113000
Save Form_Main
Save Module1
Updated Latest: C:\work\sample\Definesapp.accdb\Latest
ExportAnalysisInfo finished: C:\work\sample\Definesapp.accdb\Exports\20260622_113000
```

## エクスポート実行用モジュールをCOMから追加する

VBE画面を使わず、CodexからCOMで追加する場合は、`.mdl` を使います。

```powershell
$dbPath = 'C:\work\sample\app.accdb'
$modulePath = 'C:\work\codex-access-vba-analysis-playbook\tools\GsTools_analysisinfo_loadfromtext.mdl'

$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)

$access.LoadFromText(5, 'GsTools_analysisinfo', $modulePath)
$access.RunCommand(126)
$access.Run('ExportAnalysisInfo')

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

同名モジュールが既にある場合は、作業コピー上で削除してから取り込みます。

```powershell
$access.DoCmd.DeleteObject(5, 'GsTools_analysisinfo')
$access.LoadFromText(5, 'GsTools_analysisinfo', $modulePath)
```

## エクスポート用VBAコードの中核

同梱している `GsTools_analysisinfo.bas` の入口はこのPublic Subです。

```vb
Public Sub ExportAnalysisInfo()
    Dim outputRoot As String
    Dim outputDir As String
    Dim latestDir As String
    Dim objectsDir As String
    Dim schemaDir As String
    Dim logPath As String
    Dim currentDat As Object
    Dim currentProj As Object
    Dim fso As Object
    Dim logNo As Integer

    On Error GoTo FatalError

    Set fso = CreateObject("Scripting.FileSystemObject")

    outputRoot = CurrentProject.Path & "\Defines" & CurrentProject.Name
    outputDir = outputRoot & "\Exports\" & Format$(Now, "yyyymmdd_hhnnss")
    latestDir = outputRoot & "\Latest"
    objectsDir = outputDir & "\Objects"
    schemaDir = outputDir & "\Schema"

    EnsureFolder fso, objectsDir & "\Forms"
    EnsureFolder fso, objectsDir & "\Reports"
    EnsureFolder fso, objectsDir & "\Macros"
    EnsureFolder fso, objectsDir & "\Modules"
    EnsureFolder fso, objectsDir & "\Queries"
    EnsureFolder fso, schemaDir

    Set currentDat = Application.CurrentData
    Set currentProj = Application.CurrentProject

    ExportObjectType acForm, currentProj.AllForms, objectsDir & "\Forms", ".frm", logNo
    ExportObjectType acReport, currentProj.AllReports, objectsDir & "\Reports", ".rpt", logNo
    ExportObjectType acMacro, currentProj.AllMacros, objectsDir & "\Macros", ".mcr", logNo
    ExportObjectType acModule, currentProj.AllModules, objectsDir & "\Modules", ".mdl", logNo
    ExportObjectType acQuery, currentDat.AllQueries, objectsDir & "\Queries", ".qry", logNo

    ExportTableDefinitions schemaDir & "\Tables.txt", logNo
    ExportRelations schemaDir & "\Relations.txt", logNo
    ExportReferences schemaDir & "\References.txt", logNo
    ExportDatabaseProperties schemaDir & "\DatabaseProperties.txt", logNo

    UpdateLatestFolder fso, outputDir, latestDir, 0
End Sub
```

完全なコードは [`tools/GsTools_analysisinfo.bas`](../../tools/GsTools_analysisinfo.bas) を参照してください。

## 推奨出力先

```text
C:\work\sample\
  app.accdb
  exports\
    app.accdb\
      20260622_113000\
      Latest\
```

AIには、基本的に `Latest` フォルダを読ませます。

## 出力構成例

```text
Latest\
  modules\
    StandardModule1.bas
    ClassModule1.cls
  forms\
    Form_Main.frm
    Form_Main.code.bas
  reports\
    Report_Invoice.rpt
    Report_Invoice.code.bas
  queries\
    Query_Sample.sql
  tables\
    TableDefinitions.csv
  relations\
    Relations.csv
  references\
    References.csv
  database\
    Properties.csv
    StartupSettings.csv
  manifest.json
```

`manifest.json` には、出力日時、対象DB名、Accessバージョン、出力件数を入れます。

## 検索例

```powershell
rg -n "btnSave_Click|AfterUpdate|BeforeUpdate|Application.Run|LoadFromText|LongPtr" "C:\work\exports\app.accdb\Latest"
```

フォームイベントを探す例:

```powershell
rg -n "Private Sub .*_(Click|AfterUpdate|BeforeUpdate|Load|Open|Current)" "C:\work\exports\app.accdb\Latest\forms"
```

SQLを探す例:

```powershell
rg -n "SELECT|INSERT|UPDATE|DELETE|FROM|JOIN" "C:\work\exports\app.accdb\Latest"
```

## 注意

- `outputDir` に `.accdb` ファイルのパスを入れない。
- 出力先はフォルダにする。
- `Latest` は最新の解析結果として扱う。
- 日本語ファイルは文字コードに注意する。
