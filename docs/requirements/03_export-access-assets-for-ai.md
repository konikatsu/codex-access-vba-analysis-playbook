# Access資産をAI向けにエクスポートする

## 目的

Access DB内の資産をテキスト化し、AIエージェントが検索・解析できる状態にします。

このページは、人間が手作業でAccessを操作するための手順ではありません。人間はCodexへ対象DBと目的を指示し、Codexが作業コピーの準備、Access操作、`ExportAnalysisInfo` の実行、出力結果の確認を行う前提です。

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

このページは、対象DBで `ExportAnalysisInfo` が実行できる状態になっている前提のエクスポート手順です。

まだ実行できない場合は、先に [ExportAnalysisInfoを使える状態にする手順](02_import-analysis-module.md) を実施します。

Codexは、まず次の流れで実行します。

```text
1. 対象DBの作業コピーを用意する
2. Access DBを開発モードで開く
3. ExportAnalysisInfoを実行する
4. DBと同じフォルダに Defines<DB名>\Latest ができる
5. Codexが Latest フォルダを確認し、AI解析に使う
```

## CodexがExportAnalysisInfoを実行する

可能であれば、CodexはAccess COMや自動操作で `ExportAnalysisInfo` を実行します。

環境によりCOM実行が難しい場合だけ、CodexはVBEでイミディエイトウィンドウを開きます。

```text
Ctrl + G
```

ログを見やすくするため、実行前にイミディエイトウィンドウをクリアします。

次を入力してEnterを押します。これは人間に作業を依頼する意味ではなく、Codexが必要に応じて実行する操作です。

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
- 典型的には、VBAモジュール (`acModule`) はCP932/SJIS系、フォーム (`acForm`) / レポート (`acReport`) はUTF-16 LE with BOMで出ることが多い。
- `SaveAsText` 出力をSJIS/CP932やUTF-8と決め打ちしない。`FF FE` で始まる場合はUTF-16 LEとして読む。
- AI向けにUTF-8へ変換する場合は、生の出力を上書きせず、`*_utf8_fixed` のような別フォルダへ作成する。
- `_utf8` というフォルダ名だけで変換済みと判断しない。変換後にNULLバイトが残っていないか確認する。
- 詳しくは [Accessテキスト資産の文字コード](../10_access-text-encoding.md) を参照する。
