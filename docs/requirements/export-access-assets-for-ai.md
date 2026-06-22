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

## 実行方法

対象DBに解析用モジュールを取り込み、VBEのイミディエイトウィンドウから実行します。

```text
Ctrl + G
```

```vb
ExportAnalysisInfo
```

実行前にイミディエイトウィンドウをクリアしておくと、ログが読みやすくなります。

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
