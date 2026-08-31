# Access資産のエクスポートをAI(Codex)に依頼するナレッジ

Access / VBA の既存システムを Codex などのAIに解析させるには、まずAccess DB内の資産をテキストとして取り出す必要があります。

この文書は、人間がAccessやVBEを手作業で操作するための手順ではありません。依頼者がCodexへ指示し、Codexが解析用ツールの作成・調整、対象DBへの組み込み、Access上での実行、出力結果の確認まで行った流れを、次回以降に再利用するためのナレッジです。

特定の顧客名、DB名、サーバ名、実データは含めません。

## 人間がやること

人間が行うのは、原則としてCodexへの指示と結果確認だけです。

- 対象Access DBの場所を伝える。
- 本体DBではなく作業コピーを使うよう指示する。
- Access資産をAI解析用にエクスポートしてほしい、と依頼する。
- 出力結果や申し送り内容を確認する。
- 必要に応じて、Accessの信頼済み場所、コンテンツ有効化、ファイル権限など、人の判断が必要な点だけ確認する。

## Codexがやること

Codexは、対象DBと目的を受け取ったあと、次の作業を行います。

```text
1. 対象DBの作業コピーを用意する
2. 起動処理を止めて対象DBを開く方法を確認する
3. 解析用VBAツールを作成または調整する
4. 解析用VBAツールを対象DBへ組み込む
5. ExportAnalysisInfoを実行する
6. Latestフォルダの出力内容を確認する
7. SaveAsText出力の文字コードを確認する
8. AIが読みやすい形で出力結果、注意点、残課題をMarkdownにまとめる
```

## Codexに依頼する例

```text
対象Access DBを作業コピーに複製し、起動処理を止めて開いてください。
Access資産をAI解析用にエクスポートするためのVBAツールを作成または調整し、対象DBへ組み込んで実行してください。

出力後はLatestフォルダを確認し、フォーム、レポート、モジュール、クエリ、テーブル定義、リレーション、参照設定、DBプロパティをAIが読める状態にしてください。

本体DBは直接変更しないでください。
顧客名、実DB名、サーバ名、実データは公開用メモに含めないでください。
```

## 同梱コードの位置づけ

このリポジトリの解析用コードは、Codexに作成・調整させたものです。人間が手でVBAを書く前提ではありません。

対象DBやAccess環境に合わない場合は、同梱コードを人間が手直しするのではなく、Codexに原因確認、修正、再実行まで依頼します。

- エクスポート実行用モジュール: [`tools/GsTools_analysisinfo.bas`](../tools/GsTools_analysisinfo.bas)
- VBE貼り付け用コード: [`tools/GsTools_analysisinfo_for_vbe_paste.bas`](../tools/GsTools_analysisinfo_for_vbe_paste.bas)
- COM / `LoadFromText` 用コード: [`tools/GsTools_analysisinfo_loadfromtext.mdl`](../tools/GsTools_analysisinfo_loadfromtext.mdl)

## 出力される主な情報

現行の `ExportAnalysisInfo` では、主に次を出力します。

```text
Latest\
  Objects\
    Forms\
    Reports\
    Macros\
    Modules\
    Queries\
  Schema\
    Tables.txt
    Relations.txt
    References.txt
    DatabaseProperties.txt
  ExportAnalysisInfo.log
```

テーブル定義は `Latest\Schema\Tables.txt` に出力します。`TableDefs` を走査し、テーブル名、フィールド名、型、サイズ、必須、既定値、検証ルール、インデックス、主キー、リンクテーブル情報などを出力します。

ただし、これはテーブル定義の解析用出力であり、実データや完全なDDLを出力するものではありません。必要であれば、Codexに出力形式の追加やCSV化、サンプルデータ取得の可否確認を依頼します。

## リンクテーブルのメタデータを安全に出力する

ODBCなどのリンクテーブルを含むDBでは、資産の`SaveAsText`出力と、リンク先を読まないメタデータ抽出を分けます。
テーブル定義の出力は、テーブル名、属性、リンク種別、リンク元の論理名など、ローカルカタログから取得できる情報に限定します。

メタデータ出力中に、次のデータアクセスやリンク更新を行ってはいけません。

- `TableDef.RecordCount`
- `OpenRecordset`、保存クエリ実行、`SELECT`
- `DCount`、`DSum`などの定義域集計
- `RefreshLink`
- フォームやレポートを開いてレコードソースを評価する操作

`TableDef.Fields`や`Indexes`でも、AccessのバージョンやODBCドライバーによりリンク先のスキーマ確認が発生する可能性があります。メタデータだけを読む実装は「レコードを読まない」ための境界であり、ネットワーク接続が絶対に起きない保証ではありません。接続自体を禁止する検証では、ネットワークを観測または遮断できる環境で実行し、実行経路と結果を記録します。

接続文字列は、出力、ログ、申し送り、公開資料に生のまま残しません。サーバー名、パス、ユーザー名、パスワード、トークンを含み得るため、リンク種別だけを記録するか、詳細を完全にマスクします。

推奨ゲート:

1. 本体ではなく名前付き作業コピーを対象にする。
2. メタデータ出力と`SaveAsText`出力の対象種別・成功件数・失敗件数を別々に記録する。
3. タイムアウトや接続待ちが起きた場合、同じプロパティ参照を繰り返さない。直前の操作、対象テーブル、リンク種別を記録して停止する。
4. 出力後に、接続文字列、資格情報、実サーバー名が出力・ログ・報告に残っていないことを検索で確認する。
5. フォーム、レポート、マクロ、モジュール、クエリの`SaveAsText`件数は、開始前に列挙した件数と照合する。

## 文字コードの注意

Accessの `SaveAsText` 出力は、オブジェクト種別によって文字コードが異なることがあります。

- VBAモジュール (`acModule`): CP932/SJIS系で出ることが多い。
- フォーム (`acForm`) / レポート (`acReport`): UTF-16 LE with BOMで出ることが多い。
- `FF FE` で始まるファイルは、UTF-16 LEとして読む。
- フォーム/レポートをCP932やUTF-8として決め打ちで読むと、文字化けやNULL文字混入の原因になる。
- AI向けにUTF-8コピーを作る場合は、生の出力を上書きせず、別フォルダに変換済みコピーを作る。

詳しくは [Accessテキスト資産の文字コード](10_access-text-encoding.md) を参照してください。

## 要件別ドキュメント

- [コンテンツの有効化を確認する](requirements/00_enable-active-content.md)
- [起動処理を止めて開発モードで開く](requirements/01_startup-bypass.md)
- [ExportAnalysisInfoを使える状態にする手順](requirements/02_import-analysis-module.md)
- [Access資産をAI向けにエクスポートする](requirements/03_export-access-assets-for-ai.md)
- [別PC・別Codexへ申し送りする](requirements/06_handoff-to-another-ai-agent.md)

## Codexへ渡す最小情報

```text
対象DBは C:\work\sample\app.accdb です。
本体ではなく作業コピーを使ってください。
Access資産をAI解析用にエクスポートしてください。
出力後は Latest フォルダを確認し、まず manifest またはログ、次に Objects と Schema を確認してください。
フォームのイベント処理を調べる場合は、フォーム定義とコードビハインドの両方を確認してください。
修正提案をする前に、現象と原因を分けて報告してください。
```

## 公開時の注意

公開用メモには、次を入れません。

- 顧客名
- 実DB名
- 実サーバ名
- 実ユーザー名
- パスワード
- 業務上の個人名
- 実データ

固有名は `app.accdb`, `sample.accdb`, `localhost,14333`, `test_db` のような一般名に置き換えます。
