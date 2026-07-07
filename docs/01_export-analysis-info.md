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
7. AIが読みやすい形で出力結果、注意点、残課題をMarkdownにまとめる
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
