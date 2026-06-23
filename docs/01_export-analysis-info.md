# Access資産をAI(Codex)にエクスポートさせる手順です

Access / VBA の既存システムを Codex などのAIに解析させるには、まずAccess DB内の資産をテキストとして取り出す必要があります。

この文書は、そのための入口です。具体的な作業は要件別に分冊しています。

特定の顧客名、DB名、サーバ名、実データは含めません。

## 要件別ドキュメント

- [コンテンツの有効化を確認する](requirements/00_enable-active-content.md)
- [起動処理を止めて開発モードで開く](requirements/01_startup-bypass.md)
- [ExportAnalysisInfoを使える状態にする手順](requirements/02_import-analysis-module.md)
- [Access資産をAI向けにエクスポートする](requirements/03_export-access-assets-for-ai.md)
- [別PC・別Codexへ申し送りする](requirements/06_handoff-to-another-ai-agent.md)

## 全体像

```text
1. 起動処理を止めて対象DBを開く
2. ExportAnalysisInfoを実行する
3. LatestフォルダをAIに読ませる
4. 結果をMarkdownで申し送る
```

## 同梱コード

- エクスポート実行用モジュール: [`tools/GsTools_analysisinfo.bas`](../tools/GsTools_analysisinfo.bas)
- VBE貼り付け用コード: [`tools/GsTools_analysisinfo_for_vbe_paste.bas`](../tools/GsTools_analysisinfo_for_vbe_paste.bas)

## AIに渡す最小情報

```text
解析結果は C:\work\exports\app.accdb\Latest にあります。
まず manifest.json を読み、次に forms, modules, queries, tables, relations を確認してください。
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
