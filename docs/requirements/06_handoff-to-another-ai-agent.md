# 別PC・別Codexへ申し送りする

## 目的

別PCや別スレッドのAIエージェントへ、Access解析作業を安全に引き継ぐためのメモ形式です。

## 申し送りテンプレート

```text
目的:
- 何を完成させたいか

対象:
- 本体DB:
- 元ACCDB SHA-256:
- 自動起動: あり / なし:
- 自動起動経路: AutoExec、起動フォーム、イベント、最初の呼出先関数
- 自動起動の無効化方法:
- 自動起動無効化状態: not-required / already-supported / added-and-verified
- 通常起動と無効化起動の検証証跡:
- 検証済みbaseline:
- baseline SHA-256:
- 対応INI SHA-256:
- 実装候補:
- テストDB:
- before/after Export:
- exporter SHA-256:
- before/after manifest SHA-256:

現在できていること:
- ...

未完了:
- ...

作業ルール:
- 本体DBは直接触らない
- 自動起動の有無、経路、無効化方法、検証証跡、採用baselineを依頼元へ開始前に宣言する
- 元SHA-256と記録が一致する検証済みbaselineを土台にする
- 失敗コピーは破棄する
- フォーム差し替えは作業コピー上で DeleteObject -> LoadFromText -> Compile

次にやること:
- 1項目だけ書く

詰まり:
- なし / 内容を書く
```

## 最小セット

別のAIに最低限渡す情報:

```text
1. 元ACCDB、baseline、候補の役割と各SHA-256
2. 自動起動経路、`SKIP_AUTOEXEC`分岐の状態、通常/スキップ起動の検証結果
3. 対応INI、Access版・build、参照設定
4. exporter SHA-256、before/after manifest SHA-256、各Exportの判定
5. 変更対象、before/after/diff、セルフレビューと独立レビューの指摘処理
6. Access COMで使うAutomationSecurityと起動バイパス
7. compile、reopen compile、自己テスト、GUI/帳票確認の結果
8. 最後に成功したコピーと、失敗コピーを再利用しないこと
9. 詰まり時の操作、エラー全文、確認済み事実、推定、次の安全な一手
```

## 注意

- 公開用には実名、顧客名、実サーバ名、DB名、資格情報、接続文字列、実レコードを書かない。
- 固有名は公開時に一般名へ置換する。
- 進捗は分母・分子で書く。
- 詰まりがある場合は、試したことと失敗結果を分けて書く。
- `Latest`や解析モジュール名だけを渡さず、正式baselineを再利用できる証跡を渡す。
