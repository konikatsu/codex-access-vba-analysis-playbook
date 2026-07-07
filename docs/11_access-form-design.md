# Accessフォームデザイン確認

AccessフォームをAIエージェントが新規作成・複製・改修するときの、見た目と操作性の再発防止ルールです。

Accessの既存システムでは、業務ロジックが合っていても、既存画面と比べて行高、余白、色、枠線、フォント、列幅がずれると、利用者には「別物」「読みにくい」「壊れている」ように見えます。

## 原則

- 見た目を推測で作らない。
- 必ず基準フォームを1つ決め、キャプチャと `SaveAsText` の数値を見て合わせる。
- 文字が読めない、列が見切れる、ラベルが折り返される状態はNG。
- 既存フォームにない色や行高を新しく発明しない。
- 機能テストがOKでも、画面キャプチャ比較がNGなら完了にしない。

## 作業前ゲート

作業前に次をメモへ書く。

```text
Target form:
Reference form:
Reference capture:
Expected row height:
Expected font:
Expected header color:
Expected detail background:
Expected border style:
Expected footer/total style:
Expected buttons:
```

例:

```text
Target form: EstimateList
Reference form: CustomerList
Expected row height: same as CustomerList detail row
Expected header color: copy from CustomerList header BackColor
Expected buttons: captions must be readable, fixed footer buttons for AI tests
```

## 基準フォームを採寸する

基準フォームは見た目のスクリーンショットだけでなく、`SaveAsText` の値で採寸する。

確認する主なプロパティ:

```text
Left
Top
Width
Height
BackColor
ForeColor
BorderColor
OldBorderStyle
SpecialEffect
FontName
FontSize
FontWeight
TextAlign
TopMargin
LayoutCachedLeft
LayoutCachedTop
LayoutCachedWidth
LayoutCachedHeight
```

注意:

- Accessのフォーム/レポート `SaveAsText` はUTF-16 LE with BOMのことが多い。
- 読む前にBOMを確認する。
- UTF-8やCP932決め打ちでフォームを読むと、文字化けやNULL文字混入の原因になる。
- 詳細は [Accessテキスト資産の文字コード](10_access-text-encoding.md) を参照する。

## 連続フォームの行高

連続フォームの行高は、見えている列だけを直しても揃わない。

Detailセクション内のすべてのコントロールを確認する。

- 表示列
- 非表示ID
- 幅1の隠し値保持コントロール
- ボタン
- ラベル
- サブフォーム

行高を揃えるときは、少なくとも次を一括で合わせる。

```text
Detail.Height
Control.Top
Control.Height
Control.LayoutCachedTop
Control.LayoutCachedHeight
```

隠しコントロールの `Height` や `LayoutCachedHeight` が古いままだと、見た目の行高が揃わないことがある。

## サブフォームの合計行

明細グリッドに属する合計行は、親フォームではなくサブフォームのフッターへ置く。

NG例:

- 明細はサブフォーム内にある。
- 合計ラベルと合計金額だけ親フォームに置く。
- 親フォーム側の座標で合わせたつもりになっている。

この場合、サブフォーム内の列幅、スクロールバー、罫線、余白と座標系が別になるため、合計行がずれて見える。

推奨:

```text
Subform.FormFooter
  Total label
  Total amount textbox
  Blank cells for remaining columns
```

親フォーム側に計算用コントロールを残す場合でも、利用者に見せる合計行はサブフォーム内に置く。

```text
Parent hidden value holder: txtTotalAmount
Subform footer display: =[Parent]![txtTotalAmount]
```

## 配色

配色は既存フォームからコピーする。

特に次は推測で決めない。

- ヘッダー色
- 入力不可セルの色
- 合計行の色
- 小計/税額/税込行の色
- 選択行の色
- 削除ボタンや警告色

基準フォームから `BackColor` を確認し、同じ値を使う。

## 幅と見切れ

文字が見切れる場合、縮小して詰め込むより、フォーム幅・列幅を広げる。

確認すること:

- 名称列や摘要列の先頭がボタンに隠れていない。
- 金額列の右端が見切れていない。
- ラベルが折り返されていない。
- ボタンキャプションが省略されていない。
- 右側に不自然な余白がない。
- 追加したボタンの分だけ、既存列を押し込んでいない。

## ボタンとAIテスト性

AIエージェントがGUI確認する画面では、小さすぎる行内ボタンだけに依存しない。

推奨:

- 行内ボタンは残す。
- 画面下部に固定位置の大きめボタンを追加する。
- キャプションは省略しない。
- `編` ではなく `編集` のように意味が分かる表記にする。
- `ControlTipText` を設定する。
- 行内ボタンと固定ボタンは同じ実処理を呼ぶ。

## 画面位置

ポップアップフォームは、既存UIが別方針でない限り中央表示にする。

注意:

- `DoCmd.MoveSize 900, 900` のような固定座標は中央表示ではない。
- `AutoCenter` や中央表示関数を使う場合でも、必ずキャプチャで実際の位置を確認する。
- 「中央表示済み」と書く前に、背後の親フォームとの位置関係を見て確認する。

## 0件表示

一覧フォームで0件表示を扱うときは、空の新規入力行が出ていないか確認する。

連続フォームでは `AllowAdditions=True` のままだと、データ0件でも空行が1行表示されることがある。

一覧専用なら、必要に応じて次を設定する。

```text
AllowAdditions = False
AllowDeletions = False
```

0件時には、行内の編集/削除ボタンが押せない、または表示されない状態にする。

## 完了ゲート

完了前に次を確認する。

```text
1. 作業コピーへ取り込み済み
2. Compile OK
3. Accessを閉じて再オープン後もCompile OK
4. 画面キャプチャを保存済み
5. 基準フォームのキャプチャと並べて確認済み
6. 文字が読める
7. ラベルが折り返されていない
8. 明細行の高さと隙間が基準フォームに近い
9. 見出し色、枠線、フォントサイズが基準フォームに近い
10. 右端の金額や電話番号が見切れていない
11. 合計行が明細列と揃っている
12. 右側に不自然な余白がない
13. ボタンが既存列に重なっていない
14. 0件表示時の空行/誤操作がない
15. 残課題があれば「完了」ではなく「保留」と明記した
```

## Codexへの指示例

```text
対象フォームは EstimateList です。
基準フォームは CustomerList です。
見た目は基準フォームに合わせてください。

特に、行高、行間、見出し色、枠線、フォントサイズ、ボタン幅、文字の見切れを確認してください。
非表示コントロールも含めて、行高を押し広げる要素がないか確認してください。

変更後は作業コピーへ取り込み、コンパイル、再オープン、キャプチャ比較まで実施してください。
文字が読めない状態、列が重なる状態、右端が見切れる状態、合計行が明細列とずれる状態はNGです。
```
