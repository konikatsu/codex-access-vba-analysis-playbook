# Access/VBA作業プレイブック

Access / VBA の既存システムをAIエージェントで扱うときの、実務寄りの作業ルールです。

この文書では、固有の顧客名、DB名、サーバ名、実データは使いません。

## 基本姿勢

Access DBは、フォーム、レポート、標準モジュール、クラスモジュール、マクロ、リンクテーブル、外部SQL Serverが絡みます。AIエージェントで作業するときは、ひとつの失敗をDB全体の破損に広げない運用が重要です。

原則:

- 本体DBを直接触らない。
- 作業コピーで変更する。
- 成功した作業コピーを次の土台にする。
- 失敗した作業コピーは破棄する。
- 進捗をMarkdownに残す。
- 迷ったら実装を止め、詰まり内容を書く。

## 作業コピーの積み上げ方

おすすめは、機能単位で作業コピーと作業フォルダを分ける方法です。

```text
work/
  estimate_progress.md
  step001_list/
    copies/
    exports/
    logs/
    screenshots/
    notes.md
  step002_entry/
    copies/
    exports/
    logs/
    screenshots/
    notes.md
  step003_report/
    copies/
    exports/
    logs/
    screenshots/
    notes.md
```

作業の流れ:

```text
本体DBまたは直近の成功コピー
-> 新しい作業コピーを作る
-> そのコピーだけ変更する
-> コンパイルする
-> GUIまたは静的確認をする
-> 成功したら次の作業の土台にする
-> 失敗したらコピーごと破棄する
```

この運用では、フォーム単体の退避よりも、DBコピー単位の戻しを優先します。

## 進捗メモ

AIエージェントが長時間作業すると、現在地が曖昧になりがちです。進捗メモは必ず作ります。

例:

```markdown
# 見積機能 進捗メモ

更新日: 2026-06-22

## 全体進捗

- 5 / 8 完了
- 約65%

## 進捗表

| No | 項目 | 状態 | 完了したこと | 残っていること | 次にやること |
|---:|---|---|---|---|---|
| 1 | Docker SQL Server環境 | 完了 | テストDBを起動 | なし | 維持 |
| 2 | DB接続・作業コピー運用 | 完了 | 作業コピー方針を決定 | なし | 次コピー作成 |
| 3 | 一覧画面 | 作業中 | 基本表示 | 固定編集ボタン | GUI確認 |
```

書く内容:

- 全体進捗
- 分母・分子
- 作業項目ごとの状態
- 現在の詰まり
- 次の30分でやること

## フォーム差し替え

作業コピーであれば、フォーム差し替えは次の手順でよいです。

```text
DeleteObject
-> LoadFromText
-> RunCommand(126)
-> Accessを閉じる
-> 再オープンしてフォーム確認
```

ポイント:

- 本体DBでは直接やらない。
- 既存フォームを変名してから取り込む必要は通常ない。
- 失敗したコピーは破棄する。
- 成功したコピーを次の土台にする。

差し替え前の `SaveAsText` は必須ではありません。DBコピー単位で戻れるなら省略できます。

ただし、次の場合は `SaveAsText` 退避が有効です。

- 差分を見たい。
- フォーム単体で戻したい。
- 証跡を残したい。
- 作業コピーを再利用せざるを得ない。

## LoadFromTextは不安定なのか

`LoadFromText` 自体が常に不安定というより、失敗条件が分かりにくいAPIです。

よくある原因:

- 同名オブジェクトが残っている。
- 直前の削除失敗やロックで作業コピーが中途半端になっている。
- Accessプロセスや `.laccdb` が残っている。
- フォームのコードビハインドが残骸として残っている。
- `AutomationSecurity` の用途が合っていない。
- サンドボックスや権限不足で外部COM操作が制限されている。
- 文字コードやテキスト定義の形式が崩れている。

失敗したら、そのDB上で修復を続けるより、新しい作業コピーからやり直します。

## AutomationSecurityと起動方法

用途別の目安:

| 用途 | 推奨 |
|---|---|
| モジュール取り込み | `AutomationSecurity = 1` |
| `Application.Run` 実行 | `AutomationSecurity = 1` |
| `RunCommand(126)` コンパイル | `AutomationSecurity = 1` |
| GUI確認でマクロを止めたい | `AutomationSecurity = 3` |
| 開発モードで通常起動したい | `/cmd SKIP_AUTOEXEC` |

開発モード起動の例:

```powershell
Start-Process msaccess.exe "`"C:\path\to\app.accdb`" /cmd SKIP_AUTOEXEC"
```

DB側には、AutoExecから呼ばれる入口関数に次のような分岐を用意しておくと扱いやすくなります。

```vb
Public Function AutoExecMain()
    If InStr(1, Nz(Command(), ""), "SKIP_AUTOEXEC", vbTextCompare) > 0 Then
        Debug.Print "AutoExec skipped."
        Exit Function
    End If

    Call StartUp
End Function
```

## コンパイル

外部COMからのコンパイルは、VBEメニュー操作より `RunCommand(126)` が簡単です。

```powershell
$access.RunCommand(126)
```

`126` は `acCmdCompileAndSaveAllModules` です。

## GUIテストしやすい画面にする

AIエージェントは、Accessの連続フォーム内にある小さなボタンを押しづらいことがあります。

テストしやすくする工夫:

- 行内の小さな `編` ボタンを `編集` にする。
- ボタン幅を広げる。
- `ControlTipText` を設定する。
- 画面下部に固定の `選択行を編集` ボタンを置く。
- 行内ボタンと固定ボタンは同じ処理を呼ぶ。
- 未選択時はメッセージを出して何もしない。

これはAIのためだけではなく、利用者にとっても分かりやすいUIになります。

## ボタン処理をテストしやすくする

AIエージェントにAccessフォームをテストさせる場合、実際のマウスクリック、座標クリック、`SendKeys` に依存すると不安定になりやすいです。

ボタンの `Click` イベントは `Private` のままにし、クリック時に実行する処理をフォーム内の `Public Sub` / `Public Function` に切り出します。

```vb
Private Sub btnSave_Click()
    Call ExecuteSave
End Sub

Public Sub ExecuteSave()
    ' 保存処理本体
End Sub
```

テスト側はフォームを開いてから、クリックイベントではなく実処理を直接呼びます。

```vb
DoCmd.OpenForm "mnt_Estimate_Entry", acNormal
Call Forms("mnt_Estimate_Entry").ExecuteSave
```

この形にすると、フォーカス、画面位置、リボン状態、確認ダイアログの有無に左右されにくくなります。

注意点:

- `btnSave_Click` 自体を `Public` にする必要はありません。
- 業務処理のテストには強いですが、ボタンと処理の配線ミスは検出しにくいです。
- 配線確認として、最後に1回はGUI上でボタンを押して、`Click` イベントから同じPublic処理が呼ばれることを確認します。
- 入力チェックが `BeforeUpdate` や `AfterUpdate` に分散している場合は、Public処理側にも必要な前提チェックを置きます。

## SQL Serverテスト環境

本物のSQL Serverへ直接接続して開発するより、Docker SQL Serverへ複製したテストDBを使う方が安全です。

方針:

- 本物のサーバは読み取りまたは移行元にする。
- Docker SQL Serverをローカルに立てる。
- AccessリンクテーブルをテストDBへ張り替える。
- DDLやデータ変更はDocker側で検証する。

名前付きインスタンス形式はDockerでは扱いづらいので、通常は `localhost,ポート番号` で接続します。

## 文字コード

Access / VBA / PowerShell / Markdown が混ざると、文字化けが起きやすくなります。

対策:

- 公開用MarkdownはUTF-8で保存する。
- PowerShell 5.1で日本語リテラルを多用しない。
- スクリプトは可能ならASCII中心にする。
- 日本語の本文はMarkdown側に寄せる。
- 文字化けしたファイルは、早めに正しい文字コードで書き直す。

## 推論設定

長時間の実装を常に最高推論にする必要はありません。

目安:

- 通常の実装、メモ更新、単純な確認: medium
- 複雑なAccess/VBE/COMの切り分け: high
- DB更新、フォーム差し替え、本体反映前レビュー: high以上
- 何度も同じ失敗をしている場合: 一時的に最高レベル

推論を上げるより、作業単位を小さくし、進捗メモを更新し、失敗コピーを破棄する方が効く場面も多いです。

## 申し送りテンプレート

別PCや別スレッドのCodexへ渡すときは、次の形式にします。

```text
目的:
- 何を完成させたいか

対象:
- 本体DB:
- 作業コピー:
- テストDB:

現在できていること:
- ...

未完了:
- ...

作業ルール:
- 本体DBは直接触らない
- 成功コピーを次の土台にする
- 失敗コピーは破棄する
- フォーム差し替えは作業コピー上で DeleteObject -> LoadFromText -> Compile

次にやること:
- 1項目だけ書く

詰まり:
- なし / 内容を書く
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

固有名は、`app.accdb`, `test_db`, `localhost,14333` のような一般名に置き換えます。
