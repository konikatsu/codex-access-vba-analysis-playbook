# Access Web化UI表示崩れの切り分け

Access資産をWeb化した画面で表示崩れが起きた場合、見た目だけで原因を断定しないでください。
特に、Access由来の長文メモ、コメント、備考、履歴を一覧プレビューする画面では、データ、HTML、CSSのどれが原因かを分けて確認します。

## 1. 推定と確認済みを分ける

ユーザーへ説明するときは、次を分けます。

```text
推定:
確認済み:
未確認:
次に確認すること:
```

画面上で本文がずれて見えても、先頭空白や改行がDBに入っているとは限りません。
データ起因と説明する前に、DB実値を確認します。

## 2. DB実値、HTML、CSSの順で切り分ける

表示崩れでは、次の順で確認します。

1. DB実値
2. HTML構造
3. CSS

DB実値では、対象本文の先頭文字、文字数、trim後の文字数を確認します。

例:

```sql
SELECT
  LEN([本文]) AS body_length,
  LEN(LTRIM(RTRIM([本文]))) AS trim_length,
  LEFT([本文], 20) AS body_head
FROM ...
WHERE ...
```

`body_length` と `trim_length` が同じで、`body_head` も期待どおりなら、先頭空白や改行が原因ではない可能性が高いです。

## 3. 日本語テーブル名/列名が読めない場合

Access由来のSQL Server DBには、日本語テーブル名や日本語列名が残ることがあります。
PowerShellやコンソール表示で文字化けしたり、`????` のように表示されたりする場合、見えている文字列を正として扱わないでください。

有効な切り分け:

- 既存ソースで正しく動いているSQLや定数から対象名を確認する。
- コンソール表示ではなく、実ファイルの文字列やUnicodeコードポイントを使ってクエリ文字列を組み立てる。
- 可能なら、ASCIIの別名、ビュー、または取得済みの列メタデータを使って確認する。

文字化けした日本語をそのままクエリやpatch文脈へ貼ると、原因を増やします。

## 4. 一覧プレビューCSSの注意

業務一覧やコメント一覧で、長文を2行程度に省略したい場合があります。
`display: -webkit-box` と `-webkit-line-clamp` は便利ですが、次の組み合わせでは期待外の見え方になることがあります。

- table cell内
- `white-space: pre-wrap`
- 日本語長文
- 固定幅または狭い列
- 既存CSSで `vertical-align` や `text-align` が複雑に指定されている

まずは単純な指定を優先します。

```css
.list-preview {
  display: block;
  width: 100%;
  margin: 0;
  max-height: 3.2em;
  overflow: hidden;
  text-align: left;
  white-space: pre-wrap;
}
```

詳細画面や履歴画面の本文は完全表示のままにし、一覧プレビューだけ省略します。

## 5. 直した後に残すこと

修正メモには、次を残します。

- 最初の推定
- DB実値で確認したこと
- HTML/CSSで確認したこと
- 実際の原因または可能性が高い原因
- 一覧だけ直したのか、詳細本文にも影響するのか

誤った推定をした場合は、確認結果と訂正を残してください。
同じ表示崩れを次の担当がデータ起因と誤認しないようにします。

## 6. PHP構文確認は実行環境に合わせる

Web化プロジェクトがDockerで動いている場合、WindowsホストのPATHに `php` が無いことは異常ではありません。
毎回ローカルPATHやApache配下のPHPを探すのではなく、プロジェクトの実行環境で構文確認します。

例:

```powershell
docker ps --format "{{.Names}}\t{{.Image}}\t{{.Status}}"
docker exec <php-container> php -v
docker exec <php-container> php -l /path/in/container/file.php
```

記録すること:

- PHPを実行した場所: Windowsホスト / Dockerコンテナ / サーバ
- コンテナ名またはcomposeサービス名
- ホスト側パスとコンテナ内パスの対応
- `php -l` の結果

既知の環境前提がある場合は、「PATHにphpがない」を毎回報告せず、既定の構文確認コマンドへ直行してください。
