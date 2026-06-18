# 解析情報をエクスポートする考え方

Access DBの解析では、フォームやモジュールをテキスト化してから検索するのが有効です。

## 出力したい情報

- フォーム定義
- レポート定義
- マクロ
- 標準モジュール
- クラスモジュール
- クエリSQL
- テーブル定義
- リレーション
- 参照設定
- DBプロパティ

## SaveAsText

Accessには、オブジェクトをテキスト出力する `SaveAsText` があります。

```powershell
$access.SaveAsText(5, 'ModuleName', 'C:\work\exports\ModuleName.mdl')
```

種別は `LoadFromText` と同じです。

```text
1 = Query
2 = Form
3 = Report
4 = Macro
5 = Module
```

## Latestフォルダ

解析出力は、日時付きフォルダと `Latest` の両方を作ると扱いやすくなります。

```text
DefinesSample.accdb\
  Exports\
    20260618_153000\
  Latest\
```

AIエージェントやエディタには、基本的に `Latest` を見せます。

## 検索例

```powershell
rg -n "btnSave_Click|Application.Run|LoadFromText|LongPtr" "C:\work\DefinesSample.accdb\Latest"
```

## 注意

AccessのエクスポートファイルはCP932の場合があります。  
日本語コメントやフォーム定義を扱うときは、エンコーディングに注意してください。
