# Access作業共通ルール

CodexなどのAIエージェントでAccess DBを解析・修正・検証するときの共通ルールです。

この文書は、個別案件の作業に入る前に最初に確認する入口として使います。

## 0. stage0を作る

まず本体DBを直接触らず、`stage0_xxx` のような作業コピーを作ります。

例:

```text
stage0_original_copy.accdb
stage0_startup_bypass.accdb
stage0_assets_export_success.accdb
```

方針:

- 本体DBは直接変更しない。
- 成功したコピーを次の土台にする。
- 失敗したコピーは `failed001` などの名前で残すか破棄し、同じコピーを修復し続けない。
- `.accdb` が `.ini` や外部設定に依存する場合は、設定ファイルも一緒に確認する。

## 1. 最初は `AutomationSecurity = 3` で安全確認する

Access DBをCOMから安全に開いて、AutoExec系の起動マクロを止めたい場合は、
`OpenCurrentDatabase` より前に `AutomationSecurity = 3` を設定します。

```powershell
$access = New-Object -ComObject Access.Application
$access.AutomationSecurity = 3
$access.Visible = $true
$access.OpenCurrentDatabase($dbPath)
```

注意:

- `AutomationSecurity = 3` はCOMで開く場合だけ効きます。
- ダブルクリックや `Start-Process app.accdb` には効きません。
- 起動フォームのLoadイベントなど、マクロ以外の起動処理は止まらない場合があります。
- `Application.Run` でVBAを実行したい作業には向きません。

## 1.5. Accessの信頼設定をDB破損より先に疑う

VBE/COM操作で不可解なエラーが出た場合、いきなりDB破損やVBAプロジェクト破損と決めつけないでください。

先に確認すること:

- 画面上部の「コンテンツの有効化」をクリック済みか。
- VBEで `デバッグ -> コンパイル` が通るか。
- `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` がONか。
- COMでVBAを実行する確認に `AutomationSecurity = 3` を使っていないか。

典型的なNG:

- 「コンテンツの有効化」を確認せず、`VBAプロジェクトが破損しています` をDB破損扱いする。
- `ActiveVBProject = null` を見て、すぐにVBE自動編集不能と判断する。
- 非表示COMで重いフォームを開き、タイムアウト原因を見えないまま推測する。

正しい順番:

```text
1. 作業コピーをGUIで開く
2. コンテンツの有効化を確認する
3. VBEで手動コンパイルする
4. VBAプロジェクト オブジェクトモデルへのアクセス信頼を確認する
5. COMでは AutomationSecurity=1 で短い確認だけ行う
6. それでも失敗する場合に、DB破損やフォーム定義破損を疑う
```

詳しくは [コンテンツの有効化を確認する](requirements/00_enable-active-content.md) を参照してください。

## 2. `/cmd SKIP_AUTOEXEC` を組み込む

GUI/VBEで開発モード起動したい場合は、DB側に `/cmd SKIP_AUTOEXEC` を受ける入口を作ります。

起動例:

```powershell
Start-Process msaccess.exe "`"C:\work\sample.accdb`" /cmd SKIP_AUTOEXEC"
```

VBA例:

```vb
Public Function StartUp()
    If InStr(1, Nz(Command(), ""), "SKIP_AUTOEXEC", vbTextCompare) > 0 Then
        Debug.Print "StartUp skipped."
        Exit Function
    End If

    ' Existing startup work.
End Function
```

ポイント:

- `/cmd SKIP_AUTOEXEC` はDB側の `Command()` 判定で起動処理を抜ける仕組みです。
- `AutomationSecurity = 1` とは役割が違います。
- 起動処理を直接コメントアウトするより、戻し忘れが起きにくいです。

## 3. 資産出力ツールを組み込む

起動処理を止められる状態にしてから、解析用モジュールや資産出力ツールを組み込みます。

推奨:

- 標準モジュールを `LoadFromText` で取り込む。
- `LoadFromText` が不安定な場合は、VBEから手動インポートする。
- 取り込み後にコンパイルする。
- 出力先は日時付きフォルダと `Latest` を分ける。

## 4. 資産出力ツールは `AutomationSecurity = 1` で実行する

`Application.Run` や `RunCommand(126)` を実行する場合は、
`OpenCurrentDatabase` より前に `AutomationSecurity = 1` を設定します。

```powershell
$access = New-Object -ComObject Access.Application
$access.AutomationSecurity = 1
$access.Visible = $false
$access.OpenCurrentDatabase($dbPath)

$access.Run("ExportAnalysisInfo")
$access.RunCommand(126)
```

注意:

- `AutomationSecurity = 1` はVBA実行を許可するための設定です。
- `AutomationSecurity = 1` は起動処理をスキップしません。
- 起動処理を止めたい場合は `/cmd SKIP_AUTOEXEC`、Shift-bypass、または
  `AutomationSecurity = 3` の安全確認ルートと分けて考えます。

## 5. 標準フロー

Access解析の最初の流れは、次を標準にします。

```text
0. stage0_xxx 作業コピーを作成する
1. AutomationSecurity = 3 で安全に開き、起動時に何が動くか確認する
2. GUIでコンテンツの有効化とVBE手動コンパイルを確認する
3. VBAプロジェクト オブジェクトモデルへのアクセス信頼を確認する
4. stage0_xxx に /cmd SKIP_AUTOEXEC 分岐を組み込む
5. /cmd SKIP_AUTOEXEC で開き、起動処理が止まることを確認する
6. 資産出力ツールを組み込む
7. AutomationSecurity = 1 で開き、資産出力ツールを実行する
8. 出力結果を確認し、stage0_success として保存する
```

## 6. 危険操作前に書くこと

フォーム/レポート/モジュールの差し替え、DDL、データ更新を行う前に、必ず次を整理します。

```text
対象DB:
作業コピーか本体か:
対象オブジェクト/テーブル:
実行予定コマンド/SQL:
事前バックアップまたは直前成功コピー:
成功確認方法:
戻し方:
```

## 7. フォーム/レポート差し替え

作業コピーであれば、フォームやレポートの差し替えは
`DeleteObject -> LoadFromText` で進めてよいです。

標準手順:

```text
作業コピー作成
-> 必要なら SaveAsText で退避
-> DeleteObject
-> LoadFromText
-> RunCommand(126) でコンパイル
-> Accessを閉じる
-> 再オープンしてコンパイル/フォーム表示確認
-> 成功コピーとして保存
```

注意:

- 本体DBで直接実行しない。
- 失敗したコピーを修復し続けない。
- 成功コピーを次の作業の土台にする。

## 8. 文字コード

Access/VBA資産は文字化けしやすいため、文字コードを明示します。

推奨:

- Accessの `SaveAsText` / `LoadFromText` 対象はSJIS/CP932前提で扱う。
- 日本語を含む `.ps1` もSJIS/CP932で保存する運用に統一する。
- PowerShellでAccess資産を読む/書く場合は
  `[System.Text.Encoding]::GetEncoding(932)` を明示する。
- MarkdownやPythonはUTF-8でよい。

例:

```powershell
$encoding = [System.Text.Encoding]::GetEncoding(932)
$text = [System.IO.File]::ReadAllText($path, $encoding)
[System.IO.File]::WriteAllText($path, $text, $encoding)
```

## 9. SQL Server接続

テスト環境がある場合は、本番SQL Serverではなくローカル/Docker SQL Serverを優先します。

方針:

- 接続先は作業メモに明記する。
- パスワードや資格情報をログやチャットに出さない。
- DDLやデータ更新SQLは、実行前にSQL全文と対象DBを確認する。
- 既存の売上系など、触らないと決めたテーブルは明示して守る。

## 10. 詰まった時

同じ操作を繰り返す前に、次を整理して相談します。

```text
何をしようとしているか:
対象DB/ファイル/オブジェクト:
作業コピーか本体か:
実行したコマンド/操作:
エラー全文:
最後に成功したコピー:
次に試したいこと:
```
