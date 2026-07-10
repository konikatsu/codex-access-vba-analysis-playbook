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

## 0.5. 作業前ゲートを必ず書く

Access作業では、目的達成を急ぐほど共通ルールを飛ばしがちです。
特にCOM自動化、VBE編集、フォーム保存、DDL、データ更新の前には、作業を始める前に次を短く書いてから進めます。

```text
作業前ゲート:
1. 今触るDB:
2. 本体DBか作業コピーか:
3. 最後に成功したコピー:
4. ユーザー確認対象:
5. この操作が失敗したらコピーを破棄できるか:
6. コンテンツの有効化は済んでいるか:
7. VBAプロジェクト オブジェクトモデルへのアクセス信頼は必要か:
8. AutomationSecurity は 1 / 3 のどちらを、何の目的で使うか:
9. hidden COMではなく visible UI で確認すべき操作ではないか:
10. 実行後の成功判定:
11. 戻し方:
```

このゲートを書けない場合は、まだ実行しません。
実行前に足りない前提を確認するか、ユーザーへ相談します。

典型的なルール違反:

- 本体DBか作業コピーかを書かずに操作する。
- 最後に成功したコピーを確認せず、失敗コピーを修復し続ける。
- 作業コピー運用中なのに、ユーザー確認対象として元DBを案内する。
- Accessの信頼設定を確認せず、DB破損やCOM不安定と判断する。
- `AutomationSecurity = 3` のままVBA実行可否を判断する。
- 画面が見えないhidden COMで重いフォームを開き、タイムアウト原因を推測する。
- ユーザーへの手動依頼で、貼る場所や貼る範囲を曖昧にする。

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
- `/cmd SKIP_AUTOEXEC` はAccessのセキュリティ警告を解除する仕組みではありません。
- VBA/マクロ/フォームイベントの動作テストでは、作業フォルダをAccessの信頼済み場所に追加するか、画面の「コンテンツの有効化」を明示的に行います。
- 起動処理を直接コメントアウトするより、戻し忘れが起きにくいです。

推奨する信頼済み場所:

```text
プロジェクトルート、または作業コピーを置く親フォルダ
例: C:\dev\project-name\
```

設定場所:

```text
Access
-> ファイル
-> オプション
-> トラスト センター
-> トラスト センターの設定
-> 信頼できる場所
-> 新しい場所の追加
```

可能なら「サブフォルダーも信頼する」をONにします。

使い分け:

```text
/cmd SKIP_AUTOEXEC:
  DB側の起動処理だけをスキップする。
  VBA/フォームイベントは有効な状態で動作テストしたいときに使う。

Accessの信頼済み場所 / コンテンツの有効化:
  セキュリティ警告でVBAやフォームイベントが止まらないようにする。

AutomationSecurity = 3:
  COMからGUI確認だけを安全に行い、マクロを強制無効化したいときに使う。
  動作テストやApplication.Runの確認には使わない。
```

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
4. 動作テスト用フォルダをAccessの信頼済み場所に追加する、またはコンテンツ有効化手順を確認する
5. stage0_xxx に /cmd SKIP_AUTOEXEC 分岐を組み込む
6. /cmd SKIP_AUTOEXEC で開き、起動処理が止まることを確認する
7. VBA/フォームイベントの動作テストは、信頼済み場所またはコンテンツ有効化済みの状態で行う
8. 資産出力ツールを組み込む
9. AutomationSecurity = 1 で開き、資産出力ツールを実行する
10. 出力結果を確認し、stage0_success として保存する
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

## 6.5. ユーザーへ確認対象MDBを案内する前に確認する

Access作業でユーザーから「どのMDBを確認すればよいか」と聞かれた場合は、回答前に進捗メモを確認します。

確認する項目:

- 今触るDB
- 最後に成功したコピー
- ユーザー確認対象
- 元DB
- バックアップ
- データDB

回答の先頭には、確認で開くMDBを1つだけ書きます。
元DB、成功コピー、バックアップ、データDBは補足として分けます。

```text
確認で開くMDB:
  ...\work_xxx\copies\step002_check.mdb

補足:
- 最後に成功したコピー:
- 元DB:
- バックアップ:
- データDB:
```

注意:

- 作業コピー運用へ移行した後は、元DBをユーザー確認対象として案内しない。
- 元DBやバックアップを、確認で開くMDBのように書かない。
- フロントMDBの作業コピーと、リンク先データMDBは別項目として明記する。
- 進捗メモにユーザー確認対象がない場合は、案内前に確認対象を確定する。

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

最重要:

- Accessの `SaveAsText` / `LoadFromText` 対象を、SJIS/CP932前提で決め打ちしない。
- 先頭バイトを確認してから読む。
- 典型的には、VBAモジュール (`acModule`) はCP932/SJIS系、フォーム (`acForm`) / レポート (`acReport`) はUTF-16 LE with BOM。
- `FF FE` で始まる `SaveAsText` 出力は UTF-16 LE として扱う。
- `V<NUL>e<NUL>r<NUL>s<NUL>i<NUL>o<NUL>n<NUL>` や `V<NUL>e<NUL>r...` のように見える場合は、UTF-16系を別エンコードとして誤読している可能性が高い。
- AI向けにUTF-8へ変換する場合は、生の出力を上書きせず、別フォルダに変換済みコピーを作る。
- バイト確認は、長い `powershell.exe -Command ... ReadAllBytes ... ToString('X2')` ワンライナーではなく、名前付きスクリプトや読みやすい短いコマンドで行う。
- Access由来の日本語列名やラベルを含むPHP/HTMLを編集する場合、コンソール上の文字化け表示を正しい文字列として扱わない。

推奨:

- PowerShellでAccess資産を読む/書く場合は、ファイルごとにBOM/NULLバイトを見てエンコードを選ぶ。
- `FF FE` の場合は `[System.Text.Encoding]::Unicode` を明示する。
- BOMなしでCP932が疑われる場合だけ `[System.Text.Encoding]::GetEncoding(932)` を使う。
- MarkdownやPythonはUTF-8でよい。
- パッチ文脈には、文字化けした日本語ではなく、ASCIIの関数名、変数名、HTML class/id、SQLの英数字テーブル名など安定したアンカーを使う。

判定例:

```powershell
powershell -ExecutionPolicy Bypass -File ".\examples\inspect-access-text-encoding.ps1" -Path "C:\work\exports\Forms\Form_Main.txt"
```

詳しくは [Accessテキスト資産の文字コード](10_access-text-encoding.md) を参照してください。

## 9. SQL Server接続

テスト環境がある場合は、本番SQL Serverではなくローカル/Docker SQL Serverを優先します。

方針:

- 接続先は作業メモに明記する。
- パスワードや資格情報をログやチャットに出さない。
- DDLやデータ更新SQLは、実行前にSQL全文と対象DBを確認する。
- 既存の売上系など、触らないと決めたテーブルは明示して守る。
- `sqlcmd` がTLSや証明書まわりで失敗しても、SQL Server自体へ接続不能と断定しない。
- Web/PHP/ODBCで成功している接続は、同じ接続文字列をPowerShell ODBCで再現して実データ確認を優先する。
- 接続失敗の説明は、確認済み事実、推定、暫定回避を分ける。

詳しくは [sqlcmdを使えるようにする](07_sqlcmd-setup.md) を参照してください。

## 9.5. Web化画面の表示崩れ

Web化後の画面で表示崩れが起きた場合、見た目だけでデータ起因と断定しません。
原因説明は「推定」と「確認済み」を分け、DB実値、HTML構造、CSSの順で切り分けます。

注意:

- 先頭空白や改行が原因に見えても、DB実値を確認するまで断定しない。
- 日本語テーブル名/列名がコンソールで文字化けする場合、表示された文字列を正として扱わない。
- 一覧プレビューの2行省略では、`display: -webkit-box` / `-webkit-line-clamp` が table cell や `white-space: pre-wrap` と相性悪く見える場合がある。
- 業務一覧では、まず `display: block`、`max-height`、`overflow: hidden` のような単純なCSSで切り分ける。

詳しくは [Access Web化UI表示崩れの切り分け](13_access-web-ui-troubleshooting.md) を参照してください。

## 10. 詰まった時

同じ操作を繰り返す前に、Access共通playbook担当へ報告します。
原因が確定していなくても、詰まりの事実、試したこと、失敗結果を先に残します。

詰まり報告フォーマット:

```text
何をしようとしているか:
対象DB/ファイル/オブジェクト:
作業コピーか本体か:
最後に成功したコピー:
実行したコマンド/操作:
エラー全文:
疑っている原因:
次に試したいこと:
共通playbookへ残すべき知見候補:
```

特に次の場合は、自己判断で同じ操作を繰り返さず、早めに報告します。

- SaveAsText出力の文字コードで読めない、文字化けする、NULL文字が多い。
- Access COM / VBE操作が不安定、または同じエラーを繰り返す。
- 元DB、作業コピー、成功コピー、ユーザー確認対象を混同しそう。
- 64bit対応、参照設定、Declare修正で判断に迷う。
- Web化時にフォーム、クエリ、VBAイベント、テーブル更新の読み替えで迷う。
- ユーザーへ確認対象や作業手順を案内する前に不安がある。
