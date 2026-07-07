# Access VBA 64bit対応プレイブック

Access `.accdb` を 64bit Office 環境で動かすための、VBA API宣言修正と安全な作業手順のメモです。

このメモは公開用の一般化ナレッジです。実案件のDB名、サーバ名、ユーザー名、パスワード、業務固有名は入れないでください。

## 基本方針

- 本体DBを直接触らず、必ず作業コピーを対象にする。
- 1ファイルずつ処理する。
- 作業前にDB単位のバックアップを作る。
- AutoExec、起動フォーム、起動時マクロを実行させない開き方を先に確認する。
- VBEの自動コンパイルはダイアログを出すことがあるため、無人処理では慎重に扱う。
- 変換後は、残修正候補が0件であることを確認してから次のファイルへ進む。

## 推奨チェックリスト

1. 作業コピーを作る。
2. 作業コピーのさらに直前バックアップを作る。
3. AutoExec、起動フォーム、起動プロパティの有無を確認する。
4. 非実行で開ける方法を確認する。
5. VBAモジュールを `SaveAsText` で退避する。
6. API宣言とポインタ/ハンドル変数を64bit向けに修正する。
7. `LoadFromText` で作業コピーへ戻す。
8. もう一度 `SaveAsText` して、未変換の `Declare` や型不一致候補がないか確認する。
9. Accessプロセスと `.laccdb` が残っていないことを確認する。
10. 進捗メモへ、対象ファイル、修正モジュール、確認結果、未確認事項を記録する。

## 起動処理を止める注意点

Accessの自動起動経路は複数あります。

- AutoExecマクロ
- 起動フォーム
- 起動フォームのLoad/Openイベント
- 起動時に呼ばれる共通関数
- DBを開いた瞬間に走る接続初期化処理

`Application.AutomationSecurity = 3` だけで、すべての起動処理を止められるとは限りません。

安全な手順は、対象DBごとに「本当に起動処理が走っていない」ことを小さく確認してから、変換処理へ進むことです。

## LoadFromText失敗時の復旧

`DeleteObject` の後に `LoadFromText` が失敗すると、対象モジュールが欠落した中間状態になります。

この状態で無理に続行すると原因切り分けが難しくなるため、次を守ります。

- 実DB反映前に、コピーで `LoadFromText` まで通ることを確認する。
- 反映前バックアップを必ず残す。
- モジュール欠落が起きたら、作業前バックアップまたは成功済みコピーからDBごと復旧する。
- 壊れたDBは調査用に退避し、上書き消去しない。

## Declare修正の基本

64bit VBAでは、Windows API宣言に `PtrSafe` が必要です。

ただし `PtrSafe` を付けるだけでは不十分です。

- ポインタ、ハンドル、コールバック、アドレスを表す値は `LongPtr`。
- 通常の数値、フラグ、サイズ、成否値は多くの場合 `Long` のまま。
- 戻り値がハンドルやポインタのAPIは、戻り値も `LongPtr`。
- 戻り値が成否値のAPIは、原則 `Long` のまま。

## よくあるLongPtr対象

### ウィンドウ/プロセスハンドル

次のような名前は、用途がハンドルなら `LongPtr` にします。

- `hwnd`
- `hWnd`
- `hDC`
- `hProcess`
- `hKey`
- `hInternet`
- `hConnect`
- `hFind`

戻り値がハンドルになる代表例です。

- `FindWindow`
- `GetDesktopWindow`
- `GetDC`
- `OpenProcess`
- `SetActiveWindow`

一方、`GetDeviceCaps` や `GetSystemMetrics` の戻り値は通常 `Long` です。

### レジストリAPI

`RegCreateKeyEx` / `RegOpenKeyEx` の `phkResult` はレジストリキーのハンドルを返すため `LongPtr` です。

Declare側だけでなく、受ける変数も揃えます。

```vb
Dim result As LongPtr
Dim lngKey As LongPtr
```

`ByRef 引数の型が一致しません` が出る場合、Declare引数と受け変数の片側だけが `LongPtr` になっている可能性があります。

### SHBrowse / PIDL

次は `LongPtr` 対象です。

- `SHBrowseForFolder` の戻り値
- `SHGetPathFromIDList` の `pidl`
- `CoTaskMemFree` の `pv`
- `BROWSEINFO.hwndOwner`
- `BROWSEINFO.pidlRoot`
- `BROWSEINFO.lpfn`

古いコードで `Dim Pid&` のような型宣言文字を使っている場合は、`Dim Pid As LongPtr` に直し、使用箇所の `Pid&` も残さないようにします。

### OPENFILENAME

`OPENFILENAME` 構造体では、次のようなポインタ/ハンドル用途の項目を `LongPtr` にします。

- `hwndOwner`
- `hInstance`
- `lpstrCustomFilter`
- `lCustData`
- `lpfnHook`
- `lpTemplateName`

### Winsock

Winsock系では、ソケットやホスト情報のポインタに注意します。

- `socket` の戻り値は `LongPtr`。
- `SOCKET` 引数は `LongPtr`。
- `gethostbyname` / `inet_ntoa` の戻り値は `LongPtr`。
- `HOSTENT.hName`, `hAliases`, `hAddrList` は `LongPtr`。
- `WSADATA.lpVendorInfo` は `LongPtr`。

構造体サイズやポインタコピー長を固定値 `4` にしている古いコードは、64bitで壊れる可能性があります。

## 変換スクリプトを書くときの注意

- 複数行Declareでは、関数名の行と戻り値の `As Long` 行が分かれることがある。
- 関数名を保持し、Declareの終端行まで追跡して戻り値を判断する。
- コメント行の古いDeclareは変換対象外にする。
- `Rem` コメント行も変換対象外にする。
- 日本語混じり変数名を正規表現で直接拾うと、Windows PowerShell 5で文字化けする場合がある。
- 実DBへ反映する前に、コピーで `LoadFromText` まで通す。

## 進捗メモに残す項目

- 対象DB
- 作業コピー名
- 修正したモジュール
- 主な修正内容
- DryRun結果
- コンパイル実施有無
- 起動処理の有無
- 発生したエラー
- 復旧に使えるバックアップパス

## 完了の考え方

64bit対応の「静的確認済み」と「実行確認済み」は分けて記録します。

- 静的確認済み: API宣言やLongPtr候補の修正を行い、残修正候補が0件。
- コンパイル確認済み: VBEコンパイルが通った。
- 実行確認済み: 実際の起動画面や主要機能を64bit Accessで確認した。

ダイアログを避けるためにコンパイルを未実施にした場合は、未実施であることを明記します。
