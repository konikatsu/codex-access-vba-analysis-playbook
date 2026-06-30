# Accessテキスト資産の文字コード

Accessの `SaveAsText` / `LoadFromText` と、AI向けに変換したテキストを扱うときの文字コードルールです。

## 結論

Access資産の文字コードは、オブジェクト種別ごとの典型パターンを出発点にします。

実務上よく観測されるパターン:

- VBAモジュール (`acModule`): CP932/SJIS系。
- フォーム (`acForm`) / レポート (`acReport`): UTF-16 LE with BOM。

ただし、`SaveAsText` は公式に細かな文字コード仕様が保証されているAPIではないため、最終判断は必ずファイルの先頭バイトで行います。
オブジェクト種別だけで盲信せず、BOM、NULLバイト、実際の表示結果で確認してください。

## このルールの位置づけ

この文字コード差は、Microsoftの公式ドキュメントで安定仕様として明記されているものではありません。

このプレイブックでは、次の根拠に基づく実務上の経験則として扱います。

- Access資産をGitなどで管理するVCS連携ツールの実装で、オブジェクト種別ごとの文字コード差を前提に処理している例がある。
- 開発者コミュニティや技術記事で、`SaveAsText` 出力をバイナリエディタ等で確認した結果が共有されている。
- 実案件でも、フォーム/レポートの `SaveAsText` 出力が `FF FE` で始まるUTF-16 LEとして観測されている。
- `SaveAsText` / `LoadFromText` は、歴史的にソース管理連携などで使われてきた非公式寄りのAPIであり、周辺仕様が公式に細かく保証されているとは限らない。

そのため、この文書では「典型例」は示しますが、公式仕様として盲信しません。
最終的な読み方は、対象ファイルごとのBOM、NULLバイト、実際の表示結果で決めてください。

よくあるパターン:

- `FF FE` で始まる: UTF-16 LE with BOM。PowerShellでは `-Encoding Unicode`。
- `EF BB BF` で始まる: UTF-8 with BOM。
- BOMなしで日本語が多い: CP932/SJISの可能性あり。
- `V<NUL>e<NUL>r<NUL>s<NUL>i<NUL>o<NUL>n<NUL>` のように文字間にNULLが見える: UTF-16 LEをUTF-8やCP932として誤読している可能性が高い。
- `V<NUL>e<NUL>r...` のように見える: UTF-16のBOMや本文を別エンコードとして扱った疑いが強い。

## 最初に見ること

ファイルを開く前に、先頭バイトとNULLバイトの有無を確認します。

Defenderなどのセキュリティ製品に引っかかりにくいよう、長い `powershell.exe -Command ...` ワンライナーではなく、名前付きスクリプトで確認します。

```powershell
powershell -ExecutionPolicy Bypass -File ".\examples\inspect-access-text-encoding.ps1" -Path "C:\work\exports\Forms\Form_Main.txt"
```

このスクリプトは、先頭バイト、サンプル中のNULLバイト数、推定エンコードを表示します。

やむを得ず手元で短く確認する場合も、コマンド履歴やログ上で意味が分かる形にしてください。
特に、次のような長いPowerShellワンライナーは避けます。

```powershell
powershell.exe -Command "... [IO.File]::ReadAllBytes(...) ... ToString('X2') ..."
```

判定目安:

```text
FF FE ...        -> UTF-16 LE
FE FF ...        -> UTF-16 BE
EF BB BF ...     -> UTF-8 BOM
00 が多数ある    -> UTF-16系を誤読している可能性
```

オブジェクト種別の目安:

```text
acModule          -> CP932/SJIS候補。ただしBOMがあればBOM優先。
acForm/acReport   -> UTF-16 LE with BOM候補。FF FEが出たらUTF-16 LEとして読む。
acMacro/acQuery    -> 先頭バイトで判断する。
```

## PowerShellで読む

UTF-16 LE with BOMの場合:

```powershell
$text = Get-Content -LiteralPath $path -Encoding Unicode -Raw
```

.NETで明示する場合:

```powershell
$text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::Unicode)
```

CP932/SJIS候補の場合:

```powershell
$encoding = [System.Text.Encoding]::GetEncoding(932)
$text = [System.IO.File]::ReadAllText($path, $encoding)
```

UTF-8の場合:

```powershell
$text = Get-Content -LiteralPath $path -Encoding UTF8 -Raw
```

## AI向けUTF-8コピーを作る

生の `SaveAsText` 出力は、そのまま保持します。

AI向けにUTF-8へ変換したい場合は、別フォルダを作ります。

```text
analysis\sample_accdb_raw\
analysis\sample_accdb_utf8_fixed\
```

UTF-16 LEからUTF-8へ変換する例:

```powershell
$src = 'C:\work\analysis\sample_accdb_raw\forms\Form_Main.txt'
$dst = 'C:\work\analysis\sample_accdb_utf8_fixed\forms\Form_Main.txt'

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null

$text = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::Unicode)
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($dst, $text, $utf8NoBom)
```

変換後は、通常の本文部分にNULLバイトが大量に残っていないことを確認します。

```powershell
$bytes = Get-Content -LiteralPath $dst -Encoding Byte -TotalCount 1000
($bytes | Where-Object { $_ -eq 0 }).Count
```

## やってはいけないこと

- `SaveAsText` は常にCP932/SJISだと決め打ちする。
- `acModule`、`acForm`、`acReport` の文字コード差を無視して同じ読み方をする。
- フォルダ名に `_utf8` と付けただけで、UTF-8変換済みとして扱う。
- UTF-16 LEのファイルを `-Encoding UTF8` で読み、その結果を上書き保存する。
- 生の `SaveAsText` 出力を直接変換・上書きする。
- ファイル名が正しく日本語表示されることを、本文エンコードが正しい証拠にする。
- バイト確認のために、長い `powershell.exe -Command ... ReadAllBytes ... ToString('X2')` を直接実行する。

## `LoadFromText` に戻す場合

`LoadFromText` に戻すファイルは、できるだけAccessが `SaveAsText` で出した形式に寄せます。

推奨:

1. 差し替え対象を作業コピーから `SaveAsText` で出力する。
2. そのファイルの文字コードを判定する。
3. 同じ文字コードで編集後ファイルを書き戻す。
4. `LoadFromText` は作業コピーで試す。
5. コンパイルと再オープン確認を行う。

AI向けUTF-8コピーは解析用として便利ですが、そのまま `LoadFromText` に使う前提にしないでください。

## トラブルの見分け方

文字化け例:

```text
V<NUL>e<NUL>r<NUL>s<NUL>i<NUL>o<NUL>n<NUL> <NUL>=<NUL>2<NUL>1<NUL>
```

これは、UTF-16 LEのBOMや本文を別エンコードとして読んだ可能性が高いです。

正しく読めている例:

```text
Version =21
VersionRequired =20
Begin Form
```

この層が解決するまでは、フォーム定義やVBAロジックの解析に進まないでください。
