# Accessテキスト資産の文字コード

Accessの `SaveAsText` / `LoadFromText` と、AI向けに変換したテキストを扱うときの文字コードルールです。

## 結論

Access資産の文字コードは決め打ちしません。

特に、`SaveAsText` の出力は環境や対象オブジェクトによって見え方が変わるため、先頭バイトを確認してから読みます。

典型例:

- VBAモジュール (`acModule`): CP932/SJIS系で出ることが多い。
- フォーム (`acForm`) / レポート (`acReport`): UTF-16 LE with BOMで出ることが多い。

ただし、最終判断は必ずファイルの先頭バイトで行います。オブジェクト種別だけで決め打ちしないでください。

よくあるパターン:

- `FF FE` で始まる: UTF-16 LE with BOM。PowerShellでは `-Encoding Unicode`。
- `EF BB BF` で始まる: UTF-8 with BOM。
- BOMなしで日本語が多い: CP932/SJISの可能性あり。
- `V<NUL>e<NUL>r<NUL>s<NUL>i<NUL>o<NUL>n<NUL>` のように文字間にNULLが見える: UTF-16 LEをUTF-8やCP932として誤読している可能性が高い。
- `V<NUL>e<NUL>r...` のように見える: UTF-16のBOMや本文を別エンコードとして扱った疑いが強い。

## 最初に見ること

ファイルを開く前に、先頭バイトとNULLバイトの有無を確認します。

```powershell
$path = 'C:\work\exports\Forms\Form_Main.txt'
$bytes = [System.IO.File]::ReadAllBytes($path)
$first16 = ($bytes[0..15] | ForEach-Object { $_.ToString('X2') }) -join ' '
$nulCount = ($bytes[0..([Math]::Min(199, $bytes.Length - 1))] | Where-Object { $_ -eq 0 }).Count

[pscustomobject]@{
    Path = $path
    Length = $bytes.Length
    First16 = $first16
    NulCountFirst200 = $nulCount
}
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
acForm/acReport   -> UTF-16 LE候補。FF FEが出たらUTF-16 LEとして読む。
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
$bytes = [System.IO.File]::ReadAllBytes($dst)
($bytes | Where-Object { $_ -eq 0 }).Count
```

## やってはいけないこと

- `SaveAsText` は常にCP932/SJISだと決め打ちする。
- フォルダ名に `_utf8` と付けただけで、UTF-8変換済みとして扱う。
- UTF-16 LEのファイルを `-Encoding UTF8` で読み、その結果を上書き保存する。
- 生の `SaveAsText` 出力を直接変換・上書きする。
- ファイル名が正しく日本語表示されることを、本文エンコードが正しい証拠にする。

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
