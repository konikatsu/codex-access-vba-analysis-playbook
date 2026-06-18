# RunCommand(126)でコンパイルする

Access VBAを外部COMからコンパイルする場合、次を使えます。

```powershell
$access.RunCommand(126)
```

`126` は `acCmdCompileAndSaveAllModules` です。

## 使いどころ

- 64bit対応後にVBA全体をコンパイルしたい。
- `Declare PtrSafe` / `LongPtr` 修正後にコンパイル確認したい。
- VBEの `CommandBars` からコンパイル項目が見つからない。

## 例

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.AutomationSecurity = 1
$access.OpenCurrentDatabase($dbPath)

$access.RunCommand(126)

$access.CloseCurrentDatabase()
$access.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
```

## 失敗した場合

VBEで `デバッグ > コンパイル` を実行し、最初に止まった1件だけを確認します。

必要な情報:

- エラー文
- モジュール名
- ハイライトされた行
- 前後20行程度
- 32bit Accessか64bit Accessか
- 変換ツール実行済みか

コンパイルは最初の1件で止まるため、まず最初のエラーだけ潰します。

## 64bit化でよくある型不一致

`Declare PtrSafe` と `LongPtr` への変換後は、API宣言だけでなく、その戻り値を受ける既存変数も確認します。

例:

```vb
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

Dim lngRtn As Long
Dim hMemory As LongPtr

lngRtn = GlobalFree(hMemory)
```

64bit VBAでは、`GlobalFree` の戻り値が `LongPtr` なのに `lngRtn As Long` で受けているため、型不一致になることがあります。

戻り値を使っていない場合は、代入せずに呼び出します。

```vb
Call GlobalFree(hMemory)
```

戻り値を使う必要がある場合は、受ける変数側も `LongPtr` にします。

```vb
Dim freeResult As LongPtr
freeResult = GlobalFree(hMemory)
```

64bit化では、次の両方を確認します。

- API宣言の引数と戻り値
- API呼び出し側で戻り値やハンドルを受ける変数
