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
