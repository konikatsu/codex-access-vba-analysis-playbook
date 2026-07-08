# DOCX成果物のレンダリングQA

Access Web化や移行作業では、解析結果、設計書、説明書、提案書などをDOCXで納品・共有することがあります。
DOCXは見た目の崩れに気づきにくいため、AIエージェントで作成した場合も、PDFやPNGへレンダリングして表、図、改ページ、余白を確認します。

## 1. Windowsでは `soffice.com` を優先する

Windows版LibreOfficeには、同じフォルダに `soffice.exe` と `soffice.com` が存在する場合があります。
無人変換やPowerShellからの確認では、`soffice.exe` ではなく `soffice.com` を明示的に使うと安定しやすいです。

例:

```powershell
& "C:\Program Files\LibreOffice\program\soffice.com" --headless --version
```

`soffice.exe --headless` は、環境によってはダイアログ、無出力、タイムアウトにつながることがあります。
CLIで使う場合は、まず `soffice.com --headless --version` がコンソールへ正常に出力するか確認します。

## 2. wingetインストール後に確認すること

LibreOfficeをwingetで入れた後は、PATHだけで判断せず、実体を確認します。

```powershell
winget install --id TheDocumentFoundation.LibreOffice --source winget --accept-package-agreements --accept-source-agreements --silent

Test-Path "C:\Program Files\LibreOffice\program\soffice.com"
& "C:\Program Files\LibreOffice\program\soffice.com" --headless --version
```

`where soffice` は参考にはなりますが、Windowsでは `soffice.exe` と `soffice.com` のどちらが使われるかを別途確認してください。

## 3. 変換テスト

DOCXからPDFへ変換できるか、最小の入力で確認します。

```powershell
$soffice = "C:\Program Files\LibreOffice\program\soffice.com"
$input = "C:\work\sample.docx"
$outDir = "C:\work\rendered"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
& $soffice --headless --convert-to pdf --outdir $outDir $input
```

PDF化できたら、PopplerなどでPNGへ変換し、表や図の崩れを確認します。

## 4. タイムアウトやダイアログが出る場合

LibreOfficeがGUIダイアログを出していると、AIエージェントからはタイムアウトに見えることがあります。
次を順番に確認します。

```powershell
Get-Process soffice* -ErrorAction SilentlyContinue
Stop-Process -Name soffice,soffice.bin -Force -ErrorAction SilentlyContinue
```

専用プロファイルで切り分けます。

```powershell
$soffice = "C:\Program Files\LibreOffice\program\soffice.com"
$profile = "file:///C:/work/lo-profile-docx-qa"
& $soffice "-env:UserInstallation=$profile" --headless --version
```

`bootstrap.ini` 破損などのダイアログが出る場合は、無人変換へ進む前にLibreOffice自体の起動確認を優先します。

## 5. documents skillのレンダリングで詰まる場合

DOCXレンダリング用スクリプトが `soffice` という名前でLibreOfficeを呼ぶ場合、Windowsでは意図せず `soffice.exe` を拾うことがあります。
タイムアウトする場合は、次を確認します。

- `soffice.com --headless --version` が単体で成功するか。
- PATH上で `soffice.com` を優先できているか。
- スクリプト側でLibreOffice実行ファイルを明示指定できるか。
- 専用の `-env:UserInstallation=file:///...` プロファイルで再現するか。

確認が通るまでは、成果物DOCXの中身を直し続けず、まずレンダリング環境を切り分けます。
