# Gemini CLI / Antigravity CLI連携

Access Web化や移行作業では、調査、要約、代替案、一次レビューを外部AIに分担させると、Codex側は実査、根拠確認、設計判断、編集、テストに集中できます。
この文書では、Gemini CLIまたはAntigravity CLIを補助調査役として使うときの境界を定めます。

2026-07-17時点では、個人Google OAuthでGemini CLIを使う経路は対象外です。個人向けGemini Code Assist、Google AI Pro、Google AI Ultraでは、Gemini CLIのリクエスト提供が2026-06-18に終了しており、GoogleはAntigravity CLIへの移行を案内しています。Gemini CLIそのものを削除するのではなく、認証方式に応じて使い分けます。

## 1. 基本方針

Geminiの回答は、最終判断ではなく下書きとして扱います。
Codexは、Geminiの提案をそのまま採用せず、公式資料、実ファイル、実DB、テスト結果で確認します。

役割分担:

- Gemini: 技術選択肢の洗い出し、長文要約、移行案の比較、リスク列挙、一次レビュー、テスト観点の下書き。
- Codex: 対象範囲の決定、Access資産の実査、根拠の検証、設計決定、編集、テスト、差分確認、最終報告。

## 2. CLIの選択

| 認証・契約 | 使うCLI | 扱い |
| --- | --- | --- |
| Gemini APIキーまたはVertex AIで利用できる環境 | Gemini CLI | 組織の認証・利用規約に従って利用する。 |
| Gemini Code Assist Standard / Enterprise | Gemini CLI | Googleの案内では影響対象外。組織の設定を確認して利用する。 |
| 個人Google OAuth、Google AI Pro、Google AI Ultra | Antigravity CLI | Gemini CLIの個人向け経路は対象外。Antigravityへ移行する。 |

実機ではGemini CLI 0.51.0で個人Google OAuthの認証後に `UNSUPPORTED_CLIENT` / `IneligibleTierError` が出た。このエラーだけから原因を一般化せず、上記のGoogle公式案内の適用範囲と照合する。

## 3. 呼び出し

### 3.1 Gemini CLI

直接呼び出す例:

```powershell
gemini -p "質問"
```

プロジェクトやPCごとに共通ラッパーがある場合は、そのラッパーを使ってもよいです。
ただし、公開playbookには個人環境の認証情報や秘密情報を含めません。

```powershell
& "C:\path\to\Ask-Gemini.ps1" -Prompt "質問"
```

### 3.2 Antigravity CLI

Windowsでは公式インストーラーを使う場合がある。ダウンロードして実行する前に、組織のソフトウェア導入ルールと公式の最新手順を確認する。

```powershell
irm https://antigravity.google/cli/install.ps1 | iex
```

インストール後は、デスクトップ版の導入有無に関わらず、CLIが使えることを明示的に確認する。PATHに反映されない場合は新しいターミナルを開くか、CLIの絶対パスを使う。

```powershell
agy --version
agy --help
& "$env:LOCALAPPDATA\agy\bin\agy.exe" --version
```

調査専用の標準呼び出し:

```powershell
agy --mode plan --print "質問" --print-timeout 300s
```

共通ラッパーを使う場合の例:

```powershell
& "C:\path\to\Ask-Antigravity.ps1" -Prompt "質問"
```

`--mode plan` を既定にする。`accept-edits` は、外部AIに編集を明示的に許可する必要があり、対象・戻し方・検証方法を別途確定した場合だけ使う。`--dangerously-skip-permissions` は使わない。

## 4. 背景は明示する

GeminiまたはAntigravityとCodexの会話履歴や記憶は自動共有されません。
必要な背景だけをプロンプトに明示します。

含めるもの:

- 目的
- 匿名化した対象範囲
- 前提条件
- 既に確認済みの事実
- 禁止事項
- 欲しい出力形式

含めないもの:

- 顧客情報
- 個人情報
- パスワード
- APIキー
- 接続文字列
- 実レコード
- 公開できない業務ロジック固有名

必要な場合は、匿名化したスキーマ、SQL、VBA断片、画面仕様に限定します。

## 5. 作業ディレクトリに注意する

外部AI CLIは、起動時の作業ディレクトリ内を参照または操作できる可能性があります。
一般質問では、機密を含まない安全なフォルダから実行します。

プロジェクト固有の解析を許可する場合だけ、対象ディレクトリから実行します。
その場合も、プロンプトで境界を明示します。

```text
禁止:
- ファイル変更
- コマンド実行
- 外部送信
- 個人情報、認証情報、実レコードの利用
```

原則として、外部AIには読み取りと提案だけを依頼します。
ファイル変更、DB更新、テスト実行、pushはCodexが行います。

## 6. 依頼は小分けにする

非対話実行は、回答完成まで進捗率がほぼ見えない場合があります。
大きな依頼は、次のように分けます。

```text
1. 現状整理
2. 選択肢
3. リスク
4. 検証案
5. 推奨案
```

長いAccess資産や仕様書を一度に投げず、必要な断片と質問を絞ります。

## 7. 出力形式

外部AIには、確認済み、推測、要確認を分けて書かせます。
根拠URL、根拠ファイル、または確認方法も求めます。

推奨フォーマット:

```text
【確認済み】
- ...

【推測】
- ...

【要確認】
- ...

【選択肢比較】
- A:
- B:

【根拠または確認方法】
- ...
```

ただし、外部AIが示した引用、URL、技術案も、Codexが改めて確認します。
OpenAI/Codex、Microsoft Access、SQL Serverなど、公式資料で確認すべき内容は一次情報を優先します。

## 8. 依頼テンプレート

```text
外部AI CLIを主体に次を調査してください。

目的:
[目的]

対象:
[匿名化済み情報/明示したファイルのみ]

禁止:
- ファイル変更
- コマンド実行
- 個人情報・認証情報・実レコードの送信
- 指定外ディレクトリの参照

出力:
- 【確認済み】
- 【推測】
- 【要確認】
- 選択肢比較
- 根拠または確認方法

Codexは結果を検証して最終判断だけ行ってください。
```

## 9. Access Web化で使いやすい場面

向いている:

- AccessフォームやVBAイベントの読み替え案を複数出す。
- Web化方針の選択肢を比較する。
- 長い調査メモを要約する。
- テスト観点の下書きを作る。
- リスク一覧を作る。

向いていない:

- 実DB接続文字列や実レコードを渡す。
- MDB/ACCDBを直接変更させる。
- 公式確認なしで仕様判断を確定する。
- 外部AIの案をそのままGitHubへpushする。

最終的な編集、テスト、差分確認、ユーザー報告はCodexの責任で行います。

## 10. 参照

- [Gemini Code Assist consumer accounts (Google for Developers)](https://developers.google.com/gemini-code-assist/docs/deprecations/code-assist-individuals): 個人向けGemini Code Assist、Google AI Pro、Google AI UltraでのGemini CLI提供終了とAntigravityへの移行案内。2026-07-17確認。
