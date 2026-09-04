# Access修正の標準開発手順

日常作業用の短縮版です。判断根拠や実装例は[詳細リファレンス](17_access-development-workflow-reference.md)を参照してください。

## 0. 前提条件

起動直後に外部DBへ接続するAccessは、接続失敗やモーダルダイアログで開けなくなることがあります。そのため、自動起動の把握と安全な無効化方法の確保を作業着手の前提条件にします。

対象Access資産へ触る前に、依頼元へ必ず確認します。

1. 自動起動の有無: `あり / なし / 不明`
2. 既知の起動経路: AutoExec、起動フォーム、イベント、呼出先関数
3. 既存の無効化方法: `SKIP_AUTOEXEC / Shift-bypass / 保守フラグ / なし / 不明`
4. 検証済みbaselineの有無と、パス・SHA-256・証跡
5. 既知の外部接続、ダイアログ、副作用

依頼文に回答があっても復唱確認し、日時・回答者・回答を記録します。全5項目に`不明`を含む回答が揃い、復唱確認を終えた状態がヒアリング完了です。これは実装開始の許可ではありません。完了まではコピー、ハッシュ取得、DAO確認、StartupProbe、Access起動を行いません。

回答の`不明`は認めます。ヒアリング後に安全な事前調査で確定し、最終判定が`未確認`のまま実装へ進みません。

## 1. 自動起動とbaseline

ヒアリング完了後、次の順で開始条件を確定します。

1. 原本とINIをコピーし、原本とコピーのSHA-256を記録する。
2. 既存baselineの元SHA-256、Access版、参照設定、起動経路、Export条件が一致するか確認する。
3. 再利用できるbaselineがあれば、StartupProbeやIF追加を繰り返さない。
4. baselineがなければ、DAOと`StartupProbe`でAutoExec、起動フォーム、イベント、最初の呼出先関数を調べる。
5. 自動起動がなく、無効化が不要なら`not-required`と記録する。
6. 完全一致の`SKIP_AUTOEXEC`分岐が既にあれば、重複追加しない。
7. 分岐がなければ、Shift-bypassで開いた作業コピーの起動関数へ一度だけ追加する。
8. 通常起動と`/cmd SKIP_AUTOEXEC`起動を別々に検証する。
9. 合格した自動起動無効化対応版を開発baselineとして凍結する。

```vb
If StrComp(Nz(Command(), vbNullString), "SKIP_AUTOEXEC", vbBinaryCompare) = 0 Then
    Exit Function
End If
```

AutoExecマクロ自体は削除・改名しません。最初の副作用より前に分岐できない場合は、無効化完了と判定せず作業を止めます。

## 2. 作業開始宣言

実装前に依頼元へ次を報告します。

```text
- 自動起動: あり / なし
- 起動経路:
- 無効化方法:
- 通常起動・無効化起動の検証結果:
- 使用するbaselineとSHA-256:
```

## 3. 標準フロー

1. baselineを外部Exportツールへ渡し、二次コピーから`before`を出力する。
2. Access外で修正案とdiffを作り、セルフレビューする。
3. 必要な場合だけ独立レビューし、指摘を根拠確認する。
4. baselineから新しい候補を作り、対象資産だけ変更する。
5. 即時コンパイルし、変更対象だけExportして予定diffと照合する。
6. 候補を閉じ、仮想Shift + `AutomationSecurity=1`で再オープンして正式コンパイルする。
7. 閉じた候補を同じ外部ツールで全資産Exportし、baselineと比較する。
8. 必要な自己テスト、GUI、帳票、PDF確認を行う。
9. 合格版を名前付きで保存する。本番反映は別工程・別承認とする。

用途別の入口:

| 目的 | 経路 |
| --- | --- |
| 構造メタデータだけ読む | DAO読み取り専用 |
| 全資産Export | [外部Exportツール](16_access-external-export.md) |
| GUI/VBE作業 | `/cmd SKIP_AUTOEXEC` |
| COM静的操作 | 仮想Shift + `AutomationSecurity=3`（マクロ無効） |
| COMコンパイル | 仮想Shift + `AutomationSecurity=1` |
| 直接開けないDBの救出解析 | [空DBインポート](requirements/07_empty-database-recovery-import.md) |

COMの`OpenCurrentDatabase`へ`/cmd`は渡せません。`AutomationSecurity`だけで自動起動を止めたと判断しません。

## 4. 中断条件

次の場合は同じ操作を繰り返さず中断します。

- 必須ヒアリングが未完了
- 自動起動または無効化方法が未確認
- `AllowBypassKey=False`
- モーダルダイアログまたはCOMタイムアウト
- 原本やbaselineのSHA-256が途中で変化
- 予定外の資産差分、コンパイルエラー、残留Access PID・ロック

失敗コピーは継ぎ足し修復せず、直前の成功baselineから作り直します。

## 5. 完了条件

- 原本SHA-256が不変
- 自動起動の有無、経路、無効化方法、検証証跡を記録
- baselineと候補の由来・SHA-256が明確
- compileとreopen compileが成功
- manifestと全資産diffが予定どおり
- Access PIDと`.laccdb`/`.ldb`が残っていない
- 必要な自己テスト・画面・帳票確認が成功
- 公開物に顧客情報、接続先、資格情報、実レコードがない

## 詳細

- [標準開発手順 詳細リファレンス](17_access-development-workflow-reference.md)
- [Access作業共通ルール](00_access-work-common-rules.md)
- [起動処理を止める](requirements/01_startup-bypass.md)
- [Access COM自動化](04_access-com-automation.md)
- [LoadFromTextトラブルシュート](05_loadfromtext-troubleshooting.md)
- [RunCommand(126)でコンパイル](06_compile-with-runcommand.md)
- [Access外部Exportツール](16_access-external-export.md)
