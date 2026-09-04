# 空DBインポートによる救出解析

起動処理を安全にバイパスできず、フォーム、レポート、VBAの定義を読む必要がある場合だけ使う解析専用経路です。

この方法で作ったDBは正式なbaselineや差分基準にしません。インポートで定義や内部識別子が正規化されるため、元DBとの完全一致を期待できません。

## 前提

- 読み取り元は本体ではなく、ハッシュを記録した作業コピーにする。
- 空DBと出力先は、顧客情報や別案件を含まない隔離したstageへ置く。
- DAOで先にオブジェクト名を列挙し、取り込む対象を明記する。
- 読み取り元のAutoExec、起動フォーム、起動プロパティを空DBへ持ち込まない。
- リンクテーブルや実データは既定で取り込まない。
- 空DBを開くAccess COMでは、`OpenCurrentDatabase`より前に`AutomationSecurity=3`を設定する。この経路では取り込んだVBAを実行・コンパイルしない。

## 適用できない、または制限される入力

- 暗号化・パスワード保護された読み取り元は、`TransferDatabase`がモーダル入力を要求し得る。資格情報を非対話で安全に渡す経路を実機確認できていない場合は、この方式を無人実行しない。
- ACCDE/MDEはコンパイル済み成果物であり、標準/クラスモジュールの編集可能なVBAソースを復元できない。フォームやレポートを取り込めても、コードビハインドを完全な解析元とみなさない。
- `TransferDatabase`が`AutomationSecurity=3`の空DBセッションで対象形式・対象オブジェクトを取り込めることを、最初に機密を含まないカナリアで確認する。未確認の組合せは`要実機確認`とする。

## 標準手順

1. 読み取り元ACCDB/MDBと対応INIを作業コピーへ複製し、元とコピーのSHA-256を照合する。
2. DAO `OpenDatabase(copy, False, True)`でTableDefs、QueryDefs、Containersを列挙する。
3. 新しい空ACCDBを作る。空DBにはAutoExecマクロや起動フォームを設定しない。
4. `AutomationSecurity=3`を設定してから空DBを開き、`DoCmd.TransferDatabase(acImport, "Microsoft Access", sourceCopy, ...)`を呼ぶ。
5. フォーム、レポート、標準/クラスモジュール、必要な保存クエリを名前指定で取り込む。
6. マクロは一括インポートせず、`AutoExec`を除外して必要なものだけを名前指定する。
7. テーブル構造はまずDAO出力を使う。どうしても必要な場合だけ`StructureOnly=True`でローカルテーブルを取り込み、レコードを取り込まない。
8. 取り込んだオブジェクトを`SaveAsText`し、UTF-8派生を作る。
9. 空DBを閉じ、専用PIDとロックが消えたことを確認する。
10. 解析完了後は空DBを使い捨て、製品候補やbaselineへ昇格させない。

オブジェクト種別の定数は次のとおりです。

| 対象 | 定数 | 値 |
| --- | --- | --- |
| テーブル | `acTable` | `0` |
| クエリ | `acQuery` | `1` |
| フォーム | `acForm` | `2` |
| レポート | `acReport` | `3` |
| マクロ | `acMacro` | `4` |
| モジュール | `acModule` | `5` |

概念例:

```powershell
# The Access instance has an empty destination database open.
$access.DoCmd.TransferDatabase(0, 'Microsoft Access', $sourceCopy, 2, 'Form_Main', 'Form_Main')
$access.SaveAsText(2, 'Form_Main', $exportPath)
```

## AutoExecを除外する理由

読み取り元DBのAutoExecは、`TransferDatabase`の実行中には起動しません。しかし、AutoExecを空DBへ取り込むと、その解析DBを次に通常起動した時点で実行され得ます。

空DB方式を「最初の一度だけ安全」と扱わず、AutoExecを最後まで除外します。起動関数を呼ぶ別名マクロも、必要性が確認できるまで取り込みません。

## 禁止事項

- 本体DBをインポート元として直接指定する。
- 全マクロを機械的に取り込む。
- AutoExecを改名して取り込み、解析後に戻し忘れる。
- リンクテーブルを開く、保存クエリを実行する、実データを読む。
- インポート結果を正式差分のbeforeとして使う。
- タイムアウトした同じ解析DBを再利用する。

タイムアウト時は、記録した専用Access PIDだけを停止し、ロック消失を確認します。失敗した解析DBを保存する場合も、次の試行は新しい空DBから始めます。
