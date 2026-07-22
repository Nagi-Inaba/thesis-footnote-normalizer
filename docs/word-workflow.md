# Wordでの確認手順

## 監査前

1. 本文の構成修正を終えます。終わっていない場合は、初出・再出など順序に依存する結果を暫定扱いにします。
2. 論文を保存してWordを閉じます。
3. Gitの追跡対象外になっている入力用ディレクトリへ、別名のDOCXコピーを作ります。
4. 引用方針と参考文献台帳を完成させます。
5. Wordの参考文献一覧と照合する場合は、必要に応じて別の参考文献DOCXを用意し、`bibliography_document`のマーカーを設定します。

## 原稿を書き換えずに実行する

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-review.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-001
```

参考文献DOCXも指定する場合は、次のように実行します。

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-review.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -BibliographyDocx .\input\bibliography-review.docx `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-002
```

監査ツールはOOXMLを読み取り、レポートだけを出力します。Wordのフィールド、脚注参照、スタイル、ハイパーリンク、変更履歴、入力DOCX、参考文献DOCXは変更しません。

## 確認と修正

1. 分類結果を利用する前に、`issues.csv`の各項目を確認します。
2. `citation_classification`、`adjacent_same_source`、`ibid_rewrite_candidate`は、それぞれ別の観点として確認します。
3. 文脈依存の短縮形と、初出で使われた短縮形の警告をすべて確認します。
4. `bibliography-reconciliation.csv`がある場合は、参考文献台帳と照合します。
5. コピーした論文をWordで開き、変更履歴を有効にします。
6. 承認した修正だけを、脚注ごとに反映します。
7. 修正版を別名で保存して再監査し、指摘件数を比較します。

OOXMLの`w:type`で区切り線などの特殊脚注と示されている脚注は、引用行や問題行を作らず、引用分析の対象外にします。
