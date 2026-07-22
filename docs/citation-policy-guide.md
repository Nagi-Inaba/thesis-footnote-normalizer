# 引用方針ガイド

## 優先する基準

所属研究科の規程、指導教員からの書面による指示、基準とする引用方式、プロジェクト固有の例外の順に優先します。所属機関が別の優先順位を定めている場合は、その定めに従ってください。

この監査ツールは、決定済みの方針を記録して確認に使います。引用方式そのものを自動で決定するものではありません。

## 最小方針とバージョン0.2の方針

後方互換性のため、従来の最小方針も引き続き利用できます。

```json
{
  "policy_name": "Thesis citation policy",
  "subsequent_citation": "author-short-title-page",
  "consecutive_same_source": "short-form",
  "new_sources": "prohibited_without_separate_authorization"
}
```

ただし、この互換形式では、文献種別ごとの必須項目を定義できません。完全な監査として扱う前に、`source_type_policies`を設定してください。

文献種別ごとの規則では、台帳に必要な項目を、初出と再出に分けて指定します。承認された引用方式で初出と再出の要件が異なる場合に使います。指定するのは説明文ではなく、参考文献台帳の列名です。

```json
{
  "policy_name": "Thesis citation policy",
  "subsequent_citation": "author-short-title-page",
  "consecutive_same_source": "short-form",
  "new_sources": "prohibited_without_separate_authorization",
  "source_type_policies": {
    "book": {
      "first_use_required_fields": ["author", "title", "publisher", "year"],
      "subsequent_use_required_fields": ["author", "short_title"],
      "consecutive_same_source": "ibid"
    }
  }
}
```

項目名と文献種別は、サンプル方針と参考文献台帳の見出しに合わせてください。脚注から欠けている出版情報を推測して補ってはいけません。

## 本監査の前に決める事項

- 各`source_type`について、初出と再出で必要な項目
- `ibid.`、`op. cit.`、「同上」、「前掲」を使用できるか
- 頁範囲と末尾記号の規則
- 翻訳書の記載規則
- 専門分野固有の文献に関する例外
- 脚注で引用したすべての文献を参考文献一覧に載せるか
- 本文・脚注で引用していない背景文献を参考文献一覧に残せるか

## 文脈依存の短縮形

`ibid.`、`op. cit.`、「同上」、「前掲」は、前後の文脈に依存します。監査では、これらを常に`review_required`の確認候補として扱います。短縮形が存在しても、文献同一性や書き換えの正しさは証明されません。

適用する規則が文脈依存の短縮形を求めていない場合は、著者名・短縮書名・頁数などを含む明示的な短縮形を推奨します。文献の初出で短縮形が検出された場合は、初出情報の充足を人が確認する必要があるため警告します。

## 参考文献の情報源

`bibliography.csv`を文献同定の台帳として使用します。生成される`citation-variants.csv`は、正規化した引用表記と比較根拠を記録しますが、表記の一致を文献同一性の証明には使いません。

`-BibliographyDocx`を使うと、参考文献一覧を含むDOCXを追加で指定できます。監査ツールは、方針の`bibliography_document`で指定したマーカーの範囲だけからテキストを抽出し、確認用の`bibliography-reconciliation.csv`を作成します。抽出結果で台帳を自動的に置き換えることはありません。

`bibliography_document`には、`enabled`、`start_marker`、`end_marker`、`include_heading`、`paragraph_match_mode`を設定します。この機能を有効にして`-BibliographyDocx`を省略した場合は、監査対象の論文DOCX自体を調べます。方針側の設定がない、または無効な状態で`-BibliographyDocx`を指定するとエラーになります。

翻訳書、史料、聖書、教会文書、教会法資料、歴史的版などは、一般書とは異なる`source_type_policies`が必要になる場合があります。一般書の規則へ無理に当てはめず、文献種別ごとの規則を明示してください。
