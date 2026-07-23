# Thesis Footnote Normalizer

博士論文・修士論文のWord脚注を、原本へ触れずに棚卸しする監査キットです。

このツールが決めるのは引用方式ではありません。

研究科の執筆要領、指導教員の判断、採用する基本方式から作成した規則を使い、同一文献の初出・再出、直前再出候補、未同定脚注、複数文献脚注、未使用参考文献を一貫して確認します。

## 現在の範囲

バージョン0.2は監査専用です。入力DOCXと任意の参考文献DOCXを変更しません。

- `.docx`に保存された真のWord脚注を読み取ります。
- 入力DOCXを変更しません。
- 登録済みの文献別名と脚注本文を照合します。
- 初出または再出、同一文献の隣接、文脈依存略記の書換候補を別々に分類します。
- `w:type`を持つ区切りなどの特殊脚注を引用分析から除外します。
- 人間が確認するCSV、JSON、Markdownを生成します。
- Word、Zotero、LibreOffice、外部通信を必要としません。

次のことは行いません。

- 引用内容や頁数が正しいと保証する
- 文献情報を外部検索して補う
- 不足情報を推測して埋める
- 未使用参考文献を削除する
- DOCXの脚注を書き換える
- `.doc`、`.docm`、PDF、Wordの文末脚注を処理する

## 安全モデル

入力原本は読み取り専用として扱います。

監査の前後でSHA-256を計算し、一致しなければ処理を失敗として扱います。

実論文、監査結果、個人情報はGitへ追加しないでください。

このリポジトリの`.gitignore`は`input/`、`work/`、`output/`を除外しますが、公開前には`git status`と差分を人間が確認してください。

## 必要な環境

- Windows 10または11
- PowerShell 7以上を推奨
- 入力ファイルは`.docx`

Microsoft Wordは、監査後に変更履歴付きで修正する段階だけで使用します。

## 5分で試す

### 1. リポジトリを取得する

```powershell
$repoUrl = Read-Host 'GitHubの「Code」からコピーしたHTTPS URL'
git clone $repoUrl
Set-Location .\thesis-footnote-normalizer
```

URLは、このリポジトリを公開したGitHub画面の「Code」から取得します。

### 2. 論文のコピーを用意する

```powershell
New-Item -ItemType Directory -Path .\input -Force
$manuscript = Read-Host 'コピー元DOCXのパス'
Copy-Item -LiteralPath $manuscript -Destination .\input\thesis-footnote-review-01.docx
```

元の論文を直接指定せず、必ず別名コピーを使用します。

### AIアダプターを使う場合

CodexとClaude Codeのスキル・エージェントは任意です。

アダプターは、このclone内の監査スクリプトを呼び出します。インストール後もリポジトリを移動・削除せず、監査時はリポジトリのルートから実行してください。

内容を確認してから、対象ランタイムだけをインストールします。

```powershell
.\install.ps1 -Runtime codex -WhatIf
.\install.ps1 -Runtime codex
```

`-Runtime`には`codex`、`claude`、`both`を指定できます。

既存の同名スキルまたはエージェントがある場合、既定では停止します。

`-Force`は既存内容を確認した場合だけ使用してください。既存アダプターへマージせず、同名のスキル・エージェントを完全に置き換えるため、旧版だけに存在するファイルは残りません。

インストール先は既定で`USERPROFILE`配下に限定され、途中にジャンクションやシンボリックリンクがある場合は停止します。

Codexは`CODEX_HOME`、Claude Codeは公式の`CLAUDE_CONFIG_DIR`が設定されていれば、そのディレクトリを使用します。明示的に指定する場合は`-CodexConfigDir`または`-ClaudeConfigDir`を使用できます。

意図的に別ドライブへ配置するときだけ`-AllowExternalHome`を追加してください。

### 3. 設定ファイルをコピーする

```powershell
Copy-Item .\config\citation-policy.example.json .\input\citation-policy.json
Copy-Item .\config\bibliography.example.csv .\input\bibliography.csv
```

`citation-policy.json`と`bibliography.csv`を論文用に編集します。

### 4. 監査を実行する

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-footnote-review-01.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-001
```

別のWord参考文献一覧も照合する場合は、方針の`bibliography_document`に抽出対象のマーカーを設定してから追加します。

```powershell
.\scripts\Invoke-FootnoteAudit.ps1 `
  -InputDocx .\input\thesis-footnote-review-01.docx `
  -BibliographyCsv .\input\bibliography.csv `
  -BibliographyDocx .\input\bibliography-review.docx `
  -PolicyJson .\input\citation-policy.json `
  -OutputDirectory .\work\audit-002
```

出力先がすでに存在する場合は処理を停止します。

既存結果を意図的に置き換える場合だけ`-Force`を指定してください。

監査対象の主要XML部品には64 MiBの上限、圧縮率上限、DTD禁止を適用します。

### 5. 要確認項目から読む

次の順で確認します。

1. `issues.csv`
2. `report.md`
3. `citations.csv`
4. `footnotes.csv`
5. `summary.json`

## 設定ファイル

### 引用方針

従来の最小方針は後方互換のため引き続き受理されます。

```json
{
  "policy_name": "博士論文脚注方針",
  "subsequent_citation": "author-short-title-page",
  "consecutive_same_source": "short-form",
  "new_sources": "prohibited_without_separate_authorization"
}
```

この値は監査の前提としてレポートへ記録されます。

最小方針だけでは文献種別ごとの完全性を検査できません。

利用者は`source_type_policies`を設定し、`default`規則と各`source_type`について`first_use_required_fields`、`subsequent_use_required_fields`、`consecutive_same_source`を指定してください。

バージョン0.2も指定された書式へ脚注を整形しません。

方針の決め方は[`docs/citation-policy-guide.md`](docs/citation-policy-guide.md)を参照してください。

### 文献台帳

`bibliography.csv`は一行を一つの引用実体として扱います。

| 列 | 内容 |
|---|---|
| `source_id` | 文献を一意に識別する固定ID |
| `source_type` | `book`、`article`、`translated_book`など |
| `language` | `ja`、`en`など |
| `author` | レビュー用の著者表示 |
| `short_title` | 再出表記で使う短縮書名候補 |
| `aliases` | 脚注本文で探す別名を`|`で区切ったもの |
| `bibliography_entry` | 承認済みまたは確認中の参考文献表記 |

方針が要求する構造化フィールド（例：`title`、`publisher`、`year`）も文献種別に応じて列として保持します。

`aliases`には、著者名だけのような広すぎる文字列を避け、書名、固有の略称、原綴などを登録します。

照合前にUnicodeを正規化し、文字・数字・結合記号の境界を確認するため、`Art`が`Artificial`へ一致するような単純な部分一致を抑制します。

脚注が複数文献を含む場合は複数の`source_id`が一致します。

これは異常とは限りませんが、必ず人間の確認対象になります。

## 分類の意味

### `citation_classification`

確認済み`source_id`の文書順に基づく`first`または`repeat`です。

初出として必要な書誌情報を備えているかは、人間が確認します。

### `adjacent_same_source`

現在と直前の引用対象脚注が同一の単独`source_id`へ一致したことを示す真偽値です。

これは配列上の根拠であり、略記の許可ではありません。

### `ibid_rewrite_candidate`

隣接関係と採用方針から算出する真偽値です。

真でも書換指示ではなく、必ず人間が確認します。

`ibid.`、`op. cit.`、「同上」、「前掲」などの文脈依存略記は常に`review_required`です。
これらは文献同一性や書誌事項の正しさを証明しません。

初出で短縮形らしい表記が見つかった場合も警告し、`source_type_policies`の初出必須フィールドと照合するよう求めます。

## 出力ファイル

### `footnotes.csv`

文書中の脚注参照順、OOXML上のID、抽出テキスト、一致した文献数を記録します。

参照順は、Word画面に表示される脚注番号を再現した値ではありません。

Excel等で数式として評価され得る文字列は、CSV上で先頭にアポストロフィを付けます。

### `citations.csv`

一致した文献ごとに一行を作り、`citation_classification`、`adjacent_same_source`、`ibid_rewrite_candidate`を記録します。

### `issues.csv`

次の項目を人間の確認対象として記録します。

- 文献台帳と一致しなかった脚注
- 複数文献に一致した脚注
- 脚注で使われなかった文献台帳項目
- 解析上の不足または曖昧さ
- 文脈依存略記と初出短縮形
- 方針または文献種別の必須フィールド不足
- 参考文献DOCXとの不一致

スキーマv2の`issue_type`は、`unmatched_footnote`、`multiple_source_matches`、`missing_footnote_body`、`unused_bibliography_entry`、`short_form_on_first_use`、`contextual_shorthand_candidate`、`unresolved_shorthand`、`source_type_policy_missing`、`bibliography_metadata_missing`、`citation_required_component_missing`、`citation_variant`、`terminal_mark_mismatch`、`document_bibliography_unmatched`、`document_bibliography_multiple_matches`、`document_bibliography_not_found`、`registry_missing_from_document_bibliography`です。

`w:type`付き特殊脚注は引用行とissueを作らず、分析対象から除外します。

未使用文献は削除指示ではありません。

背景文献として意図的に残している可能性があります。

### `summary.json`

スキーマv2として、件数、入力ハッシュ、引用方針、処理結果を機械可読形式で保存します。

### `citation-variants.csv`

明示的な引用について、分類、構成要素、正規化表記、句読点署名、比較対象脚注、比較状態を記録します。
これは監査出力であり、文献同一性や正しい表記の証明ではありません。

### `bibliography-reconciliation.csv`

方針の`bibliography_document`を有効にすると、`-BibliographyDocx`で指定した文書、または指定を省略した場合は入力DOCXから、マーカー範囲の参考文献本文を抽出して構造化台帳との照合結果を保存します。
抽出結果は台帳を自動更新しません。

### `report.md`

件数と監査上の制約を人間向けにまとめます。

## 修正作業

監査後の修正は、[`docs/word-workflow.md`](docs/word-workflow.md)に従います。

1. `issues.csv`の文献同定を解決します。
2. `citations.csv`を文書順に確認します。
3. Wordで変更履歴を有効にします。
4. 承認済みの変更だけを脚注へ反映します。
5. 修正版コピーに対して新しい監査を実行します。

本文を移動、追加、削除した場合は、初出順が変わるため監査からやり直します。

## AI支援

AIは必須ではありません。

AIへ渡す前に、著者の機密保持条件と利用サービスのデータ取扱条件を確認してください。

Codex用スキルは`codex/skills/normalize-thesis-footnotes/`、Claude Code用スキルは`claude/skills/normalize-thesis-footnotes/`にあります。

AIには、原則として全文ではなく、監査結果、必要な脚注行、引用方針、文献台帳の必要部分だけを渡します。

AIが担当できるのは次の判断補助です。

- 表記揺れが同一文献かを検討する
- 翻訳書と原著情報の関係を整理する
- 初出と再出の必要情報を規則と照合する
- 複数文献脚注を分解して確認事項を示す

AIは不足情報を創作せず、判断できない項目を`review_required`として残します。

## リポジトリ構成

```text
thesis-footnote-normalizer/
├─ README.md
├─ AGENTS.md
├─ CLAUDE.md
├─ SECURITY.md
├─ LICENSE
├─ config/
│  ├─ citation-policy.example.json
│  └─ bibliography.example.csv
├─ docs/
│  ├─ citation-policy-guide.md
│  ├─ word-workflow.md
│  └─ troubleshooting.md
├─ codex/
│  ├─ skills/normalize-thesis-footnotes/
│  └─ agents/footnote-normalization-reviewer.toml
├─ claude/
│  ├─ skills/normalize-thesis-footnotes/
│  └─ agents/footnote-normalization-reviewer.md
├─ scripts/
│  ├─ Invoke-FootnoteAudit.ps1
│  └─ Test-Repository.ps1
├─ templates/
├─ tests/
└─ .github/workflows/validate.yml
```

## テスト

```powershell
.\scripts\Test-Repository.ps1
```

テストは架空の人物名、書名、出版社だけを使ったDOCXを一時生成します。

次を検証します。

- 脚注8件を順番どおりに抽出できる
- 初出、直前再出候補、非連続再出を分類できる
- 未同定脚注と未使用文献を報告できる
- 入力DOCXのSHA-256が変化しない
- 再実行で意味的に同じ結果を得られる
- アダプターの全宛先を変更前に検査し、`-Force`で旧ファイルを残さず置換できる
- Claude Codeの`CLAUDE_CONFIG_DIR`へインストールできる

GitHub Actionsは`windows-2025`上のPowerShellで同じテストを実行します。

## トラブル対応

[`docs/troubleshooting.md`](docs/troubleshooting.md)を参照してください。

## 公開前チェック

- `git status`で`input/`、`work/`、`output/`が含まれていないことを確認する
- 絶対ローカルパスが残っていないことを確認する
- 実在論文、実在する引用、個人名、メールアドレス、APIキーがないことを確認する
- `scripts/Test-Repository.ps1`が成功することを確認する
- ライセンス方針を確認する

## ライセンス

MIT Licenseです。
