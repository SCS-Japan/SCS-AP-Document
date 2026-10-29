# SCS評価制度 公開資料（最新取得版）

IPA（独立行政法人情報処理推進機構）が公開する「サプライチェーン強化に向けたセキュリティ対策評価制度（SCS評価制度）」のページ本文および関連文書を保管しています。

このディレクトリには、本リポジトリで取得した最新の資料を、原本と Markdown 形式で格納しています。

> **本リポジトリは非公式のアーカイブです。**
> 取得のタイミングによっては、公式サイトの最新版と異なる場合があります。
> 制度の適用や評価・申請にあたっては、必ず [IPA公式サイト](https://www.ipa.go.jp/security/scs/index.html) の最新情報と原本をご確認ください。

## 資料を探す

| 確認したい内容 | 参照先 |
|---|---|
| 制度の全体像 | [SCS評価制度の概要](markdown/www.ipa.go.jp/security/scs/index.md) |
| 段階別評価などの詳細 | [SCS評価制度の詳細情報](markdown/www.ipa.go.jp/security/scs/details.md) |
| 要求事項・評価基準、評価用ガイド、解説書 | [要求事項・評価基準](markdown/www.ipa.go.jp/security/scs/requirements-criteria.md) |
| 制度の規程・委員会 | [制度規程・委員会](markdown/www.ipa.go.jp/security/scs/regulation-advisory-committeess.md) |
| 評価機関・技術検証事業者・研修事業者・専門家 | [関係機関・専門家の情報](markdown/www.ipa.go.jp/security/scs/security-experts-organization/index.md) |
| よくある質問 | [FAQ](markdown/www.ipa.go.jp/security/scs/faq.md) |
| 収録資料の一覧・取得状況・過去の版 | [監視対象のページと文書](../history/documents.md) |

## ディレクトリ構成

```text
documents/
├── README.md
├── originals/             # 公開文書の原本（PDF・Word・Excel など）
│   └── www.ipa.go.jp/
└── markdown/              # ページ本文・公開文書の Markdown 版
    └── www.ipa.go.jp/
```

取得元のドメイン名とパスに沿って配置しています。

### `originals/`：原本

公開元から取得したファイルを保管しています。

次のような場合は、原本を参照してください。

- 文書の正式な記載やレイアウトを確認する
- 図表・注記・ページ構成を確認する
- Excel のアセスメントシートや Word の様式を利用する

### `markdown/`：テキスト版

公開ページの本文と、公開文書から変換した Markdown を保管しています。
GitHub 上での閲覧、文字列検索、差分確認などに利用できます。

ファイル名の例：

| 取得元 | Markdown 版 |
|---|---|
| `index.html` | `index.md` |
| `SCS-410-self-assess-guide-v1.0.pdf` | `SCS-410-self-assess-guide-v1.0.pdf.md` |
| `SCS-400-criteria-v1.1.xlsx` | `SCS-400-criteria-v1.1.xlsx.md` |

> Markdown 版は閲覧・検索の補助用です。
> 変換により、表の構造、図、注記、数式、セルの書式などが正確に再現されない場合があります。
> 内容の正確な確認や様式の利用には、原本を使用してください。

## 主な資料

★3・★4の要求事項と評価に関する主要資料です。
すべての収録資料は [文書一覧](../history/documents.md) を参照してください。

| 文書番号 | 文書名 | 原本 | Markdown |
|---|---|---|---|
| SCS-400 | ★3・★4 要求事項・評価基準 | [Excel](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-400-criteria-v1.1.xlsx) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-400-criteria-v1.1.xlsx.md) |
| SCS-410 | ★3・★4 自己評価等ガイド | [PDF](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-410-self-assess-guide-v1.0.pdf) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-410-self-assess-guide-v1.0.pdf.md) |
| SCS-410AP | ★3・★4 自己評価等ガイド 別紙 アセスメントシート | [Excel](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-410AP-self-assess-sheet-v1.0.xlsx) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-410AP-self-assess-sheet-v1.0.xlsx.md) |
| SCS-420 | ★4 第三者評価ガイド | [PDF](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-420-3rdparty-assess-guide-v1.0.pdf) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-420-3rdparty-assess-guide-v1.0.pdf.md) |
| SCS-430 | ★4 技術検証ガイド | [PDF](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-430-vuln-assess-guide-v1.0.pdf) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-430-vuln-assess-guide-v1.0.pdf.md) |
| SCS-620 | ★3・★4 要求事項・評価基準 解説書 | [PDF](originals/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-620-handbook-v1.0.pdf) | [読む](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-620-handbook-v1.0.pdf.md) |

各資料の位置づけは、[要求事項・評価基準、ガイド、解説書等の位置づけ](markdown/www.ipa.go.jp/security/scs/rcu1hd0000007a2i-att/SCS-documents-overview.pdf.md) を参照してください。

## 更新履歴と過去の版

このディレクトリには最新取得版を置き、過去の版や更新履歴は別のディレクトリで管理しています。

- [更新履歴](../history/CHANGELOG.md)：検知した変更と差分へのリンク
- [監視対象のページと文書](../history/documents.md)：取得状況、版数、過去の版への案内
- [過去の版](../archive/)：更新前の原本と Markdown
- [リポジトリの概要](../README.md)：アーカイブ全体の説明

## 利用上の注意

- 本リポジトリでの取得・変更検知の日付は、公開元の公開日・更新日と一致するとは限りません。
- Markdown 版と原本に相違がある場合は、原本を確認してください。
- 本リポジトリへの収録は、各資料の権利や利用条件を変更するものではありません。転載・再配布等にあたっては、公開元の利用条件および各資料の記載を確認してください。
- 本リポジトリは、IPAによる公式の情報提供や、制度上の判断を代替するものではありません。

## 出典

独立行政法人情報処理推進機構（IPA）  
[サプライチェーン強化に向けたセキュリティ対策評価制度（SCS評価制度）](https://www.ipa.go.jp/security/scs/index.html)
