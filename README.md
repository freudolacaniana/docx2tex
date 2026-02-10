# docx2tex

Microsoft Word（.docx）ファイルを、日本語組版に最適化された LaTeX（.tex）形式に変換するオンラインツールです。

## 🚀 これは何？

Microsoft Word 等で作成された原稿を、LaTeX で処理可能なソースファイルに変換します。

- **一括ダウンロード**: 変換後のファイルは、原稿内の画像ファイル等を含めた `.zip` 形式で取得できます。
- **日本語組版への最適化**: `lualatex` + `jlreq` 環境を前提としており、[日本語組版処理の要件](https://www.w3.org/TR/jlreq/?lang=ja)に即した高品質なPDF出力が可能です。
- **CloudLaTeX 対応**: 出力される `.zip` ファイルは、そのまま [CloudLaTeX](https://cloudlatex.io/) へインポートして使用できます。

### 💡 本ツールの特徴

Pandoc をはじめとする既存の変換ツールは多数ありますが、本ツールは特に **「日本語・縦書きの人文科学系論文、書籍、小説」** のために開発されています。

- **ルビ** や **縦中横** など、日本語文書特有の書式保持に強みがあります。
- ※ 数式や `.emf` ファイルの変換には限定的にしか対応していません。理系分野の用途には [Pandoc](https://pandoc.org/) 等の利用を推奨します。

---

## ⚠️ 注意事項

- **プライバシー・セキュリティ**: 本ツールはクライアントサイドの JavaScript のみで動作します。ファイルが外部サーバーにアップロードされることはなく、すべての処理はブラウザ内で完結します。
- **PDF出力について**: 本ツール自体は PDF を出力しません。PDF を得るには、変換後のファイルをローカルの [TeX Live](https://texwiki.texjp.org/?TeX%20Live) 環境や CloudLaTeX 等でコンパイル（タイプセット）する必要があります。
- **推奨環境**: `lualatex` + `jlreq` でのコンパイルを想定しています。その他の環境で使用する場合は、適宜 `.tex` ファイルのプリアンブルを編集してください。

---

## 📝 .docx ファイルの準備

変換精度を高めるため、元の `.docx` ファイルで以下の設定を行っておくことを推奨します。

- **スタイル機能の活用**:
- タイトルには「表題」スタイル
- サブタイトルには「副題」スタイル
- 各見出しには「見出し1」「見出し2」などのスタイルを適用してください。

- **著者名**: タイトル・サブタイトルの直下にある短い行は、自動的に著者名として認識されます。
- **見出しレベル**: 各スタイルを LaTeX のどのコマンド（`\section`, `\subsection` 等）に対応させるかは、ツール上の設定で調整可能です。

---

### 💡 Licenses

This project utilizes the following libraries under the MIT License:

- **UDOC.js**: Used for reading `.emf` files embedded within `.docx` documents.
- **JSZip**: Used for parsing `.docx` files and generating `.zip` output.
- **officemath2latex**: Used for parsing oMath formulas and generating their LaTex equivalents.


