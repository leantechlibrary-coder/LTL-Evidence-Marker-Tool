# LTL-Evidence-Marker-Tool
note紹介記事：https://note.com/leantechlibrary/n/n4604bcf3e57c

ビルド版はMicrosoft storeで販売しています。https://apps.microsoft.com/detail/9pm38hpwfngj?hl=ja-JP&gl=JP

## 刻印フォント

証拠番号の刻印には同梱フォント **LTL Evidence Sans**（`fonts/LTL-Evidence-Sans.ttf`）を使用し、出力PDFにサブセット埋め込みします。閲覧環境のフォント有無に依存せず、番号の見た目が固定されます。

- Noto Sans CJK JP を証拠番号用145グリフにサブセット・TrueType化した派生フォント（SIL Open Font License 1.1）。詳細は `fonts/OFL.txt` を参照。
- フォントファイルが無い環境（MCP同梱パッケージ `ltl-evidence-mcp` 等）や、サブセット外のグリフを含むカスタム証拠種別では、従来どおり PyMuPDF 内蔵の日本語フォント（非埋め込み）へ自動フォールバックします。
