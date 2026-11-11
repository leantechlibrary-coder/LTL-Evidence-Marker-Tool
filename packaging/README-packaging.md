# MSIX パッケージング手順（証拠番号付与ツール）

`entry.py` を GUI アプリとして exe にフリーズし（PyInstaller `--windowed`）、MSIX に包む。
スタート/タイルから起動すると GUI（証拠番号付与ツール）が開く。

## 前提（ビルド環境に必要なもの）

| 必要なもの | 入手 |
|---|---|
| Python + PyInstaller | `pip install pyinstaller` |
| Windows SDK（makeappx/signtool） | Visual Studio または SDK 単体インストーラ |
| 署名証明書 | ローカル検証は自己署名 `.pfx`（Subject=マニフェストの Publisher と一致）。**ストア提出時は不要**（Store が再署名） |
| 画像アセット | `packaging/Assets/` に PNG（StoreLogo 50x50 / Square150x150 / Square44x44 / Wide310x150 等） |
| ストア識別子 | Partner Center で予約した Identity `Name` と Publisher ID → `AppxManifest.xml` の `[TODO]` を置換 |

## 手順

1. `AppxManifest.xml` の `[TODO]`（Identity Name / Publisher）を Partner Center の値に置換。
2. `packaging/Assets/` にロゴ PNG を用意。
3. `powershell -ExecutionPolicy Bypass -File packaging\build.ps1` を実行
   → `packaging/out/ltl-stamp.msix` が出来る。
4. ローカル検証：`Add-AppxPackage packaging\out\ltl-stamp.msix`（自己署名証明書を信頼済みにしておく）。
   - スタートから「証拠番号付与ツール」を起動して動作確認。
5. ストア提出：Partner Center にアップロード（Store が署名。`runFullTrust` は審査対象）。

## クラウド（GitHub Actions）でビルドする

Windows PC が無くても、`.github/workflows/windows-build.yml` が GitHub 上の Windows 環境で
上記の `build.ps1` を実行する。流れ：テスト（pytest。MCP のテストも含む）→ exe 化 → MSIX 化。
exe / MSIX を作るジョブには `mcp` を入れないので、成果物は GUI 専用になる。

- **自動で動くとき**：PR 作成時／`main` への push 時（コード・依存・packaging に変更があるとき）
- **手動で動かす**：GitHub の Actions タブ →「windows-build」→「Run workflow」
- **成果物の入手**：実行結果ページ下部の「Artifacts」→ `ltl-stamp-windows`（30日保存）
  - `ltl-stamp.msix`：ストア提出用（未署名。Store が署名する）
  - `ltl-stamp-win-x64.zip`：exe 一式（フォント同梱）。MSIX を使わず直接起動する場合
- **リリース**：`v1.0.0` のようなタグを push すると、成果物を添付したリリースの**下書き**ができる（公開は手動）
- **署名（任意）**：リポジトリの Settings → Secrets に `LTL_PFX_BASE64`（.pfx を Base64 化した文字列）と
  `LTL_PFX_PASSWORD` を登録すると、ローカル検証用に署名した MSIX を作る。未登録なら未署名。

## MCP サーバ（Store 版には含めない）

`mcp_server.py` はリポジトリに残すが、**Microsoft Store 版（exe / MSIX）には入れない**。
刻印ロジック（`stamp_core.py`）は GUI と共有し、入口だけを分けている。

- 起動：`pip install -r requirements-mcp.txt` → `python mcp_server.py`（stdio）
- クライアント設定例：
  ```json
  { "mcpServers": { "ltl-stamp": { "command": "python", "args": ["C:\\path\\to\\mcp_server.py"] } } }
  ```
- Store 版に入らないことの歯止め（3重）：
  1. `entry.py` は MCP を import しない（`tests/test_store_build_has_no_mcp.py` が検査）
  2. `build.ps1` が PyInstaller に `--exclude-module mcp --exclude-module mcp_server` を渡す
  3. `build.ps1` がビルド後、同梱モジュール一覧に MCP が無いことを確認（あればビルド失敗）
- マニフェストに実行エイリアス（`appExecutionAlias`）は付けない（同じテストが検査）

## rename アプリ（証拠ファイル名変換ツール）での差分

同じ雛形をコピーし、以下だけ差し替える：
- `App` / Executable：`ltl-stamp` → `ltl-rename`
- DisplayName/Description：証拠ファイル名変換ツール 用に
- Identity `Name`：rename 用の予約名
- `build.ps1` の `$App = "ltl-rename"`

## 刻印フォントの同梱

`build.ps1` がリポジトリ直下の `fonts/`（LTL Evidence Sans + OFL.txt）を stage の exe と同階層へコピーする。`stamp_core._find_font_file()` はフリーズ後 exe 隣の `fonts\` を最優先で探すため、`--add-data` は不要。フォントが見つからない場合、stamp_core は内蔵 "japan" へ自動フォールバックする。
