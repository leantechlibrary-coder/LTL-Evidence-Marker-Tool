# MSIX パッケージング手順（証拠番号付与ツール）

GUI と MCP を1つの MSIX に同梱する。`entry.py` を exe にフリーズし、MSIX の
**App Execution Alias** で `ltl-stamp.exe` を登録する。これにより：

- スタート/タイルから起動 → 引数なし → **GUI**
- MCPクライアントが `ltl-stamp --mcp` を起動 → **MCPサーバ（stdio）**

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
   - GUI：スタートから「証拠番号付与ツール」を起動。
   - MCP：別プロセスから `ltl-stamp --mcp` が stdio で応答するか（下記クライアント設定で）。
5. ストア提出：Partner Center にアップロード（Store が署名。`runFullTrust` は審査対象）。

## クラウド（GitHub Actions）でビルドする

Windows PC が無くても、`.github/workflows/windows-build.yml` が GitHub 上の Windows 環境で
上記の `build.ps1` を実行する。流れ：テスト（pytest）→ exe 化 → **MCP 起動確認**
（`packaging/smoke_mcp.py`）→ MSIX 化。

- **自動で動くとき**：PR 作成時／`main` への push 時（コード・依存・packaging に変更があるとき）
- **手動で動かす**：GitHub の Actions タブ →「windows-build」→「Run workflow」
- **成果物の入手**：実行結果ページ下部の「Artifacts」→ `ltl-stamp-windows`（30日保存）
  - `ltl-stamp.msix`：ストア提出用（未署名。Store が署名する）
  - `ltl-stamp-win-x64.zip`：exe 一式（フォント同梱）。MSIX を使わず直接起動・MCP 設定する場合
- **リリース**：`v1.0.0` のようなタグを push すると、成果物を添付したリリースの**下書き**ができる（公開は手動）
- **署名（任意）**：リポジトリの Settings → Secrets に `LTL_PFX_BASE64`（.pfx を Base64 化した文字列）と
  `LTL_PFX_PASSWORD` を登録すると、ローカル検証用に署名した MSIX を作る。未登録なら未署名。

開発中の MCP 起動確認は `python packaging/smoke_mcp.py python entry.py` でもできる。

## ひとつ注意：コンソール subsystem と stdio

`--mcp` は stdin/stdout で会話するため、フリーズは **`--console`** で行う（GUI subsystem だと
stdin/stdout が None になり stdio が通らない場合がある）。一方コンソールだと **GUI 起動時に
コンソール窓が一瞬出る**。これを消すには、`entry.py` の GUI 分岐の先頭でコンソールを隠す：

```python
def _hide_console():
    import ctypes
    h = ctypes.windll.kernel32.GetConsoleWindow()
    if h:
        ctypes.windll.user32.ShowWindow(h, 0)  # SW_HIDE
# GUI 分岐で gui_main() の前に _hide_console() を呼ぶ
```

（この3行を入れるかは任意。希望があればこちらで `entry.py` に追加します。）

## MCPクライアント設定（インストール後）

```json
{ "mcpServers": { "ltl-stamp": { "command": "ltl-stamp", "args": ["--mcp"] } } }
```

## ストアより前に「今すぐ」MCPを使う（MSIX不要）

App Execution Alias は MSIX の機能だが、**エージェントを動かすだけなら MSIX は要らない**。
`entry.py` をフリーズした exe（または開発中は素の Python）をフルパスで指すだけでよい：

```json
{ "mcpServers": { "ltl-stamp": { "command": "C:\\path\\to\\ltl-stamp.exe", "args": ["--mcp"] } } }
```
（開発中は `{ "command": "python", "args": ["C:\\...\\entry.py", "--mcp"] }` でも可。）
ストア配布の段で初めて、安定した短い名前のためにエイリアスへ移行すればよい。

## rename アプリ（証拠ファイル名変換ツール）での差分

同じ雛形をコピーし、以下だけ差し替える：
- `App` / Executable / Alias：`ltl-stamp` → `ltl-rename`
- DisplayName/Description：証拠ファイル名変換ツール 用に
- Identity `Name`：rename 用の予約名
- `build.ps1` の `$App = "ltl-rename"`

## 刻印フォントの同梱

`build.ps1` がリポジトリ直下の `fonts/`（LTL Evidence Sans + OFL.txt）を stage の exe と同階層へコピーする。`stamp_core._find_font_file()` はフリーズ後 exe 隣の `fonts\` を最優先で探すため、`--add-data` は不要。**MCP同梱パッケージ（ltl-evidence-mcp）にはフォントを入れない**——stamp_core はフォント不在時に内蔵 "japan" へ自動フォールバックする設計なので、そのままで従来動作になる。
