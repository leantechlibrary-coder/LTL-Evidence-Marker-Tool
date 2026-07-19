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
