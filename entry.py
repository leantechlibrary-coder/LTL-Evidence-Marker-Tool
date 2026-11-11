"""
入口（パッケージング・エントリ）
================================

PyInstaller / MSIX のエントリポイント。GUI（証拠番号付与ツール）を起動する。

  ltl-stamp            → GUI（証拠番号付与ツール）
"""


def main():
    from pdf_evidence_marker import main as gui_main
    gui_main()


if __name__ == "__main__":
    main()
