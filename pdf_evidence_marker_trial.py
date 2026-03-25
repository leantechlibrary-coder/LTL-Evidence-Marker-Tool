import sys
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QFileDialog, QMessageBox, QLabel,
    QLineEdit, QRadioButton, QButtonGroup, QSpinBox, QComboBox,
    QGroupBox, QListWidgetItem, QAbstractItemView, QDialog, QTextEdit
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QDrag, QFont
from typing import List, Tuple


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# トライアル管理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TRIAL_DAYS = 7
# 有料版のMicrosoft StoreページURL
STORE_URL = "https://apps.microsoft.com/detail/9pm38hpwfngj?hl=ja-JP&gl=JP"


class TrialManager:
    """無料トライアルの期限を管理する。
    
    初回起動日を %APPDATA%/LeanTechLibrary/trial.json に記録し、
    経過日数を返す。アンインストールしても記録は残る。
    """

    def __init__(self):
        appdata = os.environ.get("APPDATA", str(Path.home()))
        self._dir = Path(appdata) / "LeanTechLibrary"
        self._file = self._dir / "trial_evidence_marker.json"

    def _read(self) -> dict:
        try:
            return json.loads(self._file.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write(self, data: dict):
        self._dir.mkdir(parents=True, exist_ok=True)
        self._file.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def first_launch_date(self) -> datetime:
        """初回起動日を返す。未記録なら今日を記録して返す。"""
        data = self._read()
        if "first_launch" in data:
            return datetime.fromisoformat(data["first_launch"])
        now = datetime.now()
        data["first_launch"] = now.isoformat()
        self._write(data)
        return now

    def days_remaining(self) -> int:
        """トライアル残り日数を返す（0以下 = 期限切れ）。"""
        first = self.first_launch_date()
        elapsed = (datetime.now() - first).days
        return TRIAL_DAYS - elapsed

    def is_expired(self) -> bool:
        return self.days_remaining() <= 0


class TrialDialog(QDialog):
    """起動時に表示するトライアル情報ダイアログ。"""

    def __init__(self, days_remaining: int, expired: bool, parent=None):
        super().__init__(parent)
        self.expired = expired
        self.setWindowTitle("無料トライアル版 — 証拠番号付与ツール")
        self.setMinimumWidth(460)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(14)

        # --- タイトル ---
        title = QLabel("証拠番号付与ツール — 無料トライアル版")
        title.setFont(QFont("Yu Gothic UI", 13, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # --- 残り日数 / 期限切れメッセージ ---
        if expired:
            msg = QLabel(
                "無料トライアル期間が終了しました。\n"
                "ご利用いただきありがとうございました。\n\n"
                "引き続きご利用いただくには、\n"
                "Microsoft Store で有料版をご購入ください。"
            )
            msg.setStyleSheet(
                "color: #C62828; font-size: 11pt; padding: 8px;"
            )
        else:
            msg = QLabel(
                f"無料トライアル終了まで あと {days_remaining} 日 です。\n\n"
                "トライアル期間終了後もご利用いただくには、\n"
                "Microsoft Store で有料版をご購入ください。"
            )
            msg.setStyleSheet("font-size: 11pt; padding: 8px;")

        msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        msg.setWordWrap(True)
        layout.addWidget(msg)

        # --- Storeリンクボタン ---
        store_btn = QPushButton("Microsoft Store で有料版を見る")
        store_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #106EBE; }
        """)
        store_btn.clicked.connect(self._open_store)
        layout.addWidget(store_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- 閉じる / 続けるボタン ---
        if expired:
            close_btn = QPushButton("閉じる")
            close_btn.clicked.connect(self.reject)
        else:
            close_btn = QPushButton("閉じて続ける")
            close_btn.clicked.connect(self.accept)

        close_btn.setFixedWidth(160)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- フッター ---
        footer = QLabel("開発・販売：Lean Tech Library")
        footer.setStyleSheet("color: #888; font-size: 9pt;")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(footer)

    def _open_store(self):
        """Microsoft Store の有料版ページを開く。"""
        import webbrowser
        webbrowser.open(STORE_URL)


A3_WIDTH_PT = 1190.55  # A3の長辺（pt、420mm相当）
A3_HEIGHT_PT = 841.89  # A3の短辺（pt、297mm相当）
A3_TOLERANCE = 10.0    # 許容誤差（pt）


def is_page_a3(page) -> bool:
    """ページがA3サイズかどうかを判定する（回転属性を考慮）。
    
    A3横向きは842×595pt（長辺×短辺）または595×842ptで
    rotation属性が90/270のケースを含む。
    """
    rotation = page.rotation % 360
    rect = page.rect

    if rotation in (90, 270):
        effective_width = rect.height
        effective_height = rect.width
    else:
        effective_width = rect.width
        effective_height = rect.height

    long_side = max(effective_width, effective_height)
    short_side = min(effective_width, effective_height)

    return (
        abs(long_side - A3_WIDTH_PT) <= A3_TOLERANCE and
        abs(short_side - A3_HEIGHT_PT) <= A3_TOLERANCE
    )


def is_page_landscape(page) -> bool:
    """ページが横向きかどうかを判定する（回転属性を考慮）。
    
    PDFには2種類の横向き表現がある：
    1. rect自体が横長（width > height）
    2. rectは縦長だが、rotation属性（90/270度）で横向きに表示
    
    実際の見た目（表示上）が横向きかどうかを返す。
    """
    rotation = page.rotation % 360
    rect = page.rect
    
    if rotation in (90, 270):
        # 回転属性で90/270度回転 → 見た目上はwidth/heightが入れ替わる
        effective_width = rect.height
        effective_height = rect.width
    else:
        effective_width = rect.width
        effective_height = rect.height
    
    return effective_width > effective_height


class TextViewerDialog(QDialog):
    """テキスト全文表示用の子ダイアログ"""

    def __init__(self, parent, title: str, content: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(620, 500)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(content)
        text_edit.setFont(QFont("Yu Gothic UI", 9))
        text_edit.moveCursor(text_edit.textCursor().MoveOperation.Start)
        layout.addWidget(text_edit)

        close_btn = QPushButton("閉じる")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)


class AboutDialog(QDialog):
    """カスタムAboutダイアログ（操作説明書・README・ライセンス情報へのリンク付き）"""

    # --- 埋め込みテキスト定数 ---
    # MSIX / Microsoft Store 配布前提で改訂済み

    README_TEXT = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "PDF証拠整理ツール\n"
        "README\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "この度はPDF証拠整理ツールをご利用いただき、\n"
        "誠にありがとうございます。\n\n"
        "本ツールは、訴訟・紛争案件における証拠整理業務を\n"
        "効率化するために開発された専用ツールです。\n\n\n"
        "■ 収録ツール\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "・PDF分割ツール\n"
        "  PDFファイルを複数の分割ポイントで一度に分割\n\n"
        "・証拠番号付与ツール\n"
        "  PDFファイルに証拠番号（甲第○号証等）を自動付与\n\n\n"
        "■ 動作環境\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "OS：Windows 10 / 11（64bit）\n"
        "メモリ：8GB以上推奨\n"
        "ストレージ：500MB以上の空き容量\n\n\n"
        "■ 起動方法\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Microsoft Storeからインストール後、\n"
        "スタートメニューから起動してください。\n\n\n"
        "■ クイックスタート\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "＜PDF分割ツール＞\n"
        "1. 「PDFを開く」で対象ファイルを選択\n"
        "2. 分割したい先頭ページをクリック（複数選択可）\n"
        "3. 「分割実行」をクリック\n\n"
        "＜証拠番号付与ツール＞\n"
        "1. 「フォルダを開く」でPDFファイルを読み込み\n"
        "2. ファイルの順番を調整（ドラッグ＆ドロップ）\n"
        "3. 証拠種別（甲/乙等）とフォント設定\n"
        "4. 「証拠番号を付与して保存」をクリック\n\n\n"
        "■ よくある質問\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Q. 元のPDFファイルが変更されることはありますか？\n"
        "A. ありません。常に新しいファイルとして保存されます。\n\n"
        "Q. Googleドライブに保存できますか？\n"
        "A. Googleドライブデスクトップの同期フォルダを\n"
        "   出力先に指定することで可能です。\n\n"
        "Q. 何号証まで対応していますか？\n"
        "A. システム上は9999号証まで対応しています。\n\n\n"
        "■ ご注意事項\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "・本ツールは現状有姿での提供となります\n"
        "・パスワード保護されたPDFには対応していません\n"
        "・重要なファイルは必ずバックアップを取ってからご使用ください\n\n\n"
        "■ 免責事項\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "本ソフトウェアの使用により生じたいかなる損害についても、\n"
        "開発者は一切の責任を負いかねます。\n\n\n"
        "■ 著作権とライセンス\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "開発・販売：Lean Tech Library\n\n"
        "本ソフトウェアはAGPL-3.0ライセンスの下で配布されています。\n"
        "再配布の際はライセンス条件に従ってください。\n\n"
        "ソースコード：\n"
        "https://github.com/leantechlibrary-coder/LTL-Evidence-Marker-Tool\n"
    )

    MANUAL_TEXT = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "PDF証拠整理ツール 操作説明書\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "■ 目次\n"
        "  1. PDF分割ツールの使い方\n"
        "  2. 証拠番号付与ツールの使い方\n"
        "  3. よくある質問（FAQ）\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "1. PDF分割ツールの使い方\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "【基本操作】\n\n"
        "(1) PDFファイルを開く\n"
        "  ・「PDFを開く」ボタンをクリック\n"
        "  ・または、PDFファイルをウィンドウにドラッグ＆ドロップ\n\n"
        "(2) 分割ポイントを選択\n"
        "  ・サムネイル一覧が表示されます\n"
        "  ・分割したい先頭ページをクリック（青い枠が表示されます）\n"
        "  ・複数選択可能です（再クリックで解除）\n\n"
        "(3) 分割実行\n"
        "  ・「分割実行」ボタンをクリック\n"
        "  ・確認ダイアログが表示されるので「Yes」を選択\n"
        "  ・分割完了後、出力フォルダが自動的に開きます\n\n"
        "【サムネイル表示の調整】\n"
        "  ・画面右上のスライダーでサムネイルサイズと列数を調整できます\n"
        "  ・サイズ：100px～500px\n"
        "  ・列数：2列～6列\n\n"
        "【出力について】\n"
        "  ・出力先：元のPDFファイルと同じフォルダ内に\n"
        "    「元ファイル名_分割」フォルダを自動作成\n"
        "  ・ファイル名：元ファイル名_001.pdf、元ファイル名_002.pdf...\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "2. 証拠番号付与ツールの使い方\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "【基本操作】\n\n"
        "(1) PDFファイルを読み込む\n"
        "  ・「フォルダを開く」：フォルダ内の全PDFファイルを読み込み\n"
        "  ・「ファイルを追加」：個別にファイルを選択して追加\n"
        "  ・ドラッグ＆ドロップ：ファイル/フォルダをウィンドウに直接ドロップ\n\n"
        "(2) ファイルの順番を調整\n"
        "  ・ファイルリストをドラッグ＆ドロップで並び替え\n"
        "  ・または「↑上へ」「↓下へ」ボタンで移動\n\n"
        "(3) 枝番の設定（必要な場合のみ）\n"
        "  ・枝番にしたいファイルを選択\n"
        "  ・「枝番にする」ボタンをクリック\n"
        "  ・例：第2号証の後に枝番を設定すると「甲02の1」「甲02の2」\n"
        "  ・解除する場合は「枝番を解除」ボタン\n\n"
        "(4) 証拠番号の設定\n"
        "  ・証拠種別：甲/乙/その他（カスタム文字列）\n"
        "  ・開始番号：通常は1から\n"
        "  ・証拠番号を印字する：チェックONで1ページ目右上に番号を印字\n"
        "  ・フォントサイズ：8pt～72pt（デフォルト16pt）\n"
        "  ・フォント色：赤/黒/青（デフォルト赤）\n\n"
        "(5) 実行\n"
        "  ・「証拠番号を付与して保存」ボタンをクリック\n"
        "  ・確認ダイアログで内容を確認\n"
        "  ・完了後、出力フォルダが自動的に開きます\n\n"
        "【出力について】\n"
        "  ・出力先：読み込んだファイルの親フォルダ内に\n"
        "    「親フォルダ名_番号付」フォルダを自動作成\n"
        "  ・ファイル名：「甲01.pdf」「甲02.pdf」「甲03の1.pdf」など\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "3. よくある質問（FAQ）\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Q. Googleドライブに保存できますか？\n"
        "A. Googleドライブデスクトップの同期フォルダを\n"
        "   出力先に指定することで可能です。\n\n"
        "Q. 何号証まで対応していますか？\n"
        "A. システム上は9999号証まで対応しています。\n\n"
        "Q. 元のファイルが上書きされることはありますか？\n"
        "A. ありません。常に別フォルダに新規ファイルとして出力されます。\n\n"
        "Q. 証拠番号付与ツールの「ファイルを削除」ボタンを押すと\n"
        "   元のPDFが消えますか？\n"
        "A. 消えません。リスト上から取り除くだけです。\n\n"
        "Q. PDFにパスワードがかかっている場合は？\n"
        "A. パスワード保護されたPDFには対応していません。\n"
        "   事前にパスワードを解除してから処理してください。\n\n"
        "Q. 既存の証拠番号を上書きできますか？\n"
        "A. 既存の番号を自動削除する機能はありません。\n"
        "   新しく証拠番号を追記する形になります。\n"
    )

    LICENSE_TEXT = (
        "================================================================================\n"
        "THIRD-PARTY SOFTWARE LICENSES\n"
        "PDF証拠整理ツール\n"
        "================================================================================\n\n"
        "本ソフトウェアは、以下のオープンソースソフトウェアを使用しています。\n"
        "各ソフトウェアのライセンス条項に従い、ライセンス情報を記載します。\n\n\n"
        "================================================================================\n"
        "1. PyMuPDF (fitz)\n"
        "================================================================================\n\n"
        "License: GNU Affero General Public License v3.0 (AGPL-3.0)\n"
        "Copyright: Artifex Software, Inc.\n"
        "Website: https://github.com/pymupdf/PyMuPDF\n\n"
        "ライセンス全文：https://www.gnu.org/licenses/agpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "2. PyQt6\n"
        "================================================================================\n\n"
        "License: GNU General Public License v3.0 (GPL-3.0)\n"
        "Copyright: Riverbank Computing Limited\n"
        "Website: https://www.riverbankcomputing.com/software/pyqt/\n\n"
        "ライセンス全文：https://www.gnu.org/licenses/gpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "3. Python\n"
        "================================================================================\n\n"
        "License: Python Software Foundation License (PSF)\n"
        "Copyright: Python Software Foundation\n"
        "Website: https://www.python.org/\n\n"
        "ライセンス全文：https://docs.python.org/3/license.html\n\n\n"
        "================================================================================\n"
        "本ソフトウェアのライセンス\n"
        "================================================================================\n\n"
        "本ソフトウェア（PDF証拠整理ツール）は、\n"
        "GNU Affero General Public License v3.0 (AGPL-3.0) の下で配布されます。\n"
        "再配布の際はライセンス条件に従ってください。\n\n"
        "ソースコード：\n"
        "https://github.com/leantechlibrary-coder/LTL-Evidence-Marker-Tool\n\n\n"
        "================================================================================\n"
        "免責事項\n"
        "================================================================================\n\n"
        "本ソフトウェアは「現状有姿」(AS IS) で提供され、いかなる保証もありません。\n"
        "本ソフトウェアの使用により生じたいかなる損害についても、開発者は\n"
        "一切の責任を負いません。\n"
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("このソフトについて")
        self.resize(520, 480)
        self.setMinimumSize(400, 350)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        # --- タイトル ---
        title_label = QLabel("証拠番号付与ツール v1.0（無料トライアル版）")
        title_label.setFont(QFont("Yu Gothic UI", 12, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # --- 本文（スクロール可能） ---
        about_text = QTextEdit()
        about_text.setReadOnly(True)
        about_text.setFont(QFont("Yu Gothic UI", 9))
        about_text.setPlainText(
            "【動作環境】\n"
            "Windows 10 / 11 (64bit)\n\n"
            "【重要】\n"
            "本ソフトウェアは法律専門職の業務効率化を目的としており、\n"
            "専門知識を前提とした設計です。\n\n"
            "【免責事項】\n"
            "本ソフトウェアは「現状有姿」(AS IS) で提供されます。\n"
            "本ソフトウェアの使用により生じたいかなる損害についても、\n"
            "開発者は一切の責任を負いません。\n"
            "重要なファイルは必ずバックアップを取ってからご使用ください。\n\n"
            "【開発・販売】\n"
            "Lean Tech Library\n\n"
            "ご使用前に操作説明書・READMEをご確認ください。"
        )
        layout.addWidget(about_text)

        # --- 詳細情報リンクボタン群 ---
        link_layout = QHBoxLayout()
        link_layout.setSpacing(8)

        manual_btn = QPushButton("操作説明書")
        manual_btn.setToolTip("操作説明書を表示します")
        manual_btn.clicked.connect(self._show_manual)

        readme_btn = QPushButton("README")
        readme_btn.setToolTip("READMEを表示します")
        readme_btn.clicked.connect(self._show_readme)

        license_btn = QPushButton("ライセンス情報")
        license_btn.setToolTip("サードパーティライセンス情報を表示します")
        license_btn.clicked.connect(self._show_licenses)

        link_layout.addWidget(manual_btn)
        link_layout.addWidget(readme_btn)
        link_layout.addWidget(license_btn)
        layout.addLayout(link_layout)

        # --- 閉じるボタン ---
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_btn = QPushButton("閉じる")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        close_layout.addWidget(close_btn)
        close_layout.addStretch()
        layout.addLayout(close_layout)

    def _show_manual(self):
        dlg = TextViewerDialog(self, "操作説明書", self.MANUAL_TEXT)
        dlg.exec()

    def _show_readme(self):
        dlg = TextViewerDialog(self, "README", self.README_TEXT)
        dlg.exec()

    def _show_licenses(self):
        dlg = TextViewerDialog(self, "ライセンス情報", self.LICENSE_TEXT)
        dlg.exec()


def show_about_dialog():
    """Aboutダイアログを表示"""
    dlg = AboutDialog()
    dlg.exec()


class PDFFileItem(QListWidgetItem):
    """PDFファイル情報を保持するリストアイテム"""
    def __init__(self, file_path: Path):
        super().__init__()
        self.file_path = file_path
        self.is_branch = False  # 枝番フラグ
        self.update_display()
    
    def update_display(self):
        """表示テキストを更新"""
        self.setText(self.file_path.name)


class DraggableListWidget(QListWidget):
    """ドラッグ&ドロップで並び替え可能なリストウィジェット"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)


class EvidenceMarkerWindow(QMainWindow):
    """証拠番号付与ツールのメインウィンドウ"""
    def __init__(self):
        super().__init__()
        self.pdf_files: List[PDFFileItem] = []
        
        self.init_ui()
    
    def init_ui(self):
        """UIの初期化"""
        self.setWindowTitle("証拠番号付与ツール（無料トライアル版）")
        self.setGeometry(100, 100, 900, 700)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # 使い方説明
        help_label = QLabel(
            "使い方：\n"
            "1. フォルダまたはファイルを読み込む　2. リストをドラッグで並び替え　"
            "3. 枝番が必要な場合は選択して「枝番にする」　4. 設定を確認して「実行」"
        )
        help_label.setStyleSheet("""
            QLabel {
                background-color: #FFF9C4;
                padding: 10px;
                border: 1px solid #FBC02D;
                font-size: 10pt;
            }
        """)
        main_layout.addWidget(help_label)
        
        # ファイル読み込みボタン群
        file_buttons = QHBoxLayout()
        
        self.load_folder_btn = QPushButton("フォルダを開く")
        self.load_folder_btn.clicked.connect(self.load_folder)
        
        self.add_files_btn = QPushButton("ファイルを追加")
        self.add_files_btn.clicked.connect(self.add_files)

        self.remove_file_btn = QPushButton("ファイルを削除")
        self.remove_file_btn.clicked.connect(self.remove_selected_files)

        self.clear_list_btn = QPushButton("リストをクリア")
        self.clear_list_btn.clicked.connect(self.clear_list)

        file_buttons.addWidget(self.load_folder_btn)
        file_buttons.addWidget(self.add_files_btn)
        file_buttons.addWidget(self.remove_file_btn)
        file_buttons.addWidget(self.clear_list_btn)
        file_buttons.addStretch()
        
        # Aboutリンク
        about_label = QLabel('<a href="#" style="color: #888;">About</a>')
        about_label.setOpenExternalLinks(False)
        about_label.linkActivated.connect(lambda: show_about_dialog())
        file_buttons.addWidget(about_label)
        
        main_layout.addLayout(file_buttons)
        
        # ファイルリストと操作ボタンとサムネイル
        list_layout = QHBoxLayout()
        
        # 左：番号プレビュー（ラベルなし）
        self.preview_list = QListWidget()
        self.preview_list.setMaximumWidth(150)
        list_layout.addWidget(self.preview_list, stretch=1)
        
        # 中央：ファイルリスト
        self.file_list = DraggableListWidget()
        self.file_list.itemSelectionChanged.connect(self.on_selection_changed)
        self.file_list.model().rowsMoved.connect(self.on_list_reordered)
        list_layout.addWidget(self.file_list, stretch=3)
        
        # スクロール同期
        self.file_list.verticalScrollBar().valueChanged.connect(
            self.preview_list.verticalScrollBar().setValue
        )
        self.preview_list.verticalScrollBar().valueChanged.connect(
            self.file_list.verticalScrollBar().setValue
        )
        
        # リスト操作ボタン（縦長の細い列）
        list_control_layout = QVBoxLayout()
        
        self.move_up_btn = QPushButton("↑ 上へ")
        self.move_up_btn.clicked.connect(self.move_up)
        
        self.move_down_btn = QPushButton("↓ 下へ")
        self.move_down_btn.clicked.connect(self.move_down)
        
        self.set_branch_btn = QPushButton("枝番\nにする")
        self.set_branch_btn.clicked.connect(self.set_as_branch)
        
        self.unset_branch_btn = QPushButton("枝番\n解除")
        self.unset_branch_btn.clicked.connect(self.unset_branch)
        
        list_control_layout.addWidget(self.move_up_btn)
        list_control_layout.addWidget(self.move_down_btn)
        list_control_layout.addWidget(self.set_branch_btn)
        list_control_layout.addWidget(self.unset_branch_btn)
        list_control_layout.addStretch()
        
        list_layout.addLayout(list_control_layout)
        
        # 右：選択中ファイルのサムネイル
        thumbnail_layout = QVBoxLayout()
        thumbnail_label = QLabel("選択中のファイル")
        thumbnail_label.setStyleSheet("font-weight: bold; font-size: 10pt;")
        thumbnail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.thumbnail_display = QLabel()
        self.thumbnail_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail_display.setMinimumSize(250, 350)
        self.thumbnail_display.setMaximumWidth(300)
        self.thumbnail_display.setStyleSheet("""
            QLabel {
                border: 2px solid #CCCCCC;
                background-color: #F5F5F5;
            }
        """)
        self.thumbnail_display.setText("ファイルを選択してください")
        
        thumbnail_layout.addWidget(thumbnail_label)
        thumbnail_layout.addWidget(self.thumbnail_display)
        
        list_layout.addLayout(thumbnail_layout, stretch=2)
        
        main_layout.addLayout(list_layout)
        
        # 設定グループ
        settings_group = QGroupBox("証拠番号の設定")
        settings_layout = QVBoxLayout()
        
        # 証拠種別
        type_layout = QHBoxLayout()
        type_label = QLabel("証拠種別:")
        
        self.type_group = QButtonGroup()
        self.type_kou = QRadioButton("甲")
        self.type_otsu = QRadioButton("乙")
        self.type_custom = QRadioButton("その他:")
        self.type_kou.setChecked(True)
        
        self.type_group.addButton(self.type_kou, 0)
        self.type_group.addButton(self.type_otsu, 1)
        self.type_group.addButton(self.type_custom, 2)
        
        self.custom_prefix = QLineEdit()
        self.custom_prefix.setMaximumWidth(100)
        self.custom_prefix.setEnabled(False)
        self.type_custom.toggled.connect(lambda checked: self.custom_prefix.setEnabled(checked))
        
        # プレビュー更新ボタンを証拠種別の横に配置
        self.preview_update_btn = QPushButton("プレビュー更新")
        self.preview_update_btn.clicked.connect(self.update_preview)
        
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_kou)
        type_layout.addWidget(self.type_otsu)
        type_layout.addWidget(self.type_custom)
        type_layout.addWidget(self.custom_prefix)
        type_layout.addWidget(self.preview_update_btn)
        type_layout.addStretch()
        
        settings_layout.addLayout(type_layout)
        
        # 開始番号
        start_layout = QHBoxLayout()
        start_label = QLabel("開始番号:")
        self.start_number = QSpinBox()
        self.start_number.setMinimum(1)
        self.start_number.setMaximum(9999)
        self.start_number.setValue(1)
        
        start_layout.addWidget(start_label)
        start_layout.addWidget(self.start_number)
        start_layout.addStretch()
        
        settings_layout.addLayout(start_layout)
        
        # フォント設定
        font_layout = QHBoxLayout()

        # 印字ON/OFFチェックボックス
        from PyQt6.QtWidgets import QCheckBox
        self.print_number_chk = QCheckBox("証拠番号を印字する")
        self.print_number_chk.setChecked(True)

        font_size_label = QLabel("フォントサイズ:")
        self.font_size = QSpinBox()
        self.font_size.setMinimum(8)
        self.font_size.setMaximum(72)
        self.font_size.setValue(16)
        self.font_size.setSuffix(" pt")

        font_color_label = QLabel("フォント色:")
        self.font_color = QComboBox()
        self.font_color.addItems(["赤", "黒", "青"])
        self.font_color.setCurrentText("赤")

        # チェックOFF時はフォント設定をグレーアウト
        def _on_print_toggle(checked):
            font_size_label.setEnabled(checked)
            self.font_size.setEnabled(checked)
            font_color_label.setEnabled(checked)
            self.font_color.setEnabled(checked)
        self.print_number_chk.toggled.connect(_on_print_toggle)

        font_layout.addWidget(self.print_number_chk)
        font_layout.addSpacing(16)
        font_layout.addWidget(font_size_label)
        font_layout.addWidget(self.font_size)
        font_layout.addWidget(font_color_label)
        font_layout.addWidget(self.font_color)
        font_layout.addStretch()

        settings_layout.addLayout(font_layout)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        # 実行ボタン
        execute_layout = QHBoxLayout()
        
        self.execute_btn = QPushButton("証拠番号を付与して保存")
        self.execute_btn.clicked.connect(self.execute_marking)
        self.execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        
        execute_layout.addStretch()
        execute_layout.addWidget(self.execute_btn)
        
        main_layout.addLayout(execute_layout)
        
        # ドラッグ&ドロップを有効化
        self.setAcceptDrops(True)
    
    def dragEnterEvent(self, event):
        """ドラッグされたファイルを受け入れる"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """ドロップされたファイル/フォルダを処理"""
        urls = event.mimeData().urls()
        for url in urls:
            path = Path(url.toLocalFile())
            if path.is_dir():
                self.load_folder_path(path)
            elif path.suffix.lower() == '.pdf':
                self.add_file(path)
        self.update_preview()
    
    def load_folder(self):
        """フォルダを選択してPDFファイルを読み込む"""
        folder_path = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if folder_path:
            self.load_folder_path(Path(folder_path))
            self.update_preview()
    
    def load_folder_path(self, folder_path: Path):
        """指定されたフォルダからPDFを読み込む"""
        pdf_files = sorted(folder_path.glob("*.pdf"))
        
        if not pdf_files:
            QMessageBox.warning(self, "警告", "PDFファイルが見つかりませんでした")
            return
        
        self.file_list.clear()
        
        for pdf_path in pdf_files:
            item = PDFFileItem(pdf_path)
            self.file_list.addItem(item)
        
        QMessageBox.information(self, "読み込み完了", f"{len(pdf_files)}個のPDFファイルを読み込みました")
    
    def add_files(self):
        """ファイルを追加"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "PDFファイルを選択", "", "PDF Files (*.pdf)"
        )
        
        for file_path in file_paths:
            self.add_file(Path(file_path))
        
        if file_paths:
            self.update_preview()
    
    def add_file(self, file_path: Path):
        """単一ファイルを追加"""
        item = PDFFileItem(file_path)
        self.file_list.addItem(item)
    
    def remove_selected_files(self):
        """選択中のファイルをリストから削除（元ファイルは変更しない）"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "削除するファイルを選択してください")
            return
        for item in selected_items:
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
        # 先頭が枝番になっていた場合に備えてフラグを修正
        self._fix_leading_branch()
        self.update_preview()

    def clear_list(self):
        """リストをクリア"""
        reply = QMessageBox.question(
            self, "確認", "リストをクリアしますか?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.file_list.clear()
            self.preview_list.clear()
    
    def move_up(self):
        """選択項目を上に移動"""
        current_row = self.file_list.currentRow()
        if current_row > 0:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row - 1, item)
            self.file_list.setCurrentRow(current_row - 1)
            self.update_preview()
    
    def move_down(self):
        """選択項目を下に移動"""
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row + 1, item)
            self.file_list.setCurrentRow(current_row + 1)
            self.update_preview()
    
    def on_selection_changed(self):
        """ファイル選択が変更された時"""
        self.update_thumbnail_display()
    
    def on_list_reordered(self):
        """リストが並び替えられた時"""
        self._fix_leading_branch()
        self.update_preview()

    def _fix_leading_branch(self):
        """先頭アイテムが枝番になっていたら自動で枝番を解除する"""
        if self.file_list.count() == 0:
            return
        first_item = self.file_list.item(0)
        if isinstance(first_item, PDFFileItem) and first_item.is_branch:
            first_item.is_branch = False
    
    def set_as_branch(self):
        """選択項目を枝番にする
        
        枝番にすると、一つ上の非枝番ファイルが自動的に「の1」になり、
        選択したファイルが「の2」になる。
        先頭ファイル（または枝番グループの先頭になるファイル）は枝番にできない。
        """
        selected_items = self.file_list.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "警告", "枝番にするファイルを選択してください")
            return
        
        # 先頭アイテムを枝番にしようとしていないかチェック
        for item in selected_items:
            if isinstance(item, PDFFileItem):
                row = self.file_list.row(item)
                if row == 0:
                    QMessageBox.warning(self, "警告", 
                        "リストの先頭のファイルは枝番にできません。\n"
                        "枝番にするには、一つ上に親となるファイルが必要です。")
                    return
                # 一つ上のアイテムも枝番の場合はOK（既存の枝番グループに追加）
                # 一つ上が非枝番の場合もOK（新しい枝番グループの開始）
        
        for item in selected_items:
            if isinstance(item, PDFFileItem):
                item.is_branch = True
        
        self.update_preview()
    
    def unset_branch(self):
        """枝番を解除"""
        selected_items = self.file_list.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "警告", "枝番を解除するファイルを選択してください")
            return
        
        for item in selected_items:
            if isinstance(item, PDFFileItem):
                item.is_branch = False
        
        self.update_preview()
    
    def get_prefix(self) -> str:
        """証拠種別のプレフィックスを取得"""
        if self.type_kou.isChecked():
            return "甲"
        elif self.type_otsu.isChecked():
            return "乙"
        else:
            return self.custom_prefix.text() or "証"
    
    def generate_evidence_numbers(self) -> List[Tuple[PDFFileItem, str]]:
        """証拠番号を生成
        
        枝番の挙動：
        - 枝番ファイルの直前の非枝番ファイルは自動的に「の1」になる
        - 枝番ファイルは「の2」「の3」...と続く
        - 親番号（枝番なし）は存在しない（すべて並列的な枝番になる）
        例：ファイルA, ファイルB(枝番), ファイルC
          → 甲第1号証の1, 甲第1号証の2, 甲第2号証
        """
        prefix = self.get_prefix()
        start_num = self.start_number.value()
        
        # まず各アイテムを収集
        items = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if isinstance(item, PDFFileItem):
                items.append(item)
        
        if not items:
            return []
        
        # 先読み：各非枝番アイテムの直後に枝番が続くかを判定
        has_branch_after = [False] * len(items)
        for i in range(len(items) - 1):
            if not items[i].is_branch and items[i + 1].is_branch:
                has_branch_after[i] = True
        
        # 番号を生成
        result = []
        current_main = start_num - 1
        branch_count = 0
        
        for i, item in enumerate(items):
            if item.is_branch:
                branch_count += 1
                evidence_number = f"{prefix}第{current_main}号証の{branch_count}"
            else:
                current_main += 1
                if has_branch_after[i]:
                    # 直後に枝番が続く → この項目は「の1」になる
                    branch_count = 1
                    evidence_number = f"{prefix}第{current_main}号証の{branch_count}"
                else:
                    branch_count = 0
                    evidence_number = f"{prefix}第{current_main}号証"
            
            result.append((item, evidence_number))
        
        return result
    
    def update_preview(self):
        """プレビューを更新"""
        self.preview_list.clear()
        
        evidence_list = self.generate_evidence_numbers()
        
        for item, evidence_number in evidence_list:
            preview_item = QListWidgetItem(evidence_number)
            
            if item.is_branch:
                preview_item.setForeground(QColor("#FF6B6B"))
            
            self.preview_list.addItem(preview_item)
        
        # 選択中のファイルのサムネイルを表示
        self.update_thumbnail_display()
    
    def update_thumbnail_display(self):
        """選択中のファイルのサムネイルを表示"""
        selected_items = self.file_list.selectedItems()
        
        if not selected_items:
            self.thumbnail_display.setText("ファイルを選択してください")
            return
        
        # 最初に選択されたファイルのサムネイルを表示
        selected_item = selected_items[0]
        
        if isinstance(selected_item, PDFFileItem):
            try:
                # PDFを開く
                pdf_doc = fitz.open(selected_item.file_path)
                
                # 1ページ目を取得
                first_page = pdf_doc[0]
                landscape = is_page_landscape(first_page)
                
                # サムネイル生成（回転属性を反映した状態でレンダリング）
                # get_pixmapはデフォルトでページのrotation属性を反映する
                pix = first_page.get_pixmap(dpi=72)
                
                # 表示領域に収まるようにスケーリング
                max_w, max_h = 250, 350
                scale = min(max_w / pix.width, max_h / pix.height)
                if scale < 1:
                    mat = fitz.Matrix(scale, scale)
                    pix = first_page.get_pixmap(matrix=fitz.Matrix(72/72 * scale, 72/72 * scale))
                
                # QPixmapに変換
                from PyQt6.QtGui import QImage, QPixmap
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # 表示
                self.thumbnail_display.setPixmap(pixmap)
                
                pdf_doc.close()
                
            except Exception as e:
                self.thumbnail_display.setText(f"プレビュー\n読み込み失敗:\n{str(e)}")
    
    def generate_filename(self, prefix: str, number: str) -> str:
        """ファイル名を生成（簡易形式）"""
        # 例：「1」→「甲01.pdf」、「1の2」→「甲01の2.pdf」
        if 'の' in number:
            main, branch = number.split('の')
            filename_number = f"{int(main):02d}の{branch}"
        else:
            filename_number = f"{int(number):02d}"
        
        return f"{prefix}{filename_number}.pdf"
    
    def execute_marking(self):
        """証拠番号を付与して保存"""
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "PDFファイルを読み込んでください")
            return
        
        # 確認ダイアログ
        evidence_list = self.generate_evidence_numbers()
        
        message = f"{len(evidence_list)}個のPDFファイルに証拠番号を付与します。よろしいですか?\n\n"
        message += "最初の5件:\n"
        for i, (item, evidence_number) in enumerate(evidence_list[:5]):
            message += f"{evidence_number}: {item.file_path.name}\n"
        
        if len(evidence_list) > 5:
            message += f"... 他 {len(evidence_list) - 5}件"
        
        reply = QMessageBox.question(
            self, "確認", message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return

        # 出力フォルダを自動生成（元ファイルが複数フォルダにまたがる場合は
        # 最初のファイルの親フォルダを基準にする）
        first_item = evidence_list[0][0]
        base_folder = first_item.file_path.parent
        output_path = base_folder / f"{base_folder.name}_番号付"

        # 同名フォルダが既にある場合は連番を付ける
        if output_path.exists():
            suffix = 2
            while True:
                candidate = base_folder / f"{base_folder.name}_番号付_{suffix}"
                if not candidate.exists():
                    output_path = candidate
                    break
                suffix += 1

        output_path.mkdir(parents=True, exist_ok=True)
        
        # フォント設定
        do_print = self.print_number_chk.isChecked()
        font_size = self.font_size.value()
        color_map = {
            "赤": (1, 0, 0),
            "黒": (0, 0, 0),
            "青": (0, 0, 1)
        }
        font_color = color_map[self.font_color.currentText()]

        # 右端からのマージン（pt）
        RIGHT_MARGIN = 25
        # 上端からのY位置（pt）
        TOP_Y = 30

        def text_width(text, fontsize):
            """証拠番号テキストの描画幅（pt）を返す"""
            return fitz.get_text_length(text, fontname="japan", fontsize=fontsize)

        try:
            # 処理実行
            for item, evidence_number in evidence_list:
                # PDFを開く
                pdf_doc = fitz.open(item.file_path)

                # 1ページ目を取得
                first_page = pdf_doc[0]

                if do_print:
                    # テキスト幅を計算（右揃え用）
                    tw = text_width(evidence_number, font_size)

                    # 先頭ページのサイズ・向き判定
                    is_landscape = is_page_landscape(first_page)
                    is_a3 = is_page_a3(first_page)

                    if is_landscape and is_a3:
                        # === A3横向き文書の処理 ===
                        rect = first_page.rect
                        w, h = rect.width, rect.height
                        current_rotation = first_page.rotation % 360

                        if current_rotation == 0:
                            x = w - RIGHT_MARGIN - tw
                            y = TOP_Y
                            text_rotate = 0
                        elif current_rotation == 90:
                            eff_w = h
                            x = TOP_Y
                            y = h - (eff_w - RIGHT_MARGIN - tw)
                            text_rotate = 90
                        elif current_rotation == 270:
                            eff_w = h
                            x = w - TOP_Y
                            y = eff_w - RIGHT_MARGIN - tw
                            text_rotate = 270
                        else:
                            x = RIGHT_MARGIN + tw
                            y = h - TOP_Y
                            text_rotate = 180

                        first_page.insert_text(
                            (x, y),
                            evidence_number,
                            fontsize=font_size,
                            color=font_color,
                            fontname="japan",
                            rotate=text_rotate
                        )

                    elif is_landscape:
                        # === A4横向き文書の処理（270度回転してA4縦向きに変換）===
                        rect = first_page.rect
                        w, h = rect.width, rect.height
                        current_rotation = first_page.rotation % 360

                        total_rotation = (current_rotation + 270) % 360

                        if total_rotation in (90, 270):
                            effective_width = h
                        else:
                            effective_width = w

                        target_vx = effective_width - RIGHT_MARGIN - tw
                        target_vy = TOP_Y

                        if current_rotation == 0:
                            orig_x = w - target_vy
                            orig_y = target_vx
                            text_rotate = 270
                        elif current_rotation == 90:
                            orig_x = target_vx
                            orig_y = target_vy
                            text_rotate = 0
                        elif current_rotation == 270:
                            orig_x = w - target_vx
                            orig_y = h - target_vy
                            text_rotate = 180
                        else:
                            orig_x = target_vy
                            orig_y = h - target_vx
                            text_rotate = 90

                        first_page.insert_text(
                            (orig_x, orig_y),
                            evidence_number,
                            fontsize=font_size,
                            color=font_color,
                            fontname="japan",
                            rotate=text_rotate
                        )

                        for page_num in range(len(pdf_doc)):
                            page = pdf_doc[page_num]
                            if is_page_landscape(page) and not is_page_a3(page):
                                cur_rot = page.rotation % 360
                                new_rotation = (cur_rot + 270) % 360
                                page.set_rotation(new_rotation)

                    else:
                        # === 縦向き文書の処理 ===
                        x = first_page.rect.width - RIGHT_MARGIN - tw
                        y = TOP_Y

                        first_page.insert_text(
                            (x, y),
                            evidence_number,
                            fontsize=font_size,
                            color=font_color,
                            fontname="japan"
                        )

                # ファイル名を生成（簡易形式）
                prefix = self.get_prefix()
                number_part = evidence_number.replace(prefix + "第", "").replace("号証", "")
                filename = self.generate_filename(prefix, number_part)

                # 保存
                output_file = output_path / filename
                pdf_doc.save(output_file)
                pdf_doc.close()

            # 完了メッセージ
            action_label = "証拠番号を付与" if do_print else "ファイル名を変更（印字なし）"
            QMessageBox.information(
                self, "完了",
                f"{len(evidence_list)}個のPDFファイルに{action_label}しました\n\n"
                f"出力先：{output_path}"
            )
            
            # 出力フォルダを開く
            if sys.platform == 'win32':
                os.startfile(output_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_path}"')
            else:
                os.system(f'xdg-open "{output_path}"')
        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"処理中にエラーが発生しました:\n{str(e)}")


def main():
    app = QApplication(sys.argv)

    # --- トライアルチェック ---
    trial = TrialManager()
    remaining = trial.days_remaining()
    expired = trial.is_expired()

    dlg = TrialDialog(remaining, expired)

    if expired:
        # 期限切れ → ダイアログ表示後にアプリ終了
        dlg.exec()
        sys.exit(0)
    else:
        # 期限内 → 「閉じて続ける」でメインウィンドウへ
        result = dlg.exec()
        if result == QDialog.DialogCode.Rejected:
            # ×ボタンで閉じた場合もアプリ終了
            sys.exit(0)

    window = EvidenceMarkerWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
