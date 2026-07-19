import sys
import os
from pathlib import Path
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QFileDialog, QMessageBox, QLabel,
    QLineEdit, QRadioButton, QButtonGroup, QSpinBox, QComboBox,
    QGroupBox, QListWidgetItem, QAbstractItemView, QDialog, QTextEdit
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from typing import List, Tuple
from stamp_core import (
    stamp_evidence_number, verify_stamp, subset_embedded_fonts,
    format_stamp_text, make_filename as core_make_filename,
)


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

    # 免責文の正本（About・README・ライセンス情報の3か所から参照し、表記を統一）
    # ライセンス脚注（AGPL-3.0）は各箇所で別途付記する。
    DISCLAIMER_BODY = (
        "本ソフトウェアは、明示・黙示を問わず、商品性、特定目的への\n"
        "適合性、正確性、継続的動作等のいかなる保証も行わず、\n"
        "現状有姿（AS IS）で提供されます。\n"
        "開発者は、本ソフトウェアの使用・使用不能に起因する\n"
        "直接的・間接的・付随的・特別・懲罰的損害（出力結果の誤り、\n"
        "ファイルの消失、業務上の損失等を含むがこれに限らない）について、\n"
        "一切の責任を負いません。\n"
        "出力結果の確認およびバックアップの取得は、\n"
        "利用者の責任において行ってください。\n"
        "動作保証・バグ修正・機能追加・質問対応等のサポートは\n"
        "提供しません。"
    )

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
        + DISCLAIMER_BODY + "\n\n\n"
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
        "  ・例：第2号証の後に枝番を設定すると、ファイル名は「甲002-1」「甲002-2」\n"
        "  ・解除する場合は「枝番を解除」ボタン\n\n"
        "(3') 横向き資料の向き補正（必要な場合のみ）\n"
        "  ・対象ファイルを選択し「↻ 右90°」ボタンをクリック\n"
        "  ・押すたびに90度ずつ右回転（4回で元に戻ります）\n"
        "  ・リスト上では「↻」マークで回転中のファイルが分かります\n"
        "  ・回転は保存実行時にPDFへ反映されます（元ファイルは変更しません）\n"
        "  ・1ファイル内で縦横が混在する場合は全ページに同じ回転がかかります\n\n"
        "(4) 証拠番号の設定\n"
        "  ・証拠種別：甲/乙/その他（カスタム文字列）\n"
        "  ・開始番号：通常は1から\n"
        "  ・証拠番号を印字する：チェックONで1ページ目右上に番号を印字\n"
        "  ・印字書式：「001」または「第1号証」を選択（証拠種別と組み合わせて\n"
        "    「甲001」「乙第1号証」等を印字。既定は「001」。ファイル名・証拠説明書・\n"
        "    mints証拠番号欄と揃う「001」形式を推奨）\n"
        "  ・フォントサイズ：8pt～72pt（デフォルト16pt）\n"
        "  ・フォント色：赤/黒/青（デフォルト赤）\n\n"
        "(5) 実行\n"
        "  ・「証拠番号を付与して保存」ボタンをクリック\n"
        "  ・確認ダイアログで内容を確認\n"
        "  ・完了後、出力フォルダが自動的に開きます\n\n"
        "【出力について】\n"
        "  ・出力先：読み込んだファイルの親フォルダ内に\n"
        "    「親フォルダ名_番号付」フォルダを自動作成\n"
        "  ・ファイル名：「甲001.pdf」「甲002.pdf」「甲003-1.pdf」など\n\n\n"
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
        "4. LTL Evidence Sans（刻印用フォント）\n"
        "================================================================================\n\n"
        "License: SIL Open Font License, Version 1.1 (OFL-1.1)\n"
        "Copyright: © 2014-2021 Adobe (http://www.adobe.com/),\n"
        "           with Reserved Font Name 'Source'.\n"
        "           Noto is a trademark of Google Inc.\n\n"
        "本フォントは Noto Sans CJK JP を証拠番号刻印用の145グリフに\n"
        "サブセットし、TrueTypeアウトラインへ変換のうえ\n"
        "「LTL Evidence Sans」に改名した派生フォントです\n"
        "（OFLのReserved Font Name条項に基づく改名）。\n"
        "刻印時に出力PDFへサブセット埋め込みされます。\n\n"
        "ライセンス全文：同梱の fonts/OFL.txt および\n"
        "https://scripts.sil.org/OFL\n\n\n"
        "================================================================================\n"
        "5. fontTools\n"
        "================================================================================\n\n"
        "License: MIT License\n"
        "Copyright: 2017 Just van Rossum and others\n"
        "Website: https://github.com/fonttools/fonttools\n\n"
        "埋め込みフォントのサブセット化に使用しています。\n\n"
        "ライセンス全文：\n"
        "https://github.com/fonttools/fonttools/blob/main/LICENSE\n\n\n"
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
        + DISCLAIMER_BODY + "\n"
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
        title_label = QLabel("証拠番号付与ツール v1.0")
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
            + self.DISCLAIMER_BODY + "\n\n"
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
        # 回転保留量（度・時計回り。0/90/180/270）。
        # ボタンで積み増すだけの「保留状態」であり、元ファイルにも出力にも
        # この時点では一切反映しない。実体化は execute_marking（保存実行）時に
        # set_rotation で全ページへ一律に焼き込む。is_branch と同じモデル。
        self.rotation = 0
        self.update_display()

    def update_display(self):
        """表示テキストを更新（回転保留中は先頭に「↻」を付けて可視化する）"""
        mark = "↻ " if self.rotation else ""
        self.setText(f"{mark}{self.file_path.name}")


class DraggableListWidget(QListWidget):
    """ドラッグ&ドロップで並び替え可能なリストウィジェット（属性保持版）

    QListWidget標準のInternalMoveは、ドロップ時にアイテムをmimeData経由で
    再生成するため、PDFFileItemのサブクラス属性（file_path・is_branch）が
    失われる。これを避けるため、dropEventを独自実装し、元のアイテムオブジェクトを
    takeItem/insertItemで物理的に移動させる（オブジェクトの同一性を保つ）。
    並び替え完了は reordered シグナルで通知する（takeItem/insertItemは
    rowsMovedを発火しないため）。
    """
    reordered = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDropIndicatorShown(True)

    def dropEvent(self, event):
        # 内部移動以外（外部からのドロップ等）は既定の処理に委ねる
        if event.source() is not self:
            super().dropEvent(event)
            return

        selected_items = self.selectedItems()
        if not selected_items:
            event.ignore()
            return

        # ドロップ先の行を決定（アイテムの上に落ちたらその行、空白なら末尾）
        drop_item = self.itemAt(event.position().toPoint())
        drop_row = self.row(drop_item) if drop_item is not None else self.count()

        # 取り出す行（昇順）。取り出しで生じるドロップ位置のズレを補正するため、
        # drop_rowより前にある選択行の数だけ挿入位置を前へずらす。
        rows = sorted(self.row(it) for it in selected_items)
        shift = sum(1 for r in rows if r < drop_row)
        target_row = drop_row - shift

        # 後ろの行から取り出してインデックスのズレを防ぐ
        taken = [self.takeItem(r) for r in sorted(rows, reverse=True)]
        taken.reverse()  # 元の表示順に戻す

        target_row = max(0, min(target_row, self.count()))
        for i, it in enumerate(taken):
            self.insertItem(target_row + i, it)

        # 選択状態を復元
        self.clearSelection()
        for it in taken:
            it.setSelected(True)

        event.acceptProposedAction()
        self.reordered.emit()


class EvidenceMarkerWindow(QMainWindow):
    """証拠番号付与ツールのメインウィンドウ"""
    def __init__(self):
        super().__init__()
        self.pdf_files: List[PDFFileItem] = []

        self.init_ui()

    def init_ui(self):
        """UIの初期化"""
        self.setWindowTitle("証拠番号付与ツール")
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
        self.file_list.reordered.connect(self.on_list_reordered)
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

        self.rotate_btn = QPushButton("↻ 右90°")
        self.rotate_btn.setToolTip(
            "選択中（サムネイル表示中）のファイルを右に90度回転します。\n"
            "押すたびに90度ずつ回り、4回で元に戻ります。\n"
            "回転は保存実行時にPDFへ反映され、元ファイルは変更しません。"
        )
        self.rotate_btn.clicked.connect(self.rotate_selected)

        list_control_layout.addWidget(self.move_up_btn)
        list_control_layout.addWidget(self.move_down_btn)
        list_control_layout.addWidget(self.set_branch_btn)
        list_control_layout.addWidget(self.unset_branch_btn)
        list_control_layout.addWidget(self.rotate_btn)
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

        # 設定値の変更を即座にプレビューへ反映する
        # （開始番号・証拠種別を変えてもプレビューが追従しなかった不具合への対応）
        self.type_group.idToggled.connect(lambda _id, checked: self.update_preview() if checked else None)
        self.custom_prefix.textChanged.connect(lambda _text: self.update_preview())

        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_kou)
        type_layout.addWidget(self.type_otsu)
        type_layout.addWidget(self.type_custom)
        type_layout.addWidget(self.custom_prefix)
        type_layout.addStretch()

        settings_layout.addLayout(type_layout)

        # 開始番号
        start_layout = QHBoxLayout()
        start_label = QLabel("開始番号:")
        self.start_number = QSpinBox()
        self.start_number.setMinimum(1)
        self.start_number.setMaximum(9999)
        self.start_number.setValue(1)
        self.start_number.valueChanged.connect(lambda _v: self.update_preview())

        start_layout.addWidget(start_label)
        start_layout.addWidget(self.start_number)
        start_layout.addStretch()

        settings_layout.addLayout(start_layout)

        # 印字書式（右上に刻印する証拠番号の表記）
        # プレフィックス（甲/乙/その他）に依存しない汎用ラベルで提示する。
        # 実際の印字は選択中のプレフィックスと組み合わせて生成されるため、
        # 例えばプレフィックスが「乙」なら「乙001」「乙第1号証」となる。
        # 既定（先頭・index 0）は「001」（mints手引のファイル名・証拠説明書号証欄・
        # フォーム証拠番号欄と揃う半角三桁表記）。「第1号証」は従来の引用表記を
        # 好む場合・紙運用向けの選択肢。
        # なお、ファイル名は印字書式に関わらず常に三桁形式（mints手引29頁）。
        format_layout = QHBoxLayout()
        format_label = QLabel("印字書式:")
        self.print_format = QComboBox()
        # index 0 → num3（例:甲001）, index 1 → gou（例:甲第1号証）
        self.print_format.addItems(["001", "第1号証"])
        self.print_format.setCurrentIndex(0)
        self.print_format.currentIndexChanged.connect(lambda _i: self.update_preview())

        format_layout.addWidget(format_label)
        format_layout.addWidget(self.print_format)
        format_layout.addStretch()

        settings_layout.addLayout(format_layout)

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
        """ドロップされたファイル/フォルダを処理する。

        ドロップされた要素を先に「フォルダ」と「PDFファイル」へ仕分けてから処理する。
        - フォルダが1つでも含まれる場合：リストを一度だけクリアし、ドロップされた
          全フォルダ内のPDFと、同時にドロップされた個別ファイルを読み込む
          （フォルダのドロップ＝置き換え。複数フォルダでも最後の1つに化けず全部入る）。
        - PDFファイルのみの場合：既存リストへ追記する。

        旧実装はフォルダ1件ごとに load_folder_path（内部で clear）を呼んでいたため、
        ファイルとフォルダを混在ドロップすると処理順次第で先に追加したファイルが
        消えたり、複数フォルダをドロップすると最後のフォルダしか残らない問題があった。
        """
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls()]
        dirs = [p for p in paths if p.is_dir()]
        files = [p for p in paths if not p.is_dir() and p.suffix.lower() == '.pdf']

        if dirs:
            # フォルダを含むドロップ：置き換え（一度だけクリア）
            self.file_list.clear()
            added = 0
            for d in dirs:
                for pdf_path in sorted(d.glob("*.pdf")):
                    self.add_file(pdf_path)
                    added += 1
            for f in files:
                self.add_file(f)
                added += 1
            if added == 0:
                QMessageBox.warning(self, "警告", "PDFファイルが見つかりませんでした")
        else:
            # PDFファイルのみのドロップ：既存リストへ追記
            for f in files:
                self.add_file(f)

        self._fix_leading_branch()
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
            # 枝番アイテムを先頭へ動かした場合に備えてフラグを修正
            self._fix_leading_branch()
            self.update_preview()

    def move_down(self):
        """選択項目を下に移動"""
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row + 1, item)
            self.file_list.setCurrentRow(current_row + 1)
            # 直前まで先頭だった枝番アイテムが繰り上がる場合に備えてフラグを修正
            self._fix_leading_branch()
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

    def rotate_selected(self):
        """選択中のファイルを右に90度回転（保留）する。

        一括回転は行わない。サムネイルは先頭の選択ファイル1件しか映らず、
        複数同時回転は確認できないまま向きを変えてしまう事故につながるため、
        対象はサムネイル表示中のファイル（selected_items[0]）に限定する。

        ここで行うのは rotation の積み増し（+90、4回で一周＝取り消し経路を兼ねる）と
        表示更新のみ。元ファイル・出力には触れず、実体化は execute_marking で行う。
        """
        selected_items = self.file_list.selectedItems()

        if not selected_items:
            QMessageBox.warning(self, "警告", "回転するファイルを選択してください")
            return

        # サムネイルに映っているファイル（先頭選択）だけを回す
        item = selected_items[0]
        if not isinstance(item, PDFFileItem):
            return

        item.rotation = (item.rotation + 90) % 360
        item.update_display()          # リスト上の「↻」マーカーを更新
        self.update_thumbnail_display()  # プレビューを回転後の向きで描き直す

    def get_prefix(self) -> str:
        """証拠種別のプレフィックスを取得"""
        if self.type_kou.isChecked():
            return "甲"
        elif self.type_otsu.isChecked():
            return "乙"
        else:
            return self.custom_prefix.text() or "証"

    def current_print_style(self) -> str:
        """現在選択されている印字書式を内部キーで返す（"num3" または "gou"）。

        プルダウンの表示ラベル（"001"/"第1号証"）に依存せず、index で判定する
        （index 0 = num3, index 1 = gou）。
        """
        return "num3" if self.print_format.currentIndex() == 0 else "gou"

    def format_evidence_number(self, prefix: str, main: int, branch: int,
                               style: str) -> str:
        """採番情報（プレフィックス・主番号・枝番）を指定書式の文字列へ整形する。

        branch=0 は枝番なし。
        - style="num3"（甲001）：主番号は「甲001」、枝番は「甲001-2」（mints表記）
        - style="gou"（甲第1号証）：主番号は「甲第1号証」、枝番は「甲第1号証の2」

        書式ロジックは stamp_core に集約（GUIとMCPで1コピー）。出力は不変。
        """
        return format_stamp_text(prefix, main, branch, style)

    def make_filename(self, prefix: str, main: int, branch: int) -> str:
        """出力ファイル名を生成（mints命名規則準拠・印字書式に依存しない）

        mintsの手引（29頁）に従い、ファイル名の証拠番号は常に半角三桁の
        「甲001」形式とし、枝番はハイフン区切りとする（右上に刻印する書式の
        選択とは独立。ファイル名・証拠説明書号証欄・フォーム証拠番号欄は
        いずれも甲001形式で統一されるため）。

        例：(主1, 枝0)→「甲001.pdf」、(主1, 枝2)→「甲001-2.pdf」

        命名ロジックは stamp_core に集約（GUIとMCPで1コピー）。出力は不変。
        """
        return core_make_filename(prefix, main, branch)

    def generate_evidence_numbers(self) -> List[Tuple[PDFFileItem, int, int]]:
        """証拠番号を生成し、(アイテム, 主番号, 枝番) のタプル列を返す。

        枝番（戻り値の3要素目）は 0 が「枝番なし」、1 以上が「の1, の2, …」。
        表示・刻印用の文字列化は format_evidence_number、ファイル名化は
        make_filename が担い、本メソッドは書式に依存しない採番のみを行う。

        枝番の挙動：
        - 枝番ファイルの直前の非枝番ファイルは自動的に「の1」になる
        - 枝番ファイルは「の2」「の3」...と続く
        - 親番号（枝番なし）は存在しない（すべて並列的な枝番になる）
        例：ファイルA, ファイルB(枝番), ファイルC
          → (A,1,1), (B,1,2), (C,2,0)

        なお、本メソッドはあらゆる呼び出し経路（プレビュー更新・実行）の
        単一窓口になっているため、冒頭で _fix_leading_branch を呼び、
        先頭が枝番のまま採番されて「第0号証」が生成される不具合を防ぐ
        （↑↓ボタンでの移動など、_fix_leading_branch を経由しない操作の
        保険を兼ねる）。
        """
        # 先頭が枝番のまま採番されないよう正規化（全経路の保険）
        self._fix_leading_branch()

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
            else:
                current_main += 1
                # 直後に枝番が続く非枝番は「の1」、それ以外は枝番なし(0)
                branch_count = 1 if has_branch_after[i] else 0

            result.append((item, current_main, branch_count))

        return result

    def update_preview(self):
        """プレビューを更新（選択中の印字書式どおりに番号を表示）"""
        self.preview_list.clear()

        prefix = self.get_prefix()
        style = self.current_print_style()
        evidence_list = self.generate_evidence_numbers()

        for item, main, branch in evidence_list:
            text = self.format_evidence_number(prefix, main, branch, style)
            preview_item = QListWidgetItem(text)

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

        if not isinstance(selected_item, PDFFileItem):
            return

        pdf_doc = None
        try:
            pdf_doc = fitz.open(str(selected_item.file_path))
            first_page = pdf_doc[0]

            # サムネイル生成（get_pixmapはページの /Rotate を反映してレンダリングする）
            # 表示枠（パネル）サイズ
            max_w, max_h = 250, 350

            # 回転90/270では表示時に縦横が入れ替わるので、収め先の枠も入れ替えて
            # 「回転前」の縮尺を決める。こうしてPDFから目標サイズで直接レンダリングし、
            # 回転は90度単位の無損失変換（転置）で済ませる。一度大きく焼いてから
            # ラスター縮小する経路を通らないため、どの向きでもぼやけない。
            rot = selected_item.rotation
            if rot in (90, 270):
                fit_w, fit_h = max_h, max_w
            else:
                fit_w, fit_h = max_w, max_h

            pix = first_page.get_pixmap(dpi=72)
            scale = min(fit_w / pix.width, fit_h / pix.height)
            if scale < 1:
                # PDFから縮小後の縮尺で直接レンダリング（ベクター→ラスター。鮮明）
                pix = first_page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            # scale>=1（元が枠より小さいページ）は等倍のまま。拡大ボケを避ける。

            # QPixmapに変換
            from PyQt6.QtGui import QImage, QPixmap, QTransform
            img = QImage(pix.samples, pix.width, pix.height, pix.stride,
                         QImage.Format.Format_RGB888).copy()
            pixmap = QPixmap.fromImage(img)

            # 回転保留量をプレビューに反映する。
            # pix はファイル自身の /Rotate を既に反映済みなので、ここで重ねるのは
            # ユーザーが押した分（保留中の補正）のみ。保存時の見え方と一致する。
            # 90/180/270 の軸そろえ回転は FastTransformation（既定）なら画素単位で
            # 無損失（転置）になるため、SmoothTransformation は指定しない。
            if rot:
                pixmap = pixmap.transformed(QTransform().rotate(rot))

            self.thumbnail_display.setPixmap(pixmap)

        except Exception as e:
            self.thumbnail_display.setText(f"プレビュー\n読み込み失敗:\n{e}")
        finally:
            if pdf_doc is not None:
                pdf_doc.close()

    def generate_filename(self, prefix: str, number: str) -> str:
        """[非推奨] 旧API。make_filename(prefix, main, branch) を使用すること。

        互換のために残置。文字列番号（"1" / "1の2"）を受け取り、
        新ヘルパー make_filename に委譲する。
        """
        if 'の' in number:
            main, branch = number.split('の')
            return self.make_filename(prefix, int(main), int(branch))
        return self.make_filename(prefix, int(number), 0)

    def execute_marking(self):
        """証拠番号を付与して保存"""
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "PDFファイルを読み込んでください")
            return

        # 確認ダイアログ
        prefix = self.get_prefix()
        style = self.current_print_style()
        evidence_list = self.generate_evidence_numbers()

        message = f"{len(evidence_list)}個のPDFファイルに証拠番号を付与します。よろしいですか?\n\n"
        message += "最初の5件:\n"
        for item, main, branch in evidence_list[:5]:
            disp = self.format_evidence_number(prefix, main, branch, style)
            message += f"{disp}: {item.file_path.name}\n"

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

        # prefix・style は確認ダイアログ表示時に取得済み。
        # 印字テキストは選択中の書式（甲001 / 甲第1号証）で生成し、
        # ファイル名は書式に関わらず常に甲001形式（make_filename）とする。

        # 処理実行（1ファイル単位でエラーを捕捉し、1件失敗しても残りは継続する）
        errors = []
        success = 0
        for item, main, branch in evidence_list:
            output_file = None
            try:
                stamp_text = self.format_evidence_number(prefix, main, branch, style)
                filename = self.make_filename(prefix, main, branch)
                output_file = output_path / filename

                # withで開き、エラー時もファイルハンドルを確実に解放する
                with fitz.open(str(item.file_path)) as pdf_doc:
                    # 回転保留がある場合、刻印より先に全ページへ一律で焼き込む。
                    # （横向き資料の向き補正。全ページ一律のため、1ファイル内に
                    #   縦横が混在する証拠には使わない運用とする。）
                    # 刻印ヘルパーは保存時の pdf_doc[0].rotation を読んで
                    # 「表示上の右上」に正立配置するため、回転後でも番号位置は正しい。
                    if item.rotation:
                        for pg in pdf_doc:
                            pg.set_rotation((pg.rotation + item.rotation) % 360)
                    if do_print:
                        # ページの向きは変えず、表示上の右上へ正立・横書きで刻印する。
                        # （裁判所の運用：横向きはそのままアップロード／番号は右上。
                        #   ヘルパーが /Rotate 0/90/180/270 を問わず一貫して
                        #   表示上の右上に配置する。）
                        stamp_evidence_number(
                            pdf_doc[0],
                            stamp_text,
                            font_size,
                            font_color,
                        )
                        # 埋め込みフォント（LTL Evidence Sans）を実使用グリフのみへ
                        # 再サブセットし、出力サイズの増分を数KBに抑える
                        subset_embedded_fonts(pdf_doc)
                    pdf_doc.save(str(output_file))

                # 押印の実在チェック（印字ありの場合のみ）。
                # 出力PDFを再オープンし、刻印した証拠番号が「表示上の右上」
                # （刻印予定位置）に実在するかを位置ベースで検証する。
                # 検出できなければ不良出力を削除し、当該ファイルをエラー扱いに
                # して残りの処理は継続する。
                # 単純な存在チェックではなく位置検証である理由：
                # ・番号を打ち直したいユーザーが、既に番号付きのファイルを
                #   再度読み込むのは自然な操作であり、本文・既存刻印に同一
                #   文字列があっても誤判定しないため
                # ・位置・向きがずれた刻印（回転メタデータ起因の不良）も
                #   実行時に必ず捕捉するため
                if do_print:
                    with fitz.open(str(output_file)) as verify_doc:
                        if not verify_stamp(verify_doc[0], stamp_text, font_size):
                            raise RuntimeError(
                                f"刻印した証拠番号「{stamp_text}」が"
                                f"出力PDFの右上から検出できませんでした"
                                f"（押印に失敗した可能性があります）"
                            )

                success += 1
            except Exception as e:
                # 検証に失敗した不良出力（中身が伴わないファイル）を残さない
                if do_print and output_file is not None and output_file.exists():
                    try:
                        output_file.unlink()
                    except OSError:
                        pass
                errors.append(f"・{item.file_path.name}：{e}")

        # 結果メッセージ
        action_label = "証拠番号を付与" if do_print else "ファイル名を変更（印字なし）"
        if errors:
            detail = "\n".join(errors[:10])
            if len(errors) > 10:
                detail += f"\n... 他 {len(errors) - 10}件"
            QMessageBox.warning(
                self, "一部のファイルを処理できませんでした",
                f"{success}個のPDFファイルに{action_label}しました。\n"
                f"{len(errors)}個のファイルでエラーが発生しました：\n\n{detail}\n\n"
                f"出力先：{output_path}"
            )
        else:
            QMessageBox.information(
                self, "完了",
                f"{len(evidence_list)}個のPDFファイルに{action_label}しました\n\n"
                f"出力先：{output_path}"
            )

        # 出力フォルダを開く（1件でも成功していれば開く）
        if success:
            if sys.platform == 'win32':
                os.startfile(output_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_path}"')
            else:
                os.system(f'xdg-open "{output_path}"')


def main():
    app = QApplication(sys.argv)
    window = EvidenceMarkerWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
