"""合成フィクスチャ。PyQt は使わず stamp_core / mcp_server だけを叩く。

リポジトリ直下の stamp_core.py / mcp_server.py を import できるよう、親をパスに載せる。
"""
import os
import sys

import pytest
import fitz

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _make_pdf(path, pages=1, text=None):
    doc = fitz.open()
    for i in range(pages):
        pg = doc.new_page(width=595, height=842)
        if text:
            pg.insert_text((72, 120), f"{text} {i + 1}", fontsize=20)
    doc.save(str(path))
    doc.close()


def _make_rotated_pdf(path, rotate, width=842, height=595):
    doc = fitz.open()
    pg = doc.new_page(width=width, height=height)
    pg.set_rotation(rotate)
    doc.save(str(path))
    doc.close()


@pytest.fixture
def split_like_dir(tmp_path):
    """split_execute 直後を模した無名断片（号証なし・1ページ）。"""
    d = tmp_path / "split_like"
    d.mkdir()
    for i, t in enumerate(["6月診断書", "7月診断書", "契約書"], 1):
        _make_pdf(d / f"結合スキャン_{i:03d}.pdf", 1, text=t)
    return d


@pytest.fixture
def stamp_dir(tmp_path):
    """rename済み名・自動採番対象・/Rotate付き横向きを含むフォルダ。"""
    d = tmp_path / "stamp_src"
    d.mkdir()
    _make_pdf(d / "甲001 請求書.pdf", 1, text="請求書")
    _make_pdf(d / "乙A001 契約書.pdf", 1, text="契約書")
    _make_pdf(d / "メモ.pdf", 1, text="メモ")
    _make_rotated_pdf(d / "甲002 図面.pdf", 270)
    return d
