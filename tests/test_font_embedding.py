"""刻印フォント（LTL Evidence Sans）の埋め込みとフォールバックのテスト。

観点:
  1. 同梱フォントで刻印した出力PDFに LTL Evidence Sans がサブセット埋め込みされる
  2. 埋め込み後もテキスト抽出が効き、verify_stamp（位置検証）が通る
  3. フォントファイルが無い環境（MCP同梱パッケージ想定）では "japan" へフォールバックし、
     刻印・検証とも従来どおり成立する
  4. サブセット外グリフを含むテキスト（カスタム種別等）もフォールバックで刻印できる
  5. /Rotate 付きページでも埋め込みフォントで位置検証が通る
"""
import fitz
import pytest

import stamp_core as st


def _make_pdf(path, rotate=0):
    doc = fitz.open()
    pg = doc.new_page(width=595, height=842)
    if rotate:
        pg.set_rotation(rotate)
    doc.save(str(path))
    doc.close()


def _basefonts(path):
    with fitz.open(str(path)) as doc:
        return [f[3] for f in doc[0].get_fonts(full=True)]


@pytest.fixture(autouse=True)
def _fresh_font_cache():
    """各テストで lru_cache をクリアし、monkeypatch が確実に効くようにする。"""
    st._load_font.cache_clear()
    yield
    st._load_font.cache_clear()


def test_stamp_embeds_ltl_font(tmp_path):
    src, dst = tmp_path / "s.pdf", tmp_path / "甲001.pdf"
    _make_pdf(src)
    st.stamp_one(str(src), str(dst), "甲001")  # 検証失敗なら raise
    fonts = _basefonts(dst)
    assert any("LTLEvidenceSans" in bf or "LTL" in bf for bf in fonts), fonts


def test_embedded_stamp_is_searchable_and_verified(tmp_path):
    src, dst = tmp_path / "s.pdf", tmp_path / "乙A003-2.pdf"
    _make_pdf(src)
    st.stamp_one(str(src), str(dst), "乙A003-2")
    with fitz.open(str(dst)) as doc:
        assert doc[0].search_for("乙A003-2")
        assert st.verify_stamp(doc[0], "乙A003-2", 16)


def test_fallback_when_font_absent(tmp_path, monkeypatch):
    """MCP同梱パッケージ（フォント非同梱）を模す：探索が None → japan で従来動作。"""
    monkeypatch.setattr(st, "_find_font_file", lambda: None)
    st._load_font.cache_clear()
    name, file, font = st.resolve_stamp_font("甲001")
    assert (name, file, font) == ("japan", None, None)
    src, dst = tmp_path / "s.pdf", tmp_path / "甲001.pdf"
    _make_pdf(src)
    st.stamp_one(str(src), str(dst), "甲001")
    assert not any("LTL" in bf for bf in _basefonts(dst))


def test_fallback_on_uncovered_glyph(tmp_path):
    """サブセット外グリフ（ひらがな等のカスタム種別）はテキスト単位で japan へ。"""
    name, file, _ = st.resolve_stamp_font("あて001")
    assert name == "japan" and file is None
    # 収録グリフのみのテキストは LTL のまま
    name2, file2, _ = st.resolve_stamp_font("疎甲第10号証の2")
    assert name2 == st._FONT_PDFNAME and file2 is not None
    src, dst = tmp_path / "s.pdf", tmp_path / "あて001.pdf"
    _make_pdf(src)
    st.stamp_one(str(src), str(dst), "あて001")  # フォールバック刻印＋検証が通ること


def test_embedded_font_on_rotated_page(tmp_path):
    """/Rotate 270 ページでも埋め込みフォントで表示上右上に正立し、位置検証が通る。"""
    src, dst = tmp_path / "r.pdf", tmp_path / "甲005.pdf"
    doc = fitz.open()
    pg = doc.new_page(width=842, height=595)
    pg.set_rotation(270)
    doc.save(str(src))
    doc.close()
    st.stamp_one(str(src), str(dst), "甲005")
    assert any("LTL" in bf for bf in _basefonts(dst))


def test_width_consistency_between_stamp_and_verify():
    """刻印と検証が同じ幅計算を使うこと（予定矩形ズレの回帰防止）。"""
    w1 = st.stamp_text_length("甲001", 16)
    _, _, font = st.resolve_stamp_font("甲001")
    assert font is not None
    assert abs(w1 - font.text_length("甲001", fontsize=16)) < 1e-6


def test_subsetting_keeps_output_small_and_verifiable(tmp_path):
    """保存前の再サブセットで増分が数KBに収まり、検索・検証も維持される。"""
    src, dst = tmp_path / "s.pdf", tmp_path / "甲001.pdf"
    _make_pdf(src)
    st.stamp_one(str(src), str(dst), "甲001")
    size = dst.stat().st_size
    assert size < 30_000, f"出力が想定より大きい: {size}"  # フル埋め込み(≈83KB)なら失敗
    fonts = _basefonts(dst)
    assert any("LTL" in bf for bf in fonts), fonts
    with fitz.open(str(dst)) as doc:
        assert st.verify_stamp(doc[0], "甲001", 16)
