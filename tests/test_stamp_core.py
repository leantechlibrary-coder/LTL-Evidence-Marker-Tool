"""stamp_core の正規形フォーマット契約。PyQt 非依存。

GUI 採番・rename アプリとの相互運用の要となる、ファイル名・刻印文字列・番号解析の書式を固定する。
"""
import stamp_core as st


def test_canonical_filename_format():
    assert st.make_filename("甲", 1, 0) == "甲001.pdf"
    assert st.make_filename("甲", 1, 2) == "甲001-2.pdf"
    assert st.make_filename("乙A", 12, 0) == "乙A012.pdf"


def test_canonical_stamp_text_format():
    assert st.format_stamp_text("甲", 1, 0, "num3") == "甲001"
    assert st.format_stamp_text("甲", 1, 2, "num3") == "甲001-2"
    assert st.format_stamp_text("甲", 1, 2, "gou") == "甲第1号証の2"


def test_parser_roundtrip_and_range_guard():
    assert st.parse_mints_number("甲005 標目.pdf") == ("甲", 5, 0)
    assert st.parse_mints_number("乙A001-1 契約.pdf") == ("乙A", 1, 1)
    assert st.parse_mints_number("甲001の2 見積.pdf") == ("甲", 1, 2)
    assert st.parse_mints_number("甲001-1~2 一括.pdf") is None      # 範囲表記は採番しない
    assert st.parse_mints_number("甲005 取引明細（令和5年～6年）.pdf") == ("甲", 5, 0)  # 標目内チルダは弾かない
    assert st.parse_mints_number("メモ.pdf") is None
