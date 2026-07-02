"""内蔵MCP（stamp）の契約＋スモークテスト。PyQt 非依存。"""
import asyncio
import os

import fitz

import stamp_core as st
import mcp_server as M


# ---- 正規形フォーマット契約（GUI採番・rename アプリとの越境相互運用の要）----

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


# ---- MCP ツール登録（1本：plan/executeを統合）----

def test_mcp_registers_one_stamp_tool():
    names = sorted(t.name for t in asyncio.run(M.mcp.list_tools()))
    assert names == ["stamp_execute"]


# ---- パス吸収（小型モデルが崩す表記の救済）----

def test_clean_path_normalizes_llm_mess():
    c = M._clean_path
    assert c("C:\\Users\\user\\a") == "C:/Users/user/a"       # バックスラッシュ
    assert c('  "C:/Users/user/a"  ') == "C:/Users/user/a"    # 引用符＋前後空白
    assert c("'C:/Users/user/a'") == "C:/Users/user/a"        # 単一引用符
    assert c("C:／Users／user") == "C:/Users/user"           # 全角スラッシュ
    assert c("　C:/Users/user　") == "C:/Users/user"          # 全角空白
    assert c("file:///C:/Users/user/a") == "C:/Users/user/a"  # file:///
    assert c("file://C:/Users/user/a") == "C:/Users/user/a"   # file://
    # 区切り直後の余計な空白（gemmaが挿入）を成分ごとに除去。内部空白は保持。
    assert c("C:/Users/user/ Documents/a.pdf") == "C:/Users/user/Documents/a.pdf"
    assert c("C:/Users/user /Documents/a.pdf") == "C:/Users/user/Documents/a.pdf"
    assert c("C:/My Reports/2024 evidence/a.pdf") == "C:/My Reports/2024 evidence/a.pdf"


def test_check_dir_accepts_messy_folder(split_like_dir):
    messy = '"' + str(split_like_dir).replace("/", "\\") + '"'  # バックスラッシュ＋引用符
    res = M.stamp_execute(messy, numbers={
        "結合スキャン_001.pdf": "甲1", "結合スキャン_002.pdf": "甲2",
        "結合スキャン_003.pdf": "甲3",
    }, dry_run=True)
    assert res["executed"] is False
    assert len(res["plan"]) == 3


# ---- dry_run: 計画のみ返し、書き込みしない ----

def test_dry_run_returns_plan_without_writing(split_like_dir):
    res = M.stamp_execute(str(split_like_dir), numbers={
        "結合スキャン_001.pdf": "甲1の1",
        "結合スキャン_002.pdf": "甲1の2",
        "結合スキャン_003.pdf": "甲2",
    }, dry_run=True)
    assert res["executed"] is False
    assert "output_dir" not in res
    by_src = {e["src_name"]: e for e in res["plan"]}
    assert by_src["結合スキャン_001.pdf"]["out_name"] == "甲001-1.pdf"
    assert by_src["結合スキャン_001.pdf"]["source"] == "explicit"
    assert by_src["結合スキャン_003.pdf"]["evidence_number"] == "甲002"
    assert len(list(split_like_dir.glob("*.pdf"))) == 3  # 原本不変（書き込みなし）


# ---- warnings があれば force なしでは実行しない ----

def test_warnings_block_execution_without_force(stamp_dir):
    res = M.stamp_execute(str(stamp_dir), start_fallback=900)
    assert res["executed"] is False
    assert res["warnings"]  # メモ.pdf が自動採番で warning
    assert "output_dir" not in res
    assert len(list(stamp_dir.glob("*.pdf"))) == 4  # 原本不変


# ---- warnings が無ければ force なしでも自動実行 ----

def test_auto_executes_when_no_warnings(split_like_dir):
    res = M.stamp_execute(str(split_like_dir), numbers={
        "結合スキャン_001.pdf": "甲1の1",
        "結合スキャン_002.pdf": "甲1の2",
        "結合スキャン_003.pdf": "甲2",
    })
    assert res["executed"] is True
    assert res["succeeded"] == 3
    assert res["warnings"] == []


# ---- MCP境界の安全性（引数サニタイズ＋出力先の封じ込め）----

def test_prefix_fallback_rejects_path_like_values(split_like_dir):
    # prefix_fallback は出力名に入るため、パス風・自由文字列は入口で拒否する
    for bad in ("../../evil/x", "..\\..\\x", "甲/", "note", ""):
        res = M.stamp_execute(str(split_like_dir), prefix_fallback=bad, dry_run=True)
        assert "error" in res, bad


def test_prefix_fallback_accepts_fullwidth_and_lowercase(split_like_dir):
    # 全角・小文字の記号（乙ａ）は正規化して受理（乙A001.pdf になる）
    res = M.stamp_execute(str(split_like_dir), prefix_fallback="乙ａ", dry_run=True)
    assert "error" not in res
    assert res["plan"][0]["out_name"].startswith("乙A")


def test_output_containment_guard_blocks_escape(split_like_dir, monkeypatch):
    # 万一 out_name にパス区切りが混入しても、出力フォルダ外には書かない（最終ガード）
    crafted = {"plan": [{
        "src_name": "結合スキャン_001.pdf",
        "src_path": str(split_like_dir / "結合スキャン_001.pdf"),
        "evidence_number": "甲001", "out_name": "../escape.pdf",
        "source": "explicit", "rotate": 0,
    }], "counts": {}, "warnings": []}
    monkeypatch.setattr(M.st, "plan_stamps", lambda *a, **k: crafted)
    res = M.stamp_execute(str(split_like_dir))
    assert res["executed"] is True
    assert res["succeeded"] == 0
    assert res["failed"] and "拒否" in res["failed"][0]["error"]
    assert not (split_like_dir.parent / "escape.pdf").exists()


# ---- 刻印→検証の往復（/Rotate 付きページ含む）。force=True で warnings があっても実行 ----

def test_execute_stamps_and_verifies(stamp_dir):
    res = M.stamp_execute(str(stamp_dir), start_fallback=900, force=True)
    assert res["executed"] is True
    assert res["failed"] == []
    assert res["succeeded"] == 4
    assert "stamp_src_番号付" in os.path.basename(res["output_dir"])
    # /Rotate 270 付きページにも右上正立刻印が乗り、位置検証を通る
    with fitz.open(os.path.join(res["output_dir"], "甲002.pdf")) as doc:
        assert st.verify_stamp(doc[0], "甲002", 16)
    # 原本フォルダ不変
    assert len(list(stamp_dir.glob("*.pdf"))) == 4
