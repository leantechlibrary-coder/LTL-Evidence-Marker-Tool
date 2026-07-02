"""
証拠番号付与ツール — MCPサーバ（内蔵）
=====================================

GUI（pdf_evidence_marker.py）と同じ決定論コア（stamp_core.py）を共有し、
証拠番号刻印の決定論部分を MCP ツールとして公開する。サーバはLLMを呼ばず、
ネットワークにも出ない。原本フォルダは不変、出力は隣の「(元フォルダ名)_番号付」へ。

公開ツール（1本）:
  stamp_execute : 計画（ドライラン相当）を作り、warnings が無ければそのまま各PDF1枚目
                  右上に号証を刻印して出力する（plan/execute統合）。dry_run=True で
                  常に計画のみ、force=True で warnings があっても刻印を強制。

採番は rename 済みファイル名から読むか、numbers で明示割当（枝番可）。横向き補正は
rotations。GUI側の対話採番（ドラッグ順・甲乙選択・枝番トグル）とは入力経路が違うが、
出力フォーマット（甲001 / 甲001-2）は stamp_core に集約され一致する。
"""
from __future__ import annotations

import os
from pathlib import Path

import stamp_core as st

from mcp.server.fastmcp import FastMCP

_INSTRUCTIONS = (
    "証拠番号刻印（決定論）。stamp_execute は計画を作り、warningsが無ければそのまま刻印まで実行する"
    "（plan/executeを統合）。dry_run=true で計画だけ確認（絶対に刻印しない）。"
    "warningsが出た場合は内容を確認し、問題なければ force=true を付けて再度呼び出すと刻印する。"
    "採番は rename 済みファイル名／numbers の明示割当（枝番可）。原本は不変、"
    "出力は隣の「(元フォルダ名)_番号付」へ。位置検証を通った刻印だけ残す。"
)

mcp = FastMCP("ltl-stamp", instructions=_INSTRUCTIONS)

_FONT_COLORS = {"赤": (1, 0, 0), "黒": (0, 0, 0), "青": (0, 0, 1)}


# LLMが崩しがちなパス表記の吸収用（全角スラッシュ→半角。ドライブ文字脱落は復元不可）
_PATH_FIX = str.maketrans({"／": "/", "＼": "\\"})


def _clean_path(path: str) -> str:
    """小型モデルが渡すパスの表記ゆれを吸収する（原本は読むだけ・書き換えない）。
    前後の空白/全角空白/引用符を除去、file:// を剥がし、全角スラッシュや
    バックスラッシュを半角スラッシュへ寄せる。さらに各パス成分の前後空白を除去
    （小型モデルが「.../user/ Documents/...」のように区切り直後へ挿入する余計な
    空白を救済。成分内部の空白＝「My Documents」等は保持）。
    C: の脱落など情報欠落は復元不可。"""
    s = str(path).translate(_PATH_FIX).strip()
    for q in ('"', "'"):
        if len(s) >= 2 and s[0] == q and s[-1] == q:
            s = s[1:-1].strip()
    low = s.lower()
    if low.startswith("file:///"):
        s = s[8:]
    elif low.startswith("file://"):
        s = s[7:]
    s = s.replace("\\", "/")
    return "/".join(part.strip() for part in s.split("/"))


def _abs(path: str) -> str:
    return os.path.abspath(os.path.expanduser(_clean_path(path)))


def _check_dir(folder: str):
    p = _abs(folder)
    return (p, None) if os.path.isdir(p) else (None, f"フォルダが見つかりません: {p}")


@mcp.tool()
def stamp_execute(folder: str, style: str = "num3",
                  prefix_fallback: str = "甲", start_fallback: int = 1,
                  font_size: int = 16, font_color: str = "赤",
                  do_print: bool = True, rotations: dict | None = None,
                  numbers: dict | None = None,
                  dry_run: bool = False, force: bool = False) -> dict:
    """刻印計画を作り、warnings が無ければそのまま各PDFの1ページ目右上に号証を刻印して
    隣の「(元フォルダ名)_番号付」フォルダへ出力する（plan/executeを統合）。
    採番の優先順は「numbers（明示）> ファイル名 > 自動」。
    numbers は書面（準備書面／証拠説明書）から読んだ号証をファイル名キーで明示割当する辞書
    （例 {"結合スキャン_001.pdf":"甲1の1"}）。枝番（甲1の1）も指定可。
    rotations は横向きスキャンの上向き補正（キーはファイル名でも号証でも可、値0/90/180/270）。
    dry_run=True は計画だけ返し、絶対に刻印しない（executed=false）。
    warnings（自動採番／明示不正／出力名衝突／回転・明示の該当なし）が出た場合、
    dry_run=False でも force=True を付けない限り刻印しない（executed=false で計画のみ返す）。
    刻印は位置検証を通ったものだけ残し、失敗は不良出力を消して failed に記録。
    do_print=False は刻印せずリネームのみ。font_color: 赤/黒/青。"""
    d, err = _check_dir(folder)
    if err:
        return {"error": err}
    plan = st.plan_stamps(d, style, prefix_fallback, start_fallback, rotations, numbers)
    if not plan["plan"]:
        return {"error": f"PDFがありません: {d}", "warnings": plan["warnings"]}

    if dry_run or (plan["warnings"] and not force):
        return {"executed": False, "plan": plan["plan"], "warnings": plan["warnings"]}

    src = Path(d)
    base = src.name + "_番号付"
    out_dir = src.parent / base
    n = 2
    while out_dir.exists():
        out_dir = src.parent / f"{base}_{n}"
        n += 1
    out_dir.mkdir(parents=True)

    color = _FONT_COLORS.get(font_color, (1, 0, 0))
    files, failed = [], []
    for e in plan["plan"]:
        dst = out_dir / e["out_name"]
        try:
            st.stamp_one(e["src_path"], str(dst), e["evidence_number"],
                         font_size, color, e["rotate"], do_print)
            files.append(e["out_name"])
        except Exception as ex:  # noqa: BLE001 - 1件失敗で全体を止めない
            failed.append({"src": e["src_name"], "error": str(ex)})

    return {"executed": True, "output_dir": str(out_dir), "succeeded": len(files),
            "files": files, "failed": failed, "warnings": plan["warnings"]}


def run():
    """MCPサーバを stdio で起動する（entry.py --mcp から呼ばれる）。"""
    mcp.run()


if __name__ == "__main__":
    run()
