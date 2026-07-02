"""
LTL Evidence — 証拠番号刻印コア（PyQt非依存）
==============================================

証拠番号付与ツール（GUI）から、刻印に関わる決定論的部分だけを取り出したもの。
GUIの枝番トグル・サムネイル・回転UIは「番号を決めるための入力装置」であって
ロジックではないため、ここには持ち込まない。号証番号は rename 済みファイル名から
読む（schedule_number_files と同じ発想）。本コアは fitz のみに依存し、ネットワーク
にも LLM にも出ない。原本は読むだけで、書き込みは出力フォルダの新ファイルに限る。

公開関数:
  stamp_evidence_number(page, text, font_size, font_color, ...) … 表示上の右上へ正立刻印
  verify_stamp(page, stamp_text, font_size, ...)               … 刻印の実在を位置検証
  parse_mints_number(name)                                     … ファイル名→(prefix, main, branch)
  plan_stamps(folder, ...)                                     … フォルダ全PDFの刻印計画（読み取り専用）
  stamp_one(src, dst, stamp_text, ...)                         … 1ファイルを刻印して保存＋検証

刻印2関数（stamp_evidence_number / verify_stamp）は GUI 版から無改変で移植している。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple, Optional

import fitz  # PyMuPDF


# ---------------------------------------------------------------------------
# 刻印本体（GUI から無改変で移植。マージン既定値は両関数で一致させること）
# ---------------------------------------------------------------------------

def stamp_evidence_number(page, text, font_size, font_color,
                          right_margin: float = 25, top_margin: float = 14):
    """証拠番号をページ「表示上」の右上に、正立・横書きで刻印する。

    insert_text は座標点を「回転前(mediabox)座標系」で解釈し、rotate 引数はその
    座標系に対する文字の向きを指定する。/Rotate を持つページ（横長mediabox +
    /Rotate 270 で縦に立てたスキャナPDF等）に表示上の右上座標をそのまま渡すと
    位置も向きもずれるため、表示座標で目標点を求め、derotation_matrix で回転前
    座標へ変換した点を挿入位置とし、rotate=page.rotation で表示上に正立させる。
    ベースライン y はフォントサイズに連動させ、大きい文字でも上端で切れないように
    する（top_margin=14・font_size=16 のとき従来どおり 30）。
    """
    tw = fitz.get_text_length(text, fontname="japan", fontsize=font_size)
    baseline_y = top_margin + font_size
    disp_point = fitz.Point(page.rect.width - right_margin - tw, baseline_y)
    insert_point = disp_point * page.derotation_matrix
    page.insert_text(
        insert_point,
        text,
        fontsize=font_size,
        color=font_color,
        fontname="japan",
        rotate=page.rotation,
    )


def verify_stamp(page, stamp_text, font_size,
                 right_margin: float = 25, top_margin: float = 14) -> bool:
    """刻印後のページについて「表示上の右上」に証拠番号が実在するかを検証する。

    存在チェックのみだと、(a) 本文に同一文字列があると失敗を見逃す偽陽性、
    (b) 位置・向きがずれた刻印の見逃し、の穴がある。本実装は stamp_evidence_number
    と同じ計算で「表示座標系の刻印予定矩形」を求め、search_for の各ヒット（回転前
    座標）を rotation_matrix で表示座標へ写像し、予定位置と重なるヒットが1つでも
    あれば合格とする。同番号の打ち直し（新旧が重なる）は右上に番号がある事実で合格、
    本文中の同一文字列だけでは不合格、位置ズレ刻印は不合格として検出できる。
    """
    hits = page.search_for(stamp_text)
    if not hits:
        return False
    tw = fitz.get_text_length(stamp_text, fontname="japan", fontsize=font_size)
    expected = fitz.Rect(
        page.rect.width - right_margin - tw, top_margin,
        page.rect.width - right_margin, top_margin + font_size
    ) + (-5, -5, 5, 5)
    for h in hits:
        disp = h * page.rotation_matrix  # 回転前座標 → 表示座標
        disp.normalize()
        if expected.intersects(disp):
            return True
    return False


# ---------------------------------------------------------------------------
# 採番（ファイル名から読む）
#
# rename（_filename/core.py の parse_kosho）とトークン領域で一致するよう書く。
# 両者は別 core（cross-core import はしない規約）なので物理的には別コピーだが、
# 同じ規則（全角英数字の半角化・枝番区切りゆらぎ・範囲表記の不採番・記号の大文字化）
# を写経し、一致は test_stamp.py のドリフト検出テストで担保する。
# ---------------------------------------------------------------------------

# 全角英数字 → 半角（種別漢字・かな・約物は保持）。parse_kosho の normalize_alnum と同規則。
_FW_ALNUM: dict[int, int] = {}
for _i in range(10):
    _FW_ALNUM[ord("０") + _i] = ord("0") + _i
for _i in range(26):
    _FW_ALNUM[ord("Ａ") + _i] = ord("A") + _i
    _FW_ALNUM[ord("ａ") + _i] = ord("a") + _i

# 枝番区切りのゆらぎ（ハイフン各種 / 長音 / の）と範囲マーク。parse_kosho と同一集合。
_BRANCH_SEP = "-ー‐−–—－の"
_RANGE_MARK = ("~", "〜", "～")
# 範囲表記の判定: 先頭号証トークンの直後（空白のみ挟んで）に範囲マークが続くか。
# 標目内のチルダ（甲005 …令和5年～6年…）はトークンの後ろに別テキストを挟むので拾わない。
_RANGE_AFTER_TOKEN = re.compile(r"^\s*[" + re.escape("".join(_RANGE_MARK)) + r"]")
# 種別 + 記号(A-Z) + (第) + 本番号 + (号証) + (枝番区切り + 枝番)
_MINTS_RE = re.compile(
    r"^\s*([甲乙丙丁])\s*([A-Za-z]?)\s*(?:第)?\s*0*([0-9]+)\s*(?:号証)?"
    r"(?:\s*[" + _BRANCH_SEP + r"]\s*0*([0-9]+))?"
)


def parse_mints_number(name: str) -> Optional[Tuple[str, int, int]]:
    """ファイル名（または stem）の先頭から mints 号証を読む。

    戻り値は (prefix, main, branch)。prefix は当事者記号込みで大文字化（"甲" / "乙A" 等）、
    branch は 0 が枝番なし、1以上が枝番。読めなければ None。範囲表記（甲001-1~2）は
    採番しない（None＝plan 側で自動採番＋[要確認]へ回す）。範囲判定は先頭号証トークンの
    直後の範囲マークに限定するので、標目内のチルダ（甲005 …令和5年～6年…）は弾かれない。

    rename の parse_kosho（_filename/core.py）とトークン領域で一致させてあり、
    一致は test_stamp.py のドリフト検出テストで担保する（cross-core import はしない）。
    """
    norm = str(Path(name).stem).translate(_FW_ALNUM)
    m = _MINTS_RE.match(norm)
    if not m:
        return None
    # 先頭号証トークンの直後に範囲マークが続く＝範囲表記 → 採番しない（None）
    if _RANGE_AFTER_TOKEN.match(norm[m.end():]):
        return None
    prefix = m.group(1) + (m.group(2) or "").upper()
    main = int(m.group(3))
    branch = int(m.group(4)) if m.group(4) else 0
    return prefix, main, branch


def format_stamp_text(prefix: str, main: int, branch: int, style: str = "num3") -> str:
    """採番情報を刻印書式の文字列へ。num3=甲001 / 甲001-2、gou=甲第1号証 / 甲第1号証の2。"""
    if style == "num3":
        base = f"{prefix}{main:03d}"
        return f"{base}-{branch}" if branch else base
    base = f"{prefix}第{main}号証"
    return f"{base}の{branch}" if branch else base


def make_filename(prefix: str, main: int, branch: int) -> str:
    """出力ファイル名（mints命名規則：常に半角三桁・枝番ハイフン。刻印書式に依存しない）。"""
    base = f"{prefix}{main:03d}"
    stem = f"{base}-{branch}" if branch else base
    return f"{stem}.pdf"


def _normalize_rotation(deg) -> Optional[int]:
    """回転指定を 0/90/180/270 へ正規化する。90の倍数でなければ None（不正）。

    負値・360超も許容し剰余で丸める（-90→270、450→90）。set_rotation は
    90の倍数しか受け付けないため、ここで弾いて plan の warning に回す。
    """
    try:
        d = int(deg) % 360
    except (TypeError, ValueError):
        return None
    return d if d in (0, 90, 180, 270) else None


def plan_stamps(folder: str, style: str = "num3",
                prefix_fallback: str = "甲", start_fallback: int = 1,
                rotations: Optional[dict] = None,
                numbers: Optional[dict] = None) -> dict:
    """フォルダ直下の全PDFについて、刻印テキスト・出力名・回転量を決める（書き込みなし）。

    採番の優先順は「numbers（明示）> ファイル名 > 自動」:
      - numbers はファイル名キー → 号証文字列の辞書（例 {"結合スキャン_001.pdf":"甲1の1"}）。
        書面（準備書面／証拠説明書）から読んだ号証を明示割当する窓口で、枝番（甲1の1）も
        指定できる。値は parse_mints_number で解釈する（号証パーサを一本化）。source="explicit"。
      - 明示が無いPDFはファイル名から号証を読む（source="filename"＝文書由来＝grounded）。
      - どちらも読めないPDFは、ソート順で prefix_fallback + start_fallback から連番を振り、
        source="auto" として provenance 上 [要確認] を付す（号証の根拠が文書にない）。

    rotations は横向きスキャンの上向き補正（任意）。キーは元ファイル名でも号証
    （"甲003"）でもよく、両方で引いて該当した方を採用する。値は 0/90/180/270。
    各 plan エントリに解決後の rotate を載せ、execute はこの単一窓口の結果を使う
    （採番・回転の決定を1か所に集約。GUIの generate_evidence_numbers と同じ役割）。

    warnings には「明示号証が読めず自動採番」「号証を読めず自動採番」「出力名が衝突」
    「回転指定が90の倍数でない」「回転/明示指定が該当ファイルなし（キー打ち間違い）」を入れる。
    """
    folder_p = Path(folder)
    pdfs = sorted(folder_p.glob("*.pdf"))
    plan: List[dict] = []
    warnings: List[str] = []
    seen: dict = {}

    rotations = dict(rotations or {})
    used_rot_keys: set = set()
    numbers = dict(numbers or {})
    used_num_keys: set = set()

    auto_main = start_fallback
    n_explicit = 0
    n_filename = 0
    n_auto = 0

    for p in pdfs:
        if p.name in numbers:
            used_num_keys.add(p.name)
            parsed = parse_mints_number(numbers[p.name])
            if parsed:
                prefix, main, branch = parsed
                source = "explicit"
                n_explicit += 1
            else:
                warnings.append(f"明示号証が読めません（自動採番 [要確認]）: {p.name}={numbers[p.name]}")
                prefix, main, branch = prefix_fallback, auto_main, 0
                auto_main += 1
                source = "auto"
                n_auto += 1
        elif parse_mints_number(p.name):
            prefix, main, branch = parse_mints_number(p.name)
            source = "filename"
            n_filename += 1
        else:
            prefix, main, branch = prefix_fallback, auto_main, 0
            auto_main += 1
            source = "auto"
            n_auto += 1
            warnings.append(f"号証を読めませんでした（自動採番 [要確認]）: {p.name}")

        stamp_text = format_stamp_text(prefix, main, branch, style)
        out_name = make_filename(prefix, main, branch)
        if out_name in seen:
            warnings.append(f"出力名が衝突: {out_name}（{seen[out_name]} と {p.name}）")
        else:
            seen[out_name] = p.name

        # 回転指定はファイル名・号証どちらのキーでも引く
        rot_raw = None
        for key in (p.name, stamp_text):
            if key in rotations:
                rot_raw = rotations[key]
                used_rot_keys.add(key)
                break
        rotate = 0
        if rot_raw is not None:
            norm = _normalize_rotation(rot_raw)
            if norm is None:
                warnings.append(f"回転指定が90の倍数でないため無視: {p.name}={rot_raw}")
            else:
                rotate = norm

        entry = {
            "src_name": p.name,
            "src_path": str(p),
            "evidence_number": stamp_text,
            "out_name": out_name,
            "source": source,
            "rotate": rotate,
        }
        if source == "auto":
            entry["flag"] = "[要確認]"
        plan.append(entry)

    # 該当ファイルが無かった回転キー（打ち間違いの検出）
    for key in rotations:
        if key not in used_rot_keys:
            warnings.append(f"回転指定が該当ファイルなし（キー要確認）: {key}")

    # 該当ファイルが無かった明示号証キー（打ち間違いの検出）
    for key in numbers:
        if key not in used_num_keys:
            warnings.append(f"明示号証の指定が該当ファイルなし（キー要確認）: {key}")

    return {
        "plan": plan,
        "counts": {"total": len(pdfs), "explicit": n_explicit,
                   "from_filename": n_filename, "auto": n_auto},
        "warnings": warnings,
    }


# ---------------------------------------------------------------------------
# 刻印実行（1ファイル単位。原本は読むだけ・出力は新ファイル）
# ---------------------------------------------------------------------------

def stamp_one(src: str, dst: str, stamp_text: str,
              font_size: int = 16, font_color: Tuple[float, float, float] = (1, 0, 0),
              rotate: int = 0, do_print: bool = True) -> None:
    """src の1ページ目に stamp_text を刻印して dst へ保存し、刻印を位置検証する。

    rotate（0/90/180/270）が指定されれば、刻印より先に全ページへ一律で焼き込む
    （横向き資料の上向き補正。1ファイル内で縦横混在する証拠には使わない運用）。
    do_print=False なら刻印せず保存のみ（リネーム相当）。検証に失敗したら dst を
    消して RuntimeError を投げる（不良出力を残さない）。
    """
    with fitz.open(src) as doc:
        if rotate:
            for pg in doc:
                pg.set_rotation((pg.rotation + rotate) % 360)
        if do_print:
            stamp_evidence_number(doc[0], stamp_text, font_size, font_color)
        doc.save(dst)

    if do_print:
        with fitz.open(dst) as vdoc:
            if not verify_stamp(vdoc[0], stamp_text, font_size):
                Path(dst).unlink(missing_ok=True)
                raise RuntimeError(
                    f"刻印した証拠番号「{stamp_text}」が出力PDFの右上から"
                    f"検出できませんでした（押印に失敗した可能性があります）"
                )
