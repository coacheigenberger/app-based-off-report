from __future__ import annotations

import io
from pathlib import Path
from typing import Sequence, Tuple

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor

from defense_core import frequency, pct, confidence

BLANK_BLITZ = {"-", "NONE", "NO", "NO BLITZ", "0", "BASE", "NO DATA", "UNKNOWN"}


def _fmt_pct(x) -> str:
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return str(x)


def _fmt_count_pct(count: int | float, total: int | float) -> str:
    """Coach-friendly percentage format used everywhere in the deck."""
    try:
        c = int(count)
    except Exception:
        c = 0
    try:
        t = int(total)
    except Exception:
        t = 0
    return f"{c}/{t} ({_fmt_pct(pct(c, t))})"


def _shape_text(shape) -> str:
    return getattr(shape, "text", "") if hasattr(shape, "text") else ""


def _tables(slide):
    return [s for s in slide.shapes if getattr(s, "has_table", False)]


def _text_shapes(slide):
    return [s for s in slide.shapes if hasattr(s, "text_frame")]


def _copy_font(src, dst) -> None:
    """Copy explicit font settings so data replacement does not restyle the template."""
    try:
        dst.name = src.name
        dst.size = src.size
        dst.bold = src.bold
        dst.italic = src.italic
        dst.underline = src.underline
        if src.color and src.color.type is not None:
            if src.color.rgb is not None:
                dst.color.rgb = src.color.rgb
    except Exception:
        pass


def _first_run_style(text_frame):
    """Return a run to use as the formatting source for replacement text."""
    for p in text_frame.paragraphs:
        for r in p.runs:
            return r
    return None


def _set_text_preserve_textframe(text_frame, value) -> None:
    """
    Replace visible text while preserving the template's layout and explicit font style.
    This intentionally does not touch cell fills, table dimensions, borders, margins,
    shape positions, or any other formatting.
    """
    value = "" if value is None else str(value)
    src_run = _first_run_style(text_frame)
    first_para = text_frame.paragraphs[0] if text_frame.paragraphs else None
    alignment = getattr(first_para, "alignment", None) if first_para is not None else None
    level = getattr(first_para, "level", None) if first_para is not None else None

    text_frame.clear()
    lines = value.split("\n")
    for i, line in enumerate(lines if lines else [""]):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        if alignment is not None:
            p.alignment = alignment
        if level is not None:
            p.level = level
        run = p.add_run()
        run.text = line
        if src_run is not None:
            _copy_font(src_run.font, run.font)


def _set_shape_text_preserve(shape, value) -> None:
    if hasattr(shape, "text_frame"):
        _set_text_preserve_textframe(shape.text_frame, value)


def _set_cell_text_preserve(cell, value) -> None:
    _set_text_preserve_textframe(cell.text_frame, value)


def _fill_table_preserve(table_shape, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    """
    Fill only text values in an existing table.
    Does not change fonts, fills, column widths, row heights, margins, borders, or positions.
    """
    tbl = table_shape.table
    max_rows = len(tbl.rows)
    max_cols = len(tbl.columns)

    for c in range(max_cols):
        _set_cell_text_preserve(tbl.cell(0, c), headers[c] if c < len(headers) else "")

    for r in range(1, max_rows):
        for c in range(max_cols):
            val = ""
            if r - 1 < len(rows) and c < len(rows[r - 1]):
                val = rows[r - 1][c]
            _set_cell_text_preserve(tbl.cell(r, c), val)


def _replace_opponent(slide, opponent: str) -> None:
    """Keep only the opponent line. Remove Date from every slide."""
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Opponent:" in txt:
            _set_shape_text_preserve(shape, f"Opponent: {opponent}")


def _write_takeaways(slide, lines: Sequence[str]) -> None:
    # Only update an existing Key Takeaways shape. Never add or format one.
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Key Takeaways" in txt:
            _set_shape_text_preserve(shape, "Key Takeaways\n" + "\n".join(lines[:5]))
            return


def _top_value(df: pd.DataFrame, col: str, denom: int | None = None) -> Tuple[str, float, int, int]:
    f = frequency(df, col, denom=denom)
    if f.empty:
        return "No data", 0.0, 0, int(denom or len(df))
    row = f.iloc[0]
    return str(row[col]), float(row["Pct"]), int(row["Plays"]), int(denom or len(df))


def _top_two_front_text(g: pd.DataFrame) -> str:
    total = len(g)
    fronts = frequency(g, "FRONT", denom=total).head(2)
    if fronts.empty:
        return "No data"
    top_pct = float(fronts.iloc[0]["Pct"])
    if top_pct >= 50 or len(fronts) == 1:
        return f'{fronts.iloc[0]["FRONT"]} {_fmt_count_pct(fronts.iloc[0]["Plays"], total)}'
    return "\n".join(
        f'{r["FRONT"]} {_fmt_count_pct(r["Plays"], total)}' for _, r in fronts.iterrows()
    )


def _top_name_and_stat(df: pd.DataFrame, col: str, denom: int | None = None) -> str:
    total = int(denom or len(df))
    f = frequency(df, col, denom=total).head(1)
    if f.empty:
        return "No data"
    r = f.iloc[0]
    return f'{r[col]} {_fmt_count_pct(r["Plays"], total)}'


def _top_combos_text(df: pd.DataFrame, n=5) -> str:
    total = len(df)
    f = frequency(df, "COMBO", denom=total).head(n)
    if f.empty:
        return "No data"
    return "\n".join(
        [f'{i+1}. {r["COMBO"]} - {_fmt_count_pct(r["Plays"], total)}' for i, r in f.iterrows()]
    )


def _third_down_bucket(df):
    d = df[(df["DOWN"] == 3)].copy()
    if d.empty:
        return {
            "3rd & 7+": pd.DataFrame(),
            "3rd & 3-6": pd.DataFrame(),
            "3rd & 1-2": pd.DataFrame(),
        }
    return {
        "3rd & 7+": d[d["DISTANCE"] >= 7],
        "3rd & 3-6": d[(d["DISTANCE"] >= 3) & (d["DISTANCE"] <= 6)],
        "3rd & 1-2": d[(d["DISTANCE"] >= 1) & (d["DISTANCE"] <= 2)],
    }


def _red_zone_bucket(df):
    y = df["YARD_LINE"]
    return {
        "High Red Zone (25-15)": df[(y >= 15) & (y <= 25)],
        "Red Zone (15-5)": df[(y >= 5) & (y < 15)],
        "Goal Line (Inside the 5)": df[(y >= 1) & (y < 5)],
    }


def _situational_rows_for_template(buckets: dict[str, pd.DataFrame], table_shape) -> tuple[list[str], list[list[object]]]:
    """
    If the template table has a middle column, populate Total Plays there.
    If someone uploads an older 2-column template, preserve compatibility.
    """
    cols = len(table_shape.table.columns)
    if cols >= 3:
        headers = ["Situation", "Total Plays", "Top 5 Front/Stunt/Blitz/Coverage Calls"]
        rows = [[name, len(g), _top_combos_text(g, 5)] for name, g in buckets.items()]
    else:
        headers = ["Situation", "Top 5 Front/Stunt/Blitz/Coverage Calls"]
        rows = [[name, _top_combos_text(g, 5)] for name, g in buckets.items()]
    return headers, rows




def _find_slide_by_title(prs, title_substring: str, fallback_index: int | None = None):
    want = title_substring.lower()
    for slide in prs.slides:
        joined = "\n".join(_shape_text(s) for s in _text_shapes(slide)).lower()
        if want in joined:
            return slide
    if fallback_index is not None and len(prs.slides) > fallback_index:
        return prs.slides[fallback_index]
    return None


def _result_color(result: str):
    r = (result or "").strip().upper()
    if r.startswith("W"):
        return RGBColor(0, 130, 0)
    if r.startswith("L"):
        return RGBColor(190, 0, 0)
    if r.startswith("T"):
        return RGBColor(120, 120, 120)
    return None


def _set_cell_text_preserve_color(cell, value, rgb=None) -> None:
    _set_text_preserve_textframe(cell.text_frame, value)
    if rgb is None:
        return
    try:
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = rgb
                r.font.bold = True
    except Exception:
        pass


def _fill_overview_slide(slide, opponent: str, overview: dict | None) -> None:
    if not overview:
        overview = {}
    _replace_opponent(slide, opponent)

    record = overview.get("record", "") or "No data"
    sacks = overview.get("sacks", "") or "No data"
    ints = overview.get("interceptions", "") or "No data"
    fumbles = overview.get("fumble_recoveries", "") or "No data"
    source = overview.get("source_url", "") or "GoBound"

    stat_rows = [
        ["Overall Record", record],
        ["Total Sacks", sacks],
        ["Total INTs", ints],
        ["Total Fumble Recoveries", fumbles],
        ["Last 5 Games", overview.get("last_five_record", "") or "No data"],
        ["Avg. Margin", overview.get("average_margin", "") or "No data"],
        ["One-Possession Games", overview.get("one_possession_games", "") or "No data"],
        ["Wins vs Winning Records", overview.get("wins_vs_winning_records", "") or "No data"],
    ]

    games = overview.get("games") or []
    game_rows = []
    for i, g in enumerate(games[:12], start=1):
        if hasattr(g, "__dict__"):
            g = g.__dict__
        game_rows.append([
            g.get("week") or str(i),
            g.get("opponent") or "",
            g.get("opponent_record") or "",
            g.get("score") or "",
            g.get("result") or "",
        ])

    tables = _tables(slide)
    if len(tables) >= 1:
        _fill_table_preserve(tables[0], ["Statistic", "Value"], stat_rows)
    if len(tables) >= 2:
        tbl_shape = tables[1]
        _fill_table_preserve(tbl_shape, ["Game", "Opponent", "Record", "Score", "W/L"], game_rows)
        # Intentional requested formatting: color code the result column only.
        tbl = tbl_shape.table
        result_col = min(4, len(tbl.columns) - 1)
        for r in range(1, len(tbl.rows)):
            val = tbl.cell(r, result_col).text
            _set_cell_text_preserve_color(tbl.cell(r, result_col), val, _result_color(val))

    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Source:" in txt:
            _set_shape_text_preserve(shape, f"Source: {source}")

def build_defense_pptx(engine, template_path: str | Path, opponent: str = "Opponent", overview: dict | None = None) -> bytes:
    prs = Presentation(str(template_path))
    df = engine.df
    total_plays = len(df)

    for slide in prs.slides:
        _replace_opponent(slide, opponent)

    # Slide: Opponent Overview
    overview_slide = _find_slide_by_title(prs, "Opponent Overview", None)
    if overview_slide is not None:
        _fill_overview_slide(overview_slide, opponent, overview)

    # Fronts
    slide = _find_slide_by_title(prs, "Front Tendencies", 0 if overview_slide is None else 1)
    if slide is not None:
        front_freq = frequency(df, "FRONT", denom=total_plays).head(5)
        rows = [[r["FRONT"], int(r["Plays"]), _fmt_count_pct(r["Plays"], total_plays)] for _, r in front_freq.iterrows()]
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Front", "Snaps", "Usage %"], rows)
        top_front, _, top_count, top_total = _top_value(df, "FRONT", denom=total_plays)
        top_third, _, third_count, third_total = _top_value(df[df["DOWN"] == 3], "FRONT")
        rz_parts = [g for g in _red_zone_bucket(df).values() if not g.empty]
        rz = pd.concat(rz_parts, ignore_index=True) if rz_parts else pd.DataFrame()
        rz_front, _, rz_count, rz_total = _top_value(rz, "FRONT")
        _write_takeaways(slide, [
            f"🔥 Base front: {top_front} {_fmt_count_pct(top_count, top_total)}",
            f"⚠️ 3rd down front: {top_third} {_fmt_count_pct(third_count, third_total)}",
            f"🛑 Red zone front: {rz_front} {_fmt_count_pct(rz_count, rz_total)}",
        ])

    # Blitzes
    slide = _find_slide_by_title(prs, "Blitz Tendencies", 1 if overview_slide is None else 2)
    if slide is not None:
        blitz_df = df[~df["BLITZ"].isin(BLANK_BLITZ)].copy()
        blitz_freq = frequency(blitz_df, "BLITZ", denom=total_plays).head(5)
        rows = []
        for _, r in blitz_freq.iterrows():
            b = r["BLITZ"]
            g = blitz_df[blitz_df["BLITZ"] == b]
            blitz_total = len(g)
            fronts = frequency(g, "FRONT", denom=blitz_total).head(1)
            top_front = f'{fronts.iloc[0]["FRONT"]}\n{_fmt_count_pct(fronts.iloc[0]["Plays"], blitz_total)}' if not fronts.empty else "No data"
            stunts = frequency(g, "STUNT", denom=blitz_total).head(1)
            top_stunt = f'{stunts.iloc[0]["STUNT"]}\n{_fmt_count_pct(stunts.iloc[0]["Plays"], blitz_total)}' if not stunts.empty else "No data"
            rows.append([b, int(r["Plays"]), _fmt_count_pct(r["Plays"], total_plays), top_front, top_stunt])
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Blitz", "Snaps", "Usage %", "Top Front", "Top Stunt"], rows)
        top_blitz, _, blitz_count, blitz_total = _top_value(blitz_df, "BLITZ", denom=len(blitz_df))
        third_blitz, _, third_count, third_total = _top_value(blitz_df[blitz_df["DOWN"] == 3], "BLITZ")
        _write_takeaways(slide, [
            f"🔥 Top blitz: {top_blitz} {_fmt_count_pct(blitz_count, blitz_total)} of blitz snaps",
            f"⚠️ 3rd down pressure: {third_blitz} {_fmt_count_pct(third_count, third_total)}",
            f"🛑 Overall blitz rate: {_fmt_count_pct(int(df['IS_BLITZ'].sum()), total_plays)}",
        ])

    # Coverages by down
    slide = _find_slide_by_title(prs, "Coverage Tendencies", 2 if overview_slide is None else 3)
    if slide is not None:
        cover_freq = frequency(df, "COVERAGE", denom=total_plays).head(5)
        rows = []
        for _, r in cover_freq.iterrows():
            cov = r["COVERAGE"]
            row = [cov, _fmt_count_pct(r["Plays"], total_plays)]
            for down in [1, 2, 3, 4]:
                gd = df[df["DOWN"] == down]
                cov_count = int((gd["COVERAGE"] == cov).sum()) if len(gd) else 0
                row.append(_fmt_count_pct(cov_count, len(gd)))
            rows.append(row)
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Coverage", "Overall %", "1st Down", "2nd Down", "3rd Down", "4th Down"], rows)
        top_cov, _, cov_count, cov_total = _top_value(df, "COVERAGE", denom=total_plays)
        man_count = int(df["IS_DISRESPECTFUL"].sum()) if len(df) else 0
        _write_takeaways(slide, [
            f"🔥 Base coverage: {top_cov} {_fmt_count_pct(cov_count, cov_total)}",
            f"⚠️ Man/press rate: {_fmt_count_pct(man_count, total_plays)}",
            "🛑 Cover 0/Cover 1/press = shot alert.",
        ])

    # 3rd down
    slide = _find_slide_by_title(prs, "3rd Down Tendencies", 3 if overview_slide is None else 4)
    if slide is not None:
        buckets = _third_down_bucket(df)
        tables = _tables(slide)
        if tables:
            headers, rows = _situational_rows_for_template(buckets, tables[0])
            _fill_table_preserve(tables[0], headers, rows)
        all3 = df[df["DOWN"] == 3]
        combo_freq = frequency(all3, "COMBO", denom=len(all3))
        top_combo = combo_freq.iloc[0]["COMBO"] if not combo_freq.empty else "No data"
        top_combo_count = int(combo_freq.iloc[0]["Plays"]) if not combo_freq.empty else 0
        blitz_count = int(all3["IS_BLITZ"].sum()) if len(all3) else 0
        _write_takeaways(slide, [
            f"🔥 Top 3rd down call: {top_combo} {_fmt_count_pct(top_combo_count, len(all3))}",
            f"⚠️ 3rd down blitz rate: {_fmt_count_pct(blitz_count, len(all3))}",
            f"🛑 3rd down sample: {len(all3)} snaps ({confidence(len(all3))} confidence)",
        ])

    # Red Zone
    slide = _find_slide_by_title(prs, "Red Zone Tendencies", 4 if overview_slide is None else 5)
    if slide is not None:
        buckets = _red_zone_bucket(df)
        tables = _tables(slide)
        if tables:
            headers, rows = _situational_rows_for_template(buckets, tables[0])
            _fill_table_preserve(tables[0], headers, rows)
        rz_parts = [g for g in buckets.values() if not g.empty]
        rz = pd.concat(rz_parts, ignore_index=True) if rz_parts else pd.DataFrame()
        combo_freq = frequency(rz, "COMBO", denom=len(rz)) if not rz.empty else pd.DataFrame()
        top_rz = combo_freq.iloc[0]["COMBO"] if not combo_freq.empty else "No data"
        top_rz_count = int(combo_freq.iloc[0]["Plays"]) if not combo_freq.empty else 0
        rz_blitz_count = int(rz["IS_BLITZ"].sum()) if not rz.empty else 0
        _write_takeaways(slide, [
            f"🔥 Top RZ call: {top_rz} {_fmt_count_pct(top_rz_count, len(rz))}",
            f"⚠️ RZ blitz rate: {_fmt_count_pct(rz_blitz_count, len(rz))}",
            f"🛑 Red zone sample: {len(rz)} snaps ({confidence(len(rz))} confidence)",
        ])

    # Formation
    slide = _find_slide_by_title(prs, "Formation Tendencies", 5 if overview_slide is None else 6)
    if slide is not None:
        form_freq = frequency(df, "FORMATION").head(5)
        rows = []
        for _, r in form_freq.iterrows():
            form = r["FORMATION"]
            g = df[df["FORMATION"] == form]
            front_text = _top_two_front_text(g)
            top_cov, _, cov_count, cov_total = _top_value(g, "COVERAGE", denom=len(g))
            blitz_count = int(g["IS_BLITZ"].sum()) if len(g) else 0
            rows.append([form, front_text, _fmt_count_pct(blitz_count, len(g)), f"{top_cov} {_fmt_count_pct(cov_count, cov_total)}"])
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Formation", "Front %", "Blitz %", "Coverage %"], rows)
        top_form, _, form_count, form_total = _top_value(df, "FORMATION", denom=total_plays)
        _write_takeaways(slide, [
            f"🔥 Most frequent formation: {top_form} {_fmt_count_pct(form_count, form_total)}",
            "⚠️ If top front is under 50%, this slide lists the top two fronts.",
            "🛑 Check RZ formation tendencies before goal-line calls.",
        ])

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()
