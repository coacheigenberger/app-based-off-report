from __future__ import annotations

import io
import re
from copy import deepcopy
from datetime import date
from pathlib import Path
from typing import Sequence, Tuple

import pandas as pd
from pptx import Presentation

from defense_core import frequency, pct, confidence

BLANK_BLITZ = {"-", "NONE", "NO", "NO BLITZ", "0", "BASE", "NO DATA", "UNKNOWN"}


def _fmt_pct(x) -> str:
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return str(x)


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
    # Save paragraph-level settings from the first paragraph.
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


def _replace_opponent(slide, opponent: str, report_date: str) -> None:
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Opponent:" in txt:
            _set_shape_text_preserve(shape, f"Opponent: {opponent}    Date: {report_date}")


def _write_takeaways(slide, lines: Sequence[str]) -> None:
    # Only update an existing Key Takeaways shape. Never add or format one.
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Key Takeaways" in txt:
            _set_shape_text_preserve(shape, "Key Takeaways\n" + "\n".join(lines[:5]))
            return


def _top_value(df: pd.DataFrame, col: str) -> Tuple[str, float, int]:
    f = frequency(df, col)
    if f.empty:
        return "No data", 0.0, 0
    row = f.iloc[0]
    return str(row[col]), float(row["Pct"]), int(row["Plays"])


def _top_two_front_text(g: pd.DataFrame) -> str:
    fronts = frequency(g, "FRONT").head(2)
    if fronts.empty:
        return "No data"
    top_pct = float(fronts.iloc[0]["Pct"])
    if top_pct >= 50 or len(fronts) == 1:
        return f'{fronts.iloc[0]["FRONT"]} ({_fmt_pct(top_pct)})'
    return " / ".join(
        f'{r["FRONT"]} ({_fmt_pct(r["Pct"])})' for _, r in fronts.iterrows()
    )


def _top_combos_text(df: pd.DataFrame, n=5) -> str:
    f = frequency(df, "COMBO").head(n)
    if f.empty:
        return "No data"
    return "\n".join(
        [f'{i+1}. {r["COMBO"]} ({int(r["Plays"])} | {_fmt_pct(r["Pct"])})' for i, r in f.iterrows()]
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


def build_defense_pptx(engine, template_path: str | Path, opponent: str = "Opponent") -> bytes:
    prs = Presentation(str(template_path))
    df = engine.df
    report_date = date.today().strftime("%m/%d/%Y")

    for slide in prs.slides:
        _replace_opponent(slide, opponent, report_date)

    # Slide 1: Fronts
    if len(prs.slides) >= 1:
        slide = prs.slides[0]
        front_freq = frequency(df, "FRONT").head(5)
        rows = [[r["FRONT"], int(r["Plays"]), _fmt_pct(r["Pct"])] for _, r in front_freq.iterrows()]
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Front", "Snaps", "Usage %"], rows)
        top_front, top_pct, _ = _top_value(df, "FRONT")
        top_third, top_third_pct, _ = _top_value(df[df["DOWN"] == 3], "FRONT")
        rz_parts = [g for g in _red_zone_bucket(df).values() if not g.empty]
        rz = pd.concat(rz_parts, ignore_index=True) if rz_parts else pd.DataFrame()
        rz_front, rz_pct, _ = _top_value(rz, "FRONT")
        _write_takeaways(slide, [
            f"🔥 Base front: {top_front} ({_fmt_pct(top_pct)})",
            f"⚠️ 3rd down front: {top_third} ({_fmt_pct(top_third_pct)})",
            f"🛑 Red zone front: {rz_front} ({_fmt_pct(rz_pct)})",
            "AI summary: play calls should start with their primary front tendency."
        ])

    # Slide 2: Blitzes
    if len(prs.slides) >= 2:
        slide = prs.slides[1]
        blitz_df = df[~df["BLITZ"].isin(BLANK_BLITZ)].copy()
        blitz_freq = frequency(blitz_df, "BLITZ", denom=len(df)).head(5)
        rows = []
        for _, r in blitz_freq.iterrows():
            b = r["BLITZ"]
            g = blitz_df[blitz_df["BLITZ"] == b]
            stunts = frequency(g, "STUNT", denom=len(g))
            top_stunt = stunts.iloc[0]["STUNT"] if not stunts.empty else "No data"
            top_stunt_pct = stunts.iloc[0]["Pct"] if not stunts.empty else 0
            rows.append([b, int(r["Plays"]), _fmt_pct(r["Pct"]), top_stunt, _fmt_pct(top_stunt_pct)])
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Blitz", "Snaps", "Usage %", "Top Stunt", "Stunt %"], rows)
        top_blitz, top_pct, _ = _top_value(blitz_df, "BLITZ")
        third_blitz, third_pct, _ = _top_value(blitz_df[blitz_df["DOWN"] == 3], "BLITZ")
        _write_takeaways(slide, [
            f"🔥 Top blitz: {top_blitz} ({_fmt_pct(top_pct)} of blitz snaps)",
            f"⚠️ 3rd down pressure: {third_blitz} ({_fmt_pct(third_pct)})",
            f"🛑 Overall blitz rate: {_fmt_pct(df['IS_BLITZ'].mean()*100 if len(df) else 0)}",
            "AI summary: blank blitz cells are excluded from blitz rankings."
        ])

    # Slide 3: Coverages by down
    if len(prs.slides) >= 3:
        slide = prs.slides[2]
        cover_freq = frequency(df, "COVERAGE").head(5)
        rows = []
        for _, r in cover_freq.iterrows():
            cov = r["COVERAGE"]
            row = [cov, _fmt_pct(r["Pct"])]
            for down in [1, 2, 3, 4]:
                gd = df[df["DOWN"] == down]
                cov_pct = pct((gd["COVERAGE"] == cov).sum(), len(gd)) if len(gd) else 0
                row.append(_fmt_pct(cov_pct))
            rows.append(row)
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Coverage", "Overall %", "1st Down", "2nd Down", "3rd Down", "4th Down"], rows)
        top_cov, top_pct, _ = _top_value(df, "COVERAGE")
        man_pct = df["IS_DISRESPECTFUL"].mean()*100 if len(df) else 0
        _write_takeaways(slide, [
            f"🔥 Base coverage: {top_cov} ({_fmt_pct(top_pct)})",
            f"⚠️ Man/press rate: {_fmt_pct(man_pct)}",
            "🛑 Cover 0/Cover 1/press = shot alert. Disrespect will not be tolerated.",
            "AI summary: compare coverage usage by down before calling shots."
        ])

    # Slide 4: 3rd down
    if len(prs.slides) >= 4:
        slide = prs.slides[3]
        buckets = _third_down_bucket(df)
        tables = _tables(slide)
        if tables:
            headers, rows = _situational_rows_for_template(buckets, tables[0])
            _fill_table_preserve(tables[0], headers, rows)
        all3 = df[df["DOWN"] == 3]
        combo_freq = frequency(all3, "COMBO")
        top_combo = combo_freq.iloc[0]["COMBO"] if not combo_freq.empty else "No data"
        blitz_rate = all3["IS_BLITZ"].mean()*100 if len(all3) else 0
        _write_takeaways(slide, [
            f"🔥 Top 3rd down call: {top_combo}",
            f"⚠️ 3rd down blitz rate: {_fmt_pct(blitz_rate)}",
            f"🛑 3rd down sample: {len(all3)} snaps ({confidence(len(all3))} confidence)",
            "AI summary: protection and shot plans should begin here."
        ])

    # Slide 5: Red Zone
    if len(prs.slides) >= 5:
        slide = prs.slides[4]
        buckets = _red_zone_bucket(df)
        tables = _tables(slide)
        if tables:
            headers, rows = _situational_rows_for_template(buckets, tables[0])
            _fill_table_preserve(tables[0], headers, rows)
        rz_parts = [g for g in buckets.values() if not g.empty]
        rz = pd.concat(rz_parts, ignore_index=True) if rz_parts else pd.DataFrame()
        combo_freq = frequency(rz, "COMBO") if not rz.empty else pd.DataFrame()
        top_rz = combo_freq.iloc[0]["COMBO"] if not combo_freq.empty else "No data"
        rz_blitz = rz["IS_BLITZ"].mean()*100 if not rz.empty else 0
        _write_takeaways(slide, [
            f"🔥 Top RZ call: {top_rz}",
            f"⚠️ RZ blitz rate: {_fmt_pct(rz_blitz)}",
            f"🛑 Red zone sample: {len(rz)} snaps ({confidence(len(rz))} confidence)",
            "AI summary: finish drives with answers for their highest-probability call."
        ])

    # Slide 6: Formation
    if len(prs.slides) >= 6:
        slide = prs.slides[5]
        form_freq = frequency(df, "FORMATION").head(5)
        rows = []
        for _, r in form_freq.iterrows():
            form = r["FORMATION"]
            g = df[df["FORMATION"] == form]
            front_text = _top_two_front_text(g)
            top_cov, cov_pct, _ = _top_value(g, "COVERAGE")
            blitz_pct = g["IS_BLITZ"].mean()*100 if len(g) else 0
            rows.append([form, front_text, _fmt_pct(blitz_pct), f"{top_cov} ({_fmt_pct(cov_pct)})"])
        tables = _tables(slide)
        if tables:
            _fill_table_preserve(tables[0], ["Formation", "Front %", "Blitz %", "Coverage %"], rows)
        top_form, top_pct, _ = _top_value(df, "FORMATION")
        _write_takeaways(slide, [
            f"🔥 Most frequent formation: {top_form} ({_fmt_pct(top_pct)})",
            "⚠️ If top front is under 50%, this slide lists the top two fronts.",
            "🛑 Check RZ formation tendencies before goal-line calls.",
            "AI summary: formation splits create the cleanest next-call predictor."
        ])

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()
