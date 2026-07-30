
from __future__ import annotations

import io
import re
from datetime import date
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

import pandas as pd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from defense_core import frequency, pct, confidence

RED = RGBColor(190, 0, 0)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(242, 242, 242)
DARK_GRAY = RGBColor(65, 65, 65)

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


def _set_cell(cell, value, header=False):
    cell.text = "" if value is None else str(value)
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)
    if header:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(10 if not header else 10.5)
            r.font.bold = bool(header)
            r.font.color.rgb = WHITE if header else BLACK


def _fill_table(table_shape, headers: Sequence[str], rows: Sequence[Sequence[object]]):
    tbl = table_shape.table
    max_rows = len(tbl.rows)
    max_cols = len(tbl.columns)

    for c in range(max_cols):
        value = headers[c] if c < len(headers) else ""
        _set_cell(tbl.cell(0, c), value, header=True)

    for r in range(1, max_rows):
        for c in range(max_cols):
            val = ""
            if r - 1 < len(rows) and c < len(rows[r - 1]):
                val = rows[r - 1][c]
            cell = tbl.cell(r, c)
            _set_cell(cell, val, header=False)
            cell.fill.fore_color.rgb = GRAY if r % 2 else WHITE

            # Highlight strong tendencies if the value is a percent string >= 60.
            try:
                percent = float(str(val).replace("%", "").strip())
                if "%" in str(val) and percent >= 60:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            run.font.bold = True
                            run.font.color.rgb = RED
            except Exception:
                pass


def _replace_opponent(slide, opponent: str, report_date: str):
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Opponent:" in txt:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = f"Opponent: {opponent}    Date: {report_date}"
            p.font.size = Pt(11)
            p.font.color.rgb = BLACK


def _write_takeaways(slide, lines: Sequence[str]):
    for shape in _text_shapes(slide):
        txt = _shape_text(shape)
        if "Key Takeaways" in txt:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = "Key Takeaways"
            p.font.bold = True
            p.font.size = Pt(16)
            p.font.color.rgb = BLACK
            for line in lines[:5]:
                para = shape.text_frame.add_paragraph()
                para.text = line
                para.level = 0
                para.font.size = Pt(12)
                para.font.color.rgb = BLACK
            return


def _top_value(df: pd.DataFrame, col: str) -> Tuple[str, float, int]:
    f = frequency(df, col)
    if f.empty:
        return "No data", 0.0, 0
    row = f.iloc[0]
    return str(row[col]), float(row["Pct"]), int(row["Plays"])


def _combo_string(row) -> str:
    return f'{row.get("FRONT","-")} / {row.get("STUNT","-")} / {row.get("BLITZ","-")} / {row.get("COVERAGE","-")}'


def _top_combos_text(df: pd.DataFrame, n=5) -> str:
    f = frequency(df, "COMBO").head(n)
    if f.empty:
        return "No data"
    return "\n".join([f'{i+1}. {r["COMBO"]} ({int(r["Plays"])} | {_fmt_pct(r["Pct"])})' for i, r in f.iterrows()])


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


def _add_front_pie(slide, front_freq: pd.DataFrame):
    # Remove explicit "Pie Chart" placeholder shapes/text.
    for shape in list(slide.shapes):
        if "Pie Chart" in _shape_text(shape):
            el = shape._element
            el.getparent().remove(el)

    if front_freq.empty:
        return

    chart_data = ChartData()
    chart_data.categories = list(front_freq["FRONT"].astype(str).head(5))
    chart_data.add_series("Front Usage", list(front_freq["Pct"].head(5).astype(float)))

    # Right side of slide; fits current master template.
    x, y, cx, cy = Inches(6.15), Inches(3.0), Inches(2.4), Inches(2.15)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data).chart
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.plots[0].has_data_labels = True
    labels = chart.plots[0].data_labels
    labels.show_percentage = True
    labels.show_category_name = False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END


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
            _fill_table(tables[0], ["Front", "Snaps", "Usage %"], rows)
        top_front, top_pct, _ = _top_value(df, "FRONT")
        top_third, top_third_pct, _ = _top_value(df[df["DOWN"] == 3], "FRONT")
        rz = pd.concat(_red_zone_bucket(df).values()) if len(df) else pd.DataFrame()
        rz_front, rz_pct, _ = _top_value(rz, "FRONT")
        lines = [
            f"🔥 Base front: {top_front} ({_fmt_pct(top_pct)})",
            f"⚠️ 3rd down front: {top_third} ({_fmt_pct(top_third_pct)})",
            f"🛑 Red zone front: {rz_front} ({_fmt_pct(rz_pct)})",
            "AI summary: play calls should start with their primary front tendency."
        ]
        _write_takeaways(slide, lines)
        _add_front_pie(slide, front_freq)

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
            _fill_table(tables[0], ["Blitz", "Snaps", "Usage %", "Top Stunt", "Stunt %"], rows)
        top_blitz, top_pct, _ = _top_value(blitz_df, "BLITZ")
        third_blitz, third_pct, _ = _top_value(blitz_df[blitz_df["DOWN"] == 3], "BLITZ")
        lines = [
            f"🔥 Top blitz: {top_blitz} ({_fmt_pct(top_pct)} of blitz snaps)",
            f"⚠️ 3rd down pressure: {third_blitz} ({_fmt_pct(third_pct)})",
            f"🛑 Overall blitz rate: {_fmt_pct(df['IS_BLITZ'].mean()*100 if len(df) else 0)}",
            "AI summary: blank blitz cells are excluded from blitz rankings."
        ]
        _write_takeaways(slide, lines)

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
            _fill_table(tables[0], ["Coverage", "Overall %", "1st Down", "2nd Down", "3rd Down", "4th Down"], rows)
        top_cov, top_pct, _ = _top_value(df, "COVERAGE")
        man_pct = df["IS_DISRESPECTFUL"].mean()*100 if len(df) else 0
        lines = [
            f"🔥 Base coverage: {top_cov} ({_fmt_pct(top_pct)})",
            f"⚠️ Man/press rate: {_fmt_pct(man_pct)}",
            "🛑 Cover 0/Cover 1/press = shot alert. Disrespect will not be tolerated.",
            "AI summary: compare coverage usage by down before calling shots."
        ]
        _write_takeaways(slide, lines)

    # Slide 4: 3rd down
    if len(prs.slides) >= 4:
        slide = prs.slides[3]
        buckets = _third_down_bucket(df)
        rows = [[name, _top_combos_text(g, 5)] for name, g in buckets.items()]
        tables = _tables(slide)
        if tables:
            _fill_table(tables[0], ["Situation", "Top 5 Front/Stunt/Blitz/Coverage Calls"], rows)
        all3 = df[df["DOWN"] == 3]
        top_combo = frequency(all3, "COMBO").iloc[0]["COMBO"] if not frequency(all3, "COMBO").empty else "No data"
        blitz_rate = all3["IS_BLITZ"].mean()*100 if len(all3) else 0
        lines = [
            f"🔥 Top 3rd down call: {top_combo}",
            f"⚠️ 3rd down blitz rate: {_fmt_pct(blitz_rate)}",
            f"🛑 3rd down sample: {len(all3)} snaps ({confidence(len(all3))} confidence)",
            "AI summary: protection and shot plans should begin here."
        ]
        _write_takeaways(slide, lines)

    # Slide 5: Red Zone
    if len(prs.slides) >= 5:
        slide = prs.slides[4]
        buckets = _red_zone_bucket(df)
        rows = [[name, _top_combos_text(g, 5)] for name, g in buckets.items()]
        tables = _tables(slide)
        if tables:
            _fill_table(tables[0], ["Situation", "Top 5 Front/Stunt/Blitz/Coverage Calls"], rows)
        rz = pd.concat([g for g in buckets.values() if not g.empty], ignore_index=True) if any(not g.empty for g in buckets.values()) else pd.DataFrame()
        top_rz = frequency(rz, "COMBO").iloc[0]["COMBO"] if not rz.empty and not frequency(rz, "COMBO").empty else "No data"
        rz_blitz = rz["IS_BLITZ"].mean()*100 if not rz.empty else 0
        lines = [
            f"🔥 Top RZ call: {top_rz}",
            f"⚠️ RZ blitz rate: {_fmt_pct(rz_blitz)}",
            f"🛑 Red zone sample: {len(rz)} snaps ({confidence(len(rz))} confidence)",
            "AI summary: finish drives with answers for their highest-probability call."
        ]
        _write_takeaways(slide, lines)

    # Slide 6: Formation
    if len(prs.slides) >= 6:
        slide = prs.slides[5]
        form_freq = frequency(df, "FORMATION").head(5)
        rows = []
        for _, r in form_freq.iterrows():
            form = r["FORMATION"]
            g = df[df["FORMATION"] == form]
            top_front, front_pct, _ = _top_value(g, "FRONT")
            top_cov, cov_pct, _ = _top_value(g, "COVERAGE")
            blitz_pct = g["IS_BLITZ"].mean()*100 if len(g) else 0
            rows.append([form, f"{top_front} ({_fmt_pct(front_pct)})", _fmt_pct(blitz_pct), f"{top_cov} ({_fmt_pct(cov_pct)})"])
        tables = _tables(slide)
        if tables:
            _fill_table(tables[0], ["Formation", "Front %", "Blitz %", "Coverage %"], rows)
        top_form, top_pct, _ = _top_value(df, "FORMATION")
        lines = [
            f"🔥 Most frequent formation: {top_form} ({_fmt_pct(top_pct)})",
            "⚠️ Use formation slide to locate pressure tells.",
            "🛑 Check RZ formation tendencies before goal-line calls.",
            "AI summary: formation splits create the cleanest next-call predictor."
        ]
        _write_takeaways(slide, lines)

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()
