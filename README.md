# Defense Analyst Platform v1.0

A Streamlit platform for opponent defensive breakdowns.

## Features
- Upload one or more Hudl Excel/CSV files
- Automatic column normalization and typo-tolerant aliases
- Front, stunt, blitz, coverage, pressure direction/strength
- Down-and-distance, field-zone, formation, and P & 10 breakdowns
- Top defensive call combinations
- Next-call estimator with sample confidence
- Shot alerts for Cover 0, Cover 1, and press
- Screen alerts for listed 5/6-man pressure families
- Downloadable Excel report and normalized CSV

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload the contents of this folder so `app.py` is at repository root.
3. In Streamlit Community Cloud, select the repository and use `app.py` as the main file.
4. Deploy.

## Expected columns
The platform accepts variants of: Game, O/D/K, Quarter, Time Remaining, Down, Distance, Yard Line, Field Zone, Hash, Personnel, OFF FORM, OFF Play, Backfield Set, Motion, Defensive Front, Stunt, Blitz Type, Coverage, Blitz Direction, Blitz Strength, Result, GN/LS, Turnover Forced, Penalty, # of Blitzers, P & 10, Coverage Shell, Rotation, Fit 1/2/3, and Box Add.

Missing columns are shown as no data rather than blocking the report.


## v1.2 updates
- Filters every upload to rows where `ODK` / `O/D/K` equals `D` before analysis.
- The Blitz tab now displays every blitz call and every front/stunt/blitz/coverage version.
- Down-and-distance and field-zone detail now include full front distributions.


## v1.2 updates
- Front menu shows blitz percentage and stunt percentage out of each front.
- Each front includes all stunt calls, all blitz calls, and all blitz + stunt combinations.
- Blitz tab displays the overall blitz percentage.

## PowerPoint export

The Export tab now includes a **Download PowerPoint scouting deck** button.

The app uses the bundled `MASTER DEF Breakdown Template.pptx` by default and automatically fills:
- Slide 1: Top 5 fronts, percentages, key takeaways, and a front usage pie chart.
- Slide 2: Top 5 blitzes, excluding blank/no-blitz entries, plus the top stunt tied to each blitz.
- Slide 3: Coverage usage by down.
- Slide 4: Top 5 total calls for 3rd & 7+, 3rd & 3-6, and 3rd & 1-2.
- Slide 5: Top 5 total calls for High Red Zone, Red Zone, and Goal Line.
- Slide 6: Formation tendencies.

You can upload another `.pptx` in the Export tab as long as it keeps the same six-slide/table structure.


## v2.1 updates
- PowerPoint export now uses the uploaded master breakdown template as the default.
- Export logic updates text/numbers only and avoids changing table dimensions, fills, fonts, borders, margins, or slide formatting.
- Slide 4 and Slide 5 support a middle **Total Plays** column when the template table has three columns.
- Formation Tendencies now lists the top two fronts when the most-used front is below 50%.


## v2.1 updates
- PowerPoint export now uses **count/total (percentage)** everywhere a percentage is shown.
- Opponent labels now show only `Opponent: [Name]`; dates are removed from the slides.
- Slide 1 can populate an optional Opponent Overview from a local CSV/Excel file instead of relying on a website fetch.
- Front Tendencies now supports: Front, Snaps, Usage, Blitz %, Top Blitz Call.
- Blitz Tendencies now supports: Blitz, Snaps, Usage, Top Front, Top Stunt. Top Front and Top Stunt include count/total (percentage).
- 3rd Down and Red Zone tables include a Total Plays column when present in the template.
- Formation Tendencies lists the top two fronts when the most-used front is below 50%; otherwise it lists only the top front.
- Export logic still only replaces text inside existing cells/placeholders. It does not change fonts, table fills, borders, row heights, column widths, or slide formatting.

### Optional opponent overview file
Upload a CSV or Excel file in the Export tab to populate the overview slide. The file may contain either:
- A key/value table with columns `Statistic` and `Value`, using rows such as `Overall Record`, `Total Sacks`, `Total INTs`, and `Total Fumble Recoveries`.
- A schedule table with columns such as `Game`, `Opponent`, `Record`, `Score`, and `W/L`.

You can also combine snapshot and schedule data in one Excel workbook using separate sheets.


## v2.1 fixes
- Rebuilt from the clean user master template.
- Front Tendencies table includes Blitz % and Top Blitz Call.
- Every percentage in the export tables uses count/total (percentage).
- 3rd Down and Red Zone tables keep the original row heights/header styling and add Total Plays.
- Takeaway/freeform boxes are no longer edited by export logic, preventing random font-color changes.
- GoBound fetch remains disabled; optional local opponent overview files are supported.


## v2.2 PowerPoint template cleanup
- Bundled master template uses alternating white/light gray data rows for every table.
- Opponent labels in the bundled template no longer include Date.
- Data-row font color is standardized to black in the template so export values do not inherit random red text.
- Front Tendencies retains five columns: Front, Snaps, Usage, Blitz %, Top Blitz Call.
- PowerPoint export still only replaces text in existing cells/placeholders and does not resize or recreate tables.

## v2.3 fixed situational PowerPoint export
- Slide 4 (3rd Down) and Slide 5 (Red Zone) now fill five separate columns: Situation, Total Plays, Fronts, Blitz, Coverage.
- Fronts show top 2 fronts in each situation.
- Blitz shows overall blitz rate plus top 2 blitz calls.
- Coverage shows the top coverage.
- The app must use this file as `ppt_export.py`; do not leave it named `ppt_export_v2.3.py`.
