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


## v1.1 updates
- Filters every upload to rows where `ODK` / `O/D/K` equals `D` before analysis.
- The Blitz tab now displays every blitz call and every front/stunt/blitz/coverage version.
- Down-and-distance and field-zone detail now include full front distributions.
