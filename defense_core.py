
from __future__ import annotations
import io, math, re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

DND_ORDER = [
    "1st & 10","1st & 11+","1st & <10",
    "2nd & 8+","2nd & 4-7","2nd & 1-3",
    "3rd & 8+","3rd & 4-7","3rd & 1-3",
    "4th & 8+","4th & 4-7","4th & 1-3",
]
FIELD_ZONE_ORDER = [
    "Coming Out","Backed Up","Three Down Territory","Four Down Territory",
    "High Red Zone","Low Red Zone","Goal Line"
]
PRESSURE_5_6 = {"RAM","MOSS","RAW","WRAP","SNAKE","MOW","BOMB","BASS","BOW"}
MAN_COV = {"COVER 0","COVER 1","0","1","ORANGE","RED"}
SPLIT_SAFETY = {"COVER 2","COVER 4","2","4","BLUE","BLACK"}
MIDDLE_CLOSED = {"COVER 1","COVER 3","1","3","RED","GREEN"}

ALIASES = {
    "GAME":["GAME"], "ODK":["O/D/K","ODK"], "QUARTER":["QUARTER","QTR"], "TIME":["TIME REMAINING","TIME"],
    "DOWN":["DOWN","DN"], "DISTANCE":["DISTANCE","DIST"], "YARD_LINE":["YARD LINE","YARD LN","BALL ON"],
    "FIELD_ZONE":["FIELD ZONE"], "HASH":["HASH"], "PERSONNEL":["PERSONNEL","PERS"],
    "FORMATION":["OFF FORM","FORMATION","OFFENSIVE FORMATION"], "OFF_PLAY":["OFF PLAY","PLAY"],
    "BACKFIELD":["BACKFIELD SET","BACKFIELD"], "MOTION":["MOTION"],
    "FRONT":["DEFENSIVE FRONT","DEF FRONT","FRONT","ALIGNMENT"],
    "STUNT":["STUNT","DEFENSIVE STUNT"], "BLITZ":["BLITZ TYPE","BLITZ","PRESSURE"],
    "COVERAGE":["COVERAGE","COV"], "BLITZ_DIR":["BLITZ DIRECTION (FIELD/BOUNDARY)","BLITZ DIRECTION","PRESSURE DIRECTION"],
    "BLITZ_STRENGTH":["BLITZ STRENGTH (STRONG/WEAK)","BLITZ STRENGTH","PRESSURE STRENGTH"],
    "RESULT":["RESULT"], "GNLS":["GN/LS","GAIN/LOSS","YARDS","GAIN"],
    "TURNOVER":["TURNOVER FORCED","TURNOVER"], "PENALTY":["PENALTY"],
    "P10":["P & 10","P&10","P - 10","P-10","P – 10"],
    "NUM_BLITZERS":["# OF BLITZERS","NUMBER OF BLITZERS","BLITZERS","PRESSURE COUNT"],
    "SHELL":["COVERAGE SHELL","SHELL"], "ROTATION":["ROTATION","SECONDARY ROTATION"],
    "FIT1":["FIT 1"], "FIT2":["FIT 2"], "FIT3":["FIT 3"], "BOX_ADD":["BOX ADD","DEFENDER ADDED TO BOX"],
}

def clean(v: Any, blank: str="-") -> str:
    if pd.isna(v): return blank
    s=re.sub(r"\s+"," ",str(v).strip()).upper()
    return s if s else blank

def num(v: Any) -> float:
    if pd.isna(v): return math.nan
    try: return float(v)
    except Exception:
        m=re.search(r"-?\d+(?:\.\d+)?",str(v))
        return float(m.group()) if m else math.nan

def canonical_coverage(v: Any) -> str:
    s=clean(v)
    map_words={"ORANGE":"COVER 0","RED":"COVER 1","BLUE":"COVER 2","GREEN":"COVER 3","BLACK":"COVER 4"}
    if s in map_words: return map_words[s]
    m=re.search(r"(?:COVER|COV)?\s*([0-7])",s)
    if m: return f"COVER {m.group(1)}" + (" PRESS" if "PRESS" in s else "")
    return s

def dnd(d: Any, dist: Any) -> str:
    dn=num(d); y=num(dist)
    if pd.isna(dn) or pd.isna(y): return "Unknown"
    dn=int(dn)
    if dn==1:
        return "1st & 10" if y==10 else ("1st & 11+" if y>10 else "1st & <10")
    if dn in (2,3,4):
        pre={2:"2nd",3:"3rd",4:"4th"}[dn]
        if y>=8: return f"{pre} & 8+"
        if 4<=y<=7: return f"{pre} & 4-7"
        if 1<=y<=3: return f"{pre} & 1-3"
    return "Unknown"

def field_zone(v: Any) -> str:
    y=num(v)
    if pd.isna(y): return "Unknown"
    if -10<=y<=-1:return "Coming Out"
    if -20<=y<=-11:return "Backed Up"
    if -49<=y<=-21:return "Three Down Territory"
    if 21<=y<=49:return "Four Down Territory"
    if 11<=y<=20:return "High Red Zone"
    if 5<=y<=10:return "Low Red Zone"
    if 1<=y<=4:return "Goal Line"
    return "Unknown"

def _find(columns, names):
    m={str(c).strip().upper():c for c in columns}
    for n in names:
        if n.upper() in m:return m[n.upper()]
    return None

def read_files(paths: List[str]) -> pd.DataFrame:
    frames=[]
    for p in paths:
        path=Path(p)
        df=pd.read_csv(path) if path.suffix.lower()==".csv" else pd.read_excel(path)
        ren={}
        for k,names in ALIASES.items():
            c=_find(df.columns,names)
            if c is not None: ren[c]=k
        df=df.rename(columns=ren)
        df["SOURCE_FILE"]=path.name
        frames.append(df)
    if not frames: raise ValueError("No files supplied.")
    return pd.concat(frames,ignore_index=True,sort=False)

def prepare(df: pd.DataFrame) -> pd.DataFrame:
    out=df.copy()
    for c in ALIASES:
        if c not in out.columns: out[c]="-"
    for c in ["DOWN","DISTANCE","YARD_LINE","GNLS","NUM_BLITZERS"]:
        out[c]=out[c].apply(num)
    for c in ["GAME","ODK","HASH","PERSONNEL","FORMATION","OFF_PLAY","BACKFIELD","MOTION","FRONT","STUNT","BLITZ",
              "BLITZ_DIR","BLITZ_STRENGTH","RESULT","TURNOVER","PENALTY","P10","SHELL","ROTATION","FIT1","FIT2","FIT3","BOX_ADD"]:
        out[c]=out[c].apply(clean)

    # Defense-only platform: every downstream table, prediction, recommendation,
    # chart, and export uses only rows explicitly tagged D in ODK / O-D-K.
    if "ODK" not in out.columns or out["ODK"].eq("-").all():
        raise ValueError("No ODK (or O/D/K) data was found. Add an ODK column and tag defensive snaps with D.")
    out=out[out["ODK"].eq("D")].copy()
    if out.empty:
        raise ValueError("No defensive snaps were found. The ODK (or O/D/K) column must contain D for defensive plays.")

    out["COVERAGE"]=out["COVERAGE"].apply(canonical_coverage)
    out["DND"]=[dnd(a,b) for a,b in zip(out["DOWN"],out["DISTANCE"])]
    derived=out["YARD_LINE"].apply(field_zone)
    out["FIELD_ZONE"]=out["FIELD_ZONE"].apply(clean)
    out.loc[out["FIELD_ZONE"].eq("-"),"FIELD_ZONE"]=derived[out["FIELD_ZONE"].eq("-")]
    out["IS_BLITZ"]=~out["BLITZ"].isin(["-","NONE","NO","NO BLITZ","0","BASE"])
    out["IS_STUNT"]=~out["STUNT"].isin(["-","NONE","NO","NO STUNT","0","BASE"])
    out["IS_MAN"]=out["COVERAGE"].apply(lambda x:any(x.startswith(v) for v in ["COVER 0","COVER 1"]) or "PRESS" in x)
    out["IS_DISRESPECTFUL"]=out["IS_MAN"]
    out["IS_SCREEN_PRESSURE"]=out["BLITZ"].apply(lambda x:any(tok in x for tok in PRESSURE_5_6)) | out["NUM_BLITZERS"].isin([5,6])
    out["COMBO"]=out["FRONT"]+" / "+out["STUNT"]+" / "+out["BLITZ"]+" / "+out["COVERAGE"]
    # Exact P&10 only when explicitly labeled; fallback first-and-10 is separated.
    out["P10_YES"]=out["P10"].isin(["YES","Y","1","TRUE","P & 10","P&10","P - 10","P-10","P – 10"])
    return out

def pct(n,d): return round(100*n/d,1) if d else 0.0
def confidence(n:int)->str:
    return "High" if n>=20 else ("Medium" if n>=10 else "Low")

def frequency(df:pd.DataFrame,col:str,denom:Optional[int]=None,drop_blank=True)->pd.DataFrame:
    if col not in df or df.empty:return pd.DataFrame(columns=[col,"Plays","Pct"])
    s=df[col].fillna("-").astype(str)
    if drop_blank:s=s[~s.isin(["-","NONE","NO DATA","UNKNOWN"])]
    counts=s.value_counts()
    d=denom if denom is not None else len(df)
    return pd.DataFrame({col:counts.index,"Plays":counts.values,"Pct":[pct(x,d) for x in counts.values]})

def combos(df,n=5):
    return frequency(df,"COMBO").head(n)

@dataclass
class Prediction:
    context: str
    sample: int
    front: str
    stunt: str
    blitz: str
    coverage: str
    top3: pd.DataFrame
    confidence: str
    narrative: str

class DefenseEngine:
    def __init__(self,df):
        self.df=prepare(df)
        if self.df.empty: raise ValueError("No plays found.")
    @classmethod
    def from_files(cls,paths): return cls(read_files(paths))
    def slice(self,filters:Optional[Dict[str,Any]]=None):
        g=self.df
        for key,val in (filters or {}).items():
            if val in (None,"","All",[]): continue
            vals=val if isinstance(val,(list,tuple,set)) else [val]
            norm={clean(v) for v in vals}
            if key in g:g=g[g[key].astype(str).str.upper().isin(norm)]
        return g.copy()
    def overview(self,g=None):
        g=self.df if g is None else g
        n=len(g)
        return {
            "plays":n,"blitz_pct":pct(g["IS_BLITZ"].sum(),n),
            "man_pct":pct(g["IS_MAN"].sum(),n),
            "disrespect_pct":pct(g["IS_DISRESPECTFUL"].sum(),n),
            "screen_pressure_pct":pct(g["IS_SCREEN_PRESSURE"].sum(),n),
            "confidence":confidence(n),
        }
    def breakdown(self,g=None):
        g=self.df if g is None else g
        return {k:frequency(g,k) for k in ["FRONT","STUNT","BLITZ","COVERAGE","BLITZ_DIR","BLITZ_STRENGTH"]}
    def predict(self,down=None,distance=None,field_zone_value=None,formation=None):
        filters={}
        if down is not None and distance is not None: filters["DND"]=dnd(down,distance)
        if field_zone_value: filters["FIELD_ZONE"]=field_zone_value
        if formation: filters["FORMATION"]=formation
        g=self.slice(filters)
        # Back off gracefully if exact sample is thin.
        used=filters.copy()
        if len(g)<5 and formation:
            used.pop("FORMATION",None); g=self.slice(used)
        if len(g)<5 and field_zone_value:
            used.pop("FIELD_ZONE",None); g=self.slice(used)
        n=len(g)
        def top(col):
            t=frequency(g,col)
            return str(t.iloc[0][col]) if not t.empty else "No data"
        top3=combos(g,3)
        b=top("BLITZ"); c=top("COVERAGE"); s=top("STUNT")
        narrative=f"Based on {n} matching snaps, expect {b} pressure, {c}, and {s}. Confidence: {confidence(n)}."
        return Prediction(" | ".join(f"{k}: {v}" for k,v in used.items()) or "All snaps",n,top("FRONT"),s,b,c,top3,confidence(n),narrative)
    def situation_table(self,dimension):
        rows=[]
        order=DND_ORDER if dimension=="DND" else (FIELD_ZONE_ORDER if dimension=="FIELD_ZONE" else sorted(self.df[dimension].dropna().unique()))
        for value in order:
            g=self.df[self.df[dimension]==value]
            if g.empty: continue
            ov=self.overview(g)
            rows.append({
                dimension:value,"Plays":len(g),"Blitz %":ov["blitz_pct"],"Man/Press %":ov["man_pct"],
                "Top Front":frequency(g,"FRONT").iloc[0]["FRONT"] if not frequency(g,"FRONT").empty else "No data",
                "Top Stunt":frequency(g,"STUNT").iloc[0]["STUNT"] if not frequency(g,"STUNT").empty else "No data",
                "Top Blitz":frequency(g,"BLITZ").iloc[0]["BLITZ"] if not frequency(g,"BLITZ").empty else "No data",
                "Top Coverage":frequency(g,"COVERAGE").iloc[0]["COVERAGE"] if not frequency(g,"COVERAGE").empty else "No data",
                "Confidence":confidence(len(g))
            })
        return pd.DataFrame(rows)
    def recommendations(self):
        shots=[]; screens=[]
        for dims in [("DND",),("FIELD_ZONE",),("FORMATION",),("DND","FIELD_ZONE"),("DND","FORMATION")]:
            for keys,g in self.df.groupby(list(dims),dropna=False):
                if len(g)<4: continue
                keys=keys if isinstance(keys,tuple) else (keys,)
                label=" | ".join(f"{d}: {v}" for d,v in zip(dims,keys))
                man=pct(g["IS_DISRESPECTFUL"].sum(),len(g))
                scr=pct(g["IS_SCREEN_PRESSURE"].sum(),len(g))
                if man>=35: shots.append((label,len(g),man,confidence(len(g))))
                if scr>=25: screens.append((label,len(g),scr,confidence(len(g))))
        shots=sorted(shots,key=lambda x:(-x[2],-x[1]))[:12]
        screens=sorted(screens,key=lambda x:(-x[2],-x[1]))[:12]
        return (pd.DataFrame(shots,columns=["Situation","Plays","Cover 0/1 or Press %","Confidence"]),
                pd.DataFrame(screens,columns=["Situation","Plays","5/6-Man Pressure %","Confidence"]))
    def excel_bytes(self)->bytes:
        bio=io.BytesIO()
        with pd.ExcelWriter(bio,engine="openpyxl") as w:
            pd.DataFrame([self.overview()]).to_excel(w,"Overview",index=False)
            for col in ["FRONT","STUNT","BLITZ","COVERAGE","BLITZ_DIR","BLITZ_STRENGTH","COMBO"]:
                frequency(self.df,col).to_excel(w,col[:31],index=False)
            self.situation_table("DND").to_excel(w,"Down Distance",index=False)
            self.situation_table("FIELD_ZONE").to_excel(w,"Field Zone",index=False)
            self.situation_table("FORMATION").to_excel(w,"Formation",index=False)
            self.df[self.df["P10_YES"]].to_excel(w,"P and 10 Plays",index=False)
            shots,screens=self.recommendations()
            shots.to_excel(w,"Shot Alerts",index=False); screens.to_excel(w,"Screen Alerts",index=False)
            self.df.to_excel(w,"Normalized Data",index=False)
        return bio.getvalue()
