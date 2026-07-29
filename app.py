
from __future__ import annotations
import tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from defense_core import DefenseEngine, frequency, combos, DND_ORDER, FIELD_ZONE_ORDER

st.set_page_config(page_title="Defense Analyst",page_icon="🏈",layout="wide")
st.title("🏈 Defense Analyst")
st.caption("Opponent defensive tendency, pressure, coverage, and next-call platform")

def save_uploads(files,folder):
    paths=[]
    for f in files:
        p=folder/Path(f.name).name; p.write_bytes(f.getbuffer()); paths.append(str(p))
    return paths

with st.sidebar:
    st.header("Opponent data")
    opponent=st.text_input("Opponent name",placeholder="Opponent")
    files=st.file_uploader("Upload Excel/CSV cutups",type=["xlsx","xls","csv"],accept_multiple_files=True)
    load=st.button("Load defense",type="primary",use_container_width=True)
    st.caption("Only rows tagged D in ODK / O-D-K are analyzed. Missing columns are preserved as No data.")

if load:
    if not files: st.error("Upload at least one file.")
    else:
        try:
            with tempfile.TemporaryDirectory() as td:
                eng=DefenseEngine.from_files(save_uploads(files,Path(td)))
            st.session_state.engine=eng; st.session_state.opponent=opponent or "Opponent"
            st.success(f"Loaded {len(eng.df)} defensive plays.")
        except Exception as e: st.exception(e)

eng=st.session_state.get("engine")
if eng is None:
    st.info("Upload opponent files in the sidebar to begin.")
    st.stop()

tabs=st.tabs(["Dashboard","Fronts / Stunts","Blitz","Coverage","Situations","Formation","Next Call","Attack Plan","P & 10","All Combos","Export"])

with tabs[0]:
    ov=eng.overview()
    c1,c2,c3,c4,c5=st.columns(5)
    c1.metric("Defensive plays",ov["plays"]); c2.metric("Blitz",f'{ov["blitz_pct"]}%')
    c3.metric("Man / Press",f'{ov["man_pct"]}%'); c4.metric("5/6-man pressure",f'{ov["screen_pressure_pct"]}%')
    c5.metric("Sample confidence",ov["confidence"])
    if ov["disrespect_pct"]>0:
        st.error(f'DISRESPECT ALERT: Cover 0, Cover 1, or press appeared on {ov["disrespect_pct"]}% of snaps. Take the shot when the matchup is right.')
    cols=st.columns(4)
    for col,key in zip(cols,["FRONT","STUNT","BLITZ","COVERAGE"]):
        with col:
            st.subheader(key.title())
            d=frequency(eng.df,key).head(8)
            st.dataframe(d,use_container_width=True,hide_index=True)
            if not d.empty: st.bar_chart(d.set_index(key)["Pct"])

with tabs[1]:
    st.subheader("Front menu")
    fronts=frequency(eng.df,"FRONT")
    front_rows=[]
    for _, row in fronts.iterrows():
        front=row["FRONT"]
        g=eng.df[eng.df["FRONT"]==front]
        front_rows.append({
            "FRONT":front,
            "Plays":len(g),
            "Front %":round(len(g)/len(eng.df)*100,1) if len(eng.df) else 0.0,
            "Blitz Plays":int(g["IS_BLITZ"].sum()),
            "Blitz % Out of Front":round(g["IS_BLITZ"].mean()*100,1) if len(g) else 0.0,
            "Stunt Plays":int(g["IS_STUNT"].sum()),
            "Stunt % Out of Front":round(g["IS_STUNT"].mean()*100,1) if len(g) else 0.0,
        })
    st.dataframe(pd.DataFrame(front_rows),use_container_width=True,hide_index=True)

    for front in fronts["FRONT"]:
        g=eng.df[eng.df["FRONT"]==front]
        blitz_pct=round(g["IS_BLITZ"].mean()*100,1) if len(g) else 0.0
        stunt_pct=round(g["IS_STUNT"].mean()*100,1) if len(g) else 0.0
        with st.expander(f"{front} — {len(g)} plays ({len(g)/len(eng.df)*100:.1f}%)"):
            m1,m2,m3=st.columns(3)
            m1.metric("Front snaps",len(g))
            m2.metric("Blitz % out of front",f"{blitz_pct}%")
            m3.metric("Stunt % out of front",f"{stunt_pct}%")

            a,b=st.columns(2)
            with a:
                st.markdown("**All stunt calls**")
                st.dataframe(frequency(g,"STUNT"),hide_index=True,use_container_width=True)
            with b:
                st.markdown("**All blitz calls**")
                st.dataframe(frequency(g,"BLITZ"),hide_index=True,use_container_width=True)

            st.markdown("**All blitz + stunt combinations within this front**")
            x=g.copy()
            x["BLITZ + STUNT"]=x["BLITZ"]+" / "+x["STUNT"]
            st.dataframe(frequency(x,"BLITZ + STUNT"),hide_index=True,use_container_width=True)

with tabs[2]:
    overall_blitz=eng.overview()["blitz_pct"]
    st.metric("Overall blitz percentage",f"{overall_blitz}%")
    st.caption("A play counts as a blitz when the Blitz column contains a recorded blitz call.")
    blitzes=frequency(eng.df,"BLITZ")
    st.dataframe(blitzes,use_container_width=True,hide_index=True)
    for blitz in blitzes["BLITZ"]:
        g=eng.df[eng.df["BLITZ"]==blitz]
        with st.expander(f"{blitz} — {len(g)} plays ({len(g)/len(eng.df)*100:.1f}%)"):
            st.markdown("**All front / stunt / blitz / coverage versions**")
            st.dataframe(frequency(g,"COMBO"),use_container_width=True,hide_index=True)

with tabs[3]:
    cov=frequency(eng.df,"COVERAGE")
    st.dataframe(cov,use_container_width=True,hide_index=True)
    if eng.df["IS_DISRESPECTFUL"].any():
        st.warning("Cover 0, Cover 1, and press are classified as disrespect. Build shot answers into these situations.")

with tabs[4]:
    choice=st.radio("Situation",["Down & Distance","Field Zone"],horizontal=True)
    dim="DND" if choice=="Down & Distance" else "FIELD_ZONE"
    st.dataframe(eng.situation_table(dim),use_container_width=True,hide_index=True)
    order=DND_ORDER if dim=="DND" else FIELD_ZONE_ORDER
    val=st.selectbox("Detailed situation",order)
    g=eng.df[eng.df[dim]==val]
    st.caption(f"{len(g)} plays")
    for key in ["FRONT","BLITZ","COVERAGE","COMBO"]:
        st.markdown(f"#### {key.title()}")
        st.dataframe(frequency(g,key).head(20 if key!="COMBO" else 5),use_container_width=True,hide_index=True)

with tabs[5]:
    st.dataframe(eng.situation_table("FORMATION"),use_container_width=True,hide_index=True)
    forms=sorted(v for v in eng.df["FORMATION"].unique() if v!="-")
    form=st.selectbox("Formation detail",forms if forms else ["No data"])
    g=eng.df[eng.df["FORMATION"]==form]
    for key in ["FRONT","BLITZ","COVERAGE","COMBO"]:
        st.markdown(f"#### {key.title()}")
        st.dataframe(frequency(g,key).head(20 if key!="COMBO" else 5),use_container_width=True,hide_index=True)

with tabs[6]:
    c1,c2,c3,c4=st.columns(4)
    down=c1.selectbox("Down",[1,2,3,4]); distance=c2.number_input("Distance",1,30,10)
    zone=c3.selectbox("Field zone",["All"]+FIELD_ZONE_ORDER)
    forms=sorted(v for v in eng.df["FORMATION"].unique() if v!="-")
    form=c4.selectbox("Formation",["All"]+forms)
    p=eng.predict(down,distance,None if zone=="All" else zone,None if form=="All" else form)
    st.info(p.narrative)
    a,b,c,d=st.columns(4)
    a.metric("Front",p.front); b.metric("Stunt",p.stunt); c.metric("Blitz",p.blitz); d.metric("Coverage",p.coverage)
    st.markdown("#### Top 3 total calls")
    st.dataframe(p.top3,use_container_width=True,hide_index=True)

with tabs[7]:
    shots,screens=eng.recommendations()
    st.subheader("Shot alerts")
    st.error("Cover 0, Cover 1, and press are disrespectful. These are the best historical windows to take shots.")
    st.dataframe(shots,use_container_width=True,hide_index=True)
    st.subheader("Screen alerts")
    st.info("Screen alerts flag RAM, MOSS, RAW, WRAP, SNAKE, MOW, BOMB, BASS, BOW, or a recorded 5/6-man pressure.")
    st.dataframe(screens,use_container_width=True,hide_index=True)

with tabs[8]:
    g=eng.df[eng.df["P10_YES"]]
    if g.empty:
        st.info("No explicit P & 10 / P-10 YES data was found.")
    else:
        st.caption(f"{len(g)} P & 10 plays")
        for key in ["FRONT","STUNT","BLITZ","COVERAGE","COMBO"]:
            st.markdown(f"#### {key.title()}")
            st.dataframe(frequency(g,key).head(20),use_container_width=True,hide_index=True)

with tabs[9]:
    st.dataframe(frequency(eng.df,"COMBO"),use_container_width=True,hide_index=True)

with tabs[10]:
    name=(st.session_state.get("opponent") or "Opponent").replace(" ","_")
    st.download_button("Download full Excel tendency report",eng.excel_bytes(),f"{name}_Defensive_Tendency_Report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.download_button("Download normalized CSV",eng.df.to_csv(index=False).encode(),f"{name}_Normalized_Defense.csv","text/csv")
