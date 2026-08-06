
import streamlit as st
import pandas as pd
import numpy as np
import io, os, re
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Dashboard Kinerja Dosen", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.stApp {background:#f7f7f4;}
.block-container {max-width:1450px;padding-top:1.5rem;padding-bottom:4rem}
h1,h2,h3 {font-family:Georgia,serif;color:#0d2342}
[data-testid="stMetric"] {background:white;border:1px solid #dce3ea;border-radius:14px;padding:16px;box-shadow:0 2px 10px rgba(0,0,0,.03)}
div[data-testid="stDataFrame"] {background:white;border-radius:12px}
.small-note {color:#7184a0;font-size:.88rem}
.hero {padding:25px 30px;border:1px solid #dce3ea;border-radius:18px;background:white;margin-bottom:20px}
.badge-green{background:#e8f5ee;color:#147a4b;padding:5px 10px;border-radius:99px}
.badge-yellow{background:#fff4d6;color:#956400;padding:5px 10px;border-radius:99px}
.badge-red{background:#fde9e7;color:#b53b34;padding:5px 10px;border-radius:99px}
</style>
""", unsafe_allow_html=True)

PER_FILE="kinerja_perkuliahan.xlsx"
UJI_FILE="kinerja_ujian.xlsx"

def clean_cols(df):
    df=df.copy()
    df.columns=[str(c).strip() for c in df.columns]
    return df

@st.cache_data(show_spinner=False)
def read_excel_auto(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    book=pd.ExcelFile(path)
    frames=[]
    for s in book.sheet_names:
        t=pd.read_excel(path,sheet_name=s)
        if not t.empty:
            frames.append(clean_cols(t))
    return pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()

def find_col(df, candidates):
    norm=lambda x: re.sub(r'[^A-Z0-9]','',str(x).upper())
    lookup={norm(c):c for c in df.columns}
    for c in candidates:
        if norm(c) in lookup:return lookup[norm(c)]
    return None

def pct_series(s):
    x=pd.to_numeric(s.astype(str).str.replace("%","",regex=False).str.replace(",",".",regex=False),errors="coerce")
    if x.dropna().size and x.dropna().quantile(.95)<=1.5:x=x*100
    return x.clip(0,100)

def semester_key(x):
    s=str(x).lower()
    yrs=re.findall(r'20\d{2}',s)
    y=int(yrs[0]) if yrs else 0
    term=0 if "gasal" in s or "ganjil" in s else 1
    return y*2+term

def status(v):
    if pd.isna(v): return "Tidak ada data"
    if v>=85:return "Hijau • Baik"
    if v>=70:return "Kuning • Perlu perhatian"
    return "Merah • Perlu tindak lanjut"

def excel_bytes(df, sheet="Data"):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name=sheet[:31])
    return b.getvalue()

def filters(df, cols):
    out=df.copy()
    st.sidebar.subheader("Filter")
    for label,col in cols:
        if col and col in out.columns:
            vals=sorted(out[col].dropna().astype(str).unique())
            sel=st.sidebar.multiselect(label,vals)
            if sel:out=out[out[col].astype(str).isin(sel)]
    return out

def ranking_table(g, metric, n=10, ascending=False):
    x=g.sort_values(metric,ascending=ascending).head(n).copy()
    x.insert(0,"Peringkat",range(1,len(x)+1))
    return x

def trend_delta(df,name_col,sem_col,metric,name):
    z=df[df[name_col].astype(str)==str(name)].copy()
    if z.empty:return np.nan,None,None
    q=z.groupby(sem_col,as_index=False)[metric].mean()
    q["_k"]=q[sem_col].map(semester_key);q=q.sort_values("_k")
    if len(q)<2:return np.nan,q.iloc[-1][sem_col],None
    return q.iloc[-1][metric]-q.iloc[-2][metric],q.iloc[-1][sem_col],q.iloc[-2][sem_col]

def tren_label(v):
    if pd.isna(v):return "N/A"
    if v>0:return "▲ Naik"
    if v<0:return "▼ Turun"
    return "— Tetap"

def prodi_summary(df, prodi_col, sem_col, metric_col):
    if not prodi_col or not sem_col or not metric_col or prodi_col not in df.columns:
        return pd.DataFrame()
    g=df.groupby([prodi_col,sem_col],as_index=False)[metric_col].mean()
    g["_k"]=g[sem_col].map(semester_key)
    rows=[]
    for prodi,sub in g.groupby(prodi_col):
        sub=sub.sort_values("_k")
        latest=sub.iloc[-1]
        prev=sub.iloc[-2] if len(sub)>=2 else None
        delta=latest[metric_col]-prev[metric_col] if prev is not None else np.nan
        rows.append({
            "Prodi":prodi,
            "Semester Terakhir":latest[sem_col],
            "Kinerja Terakhir (%)":latest[metric_col],
            "Semester Sebelumnya":prev[sem_col] if prev is not None else "-",
            "Kinerja Sebelumnya (%)":prev[metric_col] if prev is not None else np.nan,
            "Perubahan (poin)":delta,
        })
    out=pd.DataFrame(rows)
    if out.empty:return out
    out["Tren"]=out["Perubahan (poin)"].apply(tren_label)
    out["Status"]=out["Kinerja Terakhir (%)"].apply(status)
    return out.sort_values("Kinerja Terakhir (%)",ascending=False)

def render_prodi_section(f, prodi_col, sem_col, kin_col):
    st.subheader("🏫 Kinerja per Program Studi")
    pt=prodi_summary(f,prodi_col,sem_col,kin_col)
    if pt.empty:
        st.info("Data program studi tidak tersedia untuk filter saat ini.")
        return
    show=pt[["Prodi","Semester Terakhir","Kinerja Terakhir (%)","Semester Sebelumnya","Kinerja Sebelumnya (%)","Perubahan (poin)","Tren","Status"]].round(1)
    st.dataframe(show,use_container_width=True,hide_index=True)
    tren_colors={"▲ Naik":"#16a34a","▼ Turun":"#dc2626","— Tetap":"#94a3b8","N/A":"#cbd5e1"}
    pt_plot=pt.copy();pt_plot["Perubahan (poin)"]=pt_plot["Perubahan (poin)"].round(1)
    figp=px.bar(pt_plot,x="Prodi",y="Kinerja Terakhir (%)",color="Tren",
                color_discrete_map=tren_colors,category_orders={"Tren":["▲ Naik","▼ Turun","— Tetap","N/A"]},
                title="Kinerja Prodi (Semester Terakhir) & Perubahan vs Semester Sebelumnya",
                hover_data={"Perubahan (poin)":True})
    figp.update_yaxes(range=[0,100]);figp.update_layout(legend_title_text="Tren")
    st.plotly_chart(figp,use_container_width=True)
    naik=int((pt["Perubahan (poin)"]>0).sum());turun=int((pt["Perubahan (poin)"]<0).sum());tetap=int((pt["Perubahan (poin)"]==0).sum())
    st.markdown(f'<span class="small-note">📈 {naik} prodi naik · 📉 {turun} prodi turun · ➖ {tetap} prodi tetap dibanding semester sebelumnya.</span>',unsafe_allow_html=True)

def page_perkuliahan():
    df=read_excel_auto(PER_FILE)
    if df.empty: st.error(f"File {PER_FILE} tidak ditemukan.");return
    name=find_col(df,["NAMA DOSEN","Nama Dosen"]); nip=find_col(df,["NIP DOSEN","ID Dosen"])
    sem=find_col(df,["SEMESTER"]); fak=find_col(df,["FAKULTAS"]); prodi=find_col(df,["PRODI"]); kamp=find_col(df,["KAMPUS"])
    hadir=find_col(df,["% KEHADIRAN"]); tepat=find_col(df,["% KETEPATAN"]); kin=find_col(df,["% KINERJA"])
    for c in [hadir,tepat,kin]:
        if c: df[c]=pct_series(df[c])
    f=filters(df,[("Semester",sem),("Kampus",kamp),("Fakultas",fak),("Program Studi",prodi)])
    st.markdown('<div class="hero"><span class="badge-green">PERKULIAHAN</span><h1>Kinerja Perkuliahan Dosen</h1><div class="small-note">Kehadiran, ketepatan waktu mengajar, tren semester, ranking, dan profil individual dosen.</div></div>',unsafe_allow_html=True)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Dosen",f[name].nunique());c2.metric("Kehadiran",f"{f[hadir].mean():.1f}%");c3.metric("Ketepatan",f"{f[tepat].mean():.1f}%");c4.metric("Kinerja",f"{f[kin].mean():.1f}%")
    agg=f.groupby(name,as_index=False).agg(Kelas=(name,"size"),Kehadiran=(hadir,"mean"),Ketepatan=(tepat,"mean"),Kinerja=(kin,"mean"))
    st.subheader("🏆 Ranking Dosen")
    a,b=st.columns(2)
    with a: st.markdown("**Top 10 Kinerja Tertinggi**");st.dataframe(ranking_table(agg,"Kinerja")[["Peringkat",name,"Kelas","Kehadiran","Ketepatan","Kinerja"]].round(1),use_container_width=True,hide_index=True)
    with b: st.markdown("**Bottom 10 Kinerja Terendah**");st.dataframe(ranking_table(agg,"Kinerja",ascending=True)[["Peringkat",name,"Kelas","Kehadiran","Ketepatan","Kinerja"]].round(1),use_container_width=True,hide_index=True)
    st.subheader("📈 Tren Kinerja Semester")
    tr=f.groupby(sem,as_index=False)[[hadir,tepat,kin]].mean();tr["_k"]=tr[sem].map(semester_key);tr=tr.sort_values("_k")
    fig=px.line(tr,x=sem,y=[hadir,tepat,kin],markers=True);fig.update_yaxes(range=[0,100]);st.plotly_chart(fig,use_container_width=True)
    render_prodi_section(f,prodi,sem,kin)
    st.subheader("🔎 Cari Kinerja Dosen")
    q=st.text_input("Nama dosen",placeholder="Ketik nama dosen...",key="qper")
    names=sorted(f[name].dropna().astype(str).unique()); matches=[x for x in names if q.lower() in x.lower()] if q else []
    if matches:
        selected=st.selectbox("Pilih dosen",matches,key="sper");d=f[f[name].astype(str)==selected].copy()
        delta,latest,prev=trend_delta(df,name,sem,kin,selected)
        st.markdown(f"### 👤 {selected}")
        m1,m2,m3,m4=st.columns(4)
        m1.metric("% Kehadiran",f"{d[hadir].mean():.1f}%");m2.metric("% Ketepatan",f"{d[tepat].mean():.1f}%")
        m3.metric("% Kinerja",f"{d[kin].mean():.1f}%",None if pd.isna(delta) else f"{delta:+.1f} poin vs semester sebelumnya")
        m4.metric("Kategori",status(d[kin].mean()))
        prof=df[df[name].astype(str)==selected].groupby(sem,as_index=False)[[hadir,tepat,kin]].mean();prof["_k"]=prof[sem].map(semester_key);prof=prof.sort_values("_k")
        fig=px.bar(prof,x=sem,y=[hadir,tepat,kin],barmode="group",title="Profil Kinerja Antarsemester");fig.update_yaxes(range=[0,100]);st.plotly_chart(fig,use_container_width=True)
        st.dataframe(d,use_container_width=True,hide_index=True)
        st.download_button("⬇️ Download hasil pencarian Excel",excel_bytes(d,"Kinerja Perkuliahan"),file_name=f"kinerja_perkuliahan_{selected}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif q: st.warning("Nama dosen tidak ditemukan.")

def page_ujian():
    df=read_excel_auto(UJI_FILE)
    if df.empty:st.error(f"File {UJI_FILE} tidak ditemukan.");return
    name=find_col(df,["Nama Dosen","NAMA DOSEN"]); sem=find_col(df,["Semester","SEMESTER"]);fak=find_col(df,["Fakultas"]);prodi=find_col(df,["Prodi"]);kamp=find_col(df,["Kampus"])
    upload=find_col(df,["% Upload Soal"]);hadir=find_col(df,["% Kinerja Kehadiran"]);entry=find_col(df,["% Entry Nilai"]);kin=find_col(df,["% KINERJA"])
    for c in [upload,hadir,entry,kin]:
        if c:df[c]=pct_series(df[c])
    f=filters(df,[("Semester",sem),("Kampus",kamp),("Fakultas",fak),("Program Studi",prodi)])
    st.markdown('<div class="hero"><span class="badge-yellow">UJIAN</span><h1>Kinerja Ujian Dosen (UTS/UAS)</h1><div class="small-note">Upload soal, kehadiran ujian, entry nilai, kinerja, tren semester, ranking, dan profil individual.</div></div>',unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5)
    c1.metric("Dosen",f[name].nunique());c2.metric("Upload Soal",f"{f[upload].mean():.1f}%");c3.metric("Kehadiran",f"{f[hadir].mean():.1f}%");c4.metric("Entry Nilai",f"{f[entry].mean():.1f}%");c5.metric("Kinerja",f"{f[kin].mean():.1f}%")
    agg=f.groupby(name,as_index=False).agg(Kelas=(name,"size"),Upload=(upload,"mean"),Kehadiran=(hadir,"mean"),Entry=(entry,"mean"),Kinerja=(kin,"mean"))
    st.subheader("🏆 Ranking Dosen")
    a,b=st.columns(2)
    with a:st.markdown("**Top 10 Kinerja Tertinggi**");st.dataframe(ranking_table(agg,"Kinerja")[["Peringkat",name,"Kelas","Upload","Kehadiran","Entry","Kinerja"]].round(1),use_container_width=True,hide_index=True)
    with b:st.markdown("**Bottom 10 Kinerja Terendah**");st.dataframe(ranking_table(agg,"Kinerja",ascending=True)[["Peringkat",name,"Kelas","Upload","Kehadiran","Entry","Kinerja"]].round(1),use_container_width=True,hide_index=True)
    st.subheader("📈 Tren Kinerja Semester")
    tr=f.groupby(sem,as_index=False)[[upload,hadir,entry,kin]].mean();tr["_k"]=tr[sem].map(semester_key);tr=tr.sort_values("_k")
    fig=px.line(tr,x=sem,y=[upload,hadir,entry,kin],markers=True);fig.update_yaxes(range=[0,100]);st.plotly_chart(fig,use_container_width=True)
    render_prodi_section(f,prodi,sem,kin)
    st.subheader("🔎 Cari Kinerja Dosen")
    q=st.text_input("Nama dosen",placeholder="Ketik nama dosen...",key="quji")
    names=sorted(f[name].dropna().astype(str).unique());matches=[x for x in names if q.lower() in x.lower()] if q else []
    if matches:
        selected=st.selectbox("Pilih dosen",matches,key="suji");d=f[f[name].astype(str)==selected].copy()
        delta,latest,prev=trend_delta(df,name,sem,kin,selected)
        st.markdown(f"### 👤 {selected}")
        a,b,c,d1,e=st.columns(5)
        a.metric("% Upload Soal",f"{d[upload].mean():.1f}%");b.metric("% Kinerja Kehadiran",f"{d[hadir].mean():.1f}%");c.metric("% Entry Nilai",f"{d[entry].mean():.1f}%")
        d1.metric("% Kinerja",f"{d[kin].mean():.1f}%",None if pd.isna(delta) else f"{delta:+.1f} poin vs semester sebelumnya");e.metric("Kategori",status(d[kin].mean()))
        prof=df[df[name].astype(str)==selected].groupby(sem,as_index=False)[[upload,hadir,entry,kin]].mean();prof["_k"]=prof[sem].map(semester_key);prof=prof.sort_values("_k")
        fig=px.bar(prof,x=sem,y=[upload,hadir,entry,kin],barmode="group",title="Profil Kinerja Antarsemester");fig.update_yaxes(range=[0,100]);st.plotly_chart(fig,use_container_width=True)
        st.dataframe(d,use_container_width=True,hide_index=True)
        st.download_button("⬇️ Download hasil pencarian Excel",excel_bytes(d,"Kinerja Ujian"),file_name=f"kinerja_ujian_{selected}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif q:st.warning("Nama dosen tidak ditemukan.")

st.sidebar.title("📊 Dasbor Internal")
page=st.sidebar.radio("Pilih Dasbor",["Beranda","Kinerja Perkuliahan","Kinerja Ujian"])
st.sidebar.caption("Pembaruan data: ganti file Excel lama di repository dengan file baru menggunakan nama file yang sama.")
if page=="Beranda":
    st.markdown("""<div style="text-align:center;padding:55px 10px 25px"><div class="small-note">DASBOR INTERNAL · KINERJA DOSEN</div><h1 style="font-size:3rem">Pilih Dasbor</h1><p>Kinerja perkuliahan dan kinerja ujian dosen. Data dapat diperbarui cukup dengan mengganti file Excel sumber.</p></div>""",unsafe_allow_html=True)
    a,b=st.columns(2)
    with a:
        st.info("### 📚 Kinerja Perkuliahan\nKehadiran, ketepatan waktu, % kinerja, ranking, tren dan profil dosen.")
        if st.button("Buka Kinerja Perkuliahan",use_container_width=True):st.session_state.nav="per";st.rerun()
    with b:
        st.warning("### 📝 Kinerja Ujian (UTS/UAS)\nUpload soal, kehadiran, entry nilai, % kinerja, ranking, tren dan profil dosen.")
        if st.button("Buka Kinerja Ujian",use_container_width=True):st.session_state.nav="uji";st.rerun()
    if st.session_state.get("nav")=="per":page_perkuliahan()
    elif st.session_state.get("nav")=="uji":page_ujian()
elif page=="Kinerja Perkuliahan":page_perkuliahan()
else:page_ujian()
