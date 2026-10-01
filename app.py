import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01", page_icon="📊", layout="wide")
LEVEL=["ระดับ","ระดับการศึกษา","degree"]; Q_MAIN=["ผลงานที่ตีพิมพ์ Q1-Q2","ผลงานที่ตีพิมพ์Q1-Q2"]; Q_FALLBACK=["ฐานข้อมูล"]
GROUP=["กลุ่ม","group"]; FACULTY=["คณะ","faculty"]; PROGRAM_TOTAL=["สาขารวม","สาขา รวม","program total"]; PROGRAM=["สาขา","สาขาวิชา","หลักสูตร","program","major"]
SYSTEM=["ระบบ"]; STUDENT=["รหัสนิสิต","รหัสนักศึกษา","รหัส","student id"]; STATUS=["สถานะนิสิต","สถานะนักศึกษา","สถานะ","status"]

def clean(v): return "" if pd.isna(v) else str(v).strip()
def norm_col(s): return re.sub(r"\s+","",clean(s)).lower()
def find_col(df,aliases):
    cols={norm_col(c):c for c in df.columns}
    for a in aliases:
        if norm_col(a) in cols:return cols[norm_col(a)]
    for c in df.columns:
        if any(norm_col(a) in norm_col(c) for a in aliases):return c
    return None

def norm_level(v):
    s=clean(v).lower().replace(" ","")
    if any(x in s for x in ["ปริญญาเอก","ป.เอก","phd","ph.d","doctoral"]):return "ปริญญาเอก"
    if any(x in s for x in ["ปริญญาโท","ป.โท","master"]):return "ปริญญาโท"
    return "อื่น ๆ"
def norm_q(v):
    s=clean(v).upper().replace(" ","")
    if "Q1" in s:return "Q1"
    if "Q2" in s:return "Q2"
    return "อื่น ๆ"
def norm_system(v):
    s=clean(v).lower();s=re.sub(r"\s+","",s)
    s=s.replace("ระบบ","",1) if s.startswith("ระบบ") else s
    if not s:return ""
    if any(x in s for x in ["นอกเวลาราชการ","นอกเวลา","นอกเวล","parttime","part-time"]):return "นอกเวลา"
    if any(x in s for x in ["ในเวลาราชการ","ในเวลา","ในเวล","fulltime","full-time"]):return "ในเวลา"
    return ""
def is_success(v):
    s=clean(v).lower()
    if not s:return True
    return any(x in s for x in ["สำเร็จการศึกษา","สำเร็จ","graduate","graduated"])

def read_excel(uploaded):
    uploaded.seek(0); sheets=pd.read_excel(uploaded,sheet_name=None,header=1); best=None
    for name,df in sheets.items():
        if df is None or df.empty:continue
        df=df.copy();df.columns=[clean(c) for c in df.columns];score=0
        if find_col(df,LEVEL):score+=10
        if find_col(df,Q_MAIN):score+=10
        elif find_col(df,Q_FALLBACK):score+=5
        if find_col(df,SYSTEM):score+=5
        if find_col(df,FACULTY):score+=2
        if find_col(df,PROGRAM):score+=2
        if best is None or score>best[0]:best=(score,name,df)
    if best is None:raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    return best[1],best[2]

def prepare(df):
    out=df.copy();level_col=find_col(out,LEVEL);q_main=find_col(out,Q_MAIN);q_fallback=find_col(out,Q_FALLBACK);q_col=q_main or q_fallback
    group_col=find_col(out,GROUP);faculty_col=find_col(out,FACULTY);program_total_col=find_col(out,PROGRAM_TOTAL);program_col=find_col(out,PROGRAM);system_col=find_col(out,SYSTEM);student_col=find_col(out,STUDENT);status_col=find_col(out,STATUS)
    if not level_col:raise ValueError("ไม่พบคอลัมน์ 'ระดับ' ในแถวที่ 2")
    if not q_col:raise ValueError("ไม่พบคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2' และไม่พบ 'ฐานข้อมูล'")
    if not system_col:raise ValueError("ไม่พบคอลัมน์ 'ระบบ' ในแถวที่ 2")
    out["__ระดับ"]=out[level_col].map(norm_level);out["__Q"]=out[q_col].map(norm_q);out["__ระบบ"]=out[system_col].map(norm_system)
    out["__รหัส"]=out[student_col].map(clean) if student_col else "";out["__สถานะ"]=out[status_col].map(is_success) if status_col else True;out["__กลุ่ม"]=out[group_col].map(clean) if group_col else "";out["__คณะ"]=out[faculty_col].map(clean) if faculty_col else "";out["__สาขา"]=out[program_total_col].map(clean) if program_total_col else (out[program_col].map(clean) if program_col else "");out["__ลำดับ"]=range(len(out))
    out=out[out["__ระดับ"].isin(["ปริญญาโท","ปริญญาเอก"])].copy()
    return out,{"ระดับ":level_col,"Q1/Q2 ที่ใช้":q_col,"แหล่ง Q1/Q2":"ผลงานที่ตีพิมพ์ Q1-Q2" if q_main else "ฐานข้อมูล","ระบบ":system_col,"กลุ่ม":group_col,"คณะ":faculty_col,"สาขาที่ใช้":program_total_col or program_col,"รหัสนิสิต":student_col,"สถานะ":status_col}

def unique_students(df):
    if df.empty:return 0
    ids=df["__รหัส"].astype(str).str.strip()
    return int(ids[ids!=""].nunique()) if ids.ne("").any() else len(df)
def metrics(df):
    # ระบบใน/นอก = จำนวนแถวจากคอลัมน์ ระบบโดยตรง
    inn=int((df["__ระบบ"]=="ในเวลา").sum())
    outn=int((df["__ระบบ"]=="นอกเวลา").sum())
    # จำนวนผู้สำเร็จการศึกษา (A) = ระบบในเวลา + ระบบนอกเวลา
    a=inn+outn
    x=df[df["__สถานะ"]]
    pub=int(x["__Q"].isin(["Q1","Q2"]).sum())
    pct=round(pub*100/a,2) if a else 0
    return inn,outn,a,pub,pct

def hierarchy(data):
    cols=["ลำดับ","กลุ่ม/คณะ/สาขา","ระบบในเวลาราชการ","ระบบนอกเวลาราชการ","จำนวนผู้สำเร็จการศึกษา (A)","รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2","ร้อยละ"];rows=[];seen=set();n=1
    levels=[("__กลุ่ม",0),("__คณะ",1),("__สาขา",2)]
    for _,r in data.sort_values("__ลำดับ").iterrows():
        path=[]
        for col,indent in levels:
            val=clean(r[col])
            if not val:continue
            path.append(val);key=(col,tuple(path))
            if key in seen:continue
            seen.add(key);mask=pd.Series(True,index=data.index)
            for pcol,_ in levels[:indent+1]:
                pv=clean(r[pcol])
                if pv: mask &= data[pcol].map(clean).eq(pv)
            rows.append([n,"    "*indent+val,*metrics(data[mask])]);n+=1
    if rows:rows.append(["","รวมทั้งหมด",*metrics(data)])
    return pd.DataFrame(rows,columns=cols)

def excel_bytes(master,doctor):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as w:
        master.to_excel(w,sheet_name="ปริญญาโท",index=False,startrow=3);doctor.to_excel(w,sheet_name="ปริญญาเอก",index=False,startrow=3)
    bio.seek(0);return bio

st.title("📊 KPI01 — ตารางสถิติ")
st.caption("หัวคอลัมน์อยู่แถวที่ 2 | ระดับจาก 'ระดับ' | Q1/Q2 จาก 'ผลงานที่ตีพิมพ์ Q1-Q2' หรือ 'ฐานข้อมูล' | ระบบนับจำนวนแถวจาก 'ระบบ' | A = ในเวลา + นอกเวลา")
uploaded=st.file_uploader("อัปโหลด Excel",type=["xlsx","xls"])
if uploaded:
    try:
        sheet,raw=read_excel(uploaded);data,info=prepare(raw);master=hierarchy(data[data["__ระดับ"]=="ปริญญาโท"]);doctor=hierarchy(data[data["__ระดับ"]=="ปริญญาเอก"])
        st.success(f"อ่าน Sheet: {sheet} | ข้อมูล: {len(data):,} แถว")
        st.subheader("ปริญญาโท");st.dataframe(master,use_container_width=True,hide_index=True)
        st.subheader("ปริญญาเอก");st.dataframe(doctor,use_container_width=True,hide_index=True)
        st.download_button("📥 ดาวน์โหลด Excel KPI01",excel_bytes(master,doctor),"KPI01.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:st.error(f"ไม่สามารถอ่านไฟล์ได้: {e}")
else:st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ")
