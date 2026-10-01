import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title='KPI01', page_icon='📊', layout='wide')
LEVEL=['ระดับ','ระดับการศึกษา','degree']
Q_MAIN=['ฐานข้อมูล','ฐานข้อมูล ']
GROUP=['กลุ่ม','group']; FACULTY=['คณะ','faculty']
PROGRAM_TOTAL=['สาขารวม','สาขา รวม','program total']; PROGRAM=['สาขา','สาขาวิชา','หลักสูตร','program','major']
SYSTEM=['ระบบ']; STUDENT=['รหัสนิสิต','รหัสนักศึกษา','รหัส','student id']; STATUS=['สถานะนิสิต','สถานะนักศึกษา','สถานะ','status']

def clean(v): return '' if pd.isna(v) else str(v).strip()
def norm_col(s): return re.sub(r'\s+','',clean(s)).lower()
def find_col(df, aliases):
    cols={norm_col(c):c for c in df.columns}
    for a in aliases:
        if norm_col(a) in cols: return cols[norm_col(a)]
    for c in df.columns:
        if any(norm_col(a) in norm_col(c) for a in aliases): return c
    return None

def norm_level(v):
    s=clean(v).lower().replace(' ','')
    if any(x in s for x in ['ปริญญาเอก','ป.เอก','phd','ph.d','doctoral']): return 'ปริญญาเอก'
    if any(x in s for x in ['ปริญญาโท','ป.โท','master']): return 'ปริญญาโท'
    return 'อื่น ๆ'

def norm_q(v):
    s=clean(v).upper()
    if 'SCOPUS (Q1)' in s or 'SCOPUS(Q1)' in s: return 'Q1'
    if 'SCOPUS (Q2)' in s or 'SCOPUS(Q2)' in s: return 'Q2'
    return 'อื่น ๆ'

def norm_system(v):
    s=re.sub(r'\s+','',clean(v)).lower()
    if s.startswith('ระบบ'): s=s[4:]
    if any(x in s for x in ['นอกเวลาราชการ','นอกเวลา','นอกเวล','parttime','part-time']): return 'นอกเวลา'
    if any(x in s for x in ['ในเวลาราชการ','ในเวลา','ในเวล','fulltime','full-time']): return 'ในเวลา'
    return ''

def read_excel(uploaded):
    uploaded.seek(0); sheets=pd.read_excel(uploaded,sheet_name=None,header=1); best=None
    for name,df in sheets.items():
        if df is None or df.empty: continue
        df=df.copy(); df.columns=[clean(c) for c in df.columns]; score=0
        if find_col(df,LEVEL): score+=10
        if find_col(df,Q_MAIN): score+=10
        if find_col(df,SYSTEM): score+=5
        if find_col(df,FACULTY): score+=2
        if find_col(df,PROGRAM): score+=2
        if best is None or score>best[0]: best=(score,name,df)
    if best is None: raise ValueError('ไม่พบข้อมูลในไฟล์ Excel')
    return best[1],best[2]

def prepare(df):
    out=df.copy()
    level_col=find_col(out,LEVEL); base_col=find_col(out,Q_MAIN)
    group_col=find_col(out,GROUP); faculty_col=find_col(out,FACULTY)
    program_total_col=find_col(out,PROGRAM_TOTAL); program_col=find_col(out,PROGRAM)
    system_col=find_col(out,SYSTEM)
    if not level_col: raise ValueError("ไม่พบคอลัมน์ 'ระดับ' ในแถวที่ 2")
    if not base_col: raise ValueError("ไม่พบคอลัมน์ 'ฐานข้อมูล' ในแถวที่ 2")
    if not system_col: raise ValueError("ไม่พบคอลัมน์ 'ระบบ' ในแถวที่ 2")

    # Excel ต้นฉบับมีการ Merge Cell ในคอลัมน์กลุ่ม/คณะ ทำให้แถวถัดลงมาว่าง
    # ต้องเติมค่าจากแถวก่อนหน้า เพื่อให้แต่ละนิสิตยังอยู่ภายใต้กลุ่มและคณะที่ถูกต้อง
    if group_col: out[group_col]=out[group_col].map(clean).replace('',pd.NA).ffill().fillna('')
    if faculty_col: out[faculty_col]=out[faculty_col].map(clean).replace('',pd.NA).ffill().fillna('')

    out['__ระดับ']=out[level_col].map(norm_level)
    out['__Q']=out[base_col].map(norm_q)
    out['__ระบบ']=out[system_col].map(norm_system)
    out['__กลุ่ม']=out[group_col].map(clean) if group_col else ''
    out['__คณะ']=out[faculty_col].map(clean) if faculty_col else ''
    out['__สาขา']=out[program_total_col].map(clean) if program_total_col else (out[program_col].map(clean) if program_col else '')
    out['__ลำดับ']=range(len(out))
    out=out[out['__ระดับ'].isin(['ปริญญาโท','ปริญญาเอก'])].copy()
    return out,{'ระดับ':level_col,'Q1/Q2 ที่ใช้':'ฐานข้อมูล','ระบบ':system_col,'กลุ่ม':group_col,'คณะ':faculty_col,'สาขาที่ใช้':program_total_col or program_col}

def metrics(df):
    inn=int((df['__ระบบ']=='ในเวลา').sum()); outn=int((df['__ระบบ']=='นอกเวลา').sum())
    a=inn+outn; pub=int(df['__Q'].isin(['Q1','Q2']).sum()); pct=round(pub*100/a,2) if a else 0
    return inn,outn,a,pub,pct

def hierarchy(data):
    cols=['ลำดับ','กลุ่ม/คณะ/สาขา','ระบบในเวลาราชการ','ระบบนอกเวลาราชการ','จำนวนผู้สำเร็จการศึกษา (A)','รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2','ร้อยละ']
    rows=[]; seen=set(); n=1
    levels=[('__กลุ่ม',0),('__คณะ',1),('__สาขา',2)]
    for _,r in data.sort_values('__ลำดับ').iterrows():
        path=[]
        for col,indent in levels:
            val=clean(r[col])
            if not val: continue
            path.append(val); key=(col,tuple(path))
            if key in seen: continue
            seen.add(key)
            mask=pd.Series(True,index=data.index)
            for pcol,_ in levels[:indent+1]:
                pv=clean(r[pcol])
                if pv: mask &= data[pcol].map(clean).eq(pv)
            rows.append([n,'    '*indent+val,*metrics(data[mask])]); n+=1
    if rows: rows.append(['','รวมทั้งหมด',*metrics(data)])
    return pd.DataFrame(rows,columns=cols)

def excel_bytes(master,doctor):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as w:
        master.to_excel(w,sheet_name='ปริญญาโท',index=False,startrow=6,header=False)
        doctor.to_excel(w,sheet_name='ปริญญาเอก',index=False,startrow=6,header=False)
    wb=load_workbook(bio)
    dark='2F5D1E'; light='5A8C3A'; white='FFFFFF'; peach='FCE4D6'; green='E2F0D9'; grid='808080'
    thin=Side(style='thin',color=grid); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    title='ผลงานของนักศึกษาและผู้สำเร็จการศึกษาในระดับ {level} ปีการศึกษา 2568 รอบ 12 เดือน ข้อมูล ระหว่างเดือนกรกฎาคม 2568 ถึง มิถุนายน 2569'
    for ws,level in [(wb['ปริญญาโท'],'ปริญญาโท'),(wb['ปริญญาเอก'],'ปริญญาเอก')]:
        ws.merge_cells('A1:G1'); ws['A1']=title.format(level=level)
        for rng in ['A2:A6','B2:B6','C2:D2','C3:C6','D3:D6','E2:E6','F2:G2','F3:F6','G3:G6']: ws.merge_cells(rng)
        labels={'A2':'ลำดับ','B2':'กลุ่ม/คณะ/สาขา','C2':'ระบบ','C3':'ระบบในเวลา\nราชการ','D3':'ระบบนอกเวลา\nราชการ','E2':'จำนวน\nผู้สำเร็จ\nการศึกษา (A)','F2':'1.2.4 จำนวนผลงานระดับบัณฑิตศึกษาที่สามารถตีพิมพ์ใน\nระดับนานาชาติ SOPUS Q1-Q2','F3':'รวมจำนวนผลงานตีพิมพ์\nในวารสารระดับ\nนานาชาติ Q1-2','G3':'ร้อยละผลงานวิจัยรวมจำนวน\nผลงานตีพิมพ์ในวารสารระดับ\nนานาชาติ Q1-2'}
        for cell,val in labels.items(): ws[cell]=val
        for row in range(2,7):
            for col in range(1,8):
                c=ws.cell(row,col); c.fill=PatternFill('solid',fgColor=dark if row<=3 else light); c.font=Font(name='Tahoma',size=12,bold=True,color=white); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=border
        for row in ws.iter_rows(min_row=7,max_row=ws.max_row,min_col=1,max_col=7):
            text=str(row[1].value or ''); leading=len(text)-len(text.lstrip()); indent=2 if leading>=8 else (1 if leading>=4 else 0)
            is_group=indent==0 and text.strip() not in ['รวมทั้งหมด','']; fill=peach if is_group or text.strip()=='รวมทั้งหมด' else green
            for c in row:
                c.border=border; c.font=Font(name='Tahoma',size=12); c.alignment=Alignment(vertical='center',wrap_text=True); c.fill=PatternFill('solid',fgColor=fill)
            row[1].alignment=Alignment(horizontal='left',vertical='center',indent=indent,wrap_text=True)
            for c in row[2:]: c.alignment=Alignment(horizontal='right',vertical='center')
            row[6].number_format='0.00'
        for col,width in {'A':9,'B':52,'C':17,'D':19,'E':18,'F':23,'G':27}.items(): ws.column_dimensions[col].width=width
        ws.row_dimensions[1].height=30
        for r in range(2,7): ws.row_dimensions[r].height=38
        for r in range(7,ws.max_row+1): ws.row_dimensions[r].height=26
        ws.freeze_panes='A7'; ws.sheet_view.showGridLines=False
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

st.title('📊 KPI01 — ตารางสถิติ')
st.caption("หัวคอลัมน์อยู่แถวที่ 2 | ระดับจาก 'ระดับ' | Q1-Q2 นับจาก 'ฐานข้อมูล' เฉพาะ SCOPUS (Q1)/(Q2) | ระบบนับจำนวนแถวจาก 'ระบบ' | A = ในเวลา + นอกเวลา")
uploaded=st.file_uploader('อัปโหลด Excel',type=['xlsx','xls'])
if uploaded:
    try:
        sheet,raw=read_excel(uploaded); data,info=prepare(raw)
        master=hierarchy(data[data['__ระดับ']=='ปริญญาโท']); doctor=hierarchy(data[data['__ระดับ']=='ปริญญาเอก'])
        st.success(f'อ่าน Sheet: {sheet} | ข้อมูล: {len(data):,} แถว')
        st.subheader('ปริญญาโท'); st.dataframe(master,use_container_width=True,hide_index=True)
        st.subheader('ปริญญาเอก'); st.dataframe(doctor,use_container_width=True,hide_index=True)
        st.download_button('📥 ดาวน์โหลด Excel KPI01',excel_bytes(master,doctor),'KPI01.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e: st.error(f'ไม่สามารถอ่านไฟล์ได้: {e}')
else: st.info('อัปโหลด Excel เพื่อสร้างตารางสถิติ')
