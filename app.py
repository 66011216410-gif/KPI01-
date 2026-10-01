import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title='KPI01', page_icon='📊', layout='wide')
LEVEL=['ระดับ','ระดับการศึกษา','degree']; BASE=['ฐานข้อมูล','ฐานข้อมูล ']; GROUP=['กลุ่มสาขา']; FACULTY=['คณะ','faculty']; PROGRAM_TOTAL=['สาขารวม','สาขา รวม','program total']; PROGRAM=['สาขา','สาขาวิชา','หลักสูตร','program','major']; SYSTEM=['ระบบ']; STUDENT=['รหัสนิสิต','รหัสนักศึกษา','รหัส','student id']

def clean(v): return '' if pd.isna(v) else str(v).strip()
def norm_col(v): return re.sub(r'\s+','',clean(v)).lower()
def find_col(df,aliases):
    cols={norm_col(c):c for c in df.columns}
    for a in aliases:
        if norm_col(a) in cols:return cols[norm_col(a)]
    for c in df.columns:
        if any(norm_col(a) in norm_col(c) for a in aliases):return c
    return None

def norm_level(v):
    s=clean(v).lower().replace(' ','')
    if any(x in s for x in ['ปริญญาเอก','ป.เอก','phd','ph.d','doctoral']):return 'ปริญญาเอก'
    if any(x in s for x in ['ปริญญาโท','ป.โท','master']):return 'ปริญญาโท'
    return ''

def level_from_sheet(name):
    s=clean(name).lower().replace(' ','')
    if any(x in s for x in ['เอก','ปริญญาเอก','phd','doctoral']):return 'ปริญญาเอก'
    if any(x in s for x in ['โท','ปริญญาโท','master']):return 'ปริญญาโท'
    return ''

def norm_q(v):
    s=clean(v).upper().replace('SCOPUS(Q','SCOPUS (Q')
    if 'SCOPUS (Q1)' in s:return 'Q1'
    if 'SCOPUS (Q2)' in s:return 'Q2'
    return ''

def norm_system(v):
    s=re.sub(r'\s+','',clean(v)).lower()
    if s.startswith('ระบบ'):s=s[4:]
    if any(x in s for x in ['นอกเวลาราชการ','นอกเวลา','parttime','part-time']):return 'นอกเวลา'
    if any(x in s for x in ['ในเวลาราชการ','ในเวลา','fulltime','full-time']):return 'ในเวลา'
    return ''

def prepare_sheet(df,sheet_name):
    out=df.copy();out.columns=[clean(c) for c in out.columns]
    level_col=find_col(out,LEVEL);base_col=find_col(out,BASE);group_col=find_col(out,GROUP);faculty_col=find_col(out,FACULTY);program_total_col=find_col(out,PROGRAM_TOTAL);program_col=find_col(out,PROGRAM);system_col=find_col(out,SYSTEM);student_col=find_col(out,STUDENT)
    required=[]
    if not base_col:required.append('ฐานข้อมูล')
    if not group_col:required.append('กลุ่มสาขา')
    if not system_col:required.append('ระบบ')
    if required:raise ValueError(f"Sheet '{sheet_name}' ไม่พบคอลัมน์: {', '.join(required)}")
    forced=level_from_sheet(sheet_name)
    level=pd.Series(forced,index=out.index) if forced else (out[level_col].map(norm_level) if level_col else pd.Series('',index=out.index))
    group=out[group_col].map(clean).replace('',pd.NA).ffill().fillna('')
    if faculty_col:
        raw=out[faculty_col].map(clean).replace('',pd.NA);vals=[];lg=None;lf=''
        for g,f in zip(group,raw):
            g=clean(g)
            if g!=lg:lg,lf=g,''
            if pd.notna(f) and clean(f):lf=clean(f)
            vals.append(lf)
        faculty=pd.Series(vals,index=out.index)
    else:faculty=pd.Series('',index=out.index)
    source=out[program_total_col] if program_total_col else (out[program_col] if program_col else None)
    if source is not None:
        raw=source.map(clean).replace('',pd.NA);vals=[];lg=None;lf=None;lp=''
        for g,f,p in zip(group,faculty,raw):
            g,f=clean(g),clean(f)
            if g!=lg or f!=lf:lg,lf,lp=g,f,''
            if pd.notna(p) and clean(p):lp=clean(p)
            vals.append(lp)
        program=pd.Series(vals,index=out.index)
    else:program=pd.Series('',index=out.index)
    faculty_group={};program_group={}
    for g,f,p in zip(group,faculty,program):
        g,f,p=clean(g),clean(f),clean(p)
        if g and f and f not in faculty_group:faculty_group[f]=g
        if g and p and p not in program_group:program_group[p]=g
    canonical=[]
    for g,f,p in zip(group,faculty,program):
        g,f,p=clean(g),clean(f),clean(p)
        if f and f in faculty_group:g=faculty_group[f]
        elif p and p in program_group:g=program_group[p]
        canonical.append(g)
    out['__ระดับ']=level;out['__กลุ่ม']=canonical;out['__คณะ']=faculty.map(clean);out['__สาขา']=program.map(clean);out['__Q']=out[base_col].map(norm_q);out['__ระบบ']=out[system_col].map(norm_system);out['__รหัส']=out[student_col].map(clean) if student_col else '';out['__ลำดับ']=range(len(out))
    return out[out['__ระดับ'].isin(['ปริญญาโท','ปริญญาเอก'])].copy()

def read_excel(uploaded):
    uploaded.seek(0);sheets=pd.read_excel(uploaded,sheet_name=None,header=1);frames=[];used=[];errors=[]
    for name,df in sheets.items():
        if df is None or df.empty:continue
        try:
            p=prepare_sheet(df,name)
            if not p.empty:frames.append(p);used.append(name)
        except ValueError as e:errors.append(str(e))
    if not frames:raise ValueError('ไม่พบ Sheet ข้อมูลปริญญาโท/ปริญญาเอกที่ใช้งานได้'+(f": {'; '.join(errors[:3])}" if errors else ''))
    return used,pd.concat(frames,ignore_index=True)

def count_system(df,kind):
    x=df[df['__ระบบ']==kind]
    if x.empty:return 0
    ids=x['__รหัส'].map(clean)
    return int(ids[ids!=''].nunique()+int((ids=='').sum()))

def metrics(df):
    a=count_system(df,'ในเวลา')+count_system(df,'นอกเวลา');pub=int(df['__Q'].isin(['Q1','Q2']).sum());return count_system(df,'ในเวลา'),count_system(df,'นอกเวลา'),a,pub,round(pub*100/a,2) if a else 0

ORDER_TREE=[
('กลุ่มมนุษยศาสตร์และสังคมศาสตร์',[('การเมืองการปกครอง',['รัฐศาสตร์']),('การท่องเที่ยวและการโรงแรม',['การจัดการการท่องเที่ยวและการโรงแรม']),('การบัญชีและการจัดการ',['การจัดการสมัยใหม่','การจัดการสมาร์ตซิตี้และนวัตกรรมดิจิทัล','การบัญชี','บริหารธุรกิจและนวัตกรรมดิจิทัล']),('ดุริยางคศิลป์',['ดุริยางคศิลป์']),('มนุษยศาสตร์และสังคมศาสตร์',['การสอนภาษาอังกฤษ','ภาษาไทย','ศาสนาและภูมิปัญญาเพื่อการพัฒนา']),('ศิลปกรรมศาสตร์และวัฒนธรรมศาสตร์',['การวิจัยและสร้างสรรค์ศิลปกรรมศาสตร์','วัฒนธรรมศาสตร์']),('ศึกษาศาสตร์',['เทคโนโลยีและสื่อสารการศึกษา','การบริหารและพัฒนาการศึกษา','วิจัยและประเมินผลการศึกษา','วิทยาศาสตร์การออกกำลังกายและการกีฬา','หลักสูตรและการสอน'])]),
('กลุ่มวิทยาศาสตร์และเทคโนโลยี',[('เทคโนโลยี',['เกษตรศาสตร์','เทคโนโลยีการอาหาร']),('วิจัยวลัยรุกขเวช',['ความหลากหลายทางชีวภาพ']),('วิทยาการสารสนเทศ',['เทคโนโลยีสารสนเทศ','วิทยาการคอมพิวเตอร์','สื่อนฤมิต']),('วิทยาศาสตร์',['บรรพชีวินวิทยา','ฟิสิกส์']),('วิศวกรรมศาสตร์',['วิศวกรรมเครื่องกล','วิศวกรรมโยธา','วิศวกรรมไฟฟ้าและคอมพิวเตอร์']),('สิ่งแวดล้อมและทรัพยากรศาสตร์',['การจัดการสิ่งแวดล้อมอย่างยั่งยืน'])]),
('กลุ่มวิทยาศาสตร์สุขภาพ',[('เภสัชศาสตร์',['เภสัชศาสตร์']),('แพทยศาสตร์',['วิทยาศาสตร์สุขภาพ']),('สาธารณสุขศาสตร์',['เทคโนโลยีทางสุขภาพและความปลอดภัย','สาธารณสุขศาสตรดุษฎีบัณฑิต'])])]

def hierarchy(data):
    cols=['ลำดับ','กลุ่ม/คณะ/สาขา','ระบบในเวลาราชการ','ระบบนอกเวลาราชการ','จำนวนผู้สำเร็จการศึกษา (A)','รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2','ร้อยละ'];rows=[];seen_g=set();seen_f=set();seen_p=set();n=1
    def add_group(g):
        nonlocal n
        if not g or g in seen_g:return
        m=data['__กลุ่ม'].map(clean).eq(g)
        if m.any():seen_g.add(g);rows.append([n,g,*metrics(data[m])]);n+=1
    def add_faculty(g,f):
        nonlocal n
        key=(g,f);m=data['__กลุ่ม'].map(clean).eq(g)&data['__คณะ'].map(clean).eq(f)
        if f and key not in seen_f and m.any():seen_f.add(key);rows.append([n,'    '+f,*metrics(data[m])]);n+=1
    def add_program(g,f,p):
        nonlocal n
        key=(g,f,p);m=data['__กลุ่ม'].map(clean).eq(g)&data['__คณะ'].map(clean).eq(f)&data['__สาขา'].map(clean).eq(p)
        if p and key not in seen_p and m.any():seen_p.add(key);rows.append([n,'        '+p,*metrics(data[m])]);n+=1
    for g,fs in ORDER_TREE:
        add_group(g)
        for f,ps in fs:
            add_faculty(g,f)
            for p in ps:add_program(g,f,p)
    for _,r in data.sort_values('__ลำดับ').iterrows():
        g,f,p=clean(r['__กลุ่ม']),clean(r['__คณะ']),clean(r['__สาขา']);add_group(g)
        if g and f:add_faculty(g,f)
        if g and f and p:add_program(g,f,p)
    rows.append(['','รวมทั้งหมด',*metrics(data)])
    return pd.DataFrame(rows,columns=cols)

def q_summary(data):
    def vals(level):
        x=data[data['__ระดับ']==level];q1=int((x['__Q']=='Q1').sum());q2=int((x['__Q']=='Q2').sum());return q1,q2,q1+q2
    m,d=vals('ปริญญาโท'),vals('ปริญญาเอก');return pd.DataFrame([['ระดับปริญญาโท',*m],['ระดับปริญญาเอก',*d],['รวม',m[0]+d[0],m[1]+d[1],m[2]+d[2]]],columns=['','Scopus Q1','Scopus Q2','รวม'])

def excel_bytes(master,doctor,summary):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as w:
        master.to_excel(w,sheet_name='ปริญญาโท',index=False,startrow=6,header=False)
        doctor.to_excel(w,sheet_name='ปริญญาเอก',index=False,startrow=6,header=False)
        summary.to_excel(w,sheet_name='สรุป Scopus Q1-Q2',index=False,startrow=1)
    wb=load_workbook(bio);dark='2F5D1E';light='5A8C3A';white='FFFFFF';thin=Side(style='thin',color='808080');border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for ws,level in [(wb['ปริญญาโท'],'ปริญญาโท'),(wb['ปริญญาเอก'],'ปริญญาเอก')]:
        # หัวตารางกลับเป็นรูปแบบเดิม: หัวหลักหลายแถวและ merge เหมือนแบบตัวอย่าง
        ws.merge_cells('A1:G1');ws['A1']=f'ผลงานของนักศึกษาและผู้สำเร็จการศึกษาในระดับ{level}'
        merges=['A2:A6','B2:B6','C2:D2','C3:C6','D3:D6','E2:E6','F2:G2','F3:F6','G3:G6']
        for rng in merges:ws.merge_cells(rng)
        labels={'A2':'ลำดับ','B2':'กลุ่ม/คณะ/สาขา','C2':'ระบบ','C3':'ระบบในเวลา\nราชการ','D3':'ระบบนอกเวลา\nราชการ','E2':'จำนวนผู้สำเร็จ\nการศึกษา (A)','F2':'1.2.4 จำนวนผลงานระดับบัณฑิตศึกษาที่สามารถตีพิมพ์ใน\nระดับนานาชาติ SOPUS Q1-Q2','F3':'รวมจำนวนผลงาน\nตีพิมพ์ในวารสารระดับ\nนานาชาติ Q1-2','G3':'ร้อยละผลงานวิจัยรวมจำนวน\nผลงานตีพิมพ์ในวารสารระดับ\nนานาชาติ Q1-2'}
        for cell,val in labels.items():ws[cell]=val
        for row in ws.iter_rows(min_row=1,max_row=ws.max_row,min_col=1,max_col=7):
            for c in row:
                c.font=Font(name='Tahoma',size=11)
                c.alignment=Alignment(vertical='center',horizontal='center',wrap_text=True)
                c.border=border
        for c in ws[1]:c.fill=PatternFill('solid',fgColor=dark);c.font=Font(name='Tahoma',size=14,bold=True,color=white)
        for r in range(2,7):
            for c in ws[r]:c.fill=PatternFill('solid',fgColor=dark if r<=2 else light);c.font=Font(name='Tahoma',size=11,bold=True,color=white);c.alignment=Alignment(vertical='center',horizontal='center',wrap_text=True);c.border=border
        for r in range(7,ws.max_row+1):
            text=str(ws.cell(r,2).value or '');fill='FCE4D6' if len(text)-len(text.lstrip())==0 else 'E2F0D9'
            for c in ws[r]:c.fill=PatternFill('solid',fgColor=fill);c.border=border;c.font=Font(name='Tahoma',size=11);c.alignment=Alignment(vertical='center',wrap_text=True)
            ws.cell(r,2).alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
            for c in range(1,8):ws.cell(r,c).alignment=Alignment(horizontal='right' if c>=3 else ('center' if c==1 else 'left'),vertical='center',wrap_text=True)
            ws.cell(r,7).number_format='0.00'
        ws.column_dimensions['A'].width=9;ws.column_dimensions['B'].width=55;ws.column_dimensions['C'].width=18;ws.column_dimensions['D'].width=19;ws.column_dimensions['E'].width=20;ws.column_dimensions['F'].width=24;ws.column_dimensions['G'].width=27
        ws.row_dimensions[1].height=30
        for r in range(2,7):ws.row_dimensions[r].height=42
        ws.freeze_panes='A7';ws.sheet_view.showGridLines=False
    ws=wb['สรุป Scopus Q1-Q2'];ws['A1']='1.2.4) จำนวนผลงานระดับบัณฑิตศึกษาที่ตีพิมพ์ในระดับนานาชาติ Scopus Q1, Q2'
    for c in ws[2]:c.fill=PatternFill('solid',fgColor=dark);c.font=Font(name='Tahoma',size=12,bold=True,color=white);c.alignment=Alignment(horizontal='center');c.border=border
    for row in ws.iter_rows(min_row=3,max_row=5,min_col=1,max_col=4):
        for c in row:c.font=Font(name='Tahoma',size=12);c.alignment=Alignment(horizontal='center' if c.column>1 else 'left');c.border=border
    ws.column_dimensions['A'].width=28;ws.column_dimensions['B'].width=18;ws.column_dimensions['C'].width=18;ws.column_dimensions['D'].width=18;ws.sheet_view.showGridLines=False
    out=io.BytesIO();wb.save(out);out.seek(0);return out

st.title('📊 KPI01 — ตารางสถิติ');st.caption('ไฟล์ข้อมูลสามารถแยกเป็น 2 Sheet: ปริญญาโท และ ปริญญาเอก | ระบบไม่นับรหัสซ้ำ | A = ในเวลา + นอกเวลา | Q1-Q2 = SCOPUS (Q1)/(Q2)');uploaded=st.file_uploader('อัปโหลด Excel',type=['xlsx','xls'])
if uploaded:
    try:
        sheets,data=read_excel(uploaded);master=hierarchy(data[data['__ระดับ']=='ปริญญาโท']);doctor=hierarchy(data[data['__ระดับ']=='ปริญญาเอก']);summary=q_summary(data);st.success('อ่านข้อมูลสำเร็จ: '+', '.join(sheets));st.subheader('ปริญญาโท');st.dataframe(master,use_container_width=True,hide_index=True);st.subheader('ปริญญาเอก');st.dataframe(doctor,use_container_width=True,hide_index=True);st.subheader('สรุป Scopus Q1-Q2');st.dataframe(summary,use_container_width=True,hide_index=True);st.download_button('📥 ดาวน์โหลด Excel KPI01',excel_bytes(master,doctor,summary),'KPI01.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:st.error(f'ไม่สามารถอ่านไฟล์ได้: {e}')
else:st.info('อัปโหลด Excel เพื่อสร้างตารางสถิติ')