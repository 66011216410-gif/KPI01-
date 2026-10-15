import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

PAGE_TITLE='ฐานข้อมูลจำนวนผลงานระดับบัณฑิตศึกษาที่ตีพิมพ์ในระดับนานาชาติ Scopus Q1, Q2'
st.set_page_config(page_title=PAGE_TITLE, page_icon='📊', layout='wide')
LEVEL=['ระดับ','ระดับการศึกษา','degree']; BASE=['ฐานข้อมูล','ฐานข้อมูล ']; PUB_TITLE=['ชื่อผลงานที่ตีพิมพ์','ชื่อผลงาน','ชื่อบทความ','ชื่อเรื่อง','ชื่อสิ่งพิมพ์','publication title','article title','title']; DOI_COL=['doi','digital object identifier']; GROUP=['กลุ่มสาขา']; FACULTY=['คณะ','faculty']; PROGRAM_TOTAL=['สาขารวม','สาขา รวม','program total']; PROGRAM=['สาขา','สาขาวิชา','หลักสูตร','program','major']; SYSTEM=['ระบบ']; STUDENT=['รหัสนิสิต','รหัสนักศึกษา','รหัส','student id']

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
    # นับเฉพาะรายการที่ระบุ SCOPUS Q1 หรือ SCOPUS Q2 ชัดเจน
    s=clean(v).upper()
    matches=re.findall(r'(?<![A-Z0-9])SCOPUS\s*[\(\[]?\s*Q\s*([12])(?![0-9])',s)
    if not matches:return ''
    # หากช่องมีทั้ง Q1 และ Q2 ให้เลือก Q ที่ระบุเป็นตัวแรกเพื่อไม่ให้นับแถวเดียวซ้ำ
    return 'Q'+matches[0]
def norm_system(v):
    s=re.sub(r'\s+','',clean(v)).lower()
    if s.startswith('ระบบ'):s=s[4:]
    if any(x in s for x in ['นอกเวลาราชการ','นอกเวลา','parttime','part-time']):return 'นอกเวลา'
    if any(x in s for x in ['ในเวลาราชการ','ในเวลา','fulltime','full-time']):return 'ในเวลา'
    return ''
def prepare_sheet(df,sheet_name):
    out=df.copy();out.columns=[clean(c) for c in out.columns]
    level_col=find_col(out,LEVEL);base_col=find_col(out,BASE);title_col=find_col(out,PUB_TITLE);doi_col=find_col(out,DOI_COL);group_col=find_col(out,GROUP);faculty_col=find_col(out,FACULTY);program_total_col=find_col(out,PROGRAM_TOTAL);program_col=find_col(out,PROGRAM);system_col=find_col(out,SYSTEM);student_col=find_col(out,STUDENT)
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
    # สร้างรหัสผลงานเพื่อป้องกันการนับรายการเดียวซ้ำข้ามแถว/ชีต/ไฟล์
    def key_text(v):return re.sub(r'\s+',' ',clean(v)).strip().casefold()
    if doi_col:
        doi=out[doi_col].map(key_text).str.replace(r'^https?://(dx\.)?doi\.org/','',regex=True).str.replace(r'^doi\s*:\s*','',regex=True).str.strip()
    else:doi=pd.Series('',index=out.index)
    if title_col:title=out[title_col].map(key_text)
    else:title=pd.Series('',index=out.index)
    # ใช้ DOI ก่อน ตามด้วยชื่อผลงาน; หากไม่มีทั้งคู่ใช้ข้อมูลแถวนั้นเป็น fallback
    row_cols=[c for c in out.columns if not str(c).startswith('__')]
    fallback=out[row_cols].fillna('').astype(str).apply(lambda row:'|'.join(key_text(v) for v in row),axis=1)
    out['__ผลงาน_key']=['doi:'+d if d else ('title:'+t if t else 'row:'+f) for d,t,f in zip(doi,title,fallback)]
    return out[out['__ระดับ'].isin(['ปริญญาโท','ปริญญาเอก'])].copy()
@st.cache_data(show_spinner=False)
def read_excel(file_bytes):
    # อ่าน workbook จาก bytes เพียงครั้งเดียว แล้วใช้ DataFrame ชุดเดิมทั้งการคำนวณและ export
    sheets=pd.read_excel(io.BytesIO(file_bytes),sheet_name=None,header=1)
    frames=[];used=[];raw_sheets={};errors=[]
    for name,df in sheets.items():
        if df is None or df.empty:continue
        try:
            p=prepare_sheet(df,name)
            if not p.empty:frames.append(p);used.append(name);raw_sheets[name]=df.copy()
        except ValueError as e:errors.append(str(e))
    if not frames:raise ValueError('ไม่พบ Sheet ข้อมูลปริญญาโท/ปริญญาเอกที่ใช้งานได้'+(f": {'; '.join(errors[:3])}" if errors else ''))
    return used,pd.concat(frames,ignore_index=True),raw_sheets
def count_system(df,kind):
    x=df[df['__ระบบ']==kind]
    if x.empty:return 0
    ids=x['__รหัส'].map(clean)
    return int(ids[ids!=''].nunique()+int((ids=='').sum()))
def metrics(df):
    inn=count_system(df,'ในเวลา');out=count_system(df,'นอกเวลา');a=inn+out
    q_col=df.get('__Q_นับ',df['__Q'])
    pub=int(df.loc[q_col.isin(['Q1','Q2']),'__ผลงาน_key'].nunique())
    return inn,out,a,pub,round(pub*100/a,2) if a else 0
ORDER_TREE=[('กลุ่มมนุษยศาสตร์และสังคมศาสตร์',[('การเมืองการปกครอง',['รัฐศาสตร์']),('การท่องเที่ยวและการโรงแรม',['การจัดการการท่องเที่ยวและการโรงแรม']),('การบัญชีและการจัดการ',['การจัดการสมัยใหม่','การจัดการสมาร์ตซิตี้และนวัตกรรมดิจิทัล','การบัญชี','บริหารธุรกิจและนวัตกรรมดิจิทัล']),('ดุริยางคศิลป์',['ดุริยางคศิลป์']),('มนุษยศาสตร์และสังคมศาสตร์',['การสอนภาษาอังกฤษ','ภาษาไทย','ศาสนาและภูมิปัญญาเพื่อการพัฒนา']),('ศิลปกรรมศาสตร์และวัฒนธรรมศาสตร์',['การวิจัยและสร้างสรรค์ศิลปกรรมศาสตร์','วัฒนธรรมศาสตร์']),('ศึกษาศาสตร์',['เทคโนโลยีและสื่อสารการศึกษา','การบริหารและพัฒนาการศึกษา','วิจัยและประเมินผลการศึกษา','วิทยาศาสตร์การออกกำลังกายและการกีฬา','หลักสูตรและการสอน'])]),('กลุ่มวิทยาศาสตร์และเทคโนโลยี',[('เทคโนโลยี',['เกษตรศาสตร์','เทคโนโลยีการอาหาร']),('วิจัยวลัยรุกขเวช',['ความหลากหลายทางชีวภาพ']),('วิทยาการสารสนเทศ',['เทคโนโลยีสารสนเทศ','วิทยาการคอมพิวเตอร์','สื่อนฤมิต']),('วิทยาศาสตร์',['บรรพชีวินวิทยา','ฟิสิกส์']),('วิศวกรรมศาสตร์',['วิศวกรรมเครื่องกล','วิศวกรรมโยธา','วิศวกรรมไฟฟ้าและคอมพิวเตอร์']),('สิ่งแวดล้อมและทรัพยากรศาสตร์',['การจัดการสิ่งแวดล้อมอย่างยั่งยืน'])]),('กลุ่มวิทยาศาสตร์สุขภาพ',[('เภสัชศาสตร์',['เภสัชศาสตร์']),('แพทยศาสตร์',['วิทยาศาสตร์สุขภาพ']),('สาธารณสุขศาสตร์',['เทคโนโลยีทางสุขภาพและความปลอดภัย','สาธารณสุขศาสตรดุษฎีบัณฑิต'])])]
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
    def programs_for_faculty(g,f,preferred):
        mask=data['__กลุ่ม'].map(clean).eq(g)&data['__คณะ'].map(clean).eq(f)&data['__สาขา'].map(clean).ne('');actual=[];seen=set()
        for p in data.loc[mask,'__สาขา'].map(clean):
            if p and p not in seen:seen.add(p);actual.append(p)
        ordered=[p for p in preferred if p in seen];ordered.extend(p for p in actual if p not in ordered);return ordered
    for g,fs in ORDER_TREE:
        add_group(g)
        for f,ps in fs:
            add_faculty(g,f)
            for p in programs_for_faculty(g,f,ps):add_program(g,f,p)
    for _,r in data.sort_values('__ลำดับ').iterrows():
        g,f=clean(r['__กลุ่ม']),clean(r['__คณะ']);add_group(g)
        if g and f:
            add_faculty(g,f)
            for p in programs_for_faculty(g,f,[]):add_program(g,f,p)
    rows.append(['','รวมทั้งหมด',*metrics(data)]);return pd.DataFrame(rows,columns=cols)
def q_summary(data):
    def vals(level):
        q_col=data.get('__Q_นับ',data['__Q'])
        x=data[(data['__ระดับ']==level)&q_col.isin(['Q1','Q2'])].drop_duplicates('__ผลงาน_key')
        q1=int((x['__Q']=='Q1').sum());q2=int((x['__Q']=='Q2').sum());return q1,q2,q1+q2
    m,d=vals('ปริญญาโท'),vals('ปริญญาเอก');return pd.DataFrame([['ระดับปริญญาโท',*m],['ระดับปริญญาเอก',*d],['รวม',m[0]+d[0],m[1]+d[1],m[2]+d[2]]],columns=['','Scopus Q1','Scopus Q2','รวม'])
def safe_sheet_name(name,used):
    base='ข้อมูล_'+clean(name)[:25];base=re.sub(r'[\\/*?:\[\]]','_',base) or 'ข้อมูล';candidate=base;i=2
    while candidate in used:candidate=f'{base[:28-len(str(i))]}_{i}';i+=1
    used.add(candidate);return candidate
def make_excel(data,raw_sheets):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as w:
        for level in ['ปริญญาโท','ปริญญาเอก']:hierarchy(data[data['__ระดับ']==level]).to_excel(w,sheet_name=level,index=False,startrow=3)
        q_summary(data).to_excel(w,sheet_name='สรุป Scopus Q1-Q2',index=False)
        # รวมข้อมูลดิบจากทุกชีต/ทุกไฟล์ไว้ใน Sheet เดียว
        if raw_sheets:
            details=[]
            for name,raw in raw_sheets.items():
                detail=raw.copy()
                detail.columns=[clean(c) for c in detail.columns]
                detail.insert(0,'แหล่งข้อมูล',name)
                base_col=find_col(detail,BASE)
                if base_col:
                    detail['Q ที่ใช้ในการนับ']=detail[base_col].map(norm_q)
                    detail['สถานะการนับ Q1-Q2']=detail['Q ที่ใช้ในการนับ'].map(
                        lambda q: f'นับเป็น {q}' if q in ('Q1','Q2') else 'ไม่นับ'
                    )
                else:
                    detail['Q ที่ใช้ในการนับ']=''
                    detail['สถานะการนับ Q1-Q2']='ตรวจไม่พบคอลัมน์ฐานข้อมูล'
                details.append(detail)
            combined_detail=pd.concat(details,ignore_index=True,sort=False)
            combined_detail.to_excel(w,sheet_name='ข้อมูล',index=False)
    bio.seek(0);wb=load_workbook(bio);thin=Side(style='thin',color='000000')
    for ws in wb.worksheets:
        if ws.title in ['ปริญญาโท','ปริญญาเอก']:
            # หัวตาราง 3 ชั้น
            ws.insert_rows(1,3)
            # ลบแถวว่าง 4-6 แล้วเลื่อนข้อมูลขึ้นมาให้เริ่มที่แถว 4
            ws.delete_rows(4,3)
            # ลบแถวที่ 4 หากเป็นหัวตารางซ้ำจากข้อมูลที่เขียนโดย pandas
            header_values = [ws.cell(4,c).value for c in range(1,8)]
            expected_headers = ['ลำดับ','กลุ่ม/คณะ/สาขา','ระบบในเวลาราชการ','ระบบนอกเวลาราชการ','จำนวนผู้สำเร็จการศึกษา (A)','รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2','ร้อยละ']
            if header_values == expected_headers:
                ws.delete_rows(4,1)
            ws.merge_cells('A1:G1')
            ws['A1']=PAGE_TITLE+' ระดับ'+ws.title
            ws['A1'].font=Font(bold=True,color='FFFFFF',size=16)
            ws['A1'].fill=PatternFill('solid',fgColor='2F5D20')
            ws['A1'].alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            ws.row_dimensions[1].height=28

            # แถวหัวตารางหลัก
            ws.merge_cells('A2:A3')
            ws.merge_cells('B2:B3')
            ws.merge_cells('C2:D2')
            ws.merge_cells('E2:E3')
            ws.merge_cells('F2:G2')

            ws['A2']='ลำดับ'
            ws['B2']='กลุ่ม/คณะ/สาขา'
            ws['C2']='ระบบ'
            ws['E2']='จำนวนผู้สำเร็จ\nการศึกษา (A)'
            ws['F2']='1.2.4 จำนวนผลงานระดับบัณฑิตศึกษาที่\nสามารถตีพิมพ์ในฐานข้อมูลนานาชาติ SCOPUS\nQ1-Q2'

            # แถวหัวตารางย่อย
            ws['C3']='ระบบในเวลา\nราชการ'
            ws['D3']='ระบบนอกเวลา\nราชการ'
            ws['F3']='รวมจำนวนผลงานตีพิมพ์\nระดับนานาชาติ Q1-Q2'
            ws['G3']='ร้อยละ'

            # เติมสีและจัดรูปแบบหัวตาราง
            all_border=Border(left=thin,right=thin,top=thin,bottom=thin)
            for row in range(2,4):
                for c in range(1,8):
                    ws.cell(row,c).fill=PatternFill('solid',fgColor='2F5D20' if row == 2 else '568B3B')
                    ws.cell(row,c).font=Font(bold=True,color='FFFFFF')
                    ws.cell(row,c).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                    ws.cell(row,c).border=all_border

            ws.row_dimensions[2].height=54
            ws.row_dimensions[3].height=54
            ws.freeze_panes='A4'
            ws.column_dimensions['A'].width=9
            ws.column_dimensions['B'].width=48
            for col in range(3,8):ws.column_dimensions[chr(64+col)].width=22
            for row in range(4,ws.max_row+1):
                label=str(ws.cell(row,2).value or '')
                if label.startswith('        '):ws.cell(row,2).value=label.strip();indent=2;fill='FFFFFF'
                elif label.startswith('    '):ws.cell(row,2).value=label.strip();indent=1;fill='E2F0D9'
                elif label=='รวมทั้งหมด':fill='FCE4D6';indent=0
                else:fill='FCE4D6';indent=0
                for c in range(1,8):ws.cell(row,c).fill=PatternFill('solid',fgColor=fill);ws.cell(row,c).border=all_border
                ws.cell(row,2).alignment=Alignment(horizontal='left',vertical='center',indent=indent)
        elif ws.title=='สรุป Scopus Q1-Q2':
            # ตกแต่งตารางสรุป Scopus Q1-Q2
            ws.freeze_panes='A2'
            all_border=Border(left=thin,right=thin,top=thin,bottom=thin)
            for cell in ws[1]:
                cell.fill=PatternFill('solid',fgColor='2F5D20')
                cell.font=Font(bold=True,color='FFFFFF')
                cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                cell.border=all_border
            for row in ws.iter_rows(min_row=2,max_row=ws.max_row,min_col=1,max_col=4):
                for cell in row:
                    cell.border=all_border
                    cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                    if cell.row == ws.max_row:
                        cell.fill=PatternFill('solid',fgColor='FCE4D6')
                        cell.font=Font(bold=True)
                    elif cell.column == 1:
                        cell.fill=PatternFill('solid',fgColor='E2F0D9')
            ws.row_dimensions[1].height=30
            ws.column_dimensions['A'].width=24
            ws.column_dimensions['B'].width=16
            ws.column_dimensions['C'].width=16
            ws.column_dimensions['D'].width=14
    out=io.BytesIO();wb.save(out);out.seek(0);return out

st.title(PAGE_TITLE)
st.caption('ระบบจัดทำสถิติผลงานระดับบัณฑิตศึกษา Scopus Q1-Q2 แยกปริญญาโทและปริญญาเอก')

# เลือกไฟล์ให้ครบก่อน แล้วจึงกดปุ่มเพื่อเริ่มประมวลผล
with st.form('upload_and_run_form'):
    files = st.file_uploader(
        'อัปโหลดไฟล์ Excel (เลือกได้หลายไฟล์)',
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    run_clicked = st.form_submit_button(
        'รันประมวลผลสถิติ',
        type='primary',
        use_container_width=True
    )

if run_clicked:
    if not files:
        st.warning('กรุณาอัปโหลดไฟล์ Excel อย่างน้อยหนึ่งไฟล์ก่อนกดรัน')
    else:
        try:
            all_frames = []
            all_raw = {}
            file_errors = []
            file_success = []
            for uploaded in files:
                try:
                    used, data_part, raw = read_excel(uploaded.getvalue())
                    all_frames.append(data_part)
                    file_success.append(f"{uploaded.name}: {', '.join(used)}")
                    # ตั้งชื่อ Sheet ให้แยกกัน แม้หลายไฟล์จะมีชื่อ Sheet ซ้ำกัน
                    for sheet_name, raw_df in raw.items():
                        key = f"{uploaded.name}_{sheet_name}"
                        all_raw[key] = raw_df.copy()
                except Exception as e:
                    file_errors.append(f"{uploaded.name}: {e}")

            for msg in file_errors:
                st.warning(f"ข้ามไฟล์ที่อ่านไม่ได้ — {msg}")

            if not all_frames:
                st.error('ไม่พบไฟล์ที่มีข้อมูลปริญญาโท/ปริญญาเอกที่ใช้งานได้ กรุณาตรวจสอบไฟล์ที่อัปโหลด')
            else:
                data = pd.concat(all_frames, ignore_index=True)
                # รองรับผลลัพธ์จาก cache/ไฟล์เก่าที่อาจยังไม่มีคอลัมน์รหัสผลงาน
                if '__ผลงาน_key' not in data.columns or data['__ผลงาน_key'].isna().any():
                    def _key_text(v):
                        return re.sub(r'\s+', ' ', clean(v)).strip().casefold()
                    _doi_col = find_col(data, DOI_COL)
                    _title_col = find_col(data, PUB_TITLE)
                    _doi = data[_doi_col].map(_key_text).str.replace(r'^https?://(dx\.)?doi\.org/', '', regex=True).str.replace(r'^doi\s*:\s*', '', regex=True).str.strip() if _doi_col else pd.Series('', index=data.index)
                    _title = data[_title_col].map(_key_text) if _title_col else pd.Series('', index=data.index)
                    _source_cols = [c for c in data.columns if not str(c).startswith('__')]
                    _fallback = data[_source_cols].fillna('').astype(str).apply(lambda row: '|'.join(_key_text(v) for v in row), axis=1)
                    _generated = pd.Series(['doi:' + d if d else ('title:' + t if t else 'row:' + f) for d, t, f in zip(_doi, _title, _fallback)], index=data.index)
                    if '__ผลงาน_key' not in data.columns:
                        data['__ผลงาน_key'] = _generated
                    else:
                        _missing_key = data['__ผลงาน_key'].isna() | data['__ผลงาน_key'].astype(str).str.strip().eq('')
                        data.loc[_missing_key, '__ผลงาน_key'] = _generated.loc[_missing_key]
                # ให้แต่ละผลงาน Q1/Q2 ถูกนับเพียงครั้งเดียว แม้ซ้ำข้ามชีตหรือไฟล์
                data['__Q_นับ'] = data['__Q']
                qmask = data['__Q'].isin(['Q1','Q2'])
                duplicate_q = qmask & data.loc[qmask, '__ผลงาน_key'].duplicated(keep='first').reindex(data.index, fill_value=False)
                data.loc[duplicate_q, '__Q_นับ'] = ''
                st.info(f"พบผลงาน Scopus Q1-Q2 ที่ไม่ซ้ำ {data.loc[qmask, '__ผลงาน_key'].nunique():,} รายการ จากข้อมูลที่เข้าเกณฑ์ {int(qmask.sum()):,} แถว")
                st.success(f'อ่านข้อมูลสำเร็จ {len(file_success)} จาก {len(files)} ไฟล์')
                with st.expander('รายละเอียดไฟล์ที่อ่านสำเร็จ'):
                    for msg in file_success:
                        st.write('• ' + msg)
                st.caption(f'รวมข้อมูลทั้งหมด {len(data):,} แถว')
                for level in ['ปริญญาโท', 'ปริญญาเอก']:
                    st.subheader(level)
                    st.dataframe(
                        hierarchy(data[data['__ระดับ'] == level]),
                        use_container_width=True,
                        hide_index=True
                    )
                st.subheader('สรุป Scopus Q1-Q2')
                st.dataframe(q_summary(data), use_container_width=True, hide_index=True)
                st.download_button(
                    'ดาวน์โหลด Excel สถิติผลงานตีพิมพ์',
                    data=make_excel(data, all_raw).getvalue(),
                    file_name='สถิติผลงานตีพิมพ์.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        except Exception as e:
            st.error(f'ไม่สามารถประมวลผลไฟล์ได้: {e}')
else:
    st.info('เลือกไฟล์ Excel ให้ครบก่อน จากนั้นกด “รันประมวลผลสถิติ” เพื่อสร้างตาราง')
