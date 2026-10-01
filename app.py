import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")

LEVEL_ALIASES = ["ระดับ", "ระดับการศึกษา", "ระดับปริญญา", "วุฒิ", "degree", "level"]
PRIMARY_Q_ALIASES = ["ผลงานที่ตีพิมพ์ Q1-Q2", "ผลงานที่ตีพิมพ์Q1-Q2"]
FALLBACK_Q_ALIASES = ["ฐานข้อมูล"]
GROUP_ALIASES = ["กลุ่ม", "group"]
FACULTY_ALIASES = ["คณะ", "faculty", "school"]
PROGRAM_ALIASES = ["สาขา", "สาขาวิชา", "หลักสูตร", "program", "major"]
COMBINED_ALIASES = ["กลุ่ม/คณะ/สาขา", "กลุ่ม คณะ สาขา", "กลุ่ม/คณะ", "ชื่อคณะ", "หน่วยงาน"]
SYSTEM_ALIASES = ["ระบบ"]
STUDENT_ID_ALIASES = ["รหัสนิสิต", "รหัสนักศึกษา", "รหัส", "student id", "student_id", "id"]
STATUS_ALIASES = ["สถานะนิสิต", "สถานะนักศึกษา", "สถานะ", "status"]
PROGRAM_TOTAL_ALIASES = ["สาขารวม", "สาขา รวม", "program total"]


def clean(v):
    if pd.isna(v): return ""
    return str(v).strip()


def find_col(df, aliases):
    normalized = {re.sub(r"\s+", "", clean(c)).lower(): c for c in df.columns}
    for alias in aliases:
        key = re.sub(r"\s+", "", alias).lower()
        if key in normalized: return normalized[key]
    for c in df.columns:
        cc = re.sub(r"\s+", "", clean(c)).lower()
        if any(re.sub(r"\s+", "", a).lower() in cc for a in aliases): return c
    return None


def normalize_level(v):
    s = clean(v).lower().replace(" ", "")
    if any(x in s for x in ["ปริญญาเอก", "ป.เอก", "ph.d", "phd", "doctoral"]): return "ปริญญาเอก"
    if any(x in s for x in ["ปริญญาโท", "ป.โท", "master"]): return "ปริญญาโท"
    return "อื่น ๆ"


def normalize_q(v):
    s = clean(v).upper().replace(" ", "")
    if "Q1" in s: return "Q1"
    if "Q2" in s: return "Q2"
    return "อื่น ๆ"


def normalize_system(v):
    s = clean(v).lower().replace(" ", "").replace("-", "")
    if not s: return "ไม่ระบุ"
    # ยึดค่าจากคอลัมน์ "ระบบ" เท่านั้น
    if any(x in s for x in ["นอกเวลา", "นอกเวลาราชการ", "parttime", "parttime"]): return "นอกเวลา"
    if any(x in s for x in ["ในเวลา", "ในเวลาราชการ", "fulltime", "fulltime"]): return "ในเวลา"
    return "อื่น ๆ"


def is_success(v):
    s = clean(v).lower()
    if not s: return True
    return any(x in s for x in ["สำเร็จการศึกษา", "สำเร็จ", "graduate", "graduated"])


def read_workbook(uploaded):
    uploaded.seek(0)
    sheets = pd.read_excel(uploaded, sheet_name=None, header=1)
    candidates = []
    for name, df in sheets.items():
        if df is None or df.empty: continue
        df = df.copy(); df.columns = [clean(c) for c in df.columns]
        level_col = find_col(df, LEVEL_ALIASES)
        primary_q = find_col(df, PRIMARY_Q_ALIASES)
        fallback_q = find_col(df, FALLBACK_Q_ALIASES)
        hierarchy = sum(bool(find_col(df, a)) for a in [GROUP_ALIASES, FACULTY_ALIASES, PROGRAM_ALIASES, PROGRAM_TOTAL_ALIASES, COMBINED_ALIASES])
        score = (10 if level_col else 0) + (10 if primary_q else (5 if fallback_q else 0)) + hierarchy
        candidates.append((score, name, df))
    if not candidates: raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2]


def prepare(df):
    out = df.copy()
    level_col = find_col(out, LEVEL_ALIASES)
    primary_q = find_col(out, PRIMARY_Q_ALIASES)
    fallback_q = find_col(out, FALLBACK_Q_ALIASES)
    q_col = primary_q or fallback_q
    group_col = find_col(out, GROUP_ALIASES)
    faculty_col = find_col(out, FACULTY_ALIASES)
    program_col = find_col(out, PROGRAM_ALIASES)
    program_total_col = find_col(out, PROGRAM_TOTAL_ALIASES)
    combined_col = find_col(out, COMBINED_ALIASES)
    system_col = find_col(out, SYSTEM_ALIASES)
    student_id_col = find_col(out, STUDENT_ID_ALIASES)
    status_col = find_col(out, STATUS_ALIASES)
    if not level_col: raise ValueError("ไม่พบคอลัมน์ 'ระดับ' ในแถวที่ 2")
    if not q_col: raise ValueError("ไม่พบคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2' และไม่พบคอลัมน์สำรอง 'ฐานข้อมูล' ในแถวที่ 2")
    if not system_col: raise ValueError("ไม่พบคอลัมน์ 'ระบบ' ในแถวที่ 2")
    if not (group_col or faculty_col or program_col or program_total_col or combined_col): raise ValueError("ไม่พบคอลัมน์สำหรับ กลุ่ม / คณะ / สาขา")

    out["__ระดับ"] = out[level_col].map(normalize_level)
    out["__Q"] = out[q_col].map(normalize_q)
    # ระบบในเวลา/นอกเวลา ต้องมาจากคอลัมน์ "ระบบ" เท่านั้น
    out["__ระบบ"] = out[system_col].map(normalize_system)
    out["__รหัสนิสิต"] = out[student_id_col].map(clean) if student_id_col else ""
    out["__สถานะ"] = out[status_col].map(is_success) if status_col else True
    out["__กลุ่ม"] = out[group_col].map(clean) if group_col else ""
    out["__คณะ"] = out[faculty_col].map(clean) if faculty_col else ""
    out["__สาขา"] = out[program_total_col].map(clean) if program_total_col else (out[program_col].map(clean) if program_col else "")
    out["__รวมชื่อ"] = out[combined_col].map(clean) if combined_col else ""
    out["__ลำดับต้นฉบับ"] = range(len(out))
    out = out[out["__ระดับ"].isin(["ปริญญาโท", "ปริญญาเอก"])].copy()
    return out, {"ระดับ": level_col, "Q1/Q2 ที่ใช้": q_col, "แหล่ง Q1/Q2": "ผลงานที่ตีพิมพ์ Q1-Q2" if primary_q else "ฐานข้อมูล", "กลุ่ม": group_col, "คณะ": faculty_col, "สาขาที่ใช้": program_total_col or program_col, "ระบบ": system_col, "รหัสนิสิต": student_id_col, "สถานะ": status_col}


def unique_count(df):
    if df.empty: return 0
    ids = df["__รหัสนิสิต"].astype(str).str.strip()
    if ids.ne("").any(): return int(ids[ids != ""].nunique())
    return int(len(df))


def metrics(data):
    g = data[data["__สถานะ"]]
    p = data[data["__Q"].isin(["Q1", "Q2"])]
    # นับจากคอลัมน์ "ระบบ" ที่ normalize แล้วเท่านั้น
    in_count = unique_count(g[g["__ระบบ"] == "ในเวลา"])
    out_count = unique_count(g[g["__ระบบ"] == "นอกเวลา"])
    a = in_count + out_count
    pub = len(p)
    pct = round(pub * 100 / a, 2) if a else 0.0
    return in_count, out_count, a, pub, pct


def make_hierarchy_table(data):
    columns = ["ลำดับ", "กลุ่ม/คณะ/สาขา", "ระบบในเวลาราชการ", "ระบบนอกเวลาราชการ", "จำนวนผู้สำเร็จการศึกษา (A)", "รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2", "ร้อยละผลงานวิจัยรวมระดับนานาชาติ Q1-Q2"]
    if data.empty: return pd.DataFrame(columns=columns)
    rows, seq, seen = [], 1, set()
    levels = [("__กลุ่ม", 0), ("__คณะ", 1), ("__สาขา", 2)]
    for _, r in data.sort_values("__ลำดับต้นฉบับ").iterrows():
        path = []
        for col, indent in levels:
            val = clean(r[col])
            if not val: continue
            path.append(val)
            key = (col, tuple(path))
            if key in seen: continue
            seen.add(key)
            mask = pd.Series(True, index=data.index)
            for prev_col, _ in levels[:indent + 1]:
                prev_val = clean(r[prev_col])
                if prev_val: mask &= data[prev_col].map(clean).eq(prev_val)
            sub = data[mask]
            in_count, out_count, a, pub, pct = metrics(sub)
            rows.append([seq, ("    " * indent) + val, in_count, out_count, a, pub, pct])
            seq += 1
    result = pd.DataFrame(rows, columns=columns)
    if not result.empty:
        in_count, out_count, a, pub, pct = metrics(data)
        result.loc[len(result)] = ["", "รวมทั้งหมด", in_count, out_count, a, pub, pct]
    return result


def export_excel(master, doctor):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="ปริญญาโท", index=False, startrow=3)
        doctor.to_excel(writer, sheet_name="ปริญญาเอก", index=False, startrow=3)
    bio.seek(0)
    wb = load_workbook(bio)
    fill1 = PatternFill("solid", fgColor="355E20"); fill2 = PatternFill("solid", fgColor="548235")
    font = Font(name="Tahoma", color="FFFFFF", bold=True); thin = Side(style="thin", color="000000")
    for ws in wb.worksheets:
        ws.merge_cells("A1:A3"); ws["A1"] = "ลำดับ"
        ws.merge_cells("B1:B3"); ws["B1"] = "กลุ่ม/คณะ/สาขา"
        ws.merge_cells("C1:D1"); ws["C1"] = "ระบบ"
        ws["C2"] = "ระบบในเวลาราชการ"; ws["D2"] = "ระบบนอกเวลาราชการ"
        ws.merge_cells("E1:E3"); ws["E1"] = "จำนวนผู้สำเร็จการศึกษา (A)"
        ws.merge_cells("F1:G1"); ws["F1"] = "1.2.4 จำนวนผลงานระดับบัณฑิตศึกษาที่สามารถตีพิมพ์ในระดับนานาชาติ SCOPUS Q1-Q2"
        ws.merge_cells("F2:F3"); ws["F2"] = "รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2"
        ws.merge_cells("G2:G3"); ws["G2"] = "ร้อยละผลงานวิจัยรวมระดับนานาชาติ Q1-Q2"
        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=7):
            for cell in row:
                cell.fill = fill1 if cell.row == 1 else fill2; cell.font = font; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws.iter_rows(min_row=4):
            for cell in row:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin); cell.alignment = Alignment(vertical="center", wrap_text=True)
        widths = {"A": 8, "B": 42, "C": 18, "D": 18, "E": 20, "F": 24, "G": 24}
        for c, w in widths.items(): ws.column_dimensions[c].width = w
        ws.freeze_panes = "A4"
    out = io.BytesIO(); wb.save(out); out.seek(0); return out


st.title("📊 KPI01 — ตารางสถิติผลงานระดับบัณฑิตศึกษา")
st.caption("หัวคอลัมน์อยู่แถวที่ 2 | ระดับใช้คอลัมน์ ระดับ | Q1/Q2 ใช้ ผลงานที่ตีพิมพ์ Q1-Q2 และใช้ ฐานข้อมูล เป็นสำรอง | ระบบใน/นอกเวลา นับจากคอลัมน์ ระบบ | แถวสาขาใช้ สาขารวม")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])
if uploaded:
    try:
        sheet_name, raw = read_workbook(uploaded)
        data, cols = prepare(raw)
        master = make_hierarchy_table(data[data["__ระดับ"] == "ปริญญาโท"])
        doctor = make_hierarchy_table(data[data["__ระดับ"] == "ปริญญาเอก"])
        st.success(f"อ่าน Sheet: {sheet_name} | หัวคอลัมน์แถวที่ 2 | ใช้คอลัมน์สาขา: {cols['สาขาที่ใช้']} | ใช้ข้อมูล {len(data):,} แถว")
        with st.expander("ตรวจสอบคอลัมน์"):
            st.write(cols)
            st.write("ปริญญาโท:", int((data["__ระดับ"] == "ปริญญาโท").sum()))
            st.write("ปริญญาเอก:", int((data["__ระดับ"] == "ปริญญาเอก").sum()))
            st.write("ระบบในเวลา:", int((data["__ระบบ"] == "ในเวลา").sum()))
            st.write("ระบบนอกเวลา:", int((data["__ระบบ"] == "นอกเวลา").sum()))
        tab1, tab2 = st.tabs(["🎓 ปริญญาโท", "🎓 ปริญญาเอก"])
        with tab1: st.dataframe(master, use_container_width=True, hide_index=True)
        with tab2: st.dataframe(doctor, use_container_width=True, hide_index=True)
        excel = export_excel(master, doctor)
        st.download_button("📥 ดาวน์โหลด Excel KPI01", excel, "KPI01_ปริญญาโท_ปริญญาเอก.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e: st.error(str(e))
else: st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ")
