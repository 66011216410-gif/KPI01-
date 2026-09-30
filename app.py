import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")

LEVEL_MAP = {
    "ป.โท": ["ป.โท", "ปริญญาโท", "โท", "master", "masters", "master's"],
    "ป.เอก": ["ป.เอก", "ปริญญาเอก", "เอก", "doctor", "doctoral", "phd", "ph.d"],
}

ALIASES = {
    "level": ["ระดับ", "ระดับการศึกษา", "วุฒิ", "degree", "level"],
    "faculty": ["คณะ", "faculty", "school"],
    "program": ["สาขา", "สาขาวิชา", "หลักสูตร", "program", "major"],
    "quartile": ["quartile", "q", "ระดับ scopus", "scopus", "คุณภาพวารสาร", "ฐานข้อมูล"],
    "student_id": ["รหัสนิสิต", "รหัสนักศึกษา", "student id", "student_id", "id"],
    "title": ["ชื่อผลงาน", "ชื่อบทความ", "ชื่อเรื่อง", "title", "article title", "publication"],
}


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def find_col(df, aliases):
    normalized = {re.sub(r"\s+", "", clean(c)).lower(): c for c in df.columns}
    for alias in aliases:
        key = re.sub(r"\s+", "", alias).lower()
        if key in normalized:
            return normalized[key]
    for c in df.columns:
        cc = re.sub(r"\s+", "", clean(c)).lower()
        if any(re.sub(r"\s+", "", a).lower() in cc for a in aliases):
            return c
    return None


def normalize_level(v):
    s = clean(v).lower()
    if any(x.lower() in s for x in LEVEL_MAP["ป.เอก"]):
        return "ป.เอก"
    if any(x.lower() in s for x in LEVEL_MAP["ป.โท"]):
        return "ป.โท"
    return "อื่น ๆ"


def normalize_quartile(v):
    s = clean(v).upper().replace(" ", "")
    m = re.search(r"Q[12]", s)
    return m.group(0) if m else "อื่น ๆ"


def read_workbook(uploaded):
    uploaded.seek(0)
    sheets = pd.read_excel(uploaded, sheet_name=None)
    candidates = []
    for name, df in sheets.items():
        if df is not None and not df.empty:
            df = df.copy()
            df.columns = [clean(c) for c in df.columns]
            score = sum(find_col(df, ALIASES[k]) is not None for k in ["level", "faculty", "program", "quartile"])
            candidates.append((score, name, df))
    if not candidates:
        raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    return max(candidates, key=lambda x: x[0])[1], max(candidates, key=lambda x: x[0])[2], sheets


def prepare(df):
    out = df.copy()
    cols = {k: find_col(out, aliases) for k, aliases in ALIASES.items()}
    if not cols["level"]:
        raise ValueError("ไม่พบคอลัมน์ระดับการศึกษา เช่น ระดับ / Degree")
    if not cols["quartile"]:
        raise ValueError("ไม่พบคอลัมน์ Q1/Q2 เช่น Quartile / Scopus")

    out["__ระดับ"] = out[cols["level"]].map(normalize_level)
    out["__Quartile"] = out[cols["quartile"]].map(normalize_quartile)
    out["__คณะ"] = out[cols["faculty"]].map(clean) if cols["faculty"] else "ไม่ระบุ"
    out["__สาขา"] = out[cols["program"]].map(clean) if cols["program"] else "ไม่ระบุ"
    out["__รหัสนิสิต"] = out[cols["student_id"]].map(clean) if cols["student_id"] else ""
    out["__ชื่อผลงาน"] = out[cols["title"]].map(clean) if cols["title"] else ""
    return out, cols


def count_publications(g):
    # 1 แถว = 1 ผลงานเป็นค่าเริ่มต้น; หากมีรหัสนิสิตและผู้ใช้เลือกนับนิสิต
    return len(g)


def make_stats(g, unit="ผลงาน"):
    rows = []
    if g.empty:
        return pd.DataFrame(columns=["คณะ", "สาขา", "Q1", "Q2", "รวม Q1+Q2", "%Q1", "%Q2"])

    grouped = g.groupby(["__คณะ", "__สาขา"], dropna=False, sort=True)
    for (faculty, program), x in grouped:
        q1 = int((x["__Quartile"] == "Q1").sum())
        q2 = int((x["__Quartile"] == "Q2").sum())
        total = q1 + q2
        rows.append({
            "คณะ": faculty or "ไม่ระบุ",
            "สาขา": program or "ไม่ระบุ",
            "Q1": q1,
            "Q2": q2,
            "รวม Q1+Q2": total,
            "%Q1": round(q1 * 100 / total, 2) if total else 0,
            "%Q2": round(q2 * 100 / total, 2) if total else 0,
        })

    result = pd.DataFrame(rows)
    if not result.empty:
        total = {
            "คณะ": "รวมทั้งหมด",
            "สาขา": "",
            "Q1": int(result["Q1"].sum()),
            "Q2": int(result["Q2"].sum()),
            "รวม Q1+Q2": int(result["รวม Q1+Q2"].sum()),
        }
        total["%Q1"] = round(total["Q1"] * 100 / total["รวม Q1+Q2"], 2) if total["รวม Q1+Q2"] else 0
        total["%Q2"] = round(total["Q2"] * 100 / total["รวม Q1+Q2"], 2) if total["รวม Q1+Q2"] else 0
        result = pd.concat([result, pd.DataFrame([total])], ignore_index=True)
    return result


def make_level_summary(g):
    rows = []
    for level in ["ป.โท", "ป.เอก"]:
        x = g[g["__ระดับ"] == level]
        q1 = int((x["__Quartile"] == "Q1").sum())
        q2 = int((x["__Quartile"] == "Q2").sum())
        rows.append({"ระดับ": level, "Q1": q1, "Q2": q2, "รวม Q1+Q2": q1 + q2})
    return pd.DataFrame(rows)


def to_excel(stats_master, stats_doctor, level_summary):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        level_summary.to_excel(writer, sheet_name="สรุป", index=False)
        stats_master.to_excel(writer, sheet_name="ป.โท", index=False)
        stats_doctor.to_excel(writer, sheet_name="ป.เอก", index=False)

    bio.seek(0)
    wb = openpyxl.load_workbook(bio)
    fill = PatternFill("solid", fgColor="2F5597")
    font = Font(name="Tahoma", color="FFFFFF", bold=True)
    border = Border(*( [Side(style="thin", color="D9D9D9")] * 4 ))
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
        for col in ws.columns:
            width = min(max(len(str(c.value or "")) for c in col) + 2, 45)
            ws.column_dimensions[col[0].column_letter].width = width
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


st.title("📊 KPI01 — ตารางสถิติผลงานระดับบัณฑิตศึกษา")
st.caption("สร้างตารางแยก ป.โท และ ป.เอก โดยอ่านข้อมูล Q1/Q2 จาก Excel อัตโนมัติ")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel ตาม Sheet ที่ต้องการใช้เป็นต้นแบบ", type=["xlsx", "xls"])

if uploaded:
    try:
        sheet_name, raw, all_sheets = read_workbook(uploaded)
        data, cols = prepare(raw)

        st.success(f"อ่านข้อมูลจาก Sheet: {sheet_name} | {len(data):,} แถว")
        with st.expander("ตรวจสอบคอลัมน์ที่ระบบจับคู่"):
            st.json(cols)

        level_summary = make_level_summary(data)
        stats_master = make_stats(data[data["__ระดับ"] == "ป.โท"])
        stats_doctor = make_stats(data[data["__ระดับ"] == "ป.เอก"])

        st.subheader("สรุป ป.โท / ป.เอก")
        st.dataframe(level_summary, use_container_width=True, hide_index=True)

        tab1, tab2 = st.tabs(["🎓 ป.โท", "🎓 ป.เอก"])
        with tab1:
            st.metric("ผลงาน Q1+Q2", int(stats_master["รวม Q1+Q2"].sum()) if not stats_master.empty else 0)
            st.dataframe(stats_master, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลดตาราง ป.โท CSV", stats_master.to_csv(index=False).encode("utf-8-sig"), "kpi01_master.csv", "text/csv")
        with tab2:
            st.metric("ผลงาน Q1+Q2", int(stats_doctor["รวม Q1+Q2"].sum()) if not stats_doctor.empty else 0)
            st.dataframe(stats_doctor, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลดตาราง ป.เอก CSV", stats_doctor.to_csv(index=False).encode("utf-8-sig"), "kpi01_doctor.csv", "text/csv")

        excel = to_excel(stats_master, stats_doctor, level_summary)
        st.download_button("📥 ดาวน์โหลด Excel ตารางสถิติ KPI01", excel, "KPI01_สถิติ_ปโท_ปเอก.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with st.expander("ข้อมูลต้นฉบับ"):
            st.dataframe(raw, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(str(e))
else:
    st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ ป.โท และ ป.เอก")
