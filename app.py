import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")

LEVEL_MAP = {
    "ป.โท": ["ป.โท", "ปริญญาโท", "ปริญญาโท (ภาคปกติ)", "โท", "master", "masters", "master's"],
    "ป.เอก": ["ป.เอก", "ปริญญาเอก", "ปริญญาเอก (ภาคปกติ)", "เอก", "doctor", "doctoral", "phd", "ph.d"],
}

ALIASES = {
    "level": ["ระดับ", "ระดับการศึกษา", "ระดับปริญญา", "วุฒิ", "degree", "level", "program level"],
    "faculty": ["คณะ", "faculty", "school"],
    "program": ["สาขา", "สาขาวิชา", "หลักสูตร", "program", "major"],
    "quartile": ["quartile", "q1", "q2", "ระดับ scopus", "scopus", "คุณภาพวารสาร", "ฐานข้อมูล", "quartile (q)"],
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
    # ตรวจ ป.เอก ก่อน เพราะคำว่า "เอก" อาจอยู่ในข้อความยาว
    if any(x.lower() in s for x in LEVEL_MAP["ป.เอก"]):
        return "ป.เอก"
    if any(x.lower() in s for x in LEVEL_MAP["ป.โท"]):
        return "ป.โท"
    return "อื่น ๆ"


def normalize_quartile(v):
    s = clean(v).upper().replace(" ", "")
    m = re.search(r"Q([12])", s)
    return f"Q{m.group(1)}" if m else "อื่น ๆ"


def infer_level_from_row(row):
    """ใช้เมื่อไฟล์ไม่มีคอลัมน์ชื่อระดับโดยตรง: ค้นหาคำ ป.โท/ป.เอก ในทุกเซลล์ของแถว"""
    text = " | ".join(clean(v) for v in row.tolist()).lower()
    if any(x.lower() in text for x in LEVEL_MAP["ป.เอก"]):
        return "ป.เอก"
    if any(x.lower() in text for x in LEVEL_MAP["ป.โท"]):
        return "ป.โท"
    return "อื่น ๆ"


def infer_quartile_from_row(row):
    """ค้นหา Q1/Q2 ในทุกเซลล์ของแถว กรณีชื่อคอลัมน์ใน Excel ไม่ตรงมาตรฐาน"""
    for v in row.tolist():
        q = normalize_quartile(v)
        if q in ("Q1", "Q2"):
            return q
    return "อื่น ๆ"


def read_workbook(uploaded):
    uploaded.seek(0)
    sheets = pd.read_excel(uploaded, sheet_name=None, header=0)
    candidates = []
    for name, df in sheets.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        df.columns = [clean(c) for c in df.columns]
        # ไม่บังคับชื่อคอลัมน์ระดับอีกต่อไป
        level_col = find_col(df, ALIASES["level"])
        quartile_col = find_col(df, ALIASES["quartile"])
        row_level_hits = int(df.apply(lambda r: infer_level_from_row(r) != "อื่น ๆ", axis=1).sum())
        row_q_hits = int(df.apply(lambda r: infer_quartile_from_row(r) != "อื่น ๆ", axis=1).sum())
        score = (5 if level_col else 0) + (5 if quartile_col else 0) + row_level_hits + row_q_hits
        candidates.append((score, name, df))
    if not candidates:
        raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2], sheets


def prepare(df):
    out = df.copy()
    cols = {k: find_col(out, aliases) for k, aliases in ALIASES.items()}

    # ระดับ: ใช้คอลัมน์ตรงก่อน ถ้าไม่มีให้ตรวจทุกเซลล์ในแต่ละแถว
    if cols["level"]:
        out["__ระดับ"] = out[cols["level"]].map(normalize_level)
    else:
        out["__ระดับ"] = out.apply(infer_level_from_row, axis=1)

    # Q1/Q2: ใช้คอลัมน์ตรงก่อน ถ้าไม่มีให้ตรวจทุกเซลล์ในแต่ละแถว
    if cols["quartile"]:
        out["__Quartile"] = out[cols["quartile"]].map(normalize_quartile)
        fallback = out["__Quartile"] == "อื่น ๆ"
        if fallback.any():
            out.loc[fallback, "__Quartile"] = out.loc[fallback].apply(infer_quartile_from_row, axis=1)
    else:
        out["__Quartile"] = out.apply(infer_quartile_from_row, axis=1)

    out["__คณะ"] = out[cols["faculty"]].map(clean) if cols["faculty"] else "ไม่ระบุ"
    out["__สาขา"] = out[cols["program"]].map(clean) if cols["program"] else "ไม่ระบุ"
    out["__รหัสนิสิต"] = out[cols["student_id"]].map(clean) if cols["student_id"] else ""
    out["__ชื่อผลงาน"] = out[cols["title"]].map(clean) if cols["title"] else ""

    # ตัดแถวที่เป็นหัวตารางซ้ำหรือไม่มีทั้งระดับและ Q1/Q2
    out = out[(out["__ระดับ"].isin(["ป.โท", "ป.เอก"])) & (out["__Quartile"].isin(["Q1", "Q2"]))].copy()
    return out, cols


def make_stats(g):
    columns = ["คณะ", "สาขา", "Q1", "Q2", "รวม Q1+Q2", "%Q1", "%Q2"]
    if g.empty:
        return pd.DataFrame(columns=columns)

    rows = []
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

    result = pd.DataFrame(rows, columns=columns)
    if not result.empty:
        q1 = int(result["Q1"].sum())
        q2 = int(result["Q2"].sum())
        total = q1 + q2
        result.loc[len(result)] = [
            "รวมทั้งหมด", "", q1, q2, total,
            round(q1 * 100 / total, 2) if total else 0,
            round(q2 * 100 / total, 2) if total else 0,
        ]
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
    side = Side(style="thin", color="D9D9D9")
    border = Border(left=side, right=side, top=side, bottom=side)
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
st.caption("ระบบตรวจจับ ป.โท / ป.เอก และ Q1 / Q2 จากข้อมูลจริง แม้ชื่อคอลัมน์จะไม่ตรงมาตรฐาน")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])

if uploaded:
    try:
        sheet_name, raw, all_sheets = read_workbook(uploaded)
        data, cols = prepare(raw)

        if data.empty:
            raise ValueError("ไม่พบแถวข้อมูลที่ตรวจจับได้ว่าเป็น ป.โท/ป.เอก และ Q1/Q2")

        st.success(f"อ่านข้อมูลจาก Sheet: {sheet_name} | พบข้อมูลที่ใช้คำนวณ {len(data):,} แถว")
        with st.expander("ตรวจสอบการจับคู่ข้อมูล"):
            st.json({k: (str(v) if v is not None else None) for k, v in cols.items()})
            st.write("จำนวน ป.โท:", int((data["__ระดับ"] == "ป.โท").sum()))
            st.write("จำนวน ป.เอก:", int((data["__ระดับ"] == "ป.เอก").sum()))
            st.write("จำนวน Q1:", int((data["__Quartile"] == "Q1").sum()))
            st.write("จำนวน Q2:", int((data["__Quartile"] == "Q2").sum()))

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

        with st.expander("ข้อมูลที่ระบบนำมาคำนวณ"):
            st.dataframe(data, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(str(e))
else:
    st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ ป.โท และ ป.เอก")
