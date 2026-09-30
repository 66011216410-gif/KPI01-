import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")

LEVEL_ALIASES = ["ระดับ", "ระดับการศึกษา", "ระดับปริญญา", "วุฒิ", "degree", "level"]
PRIMARY_Q_ALIASES = ["ผลงานที่ตีพิมพ์ Q1-Q2", "ผลงานที่ตีพิมพ์Q1-Q2"]
FALLBACK_Q_ALIASES = ["ฐานข้อมูล"]
FACULTY_ALIASES = ["คณะ", "faculty", "school"]
PROGRAM_ALIASES = ["สาขา", "สาขาวิชา", "หลักสูตร", "program", "major"]


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def find_col(df, aliases):
    cols = {re.sub(r"\s+", "", clean(c)).lower(): c for c in df.columns}
    for a in aliases:
        key = re.sub(r"\s+", "", a).lower()
        if key in cols:
            return cols[key]
    for c in df.columns:
        cc = re.sub(r"\s+", "", clean(c)).lower()
        if any(re.sub(r"\s+", "", a).lower() in cc for a in aliases):
            return c
    return None


def normalize_level(v):
    s = clean(v).lower().replace(" ", "")
    if "ปริญญาเอก" in s or "ป.เอก" in s or "ph.d" in s or "phd" in s or "doctoral" in s:
        return "ปริญญาเอก"
    if "ปริญญาโท" in s or "ป.โท" in s or "master" in s:
        return "ปริญญาโท"
    return "อื่น ๆ"


def normalize_q(v):
    s = clean(v).upper().replace(" ", "")
    if re.search(r"Q1", s):
        return "Q1"
    if re.search(r"Q2", s):
        return "Q2"
    return "อื่น ๆ"


def read_workbook(uploaded):
    uploaded.seek(0)
    # ข้อมูลจริงมีหัวคอลัมน์อยู่แถวที่ 2 ของ Excel: header=1
    sheets = pd.read_excel(uploaded, sheet_name=None, header=1)
    candidates = []
    for name, df in sheets.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        df.columns = [clean(c) for c in df.columns]
        level_col = find_col(df, LEVEL_ALIASES)
        primary_q_col = find_col(df, PRIMARY_Q_ALIASES)
        fallback_q_col = find_col(df, FALLBACK_Q_ALIASES)
        q_col = primary_q_col or fallback_q_col
        score = (10 if level_col else 0) + (10 if primary_q_col else (5 if fallback_q_col else 0))
        candidates.append((score, name, df))
    if not candidates:
        raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2]


def prepare(df):
    out = df.copy()
    level_col = find_col(out, LEVEL_ALIASES)
    primary_q_col = find_col(out, PRIMARY_Q_ALIASES)
    fallback_q_col = find_col(out, FALLBACK_Q_ALIASES)
    q_col = primary_q_col or fallback_q_col
    faculty_col = find_col(out, FACULTY_ALIASES)
    program_col = find_col(out, PROGRAM_ALIASES)

    if not level_col:
        raise ValueError("ไม่พบคอลัมน์ 'ระดับ' ในแถวหัวตารางที่ 2")
    if not q_col:
        raise ValueError("ไม่พบคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2' และไม่พบคอลัมน์สำรอง 'ฐานข้อมูล' ในแถวที่ 2")

    out["__ระดับ"] = out[level_col].map(normalize_level)
    out["__Q"] = out[q_col].map(normalize_q)
    out["__คณะ"] = out[faculty_col].map(clean) if faculty_col else "ไม่ระบุ"
    out["__สาขา"] = out[program_col].map(clean) if program_col else "ไม่ระบุ"

    out = out[
        out["__ระดับ"].isin(["ปริญญาโท", "ปริญญาเอก"])
        & out["__Q"].isin(["Q1", "Q2"])
    ].copy()

    return out, {
        "ระดับ": level_col,
        "Q1/Q2 ที่ใช้": q_col,
        "แหล่ง Q1/Q2": "ผลงานที่ตีพิมพ์ Q1-Q2" if primary_q_col else "ฐานข้อมูล",
        "คณะ": faculty_col,
        "สาขา": program_col,
    }


def make_stats(g):
    columns = ["คณะ", "สาขา", "Q1", "Q2", "รวม Q1-Q2"]
    if g.empty:
        return pd.DataFrame(columns=columns)
    rows = []
    for (faculty, program), x in g.groupby(["__คณะ", "__สาขา"], dropna=False, sort=True):
        q1 = int((x["__Q"] == "Q1").sum())
        q2 = int((x["__Q"] == "Q2").sum())
        rows.append({"คณะ": faculty or "ไม่ระบุ", "สาขา": program or "ไม่ระบุ", "Q1": q1, "Q2": q2, "รวม Q1-Q2": q1 + q2})
    result = pd.DataFrame(rows, columns=columns)
    q1 = int(result["Q1"].sum())
    q2 = int(result["Q2"].sum())
    result.loc[len(result)] = ["รวมทั้งหมด", "", q1, q2, q1 + q2]
    return result


def make_summary(data):
    rows = []
    for level in ["ปริญญาโท", "ปริญญาเอก"]:
        x = data[data["__ระดับ"] == level]
        q1 = int((x["__Q"] == "Q1").sum())
        q2 = int((x["__Q"] == "Q2").sum())
        rows.append({"ระดับ": level, "Q1": q1, "Q2": q2, "รวม Q1-Q2": q1 + q2})
    return pd.DataFrame(rows)


def export_excel(summary, master, doctor):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="สรุป", index=False)
        master.to_excel(writer, sheet_name="ปริญญาโท", index=False)
        doctor.to_excel(writer, sheet_name="ปริญญาเอก", index=False)
    bio.seek(0)
    return bio


st.title("📊 KPI01 — ตารางสถิติผลงานระดับบัณฑิตศึกษา")
st.caption("หัวคอลัมน์อยู่แถวที่ 2 | ระดับใช้คอลัมน์ 'ระดับ' | Q1/Q2 ใช้ 'ผลงานที่ตีพิมพ์ Q1-Q2' และใช้ 'ฐานข้อมูล' เป็นสำรอง")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])

if uploaded:
    try:
        sheet_name, raw = read_workbook(uploaded)
        data, cols = prepare(raw)
        summary = make_summary(data)
        master = make_stats(data[data["__ระดับ"] == "ปริญญาโท"])
        doctor = make_stats(data[data["__ระดับ"] == "ปริญญาเอก"])

        st.success(f"อ่าน Sheet: {sheet_name} | หัวคอลัมน์แถวที่ 2 | ใช้ข้อมูล {len(data):,} แถว")
        with st.expander("ตรวจสอบคอลัมน์"):
            st.write(cols)
            st.write("ปริญญาโท:", int((data["__ระดับ"] == "ปริญญาโท").sum()))
            st.write("ปริญญาเอก:", int((data["__ระดับ"] == "ปริญญาเอก").sum()))
            st.write("Q1:", int((data["__Q"] == "Q1").sum()))
            st.write("Q2:", int((data["__Q"] == "Q2").sum()))

        st.subheader("สรุป ปริญญาโท / ปริญญาเอก")
        st.dataframe(summary, use_container_width=True, hide_index=True)

        tab1, tab2 = st.tabs(["🎓 ปริญญาโท", "🎓 ปริญญาเอก"])
        with tab1:
            st.dataframe(master, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลด ปริญญาโท CSV", master.to_csv(index=False).encode("utf-8-sig"), "kpi01_master.csv", "text/csv")
        with tab2:
            st.dataframe(doctor, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลด ปริญญาเอก CSV", doctor.to_csv(index=False).encode("utf-8-sig"), "kpi01_doctor.csv", "text/csv")

        excel = export_excel(summary, master, doctor)
        st.download_button("📥 ดาวน์โหลด Excel KPI01", excel, "KPI01_ปริญญาโท_ปริญญาเอก.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(str(e))
else:
    st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ")
