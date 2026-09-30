import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm_header(v):
    return re.sub(r"\s+", "", clean(v)).lower()


def find_exact_or_contains(df, names):
    headers = {norm_header(c): c for c in df.columns}
    for name in names:
        n = norm_header(name)
        if n in headers:
            return headers[n]
    for c in df.columns:
        n = norm_header(c)
        if any(norm_header(name) in n for name in names):
            return c
    return None


def normalize_level(v):
    s = clean(v).lower().replace(" ", "")
    if "ป.เอก" in s or "ปริญญาเอก" in s or "ph.d" in s or "phd" in s or "doctoral" in s:
        return "ป.เอก"
    if "ป.โท" in s or "ปริญญาโท" in s or "master" in s:
        return "ป.โท"
    return "อื่น ๆ"


def normalize_q(v):
    s = clean(v).upper().replace(" ", "")
    # ค่าที่ใช้ในคอลัมน์ "ผลงานที่ตีพิมพ์ Q1-Q2"
    if s in {"Q1", "Q1.0"}:
        return "Q1"
    if s in {"Q2", "Q2.0"}:
        return "Q2"
    # รองรับข้อความ เช่น Q1, Q1-Q2, Q2 เป็นต้น
    if re.search(r"Q1", s) and not re.search(r"Q2", s):
        return "Q1"
    if re.search(r"Q2", s) and not re.search(r"Q1", s):
        return "Q2"
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
        level_col = find_exact_or_contains(df, ["ระดับ"])
        publication_col = find_exact_or_contains(df, ["ผลงานที่ตีพิมพ์ Q1-Q2", "ผลงานที่ตีพิมพ์", "Q1-Q2"])
        score = (100 if level_col else 0) + (100 if publication_col else 0)
        candidates.append((score, name, df, level_col, publication_col))
    if not candidates:
        raise ValueError("ไม่พบข้อมูลในไฟล์ Excel")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0]


def prepare(df, level_col, publication_col):
    if not level_col:
        raise ValueError("ไม่พบคอลัมน์ 'ระดับ'")
    if not publication_col:
        raise ValueError("ไม่พบคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2'")

    out = df.copy()
    out["__ระดับ"] = out[level_col].map(normalize_level)
    out["__Q"] = out[publication_col].map(normalize_q)

    # ใช้เฉพาะข้อมูลที่เป็น ป.โท/ป.เอก และมี Q1 หรือ Q2 ในคอลัมน์ผลงานที่ตีพิมพ์
    out = out[
        out["__ระดับ"].isin(["ป.โท", "ป.เอก"])
        & out["__Q"].isin(["Q1", "Q2"])
    ].copy()
    return out


def make_stats(g):
    q1 = int((g["__Q"] == "Q1").sum())
    q2 = int((g["__Q"] == "Q2").sum())
    total = q1 + q2
    return pd.DataFrame([{
        "ระดับ": clean(g["__ระดับ"].iloc[0]) if not g.empty else "",
        "Q1": q1,
        "Q2": q2,
        "รวม Q1-Q2": total,
    }])


def make_summary(data):
    rows = []
    for level in ["ป.โท", "ป.เอก"]:
        g = data[data["__ระดับ"] == level]
        q1 = int((g["__Q"] == "Q1").sum())
        q2 = int((g["__Q"] == "Q2").sum())
        rows.append({"ระดับ": level, "Q1": q1, "Q2": q2, "รวม Q1-Q2": q1 + q2})
    return pd.DataFrame(rows)


def to_excel(summary, master, doctor):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="สรุป", index=False)
        master.to_excel(writer, sheet_name="ป.โท", index=False)
        doctor.to_excel(writer, sheet_name="ป.เอก", index=False)
    bio.seek(0)
    return bio


st.title("📊 KPI01 — ตารางสถิติผลงานระดับบัณฑิตศึกษา")
st.caption("ป.โท/ป.เอก อ่านจากคอลัมน์ 'ระดับ' และ Q1/Q2 อ่านจากคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2'")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])

if uploaded:
    try:
        score, sheet_name, raw, level_col, publication_col = read_workbook(uploaded)
        data = prepare(raw, level_col, publication_col)

        st.success(f"อ่าน Sheet: {sheet_name} | ระดับ = {level_col} | ผลงาน = {publication_col}")
        st.info(f"ใช้ข้อมูลสำหรับคำนวณทั้งหมด {len(data):,} รายการ")

        summary = make_summary(data)
        master = make_stats(data[data["__ระดับ"] == "ป.โท"])
        doctor = make_stats(data[data["__ระดับ"] == "ป.เอก"])

        st.subheader("สรุปสถิติ")
        st.dataframe(summary, use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            st.metric("ป.โท — Q1", int(summary.loc[summary["ระดับ"] == "ป.โท", "Q1"].iloc[0]))
            st.metric("ป.โท — Q2", int(summary.loc[summary["ระดับ"] == "ป.โท", "Q2"].iloc[0]))
        with c2:
            st.metric("ป.เอก — Q1", int(summary.loc[summary["ระดับ"] == "ป.เอก", "Q1"].iloc[0]))
            st.metric("ป.เอก — Q2", int(summary.loc[summary["ระดับ"] == "ป.เอก", "Q2"].iloc[0]))

        tab1, tab2 = st.tabs(["🎓 ป.โท", "🎓 ป.เอก"])
        with tab1:
            st.dataframe(master, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(doctor, use_container_width=True, hide_index=True)

        excel = to_excel(summary, master, doctor)
        st.download_button(
            "📥 ดาวน์โหลด Excel KPI01",
            excel,
            "KPI01_สถิติ_ปโท_ปเอก.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("ตรวจสอบข้อมูลที่ระบบนำมาคำนวณ"):
            st.dataframe(data, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(str(e))
else:
    st.info("อัปโหลดไฟล์ทดลองเพื่อสร้างตาราง KPI01")
