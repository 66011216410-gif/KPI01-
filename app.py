import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="KPI01 | สถิติผลงานบัณฑิตศึกษา", page_icon="📊", layout="wide")

# =========================
# ชื่อคอลัมน์ที่รองรับ
# =========================
LEVEL_ALIASES = ["ระดับ", "ระดับการศึกษา", "ระดับปริญญา", "วุฒิ", "degree", "level"]
PRIMARY_Q_ALIASES = ["ผลงานที่ตีพิมพ์ Q1-Q2", "ผลงานที่ตีพิมพ์Q1-Q2"]
FALLBACK_Q_ALIASES = ["ฐานข้อมูล"]
FACULTY_ALIASES = ["คณะ", "กลุ่ม", "faculty", "school"]
PROGRAM_ALIASES = ["สาขา", "สาขาวิชา", "หลักสูตร", "program", "major"]
GROUP_ALIASES = ["กลุ่ม/คณะ/สาขา", "กลุ่ม คณะ สาขา", "กลุ่ม/คณะ", "ชื่อคณะ", "หน่วยงาน"]
SYSTEM_ALIASES = ["ระบบ", "ประเภทระบบ", "ภาค", "แผนการศึกษา"]
STUDENT_ID_ALIASES = ["รหัสนิสิต", "รหัสนักศึกษา", "รหัส", "student id", "student_id", "id"]
STATUS_ALIASES = ["สถานะนิสิต", "สถานะนักศึกษา", "สถานะ", "status"]


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


def normalize_system(v):
    s = clean(v).lower().replace(" ", "")
    if not s:
        return "ไม่ระบุ"
    if any(x in s for x in ["นอกเวลา", "นอกเวลาราชการ", "parttime", "part-time"]):
        return "นอกเวลา"
    if any(x in s for x in ["ในเวลา", "ในเวลาราชการ", "fulltime", "full-time"]):
        return "ในเวลา"
    return clean(v)


def is_success(v):
    s = clean(v).lower()
    if not s:
        return True
    return any(x in s for x in ["สำเร็จการศึกษา", "สำเร็จ", "graduate", "graduated"])


def read_workbook(uploaded):
    uploaded.seek(0)
    # หัวคอลัมน์อยู่แถวที่ 2 ของไฟล์ Excel
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
    group_col = find_col(out, GROUP_ALIASES)
    system_col = find_col(out, SYSTEM_ALIASES)
    student_id_col = find_col(out, STUDENT_ID_ALIASES)
    status_col = find_col(out, STATUS_ALIASES)

    if not level_col:
        raise ValueError("ไม่พบคอลัมน์ 'ระดับ' ในแถวที่ 2")
    if not q_col:
        raise ValueError("ไม่พบคอลัมน์ 'ผลงานที่ตีพิมพ์ Q1-Q2' และไม่พบคอลัมน์สำรอง 'ฐานข้อมูล' ในแถวที่ 2")

    out["__ระดับ"] = out[level_col].map(normalize_level)
    out["__Q"] = out[q_col].map(normalize_q)
    out["__ระบบ"] = out[system_col].map(normalize_system) if system_col else "ไม่ระบุ"
    out["__รหัสนิสิต"] = out[student_id_col].map(clean) if student_id_col else ""
    out["__สถานะ"] = out[status_col].map(is_success) if status_col else True

    if group_col:
        out["__กลุ่ม"] = out[group_col].map(clean)
    elif faculty_col and program_col:
        out["__กลุ่ม"] = out[faculty_col].map(clean) + out[program_col].map(lambda x: " / " + clean(x) if clean(x) else "")
    elif faculty_col:
        out["__กลุ่ม"] = out[faculty_col].map(clean)
    elif program_col:
        out["__กลุ่ม"] = out[program_col].map(clean)
    else:
        out["__กลุ่ม"] = "ไม่ระบุ"

    # ข้อมูลระดับปริญญาโท/เอกทั้งหมด
    out = out[out["__ระดับ"].isin(["ปริญญาโท", "ปริญญาเอก"])].copy()
    return out, {
        "ระดับ": level_col,
        "Q1/Q2 ที่ใช้": q_col,
        "แหล่ง Q1/Q2": "ผลงานที่ตีพิมพ์ Q1-Q2" if primary_q_col else "ฐานข้อมูล",
        "กลุ่ม/คณะ/สาขา": group_col or faculty_col or program_col,
        "ระบบ": system_col,
        "รหัสนิสิต": student_id_col,
        "สถานะ": status_col,
    }


def unique_count(df):
    if df.empty:
        return 0
    if "__รหัสนิสิต" in df.columns and df["__รหัสนิสิต"].astype(str).str.strip().ne("").any():
        return int(df.loc[df["__รหัสนิสิต"].astype(str).str.strip() != "", "__รหัสนิสิต"].nunique())
    return int(len(df))


def make_kpi_table(data):
    """สร้างตารางรูปแบบเดียวกับภาพตัวอย่าง"""
    columns = [
        "ลำดับ",
        "กลุ่ม/คณะ/สาขา",
        "ระบบในเวลาราชการ",
        "ระบบนอกเวลาราชการ",
        "จำนวนผู้สำเร็จการศึกษา (A)",
        "รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2",
        "ร้อยละผลงานวิจัยรวมระดับนานาชาติ Q1-Q2",
    ]
    rows = []
    if data.empty:
        return pd.DataFrame(columns=columns)

    # ผู้สำเร็จการศึกษาใช้สถานะสำเร็จการศึกษา ส่วนผลงานนับจาก Q1/Q2
    graduates = data[data["__สถานะ"]].copy()
    publications = data[data["__Q"].isin(["Q1", "Q2"])].copy()

    groups = sorted(set(graduates["__กลุ่ม"].dropna().astype(str)) | set(publications["__กลุ่ม"].dropna().astype(str)))
    for idx, group in enumerate(groups, start=1):
        g = graduates[graduates["__กลุ่ม"].astype(str) == group]
        p = publications[publications["__กลุ่ม"].astype(str) == group]

        # A แยกระบบ โดยนับนิสิตไม่ซ้ำถ้ามีรหัสนิสิต
        in_g = g[g["__ระบบ"] == "ในเวลา"]
        out_g = g[g["__ระบบ"] == "นอกเวลา"]
        in_count = unique_count(in_g)
        out_count = unique_count(out_g)
        total_a = in_count + out_count

        pub_count = int(len(p))
        percent = round(pub_count * 100 / total_a, 2) if total_a else 0.0

        rows.append({
            "ลำดับ": idx,
            "กลุ่ม/คณะ/สาขา": group,
            "ระบบในเวลาราชการ": in_count,
            "ระบบนอกเวลาราชการ": out_count,
            "จำนวนผู้สำเร็จการศึกษา (A)": total_a,
            "รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2": pub_count,
            "ร้อยละผลงานวิจัยรวมระดับนานาชาติ Q1-Q2": percent,
        })

    result = pd.DataFrame(rows, columns=columns)
    if not result.empty:
        total_in = int(result["ระบบในเวลาราชการ"].sum())
        total_out = int(result["ระบบนอกเวลาราชการ"].sum())
        total_a = int(result["จำนวนผู้สำเร็จการศึกษา (A)"].sum())
        total_pub = int(result["รวมจำนวนผลงานตีพิมพ์ระดับนานาชาติ Q1-Q2"].sum())
        total_pct = round(total_pub * 100 / total_a, 2) if total_a else 0.0
        result.loc[len(result)] = [
            "",
            "รวมทั้งหมด",
            total_in,
            total_out,
            total_a,
            total_pub,
            total_pct,
        ]
    return result


def export_excel(master, doctor):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="ปริญญาโท", index=False, startrow=5)
        doctor.to_excel(writer, sheet_name="ปริญญาเอก", index=False, startrow=5)

    bio.seek(0)
    wb = load_workbook(bio)
    header_fill = PatternFill("solid", fgColor="355E20")
    sub_fill = PatternFill("solid", fgColor="548235")
    white_font = Font(name="Tahoma", color="FFFFFF", bold=True)
    thin = Side(style="thin", color="000000")

    for ws in wb.worksheets:
        # หัวตาราง 3 ชั้น ให้ใกล้เคียงภาพตัวอย่าง
        ws.merge_cells("A1:A3")
        ws["A1"] = "ลำดับ"
        ws.merge_cells("B1:B3")
        ws["B1"] = "กลุ่ม/คณะ/สาขา"
        ws.merge_cells("C1:D1")
        ws["C1"] = "ระบบ"
        ws["C2"] = "ระบบในเวลาราชการ"
        ws["D2"] = "ระบบนอกเวลาราชการ"
        ws.merge_cells("E1:E3")
        ws["E1"] = "จำนวนผู้สำเร็จการศึกษา (A)"
        ws.merge_cells("F1:G1")
        ws["F1"] = "1.2.4 จำนวนผลงานระดับบัณฑิตศึกษาที่สามารถตีพิมพ์ในระดับนานาชาติ SCOPUS Q1-Q2"
        ws.merge_cells("F2:F3")
        ws["F2"] = "รวมจำนวนผลงานตีพิมพ์ในวารสารระดับนานาชาติ Q1-Q2"
        ws.merge_cells("G2:G3")
        ws["G2"] = "ร้อยละผลงานวิจัยรวมจำนวนผลงานตีพิมพ์ในวารสารระดับนานาชาติ Q1-Q2"

        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=7):
            for cell in row:
                cell.fill = header_fill if cell.row == 1 else sub_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # ข้อมูลจริงเริ่มแถว 6
        for row in ws.iter_rows(min_row=6, max_row=ws.max_row, min_col=1, max_col=7):
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            row[0].alignment = Alignment(horizontal="center", vertical="center")
            for c in [2, 3, 4, 5, 6, 7]:
                row[c - 1].alignment = Alignment(horizontal="center" if c >= 3 else "left", vertical="center", wrap_text=True)

        ws.freeze_panes = "A6"
        widths = [8, 42, 18, 18, 22, 28, 32]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.row_dimensions[1].height = 42
        ws.row_dimensions[2].height = 60
        ws.row_dimensions[3].height = 60

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =========================
# หน้าเว็บ
# =========================
st.title("📊 KPI01 — ตารางสถิติแบบตัวอย่าง")
st.caption("หัวคอลัมน์ Excel อยู่แถวที่ 2 | แยกปริญญาโท/ปริญญาเอก | Q1/Q2 ใช้ 'ผลงานที่ตีพิมพ์ Q1-Q2' และใช้ 'ฐานข้อมูล' เป็นสำรอง")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])

if uploaded:
    try:
        sheet_name, raw = read_workbook(uploaded)
        data, cols = prepare(raw)

        master = make_kpi_table(data[data["__ระดับ"] == "ปริญญาโท"])
        doctor = make_kpi_table(data[data["__ระดับ"] == "ปริญญาเอก"])

        st.success(f"อ่าน Sheet: {sheet_name} | หัวคอลัมน์แถวที่ 2")

        with st.expander("ตรวจสอบการจับคอลัมน์"):
            st.write(cols)
            st.write("ข้อมูลปริญญาโท:", len(data[data["__ระดับ"] == "ปริญญาโท"]))
            st.write("ข้อมูลปริญญาเอก:", len(data[data["__ระดับ"] == "ปริญญาเอก"]))

        tab1, tab2 = st.tabs(["🎓 ปริญญาโท", "🎓 ปริญญาเอก"])
        with tab1:
            st.subheader("ตารางสถิติ ปริญญาโท")
            st.dataframe(master, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลด CSV ปริญญาโท", master.to_csv(index=False).encode("utf-8-sig"), "KPI01_ปริญญาโท.csv", "text/csv")
        with tab2:
            st.subheader("ตารางสถิติ ปริญญาเอก")
            st.dataframe(doctor, use_container_width=True, hide_index=True)
            st.download_button("ดาวน์โหลด CSV ปริญญาเอก", doctor.to_csv(index=False).encode("utf-8-sig"), "KPI01_ปริญญาเอก.csv", "text/csv")

        excel = export_excel(master, doctor)
        st.download_button(
            "📥 ดาวน์โหลด Excel รูปแบบเดียวกับตัวอย่าง",
            excel,
            "KPI01_ตารางสถิติ_ปโท_ปเอก.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(str(e))
else:
    st.info("อัปโหลด Excel เพื่อสร้างตารางสถิติ")
