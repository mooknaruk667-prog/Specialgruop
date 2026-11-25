import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.request
import os
import io
import msoffcrypto
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, date

# 🔑 ตั้งรหัสผ่านสำหรับเปิดไฟล์ Excel ตรงนี้ครับ
EXCEL_PASSWORD = "1234"

# ------------------------------------------------
# 1. ตั้งค่าหน้าเพจ & ระบบดึงข้อมูลและฟอนต์
# ------------------------------------------------
st.set_page_config(page_title="Dashboard กลุ่มเปราะบาง", layout="wide")

@st.cache_data(ttl=600)
def load_data():
    sheet_id = "1YRt4MqgSE5kS2ZHMn--COpWPDpCJjNd6AzmljG08-NI"
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    df = pd.read_csv(csv_url)
    return df

@st.cache_resource
def load_fonts():
    font_url = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf"
    font_bold_url = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Bold.ttf"
    if not os.path.exists("Sarabun-Regular.ttf"):
        urllib.request.urlretrieve(font_url, "Sarabun-Regular.ttf")
    if not os.path.exists("Sarabun-Bold.ttf"):
        urllib.request.urlretrieve(font_bold_url, "Sarabun-Bold.ttf")

try:
    df_raw = load_data()
except Exception as e:
    st.error(f"ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบการแชร์ไฟล์: {e}")
    st.stop()

# กรองเฉพาะผู้ต้องขังที่คอลัมน์ "อยู่" มีข้อความว่า "อยู่"
if 'อยู่' in df_raw.columns:
    df = df_raw[df_raw['อยู่'].astype(str).str.strip() == 'อยู่'].copy()
else:
    st.error("ไม่พบคอลัมน์ 'อยู่' ในฐานข้อมูล")
    st.stop()

if 'สถานะ' not in df.columns:
    st.error("ไม่พบคอลัมน์ 'สถานะ' ในฐานข้อมูล")
    st.stop()

df['สถานะ'] = df['สถานะ'].fillna('').astype(str)
df['เพศ'] = df['เพศ'].fillna('ไม่ระบุ').astype(str).str.strip()

# --- ธีมสีแยกเพศ ---
pie_gender_color = {'ชาย': '#3498db', 'หญิง': '#ff66b3'}
bar_gender_color = {'ชาย': '#2A9D8F', 'หญิง': '#E76F51'}

MONTHS_TH = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
]

def format_thai_date(d):
    return f"{d.day} {MONTHS_TH[d.month - 1]} {d.year}"

# ------------------------------------------------
# ฟังก์ชันช่วยคำนวณสถิติจากคอลัมน์ "สถานะ"
# ------------------------------------------------
def has_keyword(series, keywords):
    pattern = '|'.join(keywords)
    return series.str.contains(pattern, case=False, na=False)

def get_gender_summary(sub_df):
    total = sub_df.shape[0]
    m = sub_df[sub_df['เพศ'] == 'ชาย'].shape[0]
    f = sub_df[sub_df['เพศ'] == 'หญิง'].shape[0]
    pct_m = (m / total * 100) if total > 0 else 0
    pct_f = (f / total * 100) if total > 0 else 0
    return total, m, f, pct_m, pct_f

def build_two_group_df(df_base, kw_get, kw_not_get, label_get, label_not_get):
    sub_get = df_base[has_keyword(df_base['สถานะ'], kw_get)].copy()
    sub_get['กลุ่มย่อย'] = label_get
    sub_not = df_base[has_keyword(df_base['สถานะ'], kw_not_get)].copy()
    sub_not['กลุ่มย่อย'] = label_not_get
    combined = pd.concat([sub_get, sub_not], ignore_index=True)
    return combined, sub_get, sub_not

# เตรียมชุดข้อมูลทั้ง 5 ส่วน
df_elder_all, df_elder_get, df_elder_not = build_two_group_df(
    df, ['สูงอายุได้เบี้ย'], ['สูงอายุไม่ได้เบี้ย'], 'สูงอายุได้เบี้ย', 'สูงอายุไม่ได้เบี้ย'
)
df_dis_all, df_dis_get, df_dis_not = build_two_group_df(
    df, ['พิการได้เบี้ย'], ['พิการไม่ได้เบี้ย'], 'พิการได้เบี้ย', 'พิการไม่ได้เบี้ย'
)
df_pregnant = df[has_keyword(df['สถานะ'], ['ตั้งครรภ์', 'ท้อง'])].copy()

# ส่วนที่ 4: ผู้ป่วย AIDS (3 กลุ่ม: ขึ้นทะเบียนแล้ว / ประสงค์ / ไม่รับเบี้ย)
df_aids_reg = df[has_keyword(df['สถานะ'], ['AIDS ขึ้นทะเบียนแล้ว', 'AIDSขึ้นทะเบียนแล้ว'])].copy()
df_aids_reg['กลุ่มย่อย'] = 'AIDS ขึ้นทะเบียนแล้ว'

df_aids_wish = df[has_keyword(df['สถานะ'], ['AIDS ประสงค์', 'AIDSประสงค์'])].copy()
df_aids_wish['กลุ่มย่อย'] = 'AIDS ประสงค์'

df_aids_no = df[has_keyword(df['สถานะ'], ['AIDS ไม่รับเบี้ย', 'AIDSไม่รับเบี้ย'])].copy()
df_aids_no['กลุ่มย่อย'] = 'AIDS ไม่รับเบี้ย'

df_aids_all = pd.concat([df_aids_reg, df_aids_wish, df_aids_no], ignore_index=True)

# ส่วนที่ 5: บัตรประชาชน (บัตรหาย-ไม่ประสงค์ทำใหม่ และ อยู่ในทะเบียนกลาง)
kw_no_id = ['ไม่มีบัตรประชาชน', 'บัตรปชช หาย', 'ไม่มีบัตร']
kw_central = ['อยู่ในทะเบียนกลาง', 'ทะเบียนกลาง']
df_no_id = df[has_keyword(df['สถานะ'], kw_no_id)].copy()
df_central_reg = df[has_keyword(df['สถานะ'], kw_central)].copy()

# คำนวณยอดชาย-หญิงของบัตรหายและทะเบียนกลาง
_, m_no_id, f_no_id, _, _ = get_gender_summary(df_no_id)
_, m_central, f_central, _, _ = get_gender_summary(df_central_reg)

# ฟังก์ชันคำนวณยอดผู้มีบัตรประชาชนแยกเพศตามสูตรที่กำหนด
def calc_id_card_stats(total_all):
    no_id_m_sum = m_no_id + m_central
    no_id_f_sum = f_no_id + f_central
    no_id_total_sum = no_id_m_sum + no_id_f_sum

    has_id_total = max(total_all - no_id_total_sum, 0)
    # สูตรแยกเพศ: ผู้ต้องขังทั้งหมดที่มีบัตรประชาชน - (บัตรหายฯ + ทะเบียนกลาง ของเพศนั้น)
    has_id_m = max(has_id_total - no_id_m_sum, 0)
    has_id_f = max(has_id_total - no_id_f_sum, 0)

    pct_has_id = (has_id_total / total_all * 100) if total_all > 0 else 0
    pct_has_id_m = (has_id_m / has_id_total * 100) if has_id_total > 0 else 0
    pct_has_id_f = (has_id_f / has_id_total * 100) if has_id_total > 0 else 0

    pct_no_id_all = (no_id_total_sum / total_all * 100) if total_all > 0 else 0
    pct_no_id_m = (no_id_m_sum / no_id_total_sum * 100) if no_id_total_sum > 0 else 0
    pct_no_id_f = (no_id_f_sum / no_id_total_sum * 100) if no_id_total_sum > 0 else 0

    return {
        'total_all': total_all,
        'has_id_total': has_id_total, 'has_id_m': has_id_m, 'has_id_f': has_id_f,
        'pct_has_id': pct_has_id, 'pct_has_id_m': pct_has_id_m, 'pct_has_id_f': pct_has_id_f,
        'no_id_total': no_id_total_sum, 'no_id_m': no_id_m_sum, 'no_id_f': no_id_f_sum,
        'pct_no_id_all': pct_no_id_all, 'pct_no_id_m': pct_no_id_m, 'pct_no_id_f': pct_no_id_f,
    }

# เก็บค่าใน session_state สำหรับปุ่มระบุยอดผู้ต้องขังทั้งหมด
if 'total_prison_inmates' not in st.session_state:
    st.session_state.total_prison_inmates = 0
if 'inmate_as_of_date' not in st.session_state:
    st.session_state.inmate_as_of_date = date.today()

# ------------------------------------------------
# 2. ระบบสร้าง PDF
# ------------------------------------------------
def generate_pdf(total_all=0, as_of_date=None):
    load_fonts()
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("Sarabun", style="", fname="Sarabun-Regular.ttf")
    pdf.add_font("Sarabun", style="B", fname="Sarabun-Bold.ttf")

    pdf.set_font("Sarabun", "B", 16)
    pdf.cell(0, 10, "รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์", new_x="LMARGIN", new_y="NEXT", align="C")

    now = datetime.now()
    date_str = f"ข้อมูล ณ วันที่ {format_thai_date(now)} (เฉพาะผู้ต้องขังกลุ่มเปราะบางที่อยู่: {df.shape[0]} คน)"
    pdf.set_font("Sarabun", "", 11)
    pdf.cell(0, 6, date_str, new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    def draw_section(title, rows):
        pdf.set_font("Sarabun", "B", 13)
        pdf.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Sarabun", "B", 10)
        widths = [82, 22, 30, 30, 26]
        headers = ["รายการ", "รวม (คน)", "ชาย (คน / %)", "หญิง (คน / %)", "สัดส่วน"]
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1, align="C")
        pdf.ln()
        pdf.set_font("Sarabun", "", 9.5)
        for r in rows:
            for i, (val, w) in enumerate(zip(r, widths)):
                align = "L" if i == 0 else "C"
                text = f" {val}" if i == 0 else str(val)
                pdf.cell(w, 7, text, border=1, align=align)
            pdf.ln()
        pdf.ln(4)

    def make_row(label, sub_df, parent_total):
        t, m, f, pm, pf = get_gender_summary(sub_df)
        pct_parent = (t / parent_total * 100) if parent_total > 0 else 0
        return [label, f"{t:,}", f"{m:,} ({pm:.1f}%)", f"{f:,} ({pf:.1f}%)", f"{pct_parent:.1f}%"]

    t_e = df_elder_all.shape[0]
    draw_section("1. สถิติผู้สูงอายุ", [
        make_row("ผู้สูงอายุทั้งหมด", df_elder_all, t_e),
        make_row("- สูงอายุได้เบี้ย", df_elder_get, t_e),
        make_row("- สูงอายุไม่ได้เบี้ย", df_elder_not, t_e),
    ])

    t_d = df_dis_all.shape[0]
    draw_section("2. สถิติผู้พิการ", [
        make_row("ผู้พิการทั้งหมด", df_dis_all, t_d),
        make_row("- พิการได้เบี้ย", df_dis_get, t_d),
        make_row("- พิการไม่ได้เบี้ย", df_dis_not, t_d),
    ])

    pdf.set_font("Sarabun", "B", 13)
    pdf.cell(0, 8, f"3. สถิติผู้ต้องขังหญิงตั้งครรภ์: {df_pregnant.shape[0]} คน", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    t_a = df_aids_all.shape[0]
    draw_section("4. สถิติผู้ป่วย AIDS", [
        make_row("ผู้ป่วย AIDS ทั้งหมด", df_aids_all, t_a),
        make_row("- AIDS ขึ้นทะเบียนแล้ว", df_aids_reg, t_a),
        make_row("- AIDS ประสงค์", df_aids_wish, t_a),
        make_row("- AIDS ไม่รับเบี้ย", df_aids_no, t_a),
    ])

    # ส่วนที่ 5: รวมข้อมูลผู้มีบัตรประชาชนแยกชาย-หญิงลงในตาราง PDF
    id_stats = calc_id_card_stats(total_all)
    id_rows = []
    if id_stats['total_all'] > 0:
        as_of_text = format_thai_date(as_of_date) if as_of_date else format_thai_date(now)
        id_rows.append([
            f"ผู้ต้องขังทั้งหมด (ณ {as_of_text})",
            f"{id_stats['total_all']:,}",
            "-",
            "-",
            "100.00%"
        ])
        id_rows.append([
            "ผู้ต้องขังที่มีบัตรประชาชน",
            f"{id_stats['has_id_total']:,}",
            f"{id_stats['has_id_m']:,} ({id_stats['pct_has_id_m']:.1f}%)",
            f"{id_stats['has_id_f']:,} ({id_stats['pct_has_id_f']:.1f}%)",
            f"{id_stats['pct_has_id']:.2f}%"
        ])
        parent_base = id_stats['total_all']
    else:
        parent_base = df_no_id.shape[0] + df_central_reg.shape[0]

    id_rows.append(make_row("- บัตรหายและไม่ประสงค์ทำบัตรใหม่", df_no_id, parent_base))
    id_rows.append(make_row("- อยู่ในทะเบียนกลาง", df_central_reg, parent_base))

    draw_section("5. สถิติบัตรประชาชน", id_rows)

    return pdf.output()

# ------------------------------------------------
# 3. ระบบสร้าง Excel (แยกชีต + รันลำดับอัตโนมัติ + ใส่รหัสผ่าน)
# ------------------------------------------------
def generate_excel(password=EXCEL_PASSWORD, total_all=0, as_of_date=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary Dashboard"

    title_font = Font(name='Tahoma', size=13, bold=True, color='2C3E50')
    header_font = Font(name='Tahoma', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    data_font = Font(name='Tahoma', size=11, color='333333')
    thin_border = Border(
        left=Side(style='thin', color='BDC3C7'), right=Side(style='thin', color='BDC3C7'),
        top=Side(style='thin', color='BDC3C7'), bottom=Side(style='thin', color='BDC3C7')
    )

    now = datetime.now()
    ws.cell(row=2, column=2, value="รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์").font = title_font
    ws.cell(row=3, column=2, value=f"ข้อมูล ณ วันที่ {format_thai_date(now)} (เฉพาะผู้ต้องขังที่อยู่)").font = Font(name='Tahoma', size=11, italic=True)

    def make_excel_rows(items, parent_total):
        rows = []
        for label, sub_df in items:
            t, m, f, pm, pf = get_gender_summary(sub_df)
            pct_p = (t / parent_total * 100) if parent_total > 0 else 0
            rows.append({
                'รายการ': label, 'รวม (คน)': t,
                'ชาย (คน)': m, 'ชาย (%)': f"{pm:.1f}%",
                'หญิง (คน)': f, 'หญิง (%)': f"{pf:.1f}%",
                'สัดส่วนในกลุ่ม': f"{pct_p:.1f}%"
            })
        return pd.DataFrame(rows)

    current_row = [5]
    def add_table(df_summary, title):
        ws.cell(row=current_row[0], column=2, value=title).font = title_font
        current_row[0] += 1
        for col_idx, col_name in enumerate(df_summary.columns, 2):
            cell = ws.cell(row=current_row[0], column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        current_row[0] += 1
        for row in dataframe_to_rows(df_summary, index=False, header=False):
            for c_idx, val in enumerate(row, 2):
                cell = ws.cell(row=current_row[0], column=c_idx, value=val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            current_row[0] += 1
        current_row[0] += 2

    add_table(make_excel_rows([
        ('ผู้สูงอายุทั้งหมด', df_elder_all), ('สูงอายุได้เบี้ย', df_elder_get), ('สูงอายุไม่ได้เบี้ย', df_elder_not)
    ], df_elder_all.shape[0]), "1. สถิติผู้สูงอายุ")

    add_table(make_excel_rows([
        ('ผู้พิการทั้งหมด', df_dis_all), ('พิการได้เบี้ย', df_dis_get), ('พิการไม่ได้เบี้ย', df_dis_not)
    ], df_dis_all.shape[0]), "2. สถิติผู้พิการ")

    add_table(pd.DataFrame([{'รายการ': 'หญิงตั้งครรภ์', 'รวม (คน)': df_pregnant.shape[0]}]), "3. สถิติหญิงตั้งครรภ์")

    add_table(make_excel_rows([
        ('ผู้ป่วย AIDS ทั้งหมด', df_aids_all),
        ('AIDS ขึ้นทะเบียนแล้ว', df_aids_reg),
        ('AIDS ประสงค์', df_aids_wish),
        ('AIDS ไม่รับเบี้ย', df_aids_no)
    ], df_aids_all.shape[0]), "4. สถิติผู้ป่วย AIDS")

    t_id_total = df_no_id.shape[0] + df_central_reg.shape[0]
    add_table(make_excel_rows([
        ('บัตรหายและไม่ประสงค์ทำบัตรใหม่', df_no_id),
        ('อยู่ในทะเบียนกลาง', df_central_reg)
    ], t_id_total), "5. สถิติบัตรประชาชน")

    id_stats = calc_id_card_stats(total_all)
    if id_stats['total_all'] > 0:
        as_of_text = format_thai_date(as_of_date) if as_of_date else format_thai_date(now)
        add_table(pd.DataFrame([
            {
                'รายการ': f'ผู้ต้องขังทั้งหมด ณ วันที่ {as_of_text}',
                'รวม (คน)': id_stats['total_all'],
                'ชาย (คน)': '-', 'ชาย (%)': '-',
                'หญิง (คน)': '-', 'หญิง (%)': '-',
                'คิดเป็น (%)': '100.00%'
            },
            {
                'รายการ': 'ผู้ต้องขังที่มีบัตรประชาชน',
                'รวม (คน)': id_stats['has_id_total'],
                'ชาย (คน)': id_stats['has_id_m'], 'ชาย (%)': f"{id_stats['pct_has_id_m']:.1f}%",
                'หญิง (คน)': id_stats['has_id_f'], 'หญิง (%)': f"{id_stats['pct_has_id_f']:.1f}%",
                'คิดเป็น (%)': f"{id_stats['pct_has_id']:.2f}%"
            },
            {
                'รายการ': 'บัตรหายและไม่ประสงค์เสียค่าธรรมเนียมทำใหม่ / ทะเบียนกลาง',
                'รวม (คน)': id_stats['no_id_total'],
                'ชาย (คน)': id_stats['no_id_m'], 'ชาย (%)': f"{id_stats['pct_no_id_m']:.1f}%",
                'หญิง (คน)': id_stats['no_id_f'], 'หญิง (%)': f"{id_stats['pct_no_id_f']:.1f}%",
                'คิดเป็น (%)': f"{id_stats['pct_no_id_all']:.2f}%"
            },
        ]), "สรุปสัดส่วนผู้ต้องขังที่มีบัตรประชาชน (แยกเพศ)")

    for col_letter, width in [('B', 50), ('C', 14), ('D', 14), ('E', 14), ('F', 14), ('G', 14), ('H', 16)]:
        ws.column_dimensions[col_letter].width = width
    ws.sheet_view.showGridLines = False

    def add_raw_sheet(sheet_name, target_df):
        ws_sub = wb.create_sheet(title=sheet_name)
        clean_df = target_df.drop(columns=['กลุ่มย่อย', 'ลำดับ'], errors='ignore').reset_index(drop=True)
        
        if clean_df.empty:
            ws_sub.cell(row=1, column=1, value="ไม่มีข้อมูลในกลุ่มนี้").font = data_font
            return

        # รันลำดับ 1, 2, 3... อัตโนมัติในแต่ละชีต
        clean_df.insert(0, 'ลำดับ', range(1, len(clean_df) + 1))

        for r_idx, r in enumerate(dataframe_to_rows(clean_df, index=False, header=True), 1):
            ws_sub.append(r)
            for c_idx in range(1, len(r) + 1):
                cell = ws_sub.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                if r_idx == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.font = data_font
                    if c_idx == 1:
                        cell.alignment = Alignment(horizontal='center', vertical='center')

        for col in ws_sub.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws_sub.column_dimensions[col[0].column_letter].width = min(max_len + 5, 40)

    add_raw_sheet("1. ผู้สูงอายุ", df_elder_all)
    add_raw_sheet("2. ผู้พิการ", df_dis_all)
    add_raw_sheet("3. หญิงตั้งครรภ์", df_pregnant)
    add_raw_sheet("4. ผู้ป่วย AIDS", df_aids_all)
    add_raw_sheet("5.1 บัตรหาย-ไม่ประสงค์ทำใหม่", df_no_id)
    add_raw_sheet("5.2 อยู่ในทะเบียนกลาง", df_central_reg)
    add_raw_sheet("ข้อมูลดิบทั้งหมด", df)

    # บันทึกไฟล์ Excel ลงหน่วยความจำ แล้วเข้ารหัสด้วยรหัสผ่าน
    unencrypted_output = io.BytesIO()
    wb.save(unencrypted_output)
    unencrypted_output.seek(0)

    encrypted_output = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(unencrypted_output)
    office_file.encrypt(password, encrypted_output)
    return encrypted_output.getvalue()

# ------------------------------------------------
# 4. UI: ส่วนแสดงผล Streamlit (Layout 2 คอลัมน์กระชับ)
# ------------------------------------------------
col_title, col_btn1, col_btn2 = st.columns([5, 2, 2])
with col_title:
    st.title("📊 Dashboard กลุ่มเปราะบาง")
    st.caption(f"นับเฉพาะผู้ต้องขังที่สถานะ 'อยู่' จำนวนทั้งหมด {df.shape[0]} ราย")
with col_btn1:
    st.write("")
    st.download_button(
        label="📥 โหลดรายงาน (PDF)",
        data=bytes(generate_pdf(
            st.session_state.total_prison_inmates,
            st.session_state.inmate_as_of_date
        )),
        file_name="สรุปสถิติกลุ่มเปราะบาง.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True
    )
with col_btn2:
    st.write("")
    st.download_button(
        label="🔒 โหลดตาราง (Excel)",
        data=generate_excel(
            EXCEL_PASSWORD,
            st.session_state.total_prison_inmates,
            st.session_state.inmate_as_of_date
        ),
        file_name="สรุปสถิติกลุ่มเปราะบาง.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="secondary",
        use_container_width=True
    )
st.divider()

def render_group_section(title, df_all, df_get, df_not, label_get, label_not):
    st.subheader(title)
    if df_all.empty:
        st.info("ไม่มีข้อมูลในกลุ่มนี้")
        return

    t_all, m_all, f_all, pm_all, pf_all = get_gender_summary(df_all)
    t_get, m_get, f_get, pm_get, pf_get = get_gender_summary(df_get)
    t_not, m_not, f_not, pm_not, pf_not = get_gender_summary(df_not)

    pct_get = (t_get / t_all * 100) if t_all > 0 else 0
    pct_not = (t_not / t_all * 100) if t_all > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("รวมทั้งหมด", f"{t_all} คน", f"ชาย {m_all} ({pm_all:.1f}%) | หญิง {f_all} ({pf_all:.1f}%)", delta_color="off")
    c2.metric(label_get, f"{t_get} คน", f"คิดเป็น {pct_get:.1f}%", delta_color="off")
    c3.metric(label_not, f"{t_not} คน", f"คิดเป็น {pct_not:.1f}%", delta_color="inverse")

    sun_df = df_all.groupby(['กลุ่มย่อย', 'เพศ']).size().reset_index(name='จำนวน')
    fig = px.sunburst(
        sun_df,
        path=['กลุ่มย่อย', 'เพศ'],
        values='จำนวน',
        color='กลุ่มย่อย',
        color_discrete_map={label_get: '#2ecc71', label_not: '#e74c3c'},
        title="สัดส่วนการรับเบี้ยและเพศ (คลิกที่กราฟเพื่อเจาะลึก)"
    )
    fig.update_traces(textinfo="label+value+percent parent")
    fig.update_layout(margin=dict(t=35, b=10, l=10, r=10), height=320)
    st.plotly_chart(fig, use_container_width=True)

def render_aids_card(title, sub_df, total_aids, accent_color):
    t, m, f, pm, pf = get_gender_summary(sub_df)
    pct_total = (t / total_aids * 100) if total_aids > 0 else 0
    return f"""
    <div style="background-color: #f8f9fa; padding: 12px; border-radius: 8px; border-top: 4px solid {accent_color}; border-left: 1px solid #e0e0e0; border-right: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0; text-align: center; margin-bottom: 10px;">
        <div style="font-size: 14px; font-weight: bold; color: #2c3e50;">{title}</div>
        <div style="font-size: 22px; font-weight: bold; color: {accent_color}; margin: 4px 0;">
            {t} <span style="font-size: 13px; font-weight: normal; color: #7f8c8d;">คน ({pct_total:.1f}%)</span>
        </div>
        <div style="font-size: 12px; color: #34495e; line-height: 1.5;">
            <span style="color: #3498db; font-weight: 600;">ชาย {m} ({pm:.1f}%)</span><br>
            <span style="color: #ff66b3; font-weight: 600;">หญิง {f} ({pf:.1f}%)</span>
        </div>
    </div>
    """

def render_kpi_box(label, value_text, sub_text, accent_color, gender_html=""):
    gender_block = f'<div style="font-size: 12px; margin-top: 6px; padding-top: 6px; border-top: 1px dashed #dee2e6;">{gender_html}</div>' if gender_html else ""
    return f"""
    <div style="background-color: #f8f9fa; padding: 14px 12px; border-radius: 10px; border-left: 5px solid {accent_color}; border-top: 1px solid #e9ecef; border-right: 1px solid #e9ecef; border-bottom: 1px solid #e9ecef; text-align: center;">
        <div style="font-size: 13px; font-weight: 600; color: #555; margin-bottom: 4px;">{label}</div>
        <div style="font-size: 26px; font-weight: 700; color: {accent_color}; line-height: 1.2;">{value_text}</div>
        <div style="font-size: 13px; color: #6c757d; margin-top: 4px;">{sub_text}</div>
        {gender_block}
    </div>
    """

main_col1, main_col2 = st.columns(2, gap="large")

with main_col1:
    render_group_section(
        "👵👴 1. สถิติผู้สูงอายุ",
        df_elder_all, df_elder_get, df_elder_not,
        "สูงอายุได้เบี้ย", "สูงอายุไม่ได้เบี้ย"
    )
    st.divider()

    st.subheader("🤰 3. สถิติผู้ต้องขังหญิงตั้งครรภ์")
    st.metric(label="จำนวนหญิงตั้งครรภ์ทั้งหมด", value=f"{df_pregnant.shape[0]} คน")

with main_col2:
    render_group_section(
        "♿ 2. สถิติผู้พิการ",
        df_dis_all, df_dis_get, df_dis_not,
        "พิการได้เบี้ย", "พิการไม่ได้เบี้ย"
    )
    st.divider()

    # 4. ผู้ป่วย AIDS (3 กลุ่มย่อย พร้อมแยกชาย-หญิงและเปอร์เซ็นต์)
    st.subheader("🎗️ 4. สถิติผู้ป่วย AIDS")
    if not df_aids_all.empty:
        t_all, m_all, f_all, pm_all, pf_all = get_gender_summary(df_aids_all)
        st.metric(
            "ผู้ป่วย AIDS รวมทั้งหมด",
            f"{t_all} คน",
            f"ชาย {m_all} ({pm_all:.1f}%) | หญิง {f_all} ({pf_all:.1f}%)",
            delta_color="off"
        )

        c1, c2, c3 = st.columns(3)
        c1.markdown(render_aids_card("AIDS ขึ้นทะเบียนแล้ว", df_aids_reg, t_all, "#2ecc71"), unsafe_allow_html=True)
        c2.markdown(render_aids_card("AIDS ประสงค์", df_aids_wish, t_all, "#f39c12"), unsafe_allow_html=True)
        c3.markdown(render_aids_card("AIDS ไม่รับเบี้ย", df_aids_no, t_all, "#e74c3c"), unsafe_allow_html=True)

        sun_aids = df_aids_all.groupby(['กลุ่มย่อย', 'เพศ']).size().reset_index(name='จำนวน')
        fig_aids = px.sunburst(
            sun_aids,
            path=['กลุ่มย่อย', 'เพศ'],
            values='จำนวน',
            color='กลุ่มย่อย',
            color_discrete_map={
                'AIDS ขึ้นทะเบียนแล้ว': '#2ecc71',
                'AIDS ประสงค์': '#f39c12',
                'AIDS ไม่รับเบี้ย': '#e74c3c'
            },
            title="สัดส่วนสถานะผู้ป่วย AIDS และเพศ (คลิกที่กราฟเพื่อเจาะลึก)"
        )
        fig_aids.update_traces(textinfo="label+value+percent parent")
        fig_aids.update_layout(margin=dict(t=35, b=10, l=10, r=10), height=320)
        st.plotly_chart(fig_aids, use_container_width=True)
    else:
        st.info("ไม่มีข้อมูลผู้ป่วย AIDS")

st.divider()

# ------------------------------------------------
# 5. สถิติบัตรประชาชน (แบ่งกรอบและสัดส่วนสวยงาม)
# ------------------------------------------------
h_col1, h_col2 = st.columns([3, 1.5])
with h_col1:
    st.subheader("🪪 5. สถิติบัตรประชาชน")
with h_col2:
    with st.popover("📝 ระบุยอดผู้ต้องขังทั้งหมด ณ วันที่", use_container_width=True):
        st.markdown("**กรอกยอดผู้ต้องขังทั้งหมดเพื่อคำนวณ % ผู้มีบัตรประชาชน**")
        input_date = st.date_input(
            "ข้อมูล ณ วันที่:",
            value=st.session_state.inmate_as_of_date
        )
        input_total = st.number_input(
            "จำนวนผู้ต้องขังทั้งหมด (คน):",
            min_value=0,
            step=1,
            value=st.session_state.total_prison_inmates
        )
        if st.button("บันทึกและคำนวณ", type="primary", use_container_width=True):
            st.session_state.inmate_as_of_date = input_date
            st.session_state.total_prison_inmates = input_total
            st.rerun()

# แถบสรุปภาพรวมด้านบน (แสดงเมื่อมีการระบุยอดผู้ต้องขังทั้งหมด)
id_stats = calc_id_card_stats(st.session_state.total_prison_inmates)
if id_stats['total_all'] > 0:
    as_of_str = format_thai_date(st.session_state.inmate_as_of_date)

    gender_has_id_html = (
        f'<span style="color:#3498db;font-weight:600;">ชาย {id_stats["has_id_m"]:,} ({id_stats["pct_has_id_m"]:.1f}%)</span> | '
        f'<span style="color:#ff66b3;font-weight:600;">หญิง {id_stats["has_id_f"]:,} ({id_stats["pct_has_id_f"]:.1f}%)</span>'
    )
    gender_no_id_html = (
        f'<span style="color:#3498db;font-weight:600;">ชาย {id_stats["no_id_m"]:,} ({id_stats["pct_no_id_m"]:.1f}%)</span> | '
        f'<span style="color:#ff66b3;font-weight:600;">หญิง {id_stats["no_id_f"]:,} ({id_stats["pct_no_id_f"]:.1f}%)</span>'
    )

    with st.container(border=True):
        st.caption(f"📅 สรุปภาพรวมบัตรประชาชน ณ วันที่ **{as_of_str}**")
        s1, s2, s3 = st.columns(3, gap="medium")
        s1.markdown(render_kpi_box("ผู้ต้องขังทั้งหมด", f"{id_stats['total_all']:,} คน", "100.00%", "#2c3e50"), unsafe_allow_html=True)
        s2.markdown(render_kpi_box("ผู้ต้องขังที่มีบัตรประชาชน", f"{id_stats['has_id_total']:,} คน", f"คิดเป็น {id_stats['pct_has_id']:.2f}%", "#2ecc71", gender_has_id_html), unsafe_allow_html=True)
        s3.markdown(render_kpi_box("บัตรหายและไม่ประสงค์เสียค่าธรรมเนียมทำใหม่ / ทะเบียนกลาง", f"{id_stats['no_id_total']:,} คน", f"คิดเป็น {id_stats['pct_no_id_all']:.2f}%", "#e74c3c", gender_no_id_html), unsafe_allow_html=True)

# แบ่ง 2 คอลัมน์ซ้าย-ขวา พร้อมกรอบการ์ด
id_col1, id_col2 = st.columns(2, gap="medium")

with id_col1:
    with st.container(border=True):
        st.markdown("#### บัตรหายและไม่ประสงค์ทำบัตรใหม่")
        if not df_no_id.empty:
            t_n, m_n, f_n, pm_n, pf_n = get_gender_summary(df_no_id)
            c1, c2, c3 = st.columns(3)
            c1.metric("รวมทั้งหมด", f"{t_n} คน")
            c2.metric("เพศชาย", f"{m_n} คน", f"{pm_n:.1f}%", delta_color="off")
            c3.metric("เพศหญิง", f"{f_n} คน", f"{pf_n:.1f}%", delta_color="off")

            no_id_gender = df_no_id['เพศ'].value_counts().reset_index()
            no_id_gender.columns = ['เพศ', 'จำนวน']
            fig5_1 = px.pie(
                no_id_gender, names='เพศ', values='จำนวน',
                hole=0.45, color='เพศ', color_discrete_map=pie_gender_color
            )
            fig5_1.update_traces(textinfo='label+value+percent')
            fig5_1.update_layout(
                margin=dict(t=15, b=10, l=10, r=10),
                height=260,
                legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig5_1, use_container_width=True)
        else:
            st.info("ไม่มีข้อมูลบัตรหายและไม่ประสงค์ทำบัตรใหม่")

with id_col2:
    with st.container(border=True):
        st.markdown("#### อยู่ในทะเบียนกลาง")
        if not df_central_reg.empty:
            t_c, m_c, f_c, pm_c, pf_c = get_gender_summary(df_central_reg)
            st.markdown(
                f"""
                <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px; border: 1px solid #e9ecef; text-align: center; margin-bottom: 16px;">
                    <div style="font-size: 14px; color: #6c757d; font-weight: 600;">จำนวนผู้ต้องขังในทะเบียนกลางทั้งหมด</div>
                    <div style="font-size: 38px; font-weight: 700; color: #2c3e50; margin-top: 4px;">{t_c:,} <span style="font-size: 18px; font-weight: normal; color: #7f8c8d;">คน</span></div>
                </div>
                """,
                unsafe_allow_html=True
            )
            g1, g2 = st.columns(2, gap="medium")
            g1.markdown(render_kpi_box("เพศชาย", f"{m_c:,} คน", f"คิดเป็น {pm_c:.1f}%", "#3498db"), unsafe_allow_html=True)
            g2.markdown(render_kpi_box("เพศหญิง", f"{f_c:,} คน", f"คิดเป็น {pf_c:.1f}%", "#ff66b3"), unsafe_allow_html=True)

            st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
            st.markdown(
                f"""
                <div style="margin-top: 8px;">
                    <div style="display: flex; justify-content: space-between; font-size: 13px; color: #555; margin-bottom: 6px;">
                        <span>ชาย {pm_c:.1f}%</span>
                        <span>หญิง {pf_c:.1f}%</span>
                    </div>
                    <div style="display: flex; height: 14px; border-radius: 7px; overflow: hidden; background-color: #e9ecef;">
                        <div style="width: {pm_c:.1f}%; background-color: #3498db;"></div>
                        <div style="width: {pf_c:.1f}%; background-color: #ff66b3;"></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info("ไม่มีข้อมูลผู้ต้องขังที่อยู่ในทะเบียนกลาง")
