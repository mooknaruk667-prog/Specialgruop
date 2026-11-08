import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.request
import os
import io
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

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
    df = load_data()
    df['อายุ (ปี)'] = pd.to_numeric(df['อายุ (ปี)'], errors='coerce')
    total_inmates = df.shape[0] # จำนวนผู้ต้องขังทั้งหมด
except Exception as e:
    st.error(f"ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบการแชร์ไฟล์: {e}")
    st.stop()

# --- ธีมสีแยกเพศ (แยกสำหรับ Pie/Sunburst และ Bar) ---
pie_gender_color = {'ชาย': '#3498db', 'หญิง': '#ff66b3'} 
bar_gender_color = {'ชาย': '#2A9D8F', 'หญิง': '#E76F51'} 

# 📌 ฟังก์ชันดักจับค่า "ไม่มี" และ "ช่องว่าง" ขั้นเด็ดขาด
def check_no_rights(x):
    if pd.isna(x):
        return 'ยังไม่มีสิทธิ์'
    val = str(x).strip().lower()
    if val in ['ไม่มี', '', 'nan', 'none', 'null', '-', '<na>']:
        return 'ยังไม่มีสิทธิ์'
    return 'มีสิทธิ์แล้ว'

# ------------------------------------------------
# 2. ระบบสร้าง PDF (แบบหน้าเดียว 2 คอลัมน์)
# ------------------------------------------------
def generate_pdf(df):
    load_fonts()
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("Sarabun", style="", fname="Sarabun-Regular.ttf")
    pdf.add_font("Sarabun", style="B", fname="Sarabun-Bold.ttf")
    
    pdf.set_font("Sarabun", "B", 18)
    pdf.cell(0, 10, "รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์", new_x="LMARGIN", new_y="NEXT", align="C")
    
    months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
    now = datetime.now()
    date_str = f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year + 543}"
    
    pdf.set_font("Sarabun", "", 12)
    pdf.cell(0, 6, date_str, new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(4)
    
    def draw_text(x, y, text, style="", size=11):
        pdf.set_xy(x, y)
        pdf.set_font("Sarabun", style, size)
        pdf.cell(90, 6, text, ln=0)
        return y + 6
        
    def draw_table(x, y, headers, data, widths):
        pdf.set_xy(x, y)
        pdf.set_font("Sarabun", "B", 10)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1, align="C")
        y += 7
        pdf.set_font("Sarabun", "", 10)
        for row in data:
            pdf.set_xy(x, y)
            for i, (val, w) in enumerate(zip(row, widths)):
                text_val = str(val)
                if i == 0:
                    text_val = text_val.replace("ทางการได้ยินหรือสื่อความหมาย", "ทางการได้ยินฯ")
                    text_val = text_val.replace("ทางการเคลื่อนไหวหรือทางร่างกาย", "ทางการเคลื่อนไหวฯ")
                    text_val = text_val.replace("ทางจิตใจหรือพฤติกรรม", "ทางจิตใจฯ")
                    pdf.cell(w, 7, " " + text_val[:35], border=1, align="L")
                else:
                    pdf.cell(w, 7, text_val[:30], border=1, align="C")
            y += 7
        return y + 4

    col1_x = 10     
    col2_x = 110    
    y_start = pdf.get_y()
    
    # ================= คอลัมน์ 1 (ซ้าย) =================
    y = y_start
    y = draw_text(col1_x, y, "1. สถิติเด็กติดผู้ต้องขัง", "B", 13)
    df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')]
    if not df_child.empty:
        child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน')
        y = draw_table(col1_x, y, ["ช่วงอายุ", "เพศเด็ก", "จำนวน(คน)"], child_stats.values.tolist(), [36, 27, 27])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    y = draw_text(col1_x, y, "2. สถิติผู้ต้องขังหญิงตั้งครรภ์", "B", 13)
    pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'].isin(['ใช่', 'มี'])].shape[0]
    y = draw_text(col1_x, y, f"จำนวนหญิงตั้งครรภ์ทั้งหมด: {pregnant_count} คน", "", 11) + 2
    
    y = draw_text(col1_x, y, "3. สถิติผู้ต้องขังสูงอายุ (60 ปีขึ้นไป)", "B", 13)
    df_elderly = df[df['อายุ (ปี)'] >= 60].copy()
    if not df_elderly.empty:
        has_allowance = df_elderly['เบี้ยสูงอายุ'].isin(['ได้รับ', 'มี', 'ใช่'])
        no_allowance = (~has_allowance).sum()
        y = draw_text(col1_x, y, f"รวม: {df_elderly.shape[0]} คน | ยังไม่ได้เบี้ย: {no_allowance} คน", "", 11)
        bins = [59, 69, 79, float('inf')]
        labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
        df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
        elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
        elder_stats = elder_stats[elder_stats['จำนวน'] > 0]
        y = draw_table(col1_x, y, ["ช่วงอายุ", "เพศ", "จำนวน(คน)"], elder_stats.values.tolist(), [36, 27, 27])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    y = draw_text(col1_x, y, "4. สถิติผู้พิการ", "B", 13)
    df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')]
    if not df_disabled.empty:
        total_disabled = df_disabled.shape[0]
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].isin(['มี', 'ใช่'])].shape[0]
        unreg_count = total_disabled - reg_count
        y = draw_text(col1_x, y, f"รวม: {total_disabled} คน | ขึ้นทะเบียน: {reg_count} | ยังไม่ขึ้น: {unreg_count}", "", 11)
        dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col1_x, y, ["ประเภทความพิการ", "เพศ", "จำนวน(คน)"], dis_type.values.tolist(), [54, 18, 18])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    # ================= คอลัมน์ 2 (ขวา) =================
    y = y_start
    y = draw_text(col2_x, y, "5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม", "B", 13)
    df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])]
    if not df_novisit.empty:
        novisit_stats = df_novisit.groupby(['เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col2_x, y, ["เพศ", "จำนวน(คน)"], novisit_stats.values.tolist(), [45, 45])
    else:
        y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    y = draw_text(col2_x, y, "6. สถิติผู้ป่วย AIDS", "B", 13)
    df_aids = df[df['AIDS'].isin(['มี', 'ใช่'])].copy()
    if not df_aids.empty:
        total_aids = df_aids.shape[0]
        df_aids['สถานะสิทธิ์'] = df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].apply(check_no_rights)
        no_rights_count = df_aids[df_aids['สถานะสิทธิ์'] == 'ยังไม่มีสิทธิ์'].shape[0]
        has_rights_count = total_aids - no_rights_count
        y = draw_text(col2_x, y, f"รวม: {total_aids} คน | มีสิทธิ์: {has_rights_count} | ยังไม่มี: {no_rights_count}", "", 11)
        aids_stats = df_aids.groupby(['เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col2_x, y, ["เพศ", "จำนวน(คน)"], aids_stats.values.tolist(), [45, 45])
    else:
        y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    y = draw_text(col2_x, y, "7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง", "B", 13)
    df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])]
    df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])]
    
    total_valid_id = df_has_id.shape[0] + df_no_id.shape[0]
    pct_has_id = (df_has_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0
    pct_no_id = (df_no_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0
    
    y = draw_text(col2_x, y, f"ผู้ที่มีบัตรประชาชนรวม: {df_has_id.shape[0]} คน ({pct_has_id:.1f}%)", "", 11)
    y = draw_text(col2_x, y, f"ผู้ที่ไม่มีบัตรประชาชนรวม: {df_no_id.shape[0]} คน ({pct_no_id:.1f}%)", "", 11)
    
    if not df_no_id.empty:
        y += 2
        y = draw_text(col2_x, y, "กลุ่มเปราะบางที่ไม่มีบัตรประชาชน (แยกเพศ):", "B", 11)
        
        disabled_m = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        disabled_f = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        elderly_m = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        elderly_f = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        novisit_m = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        novisit_f = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        aids_m = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        aids_f = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        data = [
            ["ผู้พิการ", disabled_m, disabled_f, disabled_m + disabled_f],
            ["ผู้สูงอายุ", elderly_m, elderly_f, elderly_m + elderly_f],
            ["ไม่มีญาติ", novisit_m, novisit_f, novisit_m + novisit_f],
            ["ผู้ป่วย AIDS", aids_m, aids_f, aids_m + aids_f]
        ]
        data = [r for r in data if r[3] > 0]
        
        if data:
            y = draw_table(col2_x, y, ["กลุ่ม", "ชาย", "หญิง", "รวม"], data, [42, 16, 16, 16])
        else:
            y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 11) + 2
        
    return pdf.output()

# ------------------------------------------------
# 3. ระบบสร้าง Excel (OpenPyXL)
# ------------------------------------------------
def generate_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary Dashboard"

    title_font = Font(name='Tahoma', size=14, bold=True, color='2C3E50')
    header_font = Font(name='Tahoma', size=12, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    data_font = Font(name='Tahoma', size=11, color='333333')
    thin_border = Border(
        left=Side(style='thin', color='BDC3C7'), right=Side(style='thin', color='BDC3C7'), 
        top=Side(style='thin', color='BDC3C7'), bottom=Side(style='thin', color='BDC3C7')
    )

    months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
    now = datetime.now()
    
    ws.cell(row=2, column=2, value="รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์").font = title_font
    ws.cell(row=3, column=2, value=f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year + 543}").font = Font(name='Tahoma', size=11, italic=True)

    current_row = [5]
    def add_table(df_summary, title):
        cell = ws.cell(row=current_row[0], column=2, value=title)
        cell.font = title_font
        current_row[0] += 1
        
        col_idx = 2
        for col_name in df_summary.columns:
            cell = ws.cell(row=current_row[0], column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            col_idx += 1
        current_row[0] += 1
        
        for r_idx, row in enumerate(dataframe_to_rows(df_summary, index=False, header=False)):
            for c_idx, val in enumerate(row):
                cell = ws.cell(row=current_row[0], column=c_idx+2, value=val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            current_row[0] += 1
        current_row[0] += 2

    # 1 - 3
    df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')]
    if not df_child.empty:
        child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน (คน)')
        child_stats.columns = ['ช่วงอายุ', 'เพศเด็ก', 'จำนวน (คน)']
        add_table(child_stats, "👶 1. สถิติเด็กติดผู้ต้องขัง")

    pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'].isin(['ใช่', 'มี'])].shape[0]
    add_table(pd.DataFrame([{'กลุ่มเปราะบาง': 'หญิงตั้งครรภ์', 'จำนวน (คน)': pregnant_count}]), "🤰 2. สถิติผู้ต้องขังหญิงตั้งครรภ์")

    df_elderly = df[df['อายุ (ปี)'] >= 60].copy()
    if not df_elderly.empty:
        has_allowance = df_elderly['เบี้ยสูงอายุ'].isin(['ได้รับ', 'มี', 'ใช่'])
        no_allowance_count = (~has_allowance).sum()
        bins = [59, 69, 79, float('inf')]
        labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
        df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
        elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน (คน)')
        elder_stats = elder_stats[elder_stats['จำนวน (คน)'] > 0]
        add_table(elder_stats, f"👵👴 3. สถิติผู้ต้องขังสูงอายุ (รวม: {df_elderly.shape[0]} คน | ยังไม่ได้รับเบี้ย: {no_allowance_count} คน)")

    # 4. พิการ
    df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')]
    if not df_disabled.empty:
        total_disabled = df_disabled.shape[0]
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].isin(['มี', 'ใช่'])].shape[0]
        unreg_count = total_disabled - reg_count
        dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน (คน)')
        dis_type.columns = ['ประเภทความพิการ', 'เพศ', 'จำนวน (คน)']
        add_table(dis_type, f"♿ 4. สถิติผู้พิการ (รวม: {total_disabled} คน | ขึ้นทะเบียนแล้ว: {reg_count} คน | ยังไม่ขึ้นทะเบียน: {unreg_count} คน)")

    # 5. ไม่มีญาติ
    df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])]
    if not df_novisit.empty:
        novisit_stats = df_novisit.groupby(['เพศ']).size().reset_index(name='จำนวน (คน)')
        add_table(novisit_stats, "🚫 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม")

    # 6. AIDS 
    df_aids = df[df['AIDS'].isin(['มี', 'ใช่'])].copy()
    if not df_aids.empty:
        total_aids = df_aids.shape[0]
        df_aids['สถานะสิทธิ์'] = df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].apply(check_no_rights)
        no_rights_count = df_aids[df_aids['สถานะสิทธิ์'] == 'ยังไม่มีสิทธิ์'].shape[0]
        has_rights_count = total_aids - no_rights_count
        aids_stats = df_aids.groupby(['เพศ']).size().reset_index(name='จำนวน (คน)')
        add_table(aids_stats, f"🎗️ 6. สถิติผู้ป่วย AIDS (รวม: {total_aids} คน | มีสิทธิ์แล้ว: {has_rights_count} คน | ยังไม่มีสิทธิ์: {no_rights_count} คน)")

    # 7. บัตรประชาชน 
    df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])]
    df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])]
    
    total_valid_id = df_has_id.shape[0] + df_no_id.shape[0]
    pct_has_id = (df_has_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0
    pct_no_id = (df_no_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0
    
    add_table(pd.DataFrame([
        {'สถานะ': 'มีบัตรประจำตัวประชาชน', 'จำนวน (คน)': df_has_id.shape[0], 'เปอร์เซ็นต์': f"{pct_has_id:.1f}%"},
        {'สถานะ': 'ไม่มีบัตรประจำตัวประชาชน', 'จำนวน (คน)': df_no_id.shape[0], 'เปอร์เซ็นต์': f"{pct_no_id:.1f}%"}
    ]), "🪪 7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง")

    if not df_no_id.empty:
        disabled_m = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        disabled_f = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        elderly_m = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        elderly_f = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        novisit_m = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        novisit_f = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        aids_m = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        aids_f = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        
        df_vul_excel = pd.DataFrame([
            {'กลุ่มเปราะบางที่ไม่มีบัตร': 'ผู้พิการ', 'ชาย (คน)': disabled_m, 'หญิง (คน)': disabled_f, 'รวม (คน)': disabled_m + disabled_f},
            {'กลุ่มเปราะบางที่ไม่มีบัตร': 'ผู้สูงอายุ', 'ชาย (คน)': elderly_m, 'หญิง (คน)': elderly_f, 'รวม (คน)': elderly_m + elderly_f},
            {'กลุ่มเปราะบางที่ไม่มีบัตร': 'ไม่มีญาติเยี่ยม', 'ชาย (คน)': novisit_m, 'หญิง (คน)': novisit_f, 'รวม (คน)': novisit_m + novisit_f},
            {'กลุ่มเปราะบางที่ไม่มีบัตร': 'ผู้ป่วย AIDS', 'ชาย (คน)': aids_m, 'หญิง (คน)': aids_f, 'รวม (คน)': aids_m + aids_f}
        ])
        df_vul_excel = df_vul_excel[df_vul_excel['รวม (คน)'] > 0]
        add_table(df_vul_excel, "🔍 กลุ่มเปราะบางที่ไม่มีบัตรประชาชน (แยกเพศ)")

    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.sheet_view.showGridLines = False

    ws_raw = wb.create_sheet(title="Raw Data (ข้อมูลดิบ)")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_raw.append(r)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ------------------------------------------------
# UI: ส่วนแสดงผล Streamlit
# ------------------------------------------------
st.title("📊 Dashboard กลุ่มเปราะบาง")

col_dl1, col_dl2, _ = st.columns([1, 1, 2])
with col_dl1:
    st.download_button(
        label="📥 ดาวน์โหลดรายงาน (PDF)",
        data=bytes(generate_pdf(df)),
        file_name="สรุปสถิติกลุ่มเปราะบาง.pdf",
        mime="application/pdf",
        type="primary"
    )
with col_dl2:
    st.download_button(
        label="📥 ดาวน์โหลดตาราง (Excel)",
        data=generate_excel(df),
        file_name="สรุปสถิติกลุ่มเปราะบาง.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="secondary"
    )
st.markdown("---")

# ----- กราฟ 1. เด็กติดผู้ต้องขัง -----
st.subheader("👶 1. สถิติเด็กติดผู้ต้องขัง")
df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')].copy()
if not df_child.empty:
    child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน')
    child_stats['เปอร์เซ็นต์'] = (child_stats['จำนวน'] / child_stats['จำนวน'].sum() * 100).round(1)
    fig1 = px.bar(child_stats, x='[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', y='จำนวน', color='[1] เด็กติดผู้ต้องขัง (เพศเด็ก)', 
                  text='เปอร์เซ็นต์', barmode='group', color_discrete_map=bar_gender_color, title="แยกตามช่วงอายุและเพศเด็ก")
    fig1.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลเด็กติดผู้ต้องขัง")
st.markdown("---")

# ----- กราฟ 2. หญิงตั้งครรภ์ -----
st.subheader("🤰 2. สถิติผู้ต้องขังหญิงตั้งครรภ์")
pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'].isin(['ใช่', 'มี'])].shape[0]
st.metric(label="จำนวนหญิงตั้งครรภ์ทั้งหมด", value=f"{pregnant_count} คน")
st.markdown("---")

# ----- กราฟ 3. ผู้สูงอายุ -----
st.subheader("👵👴 3. สถิติผู้ต้องขังสูงอายุ")
df_elderly = df[df['อายุ (ปี)'] >= 60].copy()
if not df_elderly.empty:
    total_elderly = df_elderly.shape[0]
    has_allowance = df_elderly['เบี้ยสูงอายุ'].isin(['ได้รับ', 'มี', 'ใช่'])
    no_allowance_count = (~has_allowance).sum()
    no_allowance_percent = (no_allowance_count / total_elderly) * 100 if total_elderly > 0 else 0
    
    col1, col2 = st.columns(2)
    col1.metric(label="ผู้สูงอายุรวมทั้งหมด", value=f"{total_elderly} คน")
    col2.metric(label="ยังไม่ได้รับเบี้ยผู้สูงอายุ", value=f"{no_allowance_count} คน", delta=f"คิดเป็น {no_allowance_percent:.1f}%", delta_color="off")
    
    bins = [59, 69, 79, float('inf')]
    labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
    df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
    elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
    elder_stats['เปอร์เซ็นต์'] = (elder_stats['จำนวน'] / elder_stats['จำนวน'].sum() * 100).round(1)
    
    fig3 = px.bar(elder_stats, x='ช่วงอายุ', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', barmode='group',
                  color_discrete_map=bar_gender_color, title="แยกตามช่วงอายุและเพศ")
    fig3.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้สูงอายุ")
st.markdown("---")

# ----- กราฟ 4. ผู้พิการ -----
st.subheader("♿ 4. สถิติผู้พิการ")
df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')].copy()
if not df_disabled.empty:
    total_disabled = df_disabled.shape[0]
    reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].isin(['มี', 'ใช่'])].shape[0]
    unreg_count = total_disabled - reg_count
    
    reg_percent = (reg_count / total_disabled) * 100 if total_disabled > 0 else 0
    unreg_percent = (unreg_count / total_disabled) * 100 if total_disabled > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    col1.metric(label="ผู้พิการรวมทั้งหมด", value=f"{total_disabled} คน")
    col2.metric(label="ขึ้นทะเบียนสิทธิ์แล้ว", value=f"{reg_count} คน", delta=f"คิดเป็น {reg_percent:.1f}%", delta_color="off")
    col3.metric(label="ยังไม่ขึ้นทะเบียนสิทธิ์", value=f"{unreg_count} คน", delta=f"คิดเป็น {unreg_percent:.1f}%", delta_color="inverse")
    
    col_pie1, col_bar2 = st.columns(2)
    
    df_disabled['สถานะการขึ้นทะเบียน'] = df_disabled['สถานะบัตรคนพิการ'].apply(lambda x: 'ขึ้นทะเบียนแล้ว' if pd.notna(x) and str(x).strip() in ['มี', 'ใช่'] else 'ยังไม่ขึ้นทะเบียน')
    
    sunburst_data4 = df_disabled.groupby(['เพศ', 'สถานะการขึ้นทะเบียน']).size().reset_index(name='จำนวน')
    fig4_1 = px.sunburst(sunburst_data4, path=['เพศ', 'สถานะการขึ้นทะเบียน'], values='จำนวน', color='เพศ', 
                         color_discrete_map=pie_gender_color, title="สัดส่วนสถานะการขึ้นทะเบียน (👉 คลิกเพื่อเจาะลึก)")
    fig4_1.update_traces(textinfo='label+value+percent parent')
    col_pie1.plotly_chart(fig4_1, use_container_width=True)
    
    # 📌 เปลี่ยนเป็นกราฟแท่งแนวนอน (Horizontal Bar Chart)
    dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน')
    
    # เรียงลำดับเพื่อให้แท่งที่ยาวที่สุดอยู่ด้านบน
    dis_type_sort = dis_type.groupby('[3] ผู้พิการ (ระบุประเภท)')['จำนวน'].sum().reset_index()
    dis_type_sort = dis_type_sort.sort_values(by='จำนวน', ascending=True)
    dis_type['[3] ผู้พิการ (ระบุประเภท)'] = pd.Categorical(dis_type['[3] ผู้พิการ (ระบุประเภท)'], categories=dis_type_sort['[3] ผู้พิการ (ระบุประเภท)'], ordered=True)
    dis_type = dis_type.sort_values('[3] ผู้พิการ (ระบุประเภท)')
    
    fig4_2 = px.bar(dis_type, x='จำนวน', y='[3] ผู้พิการ (ระบุประเภท)', color='เพศ', orientation='h', barmode='group', 
                    color_discrete_map=bar_gender_color, title="แยกตามประเภทความพิการ", text='จำนวน')
    fig4_2.update_traces(textposition='outside')
    fig4_2.update_layout(yaxis_title=None)
    col_bar2.plotly_chart(fig4_2, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้พิการ")
st.markdown("---")

# ----- กราฟ 5. ไม่มีญาติเยี่ยม -----
st.subheader("🚫 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม")
df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])].copy()
if not df_novisit.empty:
    novisit_stats = df_novisit['เพศ'].value_counts().reset_index()
    novisit_stats.columns = ['เพศ', 'จำนวน']
    novisit_stats['เปอร์เซ็นต์'] = (novisit_stats['จำนวน'] / novisit_stats['จำนวน'].sum() * 100).round(1)
    
    fig5 = px.bar(novisit_stats, x='เพศ', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', title="ผู้ต้องขังไม่มีญาติเยี่ยม (แยกเพศ)",
                  color_discrete_map=bar_gender_color)
    fig5.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig5, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้ต้องขังที่ไม่มีญาติเยี่ยม")
st.markdown("---")

# ----- กราฟ 6. ผู้ป่วย AIDS -----
st.subheader("🎗️ 6. สถิติผู้ป่วย AIDS")
df_aids = df[df['AIDS'].isin(['มี', 'ใช่'])].copy()
if not df_aids.empty:
    total_aids = df_aids.shape[0]
    
    df_aids['สถานะสิทธิ์'] = df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].apply(check_no_rights)
    no_rights_count = df_aids[df_aids['สถานะสิทธิ์'] == 'ยังไม่มีสิทธิ์'].shape[0]
    has_rights_count = total_aids - no_rights_count
    
    no_rights_percent = (no_rights_count / total_aids) * 100 if total_aids > 0 else 0
    has_rights_percent = (has_rights_count / total_aids) * 100 if total_aids > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    col1.metric(label="จำนวนผู้ป่วย AIDS รวมทั้งหมด", value=f"{total_aids} คน")
    col2.metric(label="มีสิทธิ์แล้ว", value=f"{has_rights_count} คน", delta=f"คิดเป็น {has_rights_percent:.1f}%", delta_color="off")
    col3.metric(label="ยังไม่มีสิทธิ์", value=f"{no_rights_count} คน", delta=f"คิดเป็น {no_rights_percent:.1f}%", delta_color="inverse")
    
    _, col_pie, _ = st.columns([1, 2, 1])
    
    sunburst_data6 = df_aids.groupby(['เพศ', 'สถานะสิทธิ์']).size().reset_index(name='จำนวน')
    fig6_1 = px.sunburst(sunburst_data6, path=['เพศ', 'สถานะสิทธิ์'], values='จำนวน', color='เพศ',
                         color_discrete_map=pie_gender_color, title="สัดส่วนสถานะสิทธิ์ (👉 คลิกที่กราฟเพื่อเจาะลึก)")
    fig6_1.update_traces(textinfo='label+value+percent parent')
    col_pie.plotly_chart(fig6_1, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้ป่วย AIDS")
st.markdown("---")

# ----- กราฟ 7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง -----
st.subheader("🪪 7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง")
df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])].copy()
df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])].copy()

total_valid_id = df_has_id.shape[0] + df_no_id.shape[0]
pct_has_id = (df_has_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0
pct_no_id = (df_no_id.shape[0] / total_valid_id * 100) if total_valid_id > 0 else 0

col1, col2 = st.columns(2)
col1.metric(label="ผู้ที่มีบัตรประชาชน (รวม)", value=f"{df_has_id.shape[0]} คน", delta=f"คิดเป็น {pct_has_id:.1f}%", delta_color="off")
col2.metric(label="ผู้ที่ไม่มีบัตรประชาชน (รวม)", value=f"{df_no_id.shape[0]} คน", delta=f"คิดเป็น {pct_no_id:.1f}%", delta_color="inverse")

df_sunburst = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร', 'ไม่มี', 'ไม่มีบัตร'])].copy()
df_sunburst['สถานะบัตร'] = df_sunburst['บปชช'].apply(lambda x: 'มีบัตรประชาชน' if x in ['มี', 'ใช่', 'มีบัตร'] else 'ไม่มีบัตรประชาชน')
sunburst_data = df_sunburst.groupby(['สถานะบัตร', 'เพศ']).size().reset_index(name='จำนวน')

fig7 = px.sunburst(
    sunburst_data, 
    path=['สถานะบัตร', 'เพศ'], 
    values='จำนวน',
    color='สถานะบัตร',
    color_discrete_map={'มีบัตรประชาชน': '#2ecc71', 'ไม่มีบัตรประชาชน': '#e74c3c'},
    title="สัดส่วนสถานะบัตรและเพศ (👉 คลิกที่กราฟวงกลมแต่ละฝั่งเพื่อเจาะลึกดูเพศ)"
)
fig7.update_traces(textinfo="label+value+percent parent")
st.plotly_chart(fig7, use_container_width=True)

st.markdown("#### 🔍 กลุ่มเปราะบางที่ไม่มีบัตรประชาชน (แยกเพศ)")

if not df_no_id.empty:
    def render_vul_card(title, icon, condition):
        m = df_no_id[condition & (df_no_id['เพศ'] == 'ชาย')].shape[0]
        f = df_no_id[condition & (df_no_id['เพศ'] == 'หญิง')].shape[0]
        total = m + f
        if total == 0:
            pct_m, pct_f = 0, 0
        else:
            pct_m = (m / total) * 100
            pct_f = (f / total) * 100
            
        return f"""
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e0e0e0; text-align: center; box-shadow: 1px 1px 5px rgba(0,0,0,0.05); margin-bottom: 10px;">
            <div style="font-size: 16px; font-weight: bold; color: #2c3e50; margin-bottom: 5px;">{icon} {title}</div>
            <div style="font-size: 24px; font-weight: bold; color: #e74c3c; margin-bottom: 5px;">{total} <span style="font-size: 14px; font-weight: normal; color: #7f8c8d;">คน</span></div>
            <div style="font-size: 13px; color: #34495e;">
                <span style="color: #3498db; font-weight: 500;">ชาย {m} ({pct_m:.0f}%)</span> &nbsp;|&nbsp; 
                <span style="color: #ff66b3; font-weight: 500;">หญิง {f} ({pct_f:.0f}%)</span>
            </div>
        </div>
        """

    cond_disabled = (df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')
    cond_elderly = (df_no_id['อายุ (ปี)'] >= 60)
    cond_novisit = (df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี']))
    cond_aids = (df_no_id['AIDS'].isin(['มี', 'ใช่']))

    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    
    col_c1.markdown(render_vul_card("ผู้พิการ", "♿", cond_disabled), unsafe_allow_html=True)
    col_c2.markdown(render_vul_card("ผู้สูงอายุ", "👵👴", cond_elderly), unsafe_allow_html=True)
    col_c3.markdown(render_vul_card("ไม่มีญาติเยี่ยม", "🚫", cond_novisit), unsafe_allow_html=True)
    col_c4.markdown(render_vul_card("ผู้ป่วย AIDS", "🎗️", cond_aids), unsafe_allow_html=True)

else:
    st.info("ไม่มีข้อมูลผู้ที่ไม่มีบัตรประชาชนให้วิเคราะห์")
