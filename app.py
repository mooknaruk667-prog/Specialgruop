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
st.set_page_config(page_title="แดชบอร์ดกลุ่มเปราะบาง", layout="wide")

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
    pdf.cell(0, 10, "รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง", new_x="LMARGIN", new_y="NEXT", align="C")
    
    months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
    now = datetime.now()
    date_str = f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year}"
    
    pdf.set_font("Sarabun", "", 12)
    pdf.cell(0, 6, date_str, new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(4)
    
    def draw_text(x, y, text, style="", size=12):
        pdf.set_xy(x, y)
        pdf.set_font("Sarabun", style, size)
        pdf.cell(90, 6, text, ln=0)
        return y + 7
        
    def draw_table(x, y, headers, data, widths):
        pdf.set_xy(x, y)
        pdf.set_font("Sarabun", "B", 11)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1, align="C")
        y += 7
        pdf.set_font("Sarabun", "", 11)
        for row in data:
            pdf.set_xy(x, y)
            for val, w in zip(row, widths):
                pdf.cell(w, 7, str(val)[:30], border=1, align="C")
            y += 7
        return y + 4

    col1_x = 10     
    col2_x = 110    
    y_start = pdf.get_y()
    
    # ================= คอลัมน์ 1 (ซ้าย) =================
    y = y_start
    y = draw_text(col1_x, y, "1. สถิติเด็กติดผู้ต้องขัง", "B", 14)
    df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')]
    if not df_child.empty:
        child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน')
        y = draw_table(col1_x, y, ["ช่วงอายุ", "เพศเด็ก", "จำนวน(คน)"], child_stats.values.tolist(), [40, 25, 25])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
    y = draw_text(col1_x, y, "2. สถิติผู้ต้องขังหญิงตั้งครรภ์", "B", 14)
    pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'].isin(['ใช่', 'มี'])].shape[0]
    y = draw_text(col1_x, y, f"จำนวนหญิงตั้งครรภ์ทั้งหมด: {pregnant_count} คน", "", 12) + 2
    
    y = draw_text(col1_x, y, "3. สถิติผู้ต้องขังสูงอายุ (60 ปีขึ้นไป)", "B", 14)
    df_elderly = df[df['อายุ (ปี)'] >= 60].copy()
    if not df_elderly.empty:
        has_allowance = df_elderly['เบี้ยสูงอายุ'].isin(['ได้รับ', 'มี', 'ใช่'])
        no_allowance = (~has_allowance).sum()
        y = draw_text(col1_x, y, f"รวม: {df_elderly.shape[0]} คน | ยังไม่ได้เบี้ย: {no_allowance} คน", "", 12)
        bins = [59, 69, 79, float('inf')]
        labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
        df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
        elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
        elder_stats = elder_stats[elder_stats['จำนวน'] > 0]
        y = draw_table(col1_x, y, ["ช่วงอายุ", "เพศ", "จำนวน(คน)"], elder_stats.values.tolist(), [40, 25, 25])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
    y = draw_text(col1_x, y, "4. สถิติผู้พิการ", "B", 14)
    df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')]
    if not df_disabled.empty:
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].isin(['มี', 'ใช่'])].shape[0]
        y = draw_text(col1_x, y, f"รวม: {df_disabled.shape[0]} คน | ขึ้นทะเบียนแล้ว: {reg_count} คน", "", 12)
        dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col1_x, y, ["ประเภทความพิการ", "เพศ", "จำนวน(คน)"], dis_type.values.tolist(), [50, 20, 20])
    else:
        y = draw_text(col1_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
    # ================= คอลัมน์ 2 (ขวา) =================
    y = y_start
    y = draw_text(col2_x, y, "5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม", "B", 14)
    df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])]
    if not df_novisit.empty:
        novisit_stats = df_novisit.groupby(['เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col2_x, y, ["เพศ", "จำนวน(คน)"], novisit_stats.values.tolist(), [45, 45])
    else:
        y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
    y = draw_text(col2_x, y, "6. สถิติผู้ป่วย AIDS", "B", 14)
    df_aids = df[df['AIDS'].isin(['มี', 'ใช่'])]
    if not df_aids.empty:
        no_rights = df_aids[df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].isin(['ไม่มี'])].shape[0]
        y = draw_text(col2_x, y, f"รวม: {df_aids.shape[0]} คน | ยังไม่มีสิทธิ: {no_rights} คน", "", 12)
        aids_stats = df_aids.groupby(['เพศ']).size().reset_index(name='จำนวน')
        y = draw_table(col2_x, y, ["เพศ", "จำนวน(คน)"], aids_stats.values.tolist(), [45, 45])
    else:
        y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
    y = draw_text(col2_x, y, "7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง", "B", 14)
    df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])]
    df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])]
    pct_no_id = (df_no_id.shape[0] / total_inmates * 100) if total_inmates > 0 else 0
    
    y = draw_text(col2_x, y, f"ผู้ที่มีบัตรประชาชนรวม: {df_has_id.shape[0]} คน", "", 12)
    y = draw_text(col2_x, y, f"ผู้ที่ไม่มีบัตรประชาชนรวม: {df_no_id.shape[0]} คน (คิดเป็น {pct_no_id:.1f}%)", "", 12)
    
    if not df_no_id.empty:
        y += 2
        y = draw_text(col2_x, y, "กลุ่มเปราะบางที่ไม่มีบัตรประชาชน (แยกเพศ):", "B", 12)
        
        # คำนวณแยกเพศ
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
        # กรองเฉพาะกลุ่มที่มีคนไม่มีบัตร > 0
        data = [r for r in data if r[3] > 0]
        
        if data:
            y = draw_table(col2_x, y, ["กลุ่ม", "ชาย", "หญิง", "รวม"], data, [36, 18, 18, 18])
        else:
            y = draw_text(col2_x, y, "- ไม่มีข้อมูล -", "", 12) + 2
        
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
    ws.cell(row=2, column=2, value=f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year}").font = Font(name='Tahoma', size=11, italic=True)

    current_row = [4]
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

    # 1 - 6
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

    df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')]
    if not df_disabled.empty:
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].isin(['มี', 'ใช่'])].shape[0]
        dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน (คน)')
        dis_type.columns = ['ประเภทความพิการ', 'เพศ', 'จำนวน (คน)']
        add_table(dis_type, f"♿ 4. สถิติผู้พิการ (รวม: {df_disabled.shape[0]} คน | ขึ้นทะเบียนแล้ว: {reg_count} คน)")

    df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])]
    if not df_novisit.empty:
        novisit_stats = df_novisit.groupby(['เพศ']).size().reset_index(name='จำนวน (คน)')
        add_table(novisit_stats, "🚫 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม")

    df_aids = df[df['AIDS'].isin(['มี', 'ใช่'])]
    if not df_aids.empty:
        no_rights = df_aids[df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].isin(['ไม่มี'])].shape[0]
        aids_stats = df_aids.groupby(['เพศ']).size().reset_index(name='จำนวน (คน)')
        add_table(aids_stats, f"🎗️ 6. สถิติผู้ป่วย AIDS (รวม: {df_aids.shape[0]} คน | ยังไม่มีสิทธิ: {no_rights} คน)")

    # 7. บัตรประชาชน (เพิ่มคอลัมน์เปอร์เซ็นต์ใน Excel)
    df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])]
    df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])]
    pct_has_id = (df_has_id.shape[0] / total_inmates * 100) if total_inmates > 0 else 0
    pct_no_id = (df_no_id.shape[0] / total_inmates * 100) if total_inmates > 0 else 0
    
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
        # ซ่อนกลุ่มที่ไม่มีข้อมูล
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
st.title("📊 แดชบอร์ดสถิติผู้ต้องขังกลุ่มเปราะบาง")

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
                  text='เปอร์เซ็นต์', barmode='group', color_discrete_sequence=['#3498db', '#e74c3c'], title="แยกตามช่วงอายุและเพศเด็ก")
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
                  color_discrete_sequence=['#2ecc71', '#f1c40f'], title="แยกตามช่วงอายุและเพศ")
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
    reg_percent = (reg_count / total_disabled) * 100 if total_disabled > 0 else 0
    
    col1, col2 = st.columns(2)
    col1.metric(label="ผู้พิการรวมทั้งหมด", value=f"{total_disabled} คน")
    col2.metric(label="ขึ้นทะเบียนสิทธิ์ (มีบัตร) แล้ว", value=f"{reg_count} คน", delta=f"คิดเป็น {reg_percent:.1f}%", delta_color="off")
    
    col3, col4 = st.columns(2)
    dis_gender = df_disabled['เพศ'].value_counts().reset_index()
    dis_gender.columns = ['เพศ', 'จำนวน']
    fig4_1 = px.pie(dis_gender, names='เพศ', values='จำนวน', title="สัดส่วนผู้พิการ ชาย/หญิง", hole=0.3)
    col3.plotly_chart(fig4_1, use_container_width=True)
    
    dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน')
    dis_type['เปอร์เซ็นต์'] = (dis_type['จำนวน'] / dis_type['จำนวน'].sum() * 100).round(1)
    fig4_2 = px.bar(dis_type, x='[3] ผู้พิการ (ระบุประเภท)', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', barmode='group', title="แยกตามประเภท 1-7")
    fig4_2.update_traces(texttemplate='%{text}%', textposition='outside')
    col4.plotly_chart(fig4_2, use_container_width=True)
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
    fig5 = px.bar(novisit_stats, x='เพศ', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', title="ผู้ต้องขังไม่มีญาติเยี่ยม (แยกเพศ)")
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
    df_no_rights = df_aids[df_aids['ขึ้นทะเบียนสิทธิ(AIDS)'].isin(['ไม่มี'])].copy()
    no_rights_count = df_no_rights.shape[0]
    no_rights_percent = (no_rights_count / total_aids) * 100 if total_aids > 0 else 0
    
    col1, col2 = st.columns(2)
    col1.metric(label="จำนวนผู้ป่วย AIDS รวมทั้งหมด", value=f"{total_aids} คน")
    col2.metric(label="ผู้ป่วย AIDS ที่ยังไม่มีสิทธิ์", value=f"{no_rights_count} คน", delta=f"คิดเป็น {no_rights_percent:.1f}%", delta_color="off")
    
    col3, col4 = st.columns(2)
    aids_stats = df_aids['เพศ'].value_counts().reset_index()
    aids_stats.columns = ['เพศ', 'จำนวน']
    aids_stats['เปอร์เซ็นต์'] = (aids_stats['จำนวน'] / aids_stats['จำนวน'].sum() * 100).round(1)
    fig6_1 = px.bar(aids_stats, x='เพศ', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', title="ผู้ป่วย AIDS ทั้งหมด (แยกเพศ)")
    fig6_1.update_traces(texttemplate='%{text}%', textposition='outside')
    col3.plotly_chart(fig6_1, use_container_width=True)
    
    if not df_no_rights.empty:
        no_rights_stats = df_no_rights['เพศ'].value_counts().reset_index()
        no_rights_stats.columns = ['เพศ', 'จำนวน']
        no_rights_stats['เปอร์เซ็นต์'] = (no_rights_stats['จำนวน'] / no_rights_stats['จำนวน'].sum() * 100).round(1)
        fig6_2 = px.bar(no_rights_stats, x='เพศ', y='จำนวน', color='เพศ', text='เปอร์เซ็นต์', title="กลุ่มผู้ป่วยที่ยังไม่มีสิทธิ์ (แยกเพศ)", color_discrete_sequence=['#e74c3c', '#95a5a6'])
        fig6_2.update_traces(texttemplate='%{text}%', textposition='outside')
        col4.plotly_chart(fig6_2, use_container_width=True)
    else:
        col4.success("ผู้ป่วย AIDS ทุกคนได้รับการขึ้นทะเบียนสิทธิ์เรียบร้อยแล้ว")
else:
    st.info("ไม่มีข้อมูลผู้ป่วย AIDS")
st.markdown("---")

# ----- กราฟ 7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง -----
st.subheader("🪪 7. รายละเอียด บัตรประชาชน กลุ่มเปราะบาง")
df_has_id = df[df['บปชช'].isin(['มี', 'ใช่', 'มีบัตร'])].copy()
df_no_id = df[df['บปชช'].isin(['ไม่มี', 'ไม่มีบัตร'])].copy()

pct_has_id = (df_has_id.shape[0] / total_inmates * 100) if total_inmates > 0 else 0
pct_no_id = (df_no_id.shape[0] / total_inmates * 100) if total_inmates > 0 else 0

col1, col2 = st.columns(2)
col1.metric(label="ผู้ที่มีบัตรประชาชน (รวม)", value=f"{df_has_id.shape[0]} คน", delta=f"คิดเป็น {pct_has_id:.1f}%", delta_color="off")
col2.metric(label="ผู้ที่ไม่มีบัตรประชาชน (รวม)", value=f"{df_no_id.shape[0]} คน", delta=f"คิดเป็น {pct_no_id:.1f}%", delta_color="inverse")

col3, col4 = st.columns(2)

with col3:
    id_summary = pd.DataFrame({
        'สถานะ': ['มีบัตรประชาชน', 'ไม่มีบัตรประชาชน'],
        'จำนวน': [df_has_id.shape[0], df_no_id.shape[0]]
    })
    fig7_1 = px.pie(id_summary, names='สถานะ', values='จำนวน', title="สัดส่วนผู้มี/ไม่มีบัตร (ภาพรวม)", 
                    hole=0.4, color='สถานะ', color_discrete_map={'มีบัตรประชาชน':'#2ecc71', 'ไม่มีบัตรประชาชน':'#e74c3c'})
    fig7_1.update_traces(textinfo='percent+label')
    st.plotly_chart(fig7_1, use_container_width=True)

with col4:
    if not df_no_id.empty:
        no_id_gender = df_no_id['เพศ'].value_counts().reset_index()
        no_id_gender.columns = ['เพศ', 'จำนวน']
        fig7_2 = px.pie(no_id_gender, names='เพศ', values='จำนวน', title="สัดส่วนผู้ที่ไม่มีบัตร (แยกเพศ)", 
                        hole=0.4, color_discrete_sequence=['#e74c3c', '#c0392b'])
        fig7_2.update_traces(textinfo='percent+label')
        st.plotly_chart(fig7_2, use_container_width=True)
    else:
        st.success("ผู้ต้องขังทุกคนมีบัตรประชาชนเรียบร้อยแล้ว")

st.markdown("#### 🔍 กลุ่มเปราะบางที่ไม่มีบัตรประชาชน (แยกเพศ)")
if not df_no_id.empty:
    disabled_m = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'ชาย')].shape[0]
    disabled_f = df_no_id[(df_no_id['[3] ผู้พิการ (ระบุประเภท)'].notna()) & (df_no_id['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี') & (df_no_id['เพศ'] == 'หญิง')].shape[0]
    
    elderly_m = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
    elderly_f = df_no_id[(df_no_id['อายุ (ปี)'] >= 60) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
    
    novisit_m = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
    novisit_f = df_no_id[(df_no_id['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].isin(['ใช่', 'ไม่มีญาติเยี่ยม', 'ไม่มี'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
    
    aids_m = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'ชาย')].shape[0]
    aids_f = df_no_id[(df_no_id['AIDS'].isin(['มี', 'ใช่'])) & (df_no_id['เพศ'] == 'หญิง')].shape[0]
    
    vul_no_id_data = pd.DataFrame([
        {'กลุ่มเปราะบาง': 'ผู้พิการ', 'เพศ': 'ชาย', 'จำนวน (คน)': disabled_m},
        {'กลุ่มเปราะบาง': 'ผู้พิการ', 'เพศ': 'หญิง', 'จำนวน (คน)': disabled_f},
        {'กลุ่มเปราะบาง': 'ผู้สูงอายุ', 'เพศ': 'ชาย', 'จำนวน (คน)': elderly_m},
        {'กลุ่มเปราะบาง': 'ผู้สูงอายุ', 'เพศ': 'หญิง', 'จำนวน (คน)': elderly_f},
        {'กลุ่มเปราะบาง': 'ไม่มีญาติเยี่ยม', 'เพศ': 'ชาย', 'จำนวน (คน)': novisit_m},
        {'กลุ่มเปราะบาง': 'ไม่มีญาติเยี่ยม', 'เพศ': 'หญิง', 'จำนวน (คน)': novisit_f},
        {'กลุ่มเปราะบาง': 'ผู้ป่วย AIDS', 'เพศ': 'ชาย', 'จำนวน (คน)': aids_m},
        {'กลุ่มเปราะบาง': 'ผู้ป่วย AIDS', 'เพศ': 'หญิง', 'จำนวน (คน)': aids_f}
    ])
    
    vul_no_id_data = vul_no_id_data[vul_no_id_data['จำนวน (คน)'] > 0]
    
    if not vul_no_id_data.empty:
        fig7_3 = px.bar(vul_no_id_data, x='จำนวน (คน)', y='กลุ่มเปราะบาง', orientation='h', 
                        title="จำนวนผู้ที่ไม่มีบัตรประชาชน แยกตามกลุ่มเปราะบางและเพศ", 
                        color='เพศ', text='จำนวน (คน)', barmode='group',
                        color_discrete_map={'ชาย': '#3498db', 'หญิง': '#e74c3c'})
        fig7_3.update_traces(textposition='outside')
        fig7_3.update_layout(yaxis_title=None)
        st.plotly_chart(fig7_3, use_container_width=True)
    else:
        st.info("ไม่มีข้อมูลผู้ที่ไม่มีบัตรประชาชนในกลุ่มเปราะบางหลัก")
else:
    st.info("ไม่มีข้อมูลผู้ที่ไม่มีบัตรประชาชนให้วิเคราะห์")
