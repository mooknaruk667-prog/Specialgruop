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
    df_raw = load_data()
except Exception as e:
    st.error(f"ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบการแชร์ไฟล์: {e}")
    st.stop()

# 📌 กรองเฉพาะผู้ต้องขังที่คอลัมน์ "อยู่" มีข้อความว่า "อยู่"
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
df_aids_all, df_aids_get, df_aids_not = build_two_group_df(
    df, ['AIDS ได้เบี้ย', 'AIDSได้เบี้ย'], ['AIDS ไม่ได้เบี้ย', 'AIDSไม่ได้เบี้ย'], 'AIDS ได้เบี้ย', 'AIDS ไม่ได้เบี้ย'
)
df_no_id = df[has_keyword(df['สถานะ'], ['ไม่มีบัตรประชาชน', 'บัตรปชช หาย', 'ไม่มีบัตร'])].copy()

# ------------------------------------------------
# 2. ระบบสร้าง PDF
# ------------------------------------------------
def generate_pdf():
    load_fonts()
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("Sarabun", style="", fname="Sarabun-Regular.ttf")
    pdf.add_font("Sarabun", style="B", fname="Sarabun-Bold.ttf")

    pdf.set_font("Sarabun", "B", 16)
    pdf.cell(0, 10, "รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์", new_x="LMARGIN", new_y="NEXT", align="C")

    months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
    now = datetime.now()
    date_str = f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year + 543} (เฉพาะผู้ต้องขังที่อยู่: {df.shape[0]} คน)"
    pdf.set_font("Sarabun", "", 11)
    pdf.cell(0, 6, date_str, new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    def draw_section(title, rows):
        pdf.set_font("Sarabun", "B", 13)
        pdf.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Sarabun", "B", 10)
        widths = [65, 25, 35, 35, 30]
        headers = ["รายการ", "รวม (คน)", "ชาย (คน / %)", "หญิง (คน / %)", "สัดส่วนกลุ่ม"]
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1, align="C")
        pdf.ln()
        pdf.set_font("Sarabun", "", 10)
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
        return [label, t, f"{m} ({pm:.1f}%)", f"{f} ({pf:.1f}%)", f"{pct_parent:.1f}%"]

    # 1. ผู้สูงอายุ
    t_e = df_elder_all.shape[0]
    draw_section("1. สถิติผู้สูงอายุ", [
        make_row("ผู้สูงอายุทั้งหมด", df_elder_all, t_e),
        make_row("- สูงอายุได้เบี้ย", df_elder_get, t_e),
        make_row("- สูงอายุไม่ได้เบี้ย", df_elder_not, t_e),
    ])

    # 2. ผู้พิการ
    t_d = df_dis_all.shape[0]
    draw_section("2. สถิติผู้พิการ", [
        make_row("ผู้พิการทั้งหมด", df_dis_all, t_d),
        make_row("- พิการได้เบี้ย", df_
