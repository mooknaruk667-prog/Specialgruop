import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_gsheets import GSheetsConnection

# ตั้งค่าหน้าเพจ
st.set_page_config(page_title="Dashboard กลุ่มเปราะบาง", page_icon="📊", layout="wide")

st.title("📊 Dashboard สถิติผู้ต้องขังกลุ่มเปราะบาง")
st.markdown("---")

# 1. สร้างการเชื่อมต่อกับ Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. ฟังก์ชันโหลดข้อมูล (แคชไว้ 10 นาทีเพื่อไม่ให้โหลดซ้ำบ่อยเกินไป)
@st.cache_data(ttl=600)
def load_data():
    # ใส่ URL ของไฟล์ Google Sheets ที่เราสร้างไว้
    sheet_url = "https://docs.google.com/spreadsheets/d/1YRt4MqgSE5kS2ZHMn--COpWPDpCJjNd6AzmljG08-NI/edit"
    
    # อ่านข้อมูล
    df = conn.read(spreadsheet=sheet_url)
    # ตัดแถวที่ว่างเปล่าออก (กรณีอ่านเซลล์ว่างติดมา)
    df = df.dropna(subset=['เลขประจำตัวประชาชน', 'เพศ', 'อายุ (ปี)'], how='all')
    return df

try:
    df = load_data()

    # ----------------- การคำนวณตัวเลขสรุป (KPIs) -----------------
    elderly_df = df[pd.to_numeric(df['อายุ (ปี)'], errors='coerce') >= 60]
    elderly_count = len(elderly_df)
    pregnant_count = len(df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'] == 'ใช่'])
    disabled_count = len(df[df['[3] ผู้พิการ (ระบุประเภท)'].notna()])
    no_visitor_count = len(df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'] == 'ใช่'])
    child_count = len(df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna()])

    # แถวที่ 1: การ์ดแสดงตัวเลขสรุป
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("ผู้สูงอายุ (60+ ปี)", f"{elderly_count} คน")
    col2.metric("ผู้พิการ", f"{disabled_count} คน")
    col3.metric("ไม่มีญาติเยี่ยม", f"{no_visitor_count} คน")
    col4.metric("หญิงตั้งครรภ์", f"{pregnant_count} คน")
    col5.metric("เด็กติดผู้ต้องขัง", f"{child_count} คน")

    st.markdown("---")

    # แถวที่ 2: กราฟแสดงผล
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.subheader("ผู้สูงอายุ แยกตามช่วงวัยและเพศ")
        if not elderly_df.empty:
            bins = [59, 69, 79, 150]
            labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
            elderly_df['กลุ่มอายุ'] = pd.cut(pd.to_numeric(elderly_df['อายุ (ปี)']), bins=bins, labels=labels)
            
            age_group_counts = elderly_df.groupby(['กลุ่มอายุ', 'เพศ']).size().reset_index(name='จำนวน')
            fig_age = px.bar(age_group_counts, x='กลุ่มอายุ', y='จำนวน', color='เพศ', barmode='group', 
                             color_discrete_map={"ชาย":"#1f77b4", "หญิง":"#e377c2"})
            st.plotly_chart(fig_age, use_container_width=True)
        else:
            st.info("ยังไม่มีข้อมูลผู้สูงอายุ")

    with col_chart2:
        st.subheader("สัดส่วนประเภทผู้พิการ")
        disabled_df = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna()]
        if not disabled_df.empty:
            disability_counts = disabled_df['[3] ผู้พิการ (ระบุประเภท)'].value_counts().reset_index()
            disability_counts.columns = ['ประเภทความพิการ', 'จำนวน']
            fig_disability = px.pie(disability_counts, values='จำนวน', names='ประเภทความพิการ', hole=0.4)
            st.plotly_chart(fig_disability, use_container_width=True)
        else:
            st.info("ยังไม่มีข้อมูลผู้พิการ")

    st.subheader("ข้อมูลล่าสุด")
    st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error(f"ไม่สามารถเชื่อมต่อข้อมูลได้: {e}")
    st.info("กรุณาตรวจสอบการตั้งค่า Streamlit Secrets หรือสิทธิ์การเข้าถึงชีต")
