import streamlit as st
import pandas as pd
import plotly.express as px

# ------------------------------------------------
# 1. ตั้งค่าหน้าเพจ & ดึงข้อมูลจาก Google Sheets
# ------------------------------------------------
st.set_page_config(page_title="แดชบอร์ดกลุ่มเปราะบาง", layout="wide")

@st.cache_data(ttl=600)  # รีเฟรชข้อมูลทุกๆ 10 นาที
def load_data():
    # ลิงก์ Export CSV จาก Google Sheets ของคุณ
    sheet_id = "1YRt4MqgSE5kS2ZHMn--COpWPDpCJjNd6AzmljG08-NI"
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    
    # อ่านข้อมูล
    df = pd.read_csv(csv_url)
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"ไม่สามารถดึงข้อมูลจาก Google Sheets ได้ กรุณาตรวจสอบการแชร์ไฟล์: {e}")
    st.stop()

st.title("📊 แดชบอร์ดสถิติผู้ต้องขังกลุ่มเปราะบาง")
st.markdown("---")

# ------------------------------------------------
# 1. สถิติเด็กติดผู้ต้องขัง
# ------------------------------------------------
st.subheader("👶 1. สถิติเด็กติดผู้ต้องขัง")
# กรองข้อมูลที่มีอายุเด็ก และไม่เท่ากับคำว่า 'ไม่มี'
df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')].copy()

if not df_child.empty:
    # นับจำนวนแยกตามช่วงวัย และเพศเด็ก
    child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน')
    child_stats['เปอร์เซ็นต์'] = (child_stats['จำนวน'] / child_stats['จำนวน'].sum() * 100).round(1)
    
    fig1 = px.bar(child_stats, x='[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', y='จำนวน', color='[1] เด็กติดผู้ต้องขัง (เพศเด็ก)', 
                  text='เปอร์เซ็นต์', barmode='group', 
                  color_discrete_sequence=['#3498db', '#e74c3c'],
                  title="จำนวนเด็กติดผู้ต้องขังแยกตามช่วงอายุและเพศ")
    fig1.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.info("ยังไม่มีข้อมูลเด็กติดผู้ต้องขัง")

st.markdown("---")

# ------------------------------------------------
# 2. จำนวนผู้ต้องขังหญิงตั้งครรภ์
# ------------------------------------------------
st.subheader("🤰 2. สถิติผู้ต้องขังหญิงตั้งครรภ์")
# นับเฉพาะคนที่ระบุว่า "ใช่"
pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'] == 'ใช่'].shape[0]
st.metric(label="จำนวนหญิงตั้งครรภ์ทั้งหมด (คน)", value=pregnant_count)

st.markdown("---")

# ------------------------------------------------
# 3. สถิติผู้ต้องขังสูงอายุ (60 ปีขึ้นไป)
# ------------------------------------------------
st.subheader("👵👴 3. สถิติผู้ต้องขังสูงอายุ")
# แปลงคอลัมน์อายุให้เป็นตัวเลข และกรอง 60 ขึ้นไป
df['อายุ (ปี)'] = pd.to_numeric(df['อายุ (ปี)'], errors='coerce')
df_elderly = df[df['อายุ (ปี)'] >= 60].copy()

if not df_elderly.empty:
    st.metric(label="รวมผู้สูงอายุทั้งหมด (คน)", value=df_elderly.shape[0])
    
    # จัดกลุ่มช่วงอายุ
    bins = [59, 69, 79, float('inf')]
    labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
    df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
    
    elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
    elder_stats['เปอร์เซ็นต์'] = (elder_stats['จำนวน'] / elder_stats['จำนวน'].sum() * 100).round(1)
    
    fig3 = px.bar(elder_stats, x='ช่วงอายุ', y='จำนวน', color='เพศ', 
                  text='เปอร์เซ็นต์', barmode='group',
                  color_discrete_sequence=['#2ecc71', '#f1c40f'],
                  title="จำนวนผู้สูงอายุแยกตามช่วงอายุและเพศ")
    fig3.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("ยังไม่มีข้อมูลผู้สูงอายุ")

st.markdown("---")

# ------------------------------------------------
# 4. สถิติผู้พิการ
# ------------------------------------------------
st.subheader("♿ 4. สถิติผู้พิการ")
df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')].copy()

if not df_disabled.empty:
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="ผู้พิการรวมทั้งหมด (คน)", value=df_disabled.shape[0])
    with col2:
        # ดูสถานะว่าขึ้นทะเบียน (มีบัตร) หรือไม่ (หาคำว่า "มีบัตร" หรือ "มี")
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].astype(str).str.contains('มี', na=False)].shape[0]
        st.metric(label="ขึ้นทะเบียนสิทธิ์ (มีบัตร) แล้ว (คน)", value=reg_count)
    
    col3, col4 = st.columns(2)
    
    with
