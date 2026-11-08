import streamlit as st
import pandas as pd
import plotly.express as px

# ------------------------------------------------
# 1. ตั้งค่าหน้าเพจ & ดึงข้อมูลจาก Google Sheets
# ------------------------------------------------
st.set_page_config(page_title="แดชบอร์ดกลุ่มเปราะบาง", layout="wide")

@st.cache_data(ttl=600)  # ดึงข้อมูลใหม่ทุกๆ 10 นาที
def load_data():
    # ใส่ ID ชีตของคุณ
    sheet_id = "1YRt4MqgSE5kS2ZHMn--COpWPDpCJjNd6AzmljG08-NI"
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    
    # อ่านข้อมูล
    df = pd.read_csv(csv_url)
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบการแชร์ไฟล์: {e}")
    st.stop()

st.title("📊 แดชบอร์ดสถิติผู้ต้องขังกลุ่มเปราะบาง")
st.markdown("---")

# ------------------------------------------------
# 1. สถิติเด็กติดผู้ต้องขัง
# ------------------------------------------------
st.subheader("👶 1. สถิติเด็กติดผู้ต้องขัง")
# กรองเอาเฉพาะคนที่ช่วงวัยไม่เป็นค่าว่าง และไม่ใช่คำว่า 'ไม่มี'
df_child = df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')].copy()

if not df_child.empty:
    child_stats = df_child.groupby(['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', '[1] เด็กติดผู้ต้องขัง (เพศเด็ก)']).size().reset_index(name='จำนวน')
    child_stats['เปอร์เซ็นต์'] = (child_stats['จำนวน'] / child_stats['จำนวน'].sum() * 100).round(1)
    
    fig1 = px.bar(child_stats, x='[1] เด็กติดผู้ต้องขัง (ช่วงวัย)', y='จำนวน', color='[1] เด็กติดผู้ต้องขัง (เพศเด็ก)', 
                  text='เปอร์เซ็นต์', barmode='group', 
                  color_discrete_sequence=['#3498db', '#e74c3c'],
                  title="แยกตามช่วงอายุและเพศเด็ก")
    fig1.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลเด็กติดผู้ต้องขัง")

st.markdown("---")

# ------------------------------------------------
# 2. จำนวนผู้ต้องขังหญิงตั้งครรภ์
# ------------------------------------------------
st.subheader("🤰 2. สถิติผู้ต้องขังหญิงตั้งครรภ์")
# กรองเฉพาะคนที่ติ๊ก "ใช่" หรือ "มี" 
pregnant_count = df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'].astype(str).str.contains('ใช่|มี', na=False)].shape[0]
st.metric(label="จำนวนหญิงตั้งครรภ์ทั้งหมด", value=f"{pregnant_count} คน")

st.markdown("---")

# ------------------------------------------------
# 3. สถิติผู้ต้องขังสูงอายุ (60 ปีขึ้นไป)
# ------------------------------------------------
st.subheader("👵👴 3. สถิติผู้ต้องขังสูงอายุ")
df['อายุ (ปี)'] = pd.to_numeric(df['อายุ (ปี)'], errors='coerce')
df_elderly = df[df['อายุ (ปี)'] >= 60].copy()

if not df_elderly.empty:
    st.metric(label="ผู้สูงอายุรวมทั้งหมด", value=f"{df_elderly.shape[0]} คน")
    
    # แบ่งช่วงอายุตามที่คุณต้องการ
    bins = [59, 69, 79, float('inf')]
    labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
    df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
    
    elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
    elder_stats['เปอร์เซ็นต์'] = (elder_stats['จำนวน'] / elder_stats['จำนวน'].sum() * 100).round(1)
    
    fig3 = px.bar(elder_stats, x='ช่วงอายุ', y='จำนวน', color='เพศ', 
                  text='เปอร์เซ็นต์', barmode='group',
                  color_discrete_sequence=['#2ecc71', '#f1c40f'],
                  title="แยกตามช่วงอายุและเพศ")
    fig3.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้สูงอายุ")

st.markdown("---")

# ------------------------------------------------
# 4. สถิติผู้พิการ
# ------------------------------------------------
st.subheader("♿ 4. สถิติผู้พิการ")
df_disabled = df[df['[3] ผู้พิการ (ระบุประเภท)'].notna() & (df['[3] ผู้พิการ (ระบุประเภท)'] != 'ไม่มี')].copy()

if not df_disabled.empty:
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="ผู้พิการรวมทั้งหมด", value=f"{df_disabled.shape[0]} คน")
    with col2:
        # ขึ้นทะเบียน (สถานะบัตรคนพิการไม่ใช่คำว่า "ไม่มี" หรือค่าว่าง)
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].astype(str).str.contains('มี', na=False)].shape[0]
        st.metric(label="ขึ้นทะเบียนสิทธิ์ (มีบัตร) แล้ว", value=f"{reg_count} คน")
    
    col3, col4 = st.columns(2)
    
    with col3:
        dis_gender = df_disabled['เพศ'].value_counts().reset_index()
        dis_gender.columns = ['เพศ', 'จำนวน']
        fig4_1 = px.pie(dis_gender, names='เพศ', values='จำนวน', title="สัดส่วนผู้พิการ ชาย/หญิง", hole=0.3)
        fig4_1.update_traces(textinfo='percent+label')
        st.plotly_chart(fig4_1, use_container_width=True)
        
    with col4:
        dis_type = df_disabled.groupby(['[3] ผู้พิการ (ระบุประเภท)', 'เพศ']).size().reset_index(name='จำนวน')
        dis_type['เปอร์เซ็นต์'] = (dis_type['จำนวน'] / dis_type['จำนวน'].sum() * 100).round(1)
        
        fig4_2 = px.bar(dis_type, x='[3] ผู้พิการ (ระบุประเภท)', y='จำนวน', color='เพศ', 
                        text='เปอร์เซ็นต์', barmode='group', title="แยกตามประเภท 1-7")
        fig4_2.update_traces(texttemplate='%{text}%', textposition='outside')
        st.plotly_chart(fig4_2, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้พิการ")

st.markdown("---")

# ------------------------------------------------
# 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม
# ------------------------------------------------
st.subheader("🚫 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม")
df_novisit = df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'].astype(str).str.contains('ใช่|ไม่มี', na=False)].copy()

if not df_novisit.empty:
    novisit_stats = df_novisit['เพศ'].value_counts().reset_index()
    novisit_stats.columns = ['เพศ', 'จำนวน']
    novisit_stats['เปอร์เซ็นต์'] = (novisit_stats['จำนวน'] / novisit_stats['จำนวน'].sum() * 100).round(1)
    
    fig5 = px.bar(novisit_stats, x='เพศ', y='จำนวน', color='เพศ', 
                  text='เปอร์เซ็นต์', title="ผู้ต้องขังไม่มีญาติเยี่ยม (แยกเพศ)")
    fig5.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig5, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้ต้องขังที่ไม่มีญาติเยี่ยม")
