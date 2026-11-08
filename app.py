import streamlit as st
import pandas as pd
import plotly.express as px

# สมมติว่าคุณโหลดข้อมูลจาก Google Sheets มาเก็บในตัวแปร df แล้ว
# df = ... (โค้ดดึงข้อมูลของคุณ) ...

st.title("📊 แดชบอร์ดสถิติผู้ต้องขังกลุ่มเปราะบาง")
st.markdown("---")

# ==========================================
# 1. สถิติเด็กติดผู้ต้องขัง (แยกตามช่วงอายุและเพศ)
# ==========================================
st.subheader("👶 1. สถิติเด็กติดผู้ต้องขัง")
# กรองเฉพาะคนที่ระบุอายุเด็ก และไม่เท่ากับ 'ไม่มี' หรือค่าว่าง
df_child = df[df['เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna() & (df['เด็กติดผู้ต้องขัง (ช่วงวัย)'] != 'ไม่มี')]
if not df_child.empty:
    child_stats = df_child.groupby(['เด็กติดผู้ต้องขัง (ช่วงวัย)', 'เพศ']).size().reset_index(name='จำนวน')
    child_stats['เปอร์เซ็นต์'] = (child_stats['จำนวน'] / child_stats['จำนวน'].sum() * 100).round(1)
    
    fig1 = px.bar(child_stats, x='เด็กติดผู้ต้องขัง (ช่วงวัย)', y='จำนวน', color='เพศ', 
                  text='เปอร์เซ็นต์', barmode='group', 
                  color_discrete_sequence=['#3498db', '#e74c3c'])
    fig1.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลเด็กติดผู้ต้องขัง")

# ==========================================
# 2. จำนวนผู้ต้องขังหญิงตั้งครรภ์
# ==========================================
st.subheader("🤰 2. สถิติผู้ต้องขังหญิงตั้งครรภ์")
# ปรับชื่อคอลัมน์ "หญิงตั้งครรภ์" หรือ "สถานะการตั้งครรภ์" ตามชีตจริงของคุณ
pregnant_count = df[df['หญิงตั้งครรภ์'].astype(str).str.contains('มี', na=False)].shape[0]
st.metric(label="จำนวนหญิงตั้งครรภ์ทั้งหมด", value=f"{pregnant_count} คน")

# ==========================================
# 3. สถิติผู้ต้องขังสูงอายุ (60 ปีขึ้นไป แยกช่วงอายุและเพศ)
# ==========================================
st.subheader("👵👴 3. สถิติผู้ต้องขังสูงอายุ")
# กรองเอาเฉพาะอายุ 60 ขึ้นไป (ต้องแปลงคอลัมน์อายุเป็นตัวเลขก่อน)
df['อายุ (ปี)'] = pd.to_numeric(df['อายุ (ปี)'], errors='coerce')
df_elderly = df[df['อายุ (ปี)'] >= 60].copy()

if not df_elderly.empty:
    st.metric(label="รวมผู้สูงอายุทั้งหมด", value=f"{df_elderly.shape[0]} คน")
    
    # จัดกลุ่มช่วงอายุ
    bins = [59, 69, 79, float('inf')]
    labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
    df_elderly['ช่วงอายุ'] = pd.cut(df_elderly['อายุ (ปี)'], bins=bins, labels=labels)
    
    elder_stats = df_elderly.groupby(['ช่วงอายุ', 'เพศ']).size().reset_index(name='จำนวน')
    elder_stats['เปอร์เซ็นต์'] = (elder_stats['จำนวน'] / elder_stats['จำนวน'].sum() * 100).round(1)
    
    fig3 = px.bar(elder_stats, x='ช่วงอายุ', y='จำนวน', color='เพศ', 
                  text='เปอร์เซ็นต์', barmode='group',
                  color_discrete_sequence=['#2ecc71', '#f1c40f'])
    fig3.update_traces(texttemplate='%{text}%', textposition='outside')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้สูงอายุ")

# ==========================================
# 4. สถิติผู้พิการ (แยกเพศ, การขึ้นทะเบียน, และประเภท 1-7)
# ==========================================
st.subheader("♿ 4. สถิติผู้พิการ")
# กรองเฉพาะผู้ที่ระบุประเภทความพิการ (ไม่รวม 'ไม่มี' หรือค่าว่าง)
df_disabled = df[df['ประเภทความพิการ'].notna() & (df['ประเภทความพิการ'] != 'ไม่มี')]

if not df_disabled.empty:
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="ผู้พิการรวมทั้งหมด", value=f"{df_disabled.shape[0]} คน")
    with col2:
        reg_count = df_disabled[df_disabled['สถานะบัตรคนพิการ'].astype(str).str.contains('มี', na=False)].shape[0]
        st.metric(label="ขึ้นทะเบียนสิทธิ์ (มีบัตร) แล้ว", value=f"{reg_count} คน")
    
    # 4.1 กราฟวงกลม สัดส่วนผู้พิการ ชาย/หญิง พร้อมเปอร์เซ็นต์
    dis_gender = df_disabled['เพศ'].value_counts().reset_index()
    dis_gender.columns = ['เพศ', 'จำนวน']
    fig4_1 = px.pie(dis_gender, names='เพศ', values='จำนวน', title="สัดส่วนผู้พิการ ชาย/หญิง")
    st.plotly_chart(fig4_1, use_container_width=True)
    
    # 4.2 กราฟแท่ง แยกประเภทความพิการ (1-7) แบ่งตามเพศ
    dis_type = df_disabled.groupby(['ประเภทความพิการ', 'เพศ']).size().reset_index(name='จำนวน')
    fig4_2 = px.bar(dis_type, x='ประเภทความพิการ', y='จำนวน', color='เพศ', 
                    barmode='group', title="จำแนกตามประเภทความพิการ (1-7)")
    st.plotly_chart(fig4_2, use_container_width=True)
else:
    st.info("ไม่มีข้อมูลผู้พิการ")

# ==========================================
# 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม
# ==========================================
st.subheader("🚫 5. สถิติผู้ต้องขังไม่มีญาติเยี่ยม")
# ปรับชื่อคอลัมน์ให้ตรงกับชีตของคุณ
df_novisit = df[df['ไม่มีญาติเยี่ยม'].astype(str).str.contains('ไม่มี', na=False) | (df['ไม่มีญาติเยี่ยม'] == True)]

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
