import streamlit as st
import pandas as pd
import plotly.express as px

# ตั้งค่าหน้าเพจ
st.set_page_config(page_title="Dashboard สถิติผู้ต้องขังกลุ่มเปราะบาง", page_icon="📊", layout="wide")

st.title("📊 Dashboard สถิติผู้ต้องขังกลุ่มเปราะบาง")
st.markdown("---")

# ฟังก์ชันสำหรับโหลดข้อมูล (เปลี่ยนชื่อไฟล์ตามที่ใช้งานจริง)
@st.cache_data
def load_data():
    # สมมติว่าดาวน์โหลดไฟล์เป็นชื่อ สถิติผู้ต้องขังกลุ่มเปราะบาง.xlsx มาไว้ที่โฟลเดอร์เดียวกัน
    # หรือถ้าเชื่อมผ่าน Google Sheets API ก็สามารถดึงข้อมูลผ่าน DataFrame ได้เลย
    try:
        df = pd.read_excel("สถิติผู้ต้องขังกลุ่มเปราะบาง.xlsx")
    except:
        # ข้อมูลจำลองกรณีที่ยังไม่ได้ใส่ไฟล์
        data = {
            "เพศ": ["ชาย", "หญิง", "ชาย", "หญิง", "ชาย"],
            "อายุ (ปี)": [72, 35, 65, 28, 61],
            "[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)": [None, "ใช่", None, None, None],
            "[3] ผู้พิการ (ระบุประเภท)": ["ประเภท 3", None, None, None, "ประเภท 1"],
            "[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)": ["ใช่", None, "ใช่", None, "ใช่"],
            "[1] เด็กติดผู้ต้องขัง (ช่วงวัย)": [None, None, None, "4-6 เดือน", None]
        }
        df = pd.DataFrame(data)
    return df

df = load_data()

# ----------------- การคำนวณตัวเลขสรุป (KPIs) -----------------
# 1. ผู้สูงอายุ (60 ปีขึ้นไป)
elderly_df = df[pd.to_numeric(df['อายุ (ปี)'], errors='coerce') >= 60]
elderly_count = len(elderly_df)

# 2. หญิงตั้งครรภ์
pregnant_count = len(df[df['[2] หญิงตั้งครรภ์ (ใช่/ไม่ใช่)'] == 'ใช่'])

# 3. ผู้พิการ
disabled_count = len(df[df['[3] ผู้พิการ (ระบุประเภท)'].notna()])

# 4. ไม่มีญาติเยี่ยม
no_visitor_count = len(df[df['[4] ไม่มีญาติเยี่ยม (ใช่/ไม่ใช่)'] == 'ใช่'])

# 5. เด็กติดผู้ต้องขัง
child_count = len(df[df['[1] เด็กติดผู้ต้องขัง (ช่วงวัย)'].notna()])


# ----------------- ส่วนแสดงผล -----------------
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
    st.subheader("จำนวนผู้สูงอายุ แยกตามช่วงวัย")
    if not elderly_df.empty:
        # แบ่งกลุ่มอายุ
        bins = [59, 69, 79, 150]
        labels = ['60-69 ปี', '70-79 ปี', '80 ปีขึ้นไป']
        elderly_df['กลุ่มอายุ'] = pd.cut(elderly_df['อายุ (ปี)'], bins=bins, labels=labels)
        
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

# ส่วนแสดงข้อมูลดิบ (ตาราง)
st.subheader("ตารางข้อมูลล่าสุด")
st.dataframe(df.head(20), use_container_width=True)
