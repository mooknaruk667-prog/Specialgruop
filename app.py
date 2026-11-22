# ------------------------------------------------
# 3. ระบบสร้าง Excel (แยกชีตข้อมูลดิบตามกลุ่มเปราะบาง)
# ------------------------------------------------
def generate_excel():
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

    months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
    now = datetime.now()
    ws.cell(row=2, column=2, value="รายงานสรุปสถิติผู้ต้องขังกลุ่มเปราะบาง เรือนจำจังหวัดบุรีรัมย์").font = title_font
    ws.cell(row=3, column=2, value=f"ข้อมูล ณ วันที่ {now.day} {months[now.month-1]} {now.year + 543} (เฉพาะผู้ต้องขังที่อยู่)").font = Font(name='Tahoma', size=11, italic=True)

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
        ('ผู้ป่วย AIDS ทั้งหมด', df_aids_all), ('AIDS ได้เบี้ย', df_aids_get), ('AIDS ไม่ได้เบี้ย', df_aids_not)
    ], df_aids_all.shape[0]), "4. สถิติผู้ป่วย AIDS")

    add_table(make_excel_rows([
        ('ไม่มีบัตรประชาชนทั้งหมด', df_no_id)
    ], df_no_id.shape[0]), "5. สถิติผู้ต้องขังไม่มีบัตรประชาชน")

    for col_letter, width in [('B', 28), ('C', 14), ('D', 14), ('E', 14), ('F', 14), ('G', 14), ('H', 16)]:
        ws.column_dimensions[col_letter].width = width
    ws.sheet_view.showGridLines = False

    # 📌 ฟังก์ชันช่วยสร้างชีตข้อมูลดิบแยกตามกลุ่ม พร้อมจัดหัวตารางให้สวยงาม
    def add_raw_sheet(sheet_name, target_df):
        ws_sub = wb.create_sheet(title=sheet_name)
        clean_df = target_df.drop(columns=['กลุ่มย่อย'], errors='ignore')
        
        if clean_df.empty:
            ws_sub.cell(row=1, column=1, value="ไม่มีข้อมูลในกลุ่มนี้").font = data_font
            return

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

        # ปรับความกว้างคอลัมน์อัตโนมัติ
        for col in ws_sub.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws_sub.column_dimensions[col[0].column_letter].width = min(max_len + 5, 40)

    # สร้างชีตแยกแต่ละกลุ่ม
    add_raw_sheet("1. ผู้สูงอายุ", df_elder_all)
    add_raw_sheet("2. ผู้พิการ", df_dis_all)
    add_raw_sheet("3. หญิงตั้งครรภ์", df_pregnant)
    add_raw_sheet("4. ผู้ป่วย AIDS", df_aids_all)
    add_raw_sheet("5. ไม่มีบัตรประชาชน", df_no_id)
    add_raw_sheet("ข้อมูลดิบทั้งหมด", df)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
