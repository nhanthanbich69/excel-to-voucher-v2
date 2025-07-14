import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
import traceback
import re
from openpyxl import load_workbook
from collections import defaultdict

# ====== CẤU HÌNH GIAO DIỆN =======
st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán Chuẩn từ Excel")
tab1, tab2, tab3, tab4 = st.tabs([
    "🧾 Tạo File Hạch Toán", 
    "🔍 So sánh và Xoá dòng trùng", 
    "📊 File tuỳ chỉnh (Check thủ công)", 
    "📐 So sánh Số tiền giữa các file"
])

# ====== HÀM TIỆN ÍCH CHUNG =======
def extract_month_year_from_filename(filename):
    try:
        match = re.search(r'(\d{4})[\.\-_]?\s*(\d{2})|\s*(\d{2})[\.\-_]?\s*(\d{4})', filename)
        if match:
            year = match.group(1) or match.group(4)
            month = match.group(2) or match.group(3)
            return month, year
    except: pass
    return "Tự đặt tên nhé", "Tự đặt tên nhé"

def to_ddmmyyyy(date_val):
    try:
        if pd.isnull(date_val): return ""
        if isinstance(date_val, (pd.Timestamp, pd.DatetimeTZDtype)):
            return date_val.strftime("%d/%m/%Y")
        if isinstance(date_val, (float, int)):
            return pd.to_datetime(date_val, origin='1899-12-30', unit='D').strftime("%d/%m/%Y")
        if isinstance(date_val, str):
            parsed = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
            if pd.isnull(parsed): parsed = pd.to_datetime(date_val, errors='coerce')
            return parsed.strftime("%d/%m/%Y") if not pd.isnull(parsed) else ""
        return str(date_val)
    except: return ""

def format_name(name):
    clean = re.split(r'[\n\r\t\u00A0\u2003]+', str(name).strip())[0]
    clean = re.sub(r'\s+', ' ', clean)
    return clean.replace("-", "").title()

def classify_department(value, content_value=None):
    val = str(value).upper()
    if "VACCINE" in val or "VACXIN" in val: return "VACCINE"
    elif "THUỐC" in val: return "THUOC"
    elif "THẺ" in val: return "THE"
    if content_value:
        content_val = str(content_value).upper()
        if "VACCINE" in content_val: return "VACCINE"
        elif "THUỐC" in content_val: return "THUOC"
    return "KCB"

category_info = {
    "KCB": {"ma": "KHACHLE01", "ten": "Khách hàng lẻ - Khám chữa bệnh"},
    "THUOC": {"ma": "KHACHLE02", "ten": "Khách hàng lẻ - Bán thuốc"},
    "VACCINE": {"ma": "KHACHLE03", "ten": "Khách hàng lẻ - Tiêm vacxin"},
}

output_columns = [
    "Ngày hạch toán (*)", "Ngày chứng từ (*)", "Số chứng từ (*)",
    "Mã đối tượng", "Tên đối tượng", "Nộp vào TK", "Mở tại ngân hàng",
    "Lý do thu", "Diễn giải lý do thu", "Diễn giải (hạch toán)",
    "TK Nợ (*)", "TK Có (*)", "Số tiền"
]

# ====== XỬ LÝ 1 FILE =======
def process_single_file(uploaded_file, chu_hau_to, prefix):
    logs = []
    data_by_category = {cat: {} for cat in category_info}
    file_name = uploaded_file.name
    thang, nam = extract_month_year_from_filename(file_name)
    has_pos = int(nam) <= 2022 if nam.isdigit() else True

    def gen_so_chung_tu(date_str, category):
        try:
            d, m, y = date_str.split("/")
            return f"NVK{category}{d.zfill(2)}{m.zfill(2)}{y}{chu_hau_to}"
        except:
            return f"NVK_INVALID_{chu_hau_to}"

    xls = pd.ExcelFile(uploaded_file)
    for sheet_name in xls.sheet_names:
        if not sheet_name.replace(".", "", 1).isdigit() and not sheet_name.replace(",", "", 1).isdigit():
            logs.append(f"⏩ Bỏ qua sheet: {sheet_name}")
            continue
        df = xls.parse(sheet_name)
        df.columns = [str(col).strip().upper() for col in df.columns]
        if "KHOA/BỘ PHẬN" not in df.columns or "TRẢ THẺ" not in df.columns:
            logs.append(f"⚠️ Thiếu cột ở sheet {sheet_name}")
            continue
        date_column = 'NGÀY QUỸ' if 'NGÀY QUỸ' in df.columns else 'NGÀY KHÁM'
        if date_column not in df.columns:
            logs.append(f"⚠️ Không có cột ngày: {date_column}")
            continue

        df["TRẢ THẺ"] = pd.to_numeric(df["TRẢ THẺ"], errors="coerce")
        df = df[df["TRẢ THẺ"].notna() & (df["TRẢ THẺ"] != 0)]
        df = df[df[date_column].notna() & (df[date_column] != "-")]
        df = df[df["HỌ VÀ TÊN"].notna() & (df["HỌ VÀ TÊN"] != "-")]

        df["CATEGORY"] = df.apply(lambda row: classify_department(row["KHOA/BỘ PHẬN"], row.get("NỘI DUNG THU")), axis=1)

        for category in data_by_category:
            cat_df = df[df["CATEGORY"] == category]
            if cat_df.empty: continue
            for mode in ["PT", "PC"]:
                is_pt = mode == "PT"
                df_mode = cat_df[cat_df["TRẢ THẺ"] > 0] if is_pt else cat_df[cat_df["TRẢ THẺ"] < 0]
                if df_mode.empty: continue
                out_df = pd.DataFrame()
                out_df["Ngày hạch toán (*)"] = df_mode[date_column].apply(to_ddmmyyyy)
                out_df["Ngày chứng từ (*)"] = out_df["Ngày hạch toán (*)"]
                out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(lambda x: gen_so_chung_tu(x, category))
                out_df["Mã đối tượng"] = category_info[category]["ma"]
                out_df["Tên đối tượng"] = df_mode["HỌ VÀ TÊN"].apply(format_name)
                out_df["Nộp vào TK"] = "1290153594"
                out_df["Mở tại ngân hàng"] = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam - Hoàng Mai"
                out_df["Lý do thu"] = ""
                ten_dv = category_info[category]['ten'].split('-')[-1].strip().lower()
                pos_phrase = " qua pos" if has_pos else ""
                out_df["Diễn giải lý do thu"] = ("Thu tiền" if is_pt else "Chi tiền") + f" {ten_dv}{pos_phrase} ngày " + out_df["Ngày chứng từ (*)"]
                out_df["TK Nợ (*)"] = "1368" if has_pos else "1121"
                out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do thu"] + " " + df_mode["HỌ VÀ TÊN"].apply(format_name)
                out_df["TK Có (*)"] = "131"
                out_df["Số tiền"] = df_mode["TRẢ THẺ"].abs().apply(lambda x: f"=VALUE({x})")
                out_df = out_df[output_columns]
                data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

    pt_pc_by_category = {cat: {"PT": defaultdict(list), "PC": defaultdict(list)} for cat in category_info}
    for category, days in data_by_category.items():
        for day, data in days.items():
            for mode in ["PT", "PC"]:
                if mode in data and not data[mode].empty:
                    pt_pc_by_category[category][mode][day].append(data[mode])

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for category in pt_pc_by_category:
            for mode in ["PT", "PC"]:
                day_dict = pt_pc_by_category[category][mode]
                if not day_dict: continue
                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    for day, dfs in sorted(day_dict.items()):
                        merged_df = pd.concat(dfs).reset_index(drop=True)
                        if merged_df.empty: continue
                        merged_df.to_excel(writer, sheet_name=day.strip(), index=False)
                        workbook = writer.book
                        worksheet = writer.sheets[day.strip()]
                        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                        for col_num, col_name in enumerate(merged_df.columns):
                            worksheet.write(0, col_num, col_name, header_format)
                        for i, col in enumerate(merged_df.columns):
                            max_width = max([len(str(col))] + [len(str(v)) for v in merged_df[col].values])
                            worksheet.set_column(i, i, max_width + 2)
                        worksheet.set_tab_color('#92D050')
                output.seek(0)
                file_path = f"{prefix}_{category}/{mode}.xlsx"
                zip_file.writestr(file_path, output.read())

    cleaned_zip = BytesIO()
    with zipfile.ZipFile(zip_buffer, "r") as zin, zipfile.ZipFile(cleaned_zip, "w") as zout:
        for item in zin.infolist():
            if item.filename.endswith(".xlsx"):
                with zin.open(item.filename) as f:
                    wb = load_workbook(f, data_only=False)
                    for sheet in wb.worksheets:
                        headers = [cell.value for cell in sheet[1]]
                        if "Số tiền" in headers:
                            col_idx = headers.index("Số tiền") + 1
                            for row in sheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                                cell = row[0]
                                if cell.data_type == "f" and isinstance(cell.value, str):
                                    match = re.search(r"=VALUE\(([\d.]+)\)", cell.value)
                                    if match:
                                        cell.value = float(match.group(1))
                                        cell.data_type = 'n'
                    temp_output = BytesIO()
                    wb.save(temp_output)
                    temp_output.seek(0)
                    zout.writestr(item.filename, temp_output.read())
            else:
                zout.writestr(item, zin.read(item.filename))

    return cleaned_zip, logs


# ====== GIAO DIỆN TAB 1 =======
with tab1:
    uploaded_files = st.file_uploader("📂 Chọn nhiều file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1, NV123)").strip().upper()

    if st.button("🚀 Tạo File Zip Tổng Hợp") and uploaded_files and chu_hau_to:
        try:
            zip_master = BytesIO()
            logs_all = []
            with zipfile.ZipFile(zip_master, "w") as zip_all:
                for uploaded_file in uploaded_files:
                    file_name = uploaded_file.name
                    thang, nam = extract_month_year_from_filename(file_name)
                    prefix = f"T{thang}_{nam}" if thang != "Tự đặt tên nhé" and nam != "Tự đặt tên nhé" else "TBD"
                    zip_sub, logs = process_single_file(uploaded_file, chu_hau_to, prefix)
                    folder_name = f"{os.path.splitext(file_name)[0]}_{prefix}"
                    with zipfile.ZipFile(zip_sub, "r") as zsub:
                        for item in zsub.infolist():
                            content = zsub.read(item.filename)
                            zip_all.writestr(f"{folder_name}/{item.filename}", content)
                    logs_all.append(f"📄 {file_name}:")
                    logs_all.extend([f"  - {line}" for line in logs])

            st.success("🎉 Tạo file zip tổng hợp thành công!")
            st.download_button("📦 Tải File Zip Tổng", data=zip_master.getvalue(), file_name=f"TongHop_{chu_hau_to}.zip")
            st.markdown("### 📑 Nhật ký xử lý toàn bộ:")
            st.text("\n".join(logs_all))
        except Exception as e:
            st.error("❌ Đã xảy ra lỗi:")
            st.code(traceback.format_exc(), language="python")

# ======= TAB 2: SO SÁNH XOÁ TRÙNG =======
with tab2:
    st.header("🔍 So sánh với File Gốc và Xoá dòng trùng")

    base_file = st.file_uploader("📂 File Gốc (Base - Excel)", type=["xlsx"], key="base_file")
    zip_compare_file = st.file_uploader("📦 File ZIP đầu ra của hệ thống", type=["zip"], key="zip_compare")

    def normalize_name(name):
        try:
            name = str(name).strip().lower()
            name = re.sub(r'\s+', ' ', name)
            return name
        except:
            return str(name)

    def normalize_date(date_val):
        try:
            if pd.isna(date_val) or str(date_val).strip() in ["", "-", "NaT", "NaN"]:
                return None
            if isinstance(date_val, str):
                date_val = pd.to_datetime(date_val, dayfirst=True, errors="coerce")
            return date_val.strftime("%d/%m/%Y") if pd.notna(date_val) else None
        except:
            return None

    def normalize_columns(columns):
        return [
            str(c).strip()
            .replace('\xa0', ' ')
            .replace('\n', ' ')
            .replace('\t', ' ')
            .replace('\r', ' ')
            .strip()
            .title()
        for c in columns
    ]

    def extract_type_from_path(path):
        path = path.upper()
        if "KCB" in path:
            return "Khám chữa bệnh"
        elif "THUOC" in path:
            return "Thuốc"
        elif "VACCINE" in path:
            return "Vaccine"
        elif "THE" in path:
            return "Thẻ"
        return "Khác"

    if st.button("🚫 Xoá dòng trùng theo Tên + Ngày + Số Tiền"):
        if base_file and zip_compare_file:
            try:
                base_df = pd.read_excel(base_file)
                base_df.columns = normalize_columns(base_df.columns)

                required_cols = {"Tên Đối Tượng", "Ngày Hạch Toán", "Phát Sinh Nợ", "Số Tiền"}
                missing_cols = required_cols - set(base_df.columns)

                if missing_cols:
                    st.error(f"""❌ File gốc **{base_file.name}** thiếu cột: {', '.join(missing_cols)}
🔍 Các cột hiện có: {', '.join(base_df.columns)}""")
                    st.stop()

                base_df["Tên chuẩn"] = base_df["Tên Đối Tượng"].apply(normalize_name)
                base_df["Ngày chuẩn"] = base_df["Ngày Hạch Toán"].apply(normalize_date)
                base_df = base_df[base_df["Tên chuẩn"].notna() & base_df["Ngày chuẩn"].notna()]
                base_lookup = base_df.set_index(["Tên chuẩn", "Ngày chuẩn", "Số Tiền"])["Phát Sinh Nợ"].to_dict()

                base_pairs = set(base_lookup.keys())

                zip_in = zipfile.ZipFile(zip_compare_file, 'r')
                zip_namelist = [fn for fn in zip_in.namelist() if fn.lower().endswith(".xlsx")]
                total_files = len(zip_namelist)
                zip_buffer = BytesIO()

                progress = st.progress(0, text="🚧 Đang xử lý ZIP...")
                logs = []
                total_removed = 0
                matched_rows_summary = []

                with zipfile.ZipFile(zip_buffer, "w") as zip_out:
                    for idx, file_name in enumerate(zip_namelist):
                        with zip_in.open(file_name) as f:
                            xls = pd.ExcelFile(f)
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                for sheet in xls.sheet_names:
                                    df = pd.read_excel(xls, sheet_name=sheet)
                                    df.columns = normalize_columns(df.columns)

                                    if "Tên Đối Tượng" in df.columns and "Ngày Hạch Toán (*)" in df.columns and "Số Tiền" in df.columns:
                                        df["Tên chuẩn"] = df["Tên Đối Tượng"].apply(normalize_name)
                                        df["Ngày chuẩn"] = df["Ngày Hạch Toán (*)"].apply(normalize_date)
                                        df["Số Tiền chuẩn"] = df["Số Tiền"].apply(pd.to_numeric, errors='coerce')
                                        df["STT Gốc"] = df.index

                                        df["Trạng thái"] = df.apply(
                                            lambda row: "Trùng hoàn toàn" if (row["Tên chuẩn"], row["Ngày chuẩn"], row["Số Tiền chuẩn"]) in base_pairs else "Không trùng",
                                            axis=1
                                        )

                                        matched = df[df["Trạng thái"] == "Trùng hoàn toàn"]
                                        removed = len(matched)
                                        total_removed += removed

                                        if not matched.empty:
                                            temp_matched = matched.copy()
                                            temp_matched["Loại"] = extract_type_from_path(file_name)
                                            temp_matched["Sheet"] = sheet
                                            temp_matched["Phát Sinh Nợ (File Gốc)"] = temp_matched.apply(
                                                lambda row: base_lookup.get((row["Tên chuẩn"], row["Ngày chuẩn"], row["Số Tiền chuẩn"])), axis=1
                                            )
                                            matched_rows_summary.append(
                                                temp_matched[[
                                                    "Loại", "Sheet", "STT Gốc", "Tên Đối Tượng",
                                                    "Ngày Hạch Toán (*)", "Số Tiền", "Phát Sinh Nợ (File Gốc)"
                                                ]]
                                            )
                                            logs.append(f"- 📄 `{file_name}` | Sheet: `{sheet}` 👉 Đã xoá {removed} dòng")

                                        df = df[df["Trạng thái"] != "Trùng hoàn toàn"]
                                        df.drop(columns=["Tên chuẩn", "Ngày chuẩn", "Trạng thái", "Số Tiền chuẩn"], inplace=True)

                                    df.to_excel(writer, sheet_name=sheet, index=False)

                                    workbook = writer.book
                                    worksheet = writer.sheets[sheet]
                                    header_format = workbook.add_format({
                                        'bold': True, 'bg_color': '#FFE699', 'border': 1
                                    })

                                    for col_num, col_name in enumerate(df.columns):
                                        worksheet.write(0, col_num, col_name, header_format)
                                        max_width = max([len(str(col_name))] + [len(str(v)) for v in df[col_name]])
                                        worksheet.set_column(col_num, col_num, max_width + 2)

                                    worksheet.set_tab_color("#FFC000")

                            output.seek(0)
                            zip_out.writestr(file_name, output.read())

                        progress.progress((idx + 1) / total_files, text=f"✅ Đã xử lý {idx + 1}/{total_files} file")

                st.session_state["matched_rows_summary"] = matched_rows_summary
                st.session_state["logs"] = logs
                st.session_state["zip_buffer"] = zip_buffer.getvalue()
                st.session_state["zip_ready"] = True

                st.success(f"🎉 Đã xoá tổng cộng {total_removed} dòng trùng trong {total_files} file Excel.")

            except Exception as e:
                st.error("❌ Lỗi khi xử lý ZIP:")
                st.code(traceback.format_exc(), language="python")

# 👇 LOG chi tiết
if "logs" in st.session_state:
    st.subheader("📜 Log chi tiết đã xử lý")
    for log in st.session_state["logs"]:
        st.markdown(log)

# 👇 BẢNG preview + bộ lọc
if "matched_rows_summary" in st.session_state and st.session_state["matched_rows_summary"]:
    st.subheader("📊 Dòng trùng đã xoá (Tên + Ngày + Số Tiền):")
    combined_df = pd.concat(st.session_state["matched_rows_summary"], ignore_index=True)

    # Bộ lọc
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        filter_type = st.selectbox("🔍 Lọc theo Loại", ["Tất cả"] + sorted(combined_df["Loại"].unique()))
    with col2:
        filter_sheet = st.selectbox("📄 Lọc theo Sheet", ["Tất cả"] + sorted(combined_df["Sheet"].unique()))
    with col3:
        filter_name = st.text_input("🧍 Lọc theo Tên chứa", "")
    with col4:
        filter_date = st.text_input("📅 Lọc theo Ngày Hạch Toán", "")

    # Áp dụng filter
    filtered_df = combined_df.copy()
    if filter_type != "Tất cả":
        filtered_df = filtered_df[filtered_df["Loại"] == filter_type]
    if filter_sheet != "Tất cả":
        filtered_df = filtered_df[filtered_df["Sheet"] == filter_sheet]
    if filter_name.strip():
        filtered_df = filtered_df[filtered_df["Tên Đối Tượng"].str.contains(filter_name.strip(), case=False, na=False)]
    if filter_date.strip():
        filtered_df = filtered_df[filtered_df["Ngày Hạch Toán (*)"].astype(str).str.contains(filter_date.strip())]

    st.dataframe(filtered_df)

# 👇 Button tải file
if "zip_buffer" in st.session_state and st.session_state["zip_ready"]:
    st.download_button(
        "📥 Tải file ZIP đã xoá dòng trùng",
        data=st.session_state["zip_buffer"],
        file_name="output_cleaned.zip"
    )

with tab3:
    st.header("📊 Gộp Dữ Liệu Tháng Thành 1 File Excel Tổng Hợp")
    zip_input = st.file_uploader("📂 Tải lên file Zip đầu ra từ Tab 1", type=["zip"], key="zip_monthly")

    if zip_input:
        try:
            group_data = {
                "PT_KCB": [], "PC_KCB": [],
                "PT_THUOC": [], "PC_THUOC": [],
                "PT_VACCINE": [], "PC_VACCINE": []
            }

            # 🧠 Lấy tên tháng & năm từ file zip nếu có thể
            match = re.search(r't(\d{1,2})[_\-\.](\d{4})', zip_input.name.lower())
            if match:
                thang_text = f"T{int(match.group(1))}"
                nam_text = match.group(2)
            else:
                thang_text = "TBD"
                nam_text = "XXXX"

            with zipfile.ZipFile(zip_input, "r") as zipf:
                for filename in zipf.namelist():
                    if not filename.endswith(".xlsx"):
                        continue

                    with zipf.open(filename) as f:
                        xls = pd.ExcelFile(f)
                        for sheet_name in xls.sheet_names:
                            if sheet_name.startswith("PT") or sheet_name.startswith("PC"):
                                df = xls.parse(sheet_name)
                                if not set(["Ngày chứng từ (*)", "Tên đối tượng", "Số tiền"]).issubset(df.columns):
                                    continue

                                short_type = None
                                if "KCB" in filename.upper():
                                    short_type = "KCB"
                                elif "THUOC" in filename.upper():
                                    short_type = "THUOC"
                                elif "VACCINE" in filename.upper():
                                    short_type = "VACCINE"
                                else:
                                    continue

                                mode = "PT" if sheet_name.startswith("PT") else "PC"
                                key = f"{mode}_{short_type}"

                                df_filtered = df[["Ngày chứng từ (*)", "Tên đối tượng", "Số tiền"]].copy()
                                group_data[key].append(df_filtered)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                for key, df_list in group_data.items():
                    if not df_list:
                        continue
                    merged_df = pd.concat(df_list, ignore_index=True)
                    merged_df.columns = ["Ngày", "Tên", "Số tiền"]

                    # Thêm công thức cột Ghi chú
                    merged_df["Ghi chú"] = ""

                    merged_df.to_excel(writer, sheet_name=key, index=False, startrow=0, header=True)

                    workbook = writer.book
                    worksheet = writer.sheets[key]

                    # Format header
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1})
                    for col_num, value in enumerate(merged_df.columns):
                        worksheet.write(0, col_num, value, header_format)
                        max_width = max(len(str(value)), *(merged_df.iloc[:, col_num].astype(str).map(len)))
                        worksheet.set_column(col_num, col_num, max_width + 2)

                    # Viết công thức Ghi chú
                    for row_num in range(1, len(merged_df)+1):
                        formula = f'=IF(COUNTIFS(A:A,A{row_num+1},B:B,B{row_num+1},C:C,C{row_num+1})>1,"Lặp","")'
                        worksheet.write_formula(row_num, 3, formula)

                    worksheet.set_tab_color("#FFD966")

            file_name_out = f"TongHop_{thang_text}_{nam_text}.xlsx"
            st.success(f"🎉 Đã gộp xong dữ liệu tháng {thang_text}/{nam_text}!")
            st.download_button("📥 Tải File Tổng Hợp", data=output.getvalue(), file_name=file_name_out)

        except Exception as e:
            st.error("❌ Lỗi khi xử lý file Zip:")
            st.code(traceback.format_exc(), language="python")

with tab4:
    st.subheader("📑 So sánh 'Số tiền' giữa nhiều file Excel")

    uploaded_excels = st.file_uploader(
        "📂 Chọn nhiều file Excel để so sánh", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="multi_excel_compare"
    )

    if uploaded_excels:
        try:
            all_records = []

            for file in uploaded_excels:
                xl = pd.ExcelFile(file)
                for sheet in xl.sheet_names:
                    df = xl.parse(sheet)
                    df.columns = [str(c).strip() for c in df.columns]

                    cols_lower = [c.lower() for c in df.columns]
                    required = {"số tiền", "số chứng từ (*)", "ngày chứng từ (*)"}
                    if not required.issubset(set(cols_lower)):
                        continue

                    ten_col = next((c for c in df.columns if c.strip().lower() in ["họ và tên", "tên đối tượng"]), None)
                    if not ten_col: continue

                    df["TÊN FILE"] = file.name
                    df["TÊN SHEET"] = sheet
                    df["KEY"] = df[ten_col].astype(str).str.strip() + "_" + df["Số chứng từ (*)"].astype(str)

                    df["SỐ TIỀN GỐC"] = (
                        df["Số tiền"]
                        .astype(str)
                        .str.replace("=VALUE(", "", regex=False)
                        .str.replace(")", "", regex=False)
                        .astype(float)
                    )

                    all_records.append(df[["KEY", "SỐ TIỀN GỐC", "TÊN FILE", "TÊN SHEET"]])

            if not all_records:
                st.warning("⚠️ Không tìm thấy dữ liệu phù hợp để so sánh.")
            else:
                full_df = pd.concat(all_records)
                pivot_df = full_df.pivot_table(
                    index="KEY", 
                    columns="TÊN FILE", 
                    values="SỐ TIỀN GỐC", 
                    aggfunc="first"
                ).reset_index()

                # So sánh: những dòng có sự khác biệt giữa các file
                diff_mask = pivot_df.drop("KEY", axis=1).apply(
                    lambda row: len(set(row.dropna())) > 1, axis=1
                )
                result_df = pivot_df[diff_mask]

                st.markdown(f"""
                ### 📊 Kết quả so sánh 'Số tiền'
                - Tổng dòng dữ liệu: `{len(pivot_df)}`
                - Số dòng khác biệt: `{len(result_df)}`
                """)

                st.dataframe(result_df, use_container_width=True)

                excel_bytes = BytesIO()
                with pd.ExcelWriter(excel_bytes, engine="xlsxwriter") as writer:
                    result_df.to_excel(writer, index=False)
                excel_bytes.seek(0)

                st.download_button(
                    "⬇️ Tải kết quả so sánh (Excel)",
                    data=excel_bytes.getvalue(),
                    file_name="So_sanh_So_tien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("❌ Đã xảy ra lỗi khi xử lý các file Excel:")
            st.code(traceback.format_exc(), language="python")
