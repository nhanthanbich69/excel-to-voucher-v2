import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
import traceback
import re

from openpyxl import load_workbook  # <- cần thiết cho xử lý công thức

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán Chuẩn từ Excel")
tab1, tab2, tab3 = st.tabs([
    "🧾 Tạo File Hạch Toán", 
    "🔍 So sánh và Xoá dòng trùng",
    "📊 Gộp Dữ Liệu Tháng Thành 1 File"
])

with tab1:
    uploaded_file = st.file_uploader("📂 Chọn file Excel (.xlsx)", type=["xlsx"])
    
    def extract_month_year_from_filename(filename):
        try:
            match = re.search(r'(\d{4})[\.\-_]?\s*(\d{2})|\s*(\d{2})[\.\-_]?\s*(\d{4})', filename)
            if match:
                year = match.group(1) or match.group(4)
                month = match.group(2) or match.group(3)
                return month, year
            else:
                return "Tự đặt tên nhé", "Tự đặt tên nhé"
        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý tên file: {str(e)}")
            return "Tự đặt tên nhé", "Tự đặt tên nhé"
    
    def to_ddmmyyyy(date_val):
        try:
            if pd.isnull(date_val):
                return ""
            if isinstance(date_val, pd.Timestamp):
                return date_val.strftime("%d/%m/%Y")
            if isinstance(date_val, float) or isinstance(date_val, int):
                return pd.to_datetime(date_val, origin='1899-12-30', unit='D').strftime("%d/%m/%Y")
            if isinstance(date_val, str):
                parsed = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
                if pd.isnull(parsed):
                    parsed = pd.to_datetime(date_val, errors='coerce')
                return parsed.strftime("%d/%m/%Y") if not pd.isnull(parsed) else ""
            return str(date_val)
        except:
            return ""

    if uploaded_file:
        try:
            file_name = uploaded_file.name
            thang, nam = extract_month_year_from_filename(file_name)
            if thang != "Tự đặt tên nhé" and nam != "Tự đặt tên nhé":
                st.success(f"Đã tự động lấy tháng: {thang} và năm: {nam} từ tên file {file_name}")
            else:
                st.error("Không thể xác định tháng và năm từ tên file. Vui lòng kiểm tra lại tên file.")
        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý file tải lên: {str(e)}")
            thang, nam = "Tự đặt tên nhé", "Tự đặt tên nhé"
    else:
        thang, nam = "Tự đặt tên nhé", "Tự đặt tên nhé"

    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1, NV123)").strip().upper()
    prefix = f"T{thang}_{nam}" if thang != "Tự đặt tên nhé" and nam != "Tự đặt tên nhé" else "TBD"

    def classify_department(value, content_value=None):
        try:
            val = str(value).upper()
            if "VACCINE" in val or "VACXIN" in val:
                return "VACCINE"
            elif "THUỐC" in val:
                return "THUOC"
            elif "THẺ" in val:
                return "THE"
            if content_value:
                content_val = str(content_value).upper()
                if "VACCINE" in content_val:
                    return "VACCINE"
                elif "THUỐC" in content_val:
                    return "THUOC"
                elif "THẺ" in content_val:
                    return "THE"
        except:
            pass
        return "KCB"

    category_info = {
        "KCB": {"ma": "KHACHLE01", "ten": "Khách hàng lẻ - Khám chữa bệnh"},
        "THUOC": {"ma": "KHACHLE02", "ten": "Khách hàng lẻ - Bán thuốc"},
        "VACCINE": {"ma": "KHACHLE03", "ten": "Khách hàng lẻ - Tiêm vacxin"},
        "THE": {"ma": "KHACHLE04", "ten": "Khách hàng lẻ - Trả thẻ"}
    }

    output_columns = [
        "Ngày hạch toán (*)", "Ngày chứng từ (*)", "Số chứng từ (*)",
        "Mã đối tượng", "Tên đối tượng", "Nộp vào TK", "Mở tại ngân hàng",
        "Lý do thu", "Diễn giải lý do thu", "Diễn giải (hạch toán)",
        "TK Nợ (*)", "TK Có (*)", "Số tiền"
    ]

    def format_name(name):
        try:
            clean = re.split(r'[\n\r\t\u00A0\u2003]+', str(name).strip())[0]
            clean = re.sub(r'\s+', ' ', clean)
            return clean.replace("-", "").title()
        except:
            return str(name)

    def gen_so_chung_tu(date_str, category):
        try:
            d, m, y = date_str.split("/")
            return f"NVK{category}{d.zfill(2)}{m.zfill(2)}{y}{chu_hau_to}"
        except:
            return f"NVK_INVALID_{chu_hau_to}"

    if st.button("🚀 Tạo File Zip") and uploaded_file and chu_hau_to:
        try:
            xls = pd.ExcelFile(uploaded_file)
            st.success(f"📥 Đã đọc thành công file {uploaded_file.name} với {len(xls.sheet_names)} sheet. Đang xử lý, vui lòng đợi...")
            data_by_category = {k: {} for k in category_info}
            logs = []

            try:
                has_pos = int(nam) <= 2022
            except:
                has_pos = True

            for sheet_name in xls.sheet_names:
                if not sheet_name.replace(".", "", 1).isdigit() and not sheet_name.replace(",", "", 1).isdigit():
                    logs.append(f"⏩ Bỏ qua sheet không hợp lệ: {sheet_name}")
                    continue

                df = xls.parse(sheet_name)
                df.columns = [str(col).strip().upper() for col in df.columns]

                if "KHOA/BỘ PHẬN" not in df.columns or "TIỀN MẶT" not in df.columns:
                    logs.append(f"⚠️ Sheet {sheet_name} thiếu cột cần thiết.")
                    continue

                date_column = 'NGÀY QUỸ' if 'NGÀY QUỸ' in df.columns else 'NGÀY KHÁM'
                if date_column not in df.columns:
                    logs.append(f"⚠️ Sheet {sheet_name} thiếu cột ngày ({date_column})")
                    continue

                df["TIỀN MẶT"] = pd.to_numeric(df["TIỀN MẶT"], errors="coerce")
                df = df[df["TIỀN MẶT"].notna() & (df["TIỀN MẶT"] != 0)]
                df = df[df[date_column].notna() & (df[date_column] != "-")]
                df = df[df["HỌ VÀ TÊN"].notna() & (df["HỌ VÀ TÊN"] != "-")]

                df["CATEGORY"] = df.apply(lambda row: classify_department(row["KHOA/BỘ PHẬN"], row.get("NỘI DUNG THU")), axis=1)

                for category in data_by_category:
                    cat_df = df[df["CATEGORY"] == category]
                    if cat_df.empty:
                        continue

                    for mode in ["PT", "PC"]:
                        is_pt = mode == "PT"
                        df_mode = cat_df[cat_df["TIỀN MẶT"] > 0] if is_pt else cat_df[cat_df["TIỀN MẶT"] < 0]
                        if df_mode.empty:
                            continue

                        df_mode = df_mode.reset_index(drop=True)

                        out_df = pd.DataFrame()
                        out_df["Ngày hạch toán (*)"] = df_mode[date_column].apply(to_ddmmyyyy)
                        out_df["Ngày chứng từ (*)"] = out_df["Ngày hạch toán (*)"]
                        out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(lambda x: gen_so_chung_tu(x, category))
                        out_df["Mã đối tượng"] = category_info[category]["ma"]
                        out_df["Tên đối tượng"] = df_mode["HỌ VÀ TÊN"].apply(format_name)
                        out_df["Nộp vào TK"] = "1290153594"
                        out_df["Mở tại ngân hàng"] = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam - Hoàng Mai"
                        out_df["Lý do thu"] = ""

                        try:
                            ten_dv = category_info[category]['ten'].split('-')[-1].strip().lower()
                            pos_phrase = " qua pos" if has_pos else ""
                            out_df["Diễn giải lý do thu"] = (
                                ("Thu tiền" if is_pt else "Chi tiền") +
                                f" {ten_dv}{pos_phrase} ngày " + out_df["Ngày chứng từ (*)"]
                            )
                            out_df["TK Nợ (*)"] = "1368" if has_pos else "1121"
                        except:
                            out_df["Diễn giải lý do thu"] = ""
                            out_df["TK Nợ (*)"] = ""

                        out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do thu"] + " " + df_mode["HỌ VÀ TÊN"].apply(format_name)
                        out_df["TK Có (*)"] = "131"
                        out_df["Số tiền"] = df_mode["TIỀN MẶT"].abs().apply(lambda x: f"=VALUE({x})")

                        out_df = out_df.astype(str)
                        out_df = out_df[output_columns]

                        data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                        logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

            if all(not sheets for sheets in data_by_category.values()):
                st.warning("⚠️ Không có dữ liệu hợp lệ sau khi lọc.")
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for category, sheets in data_by_category.items():
                        for day, data in sheets.items():
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                for mode in ["PT", "PC"]:
                                    if mode in data and not data[mode].empty:
                                        full_df = data[mode]
                                        chunks = [full_df[i:i+500] for i in range(0, len(full_df), 500)]
                                        for idx, chunk in enumerate(chunks):
                                            sheet_name = mode if idx == 0 else f"{mode} {idx + 1}"
                                            chunk.to_excel(writer, sheet_name=sheet_name, index=False)

                                            workbook = writer.book
                                            worksheet = writer.sheets[sheet_name]
                                            header_format = workbook.add_format({
                                                'bold': True, 'bg_color': '#D9E1F2', 'border': 1
                                            })

                                            for col_num, col_name in enumerate(chunk.columns):
                                                worksheet.write(0, col_num, col_name, header_format)

                                            for i, col in enumerate(chunk.columns):
                                                max_width = max([len(str(col))] + [len(str(v)) for v in chunk[col].values])
                                                worksheet.set_column(i, i, max_width + 2)

                                            worksheet.set_tab_color('#92D050')
                            output.seek(0)
                            zip_path = f"{prefix}_{category}/{day.replace(',', '.').strip()}.xlsx"
                            zip_file.writestr(zip_path, output.read())
                
                # 🔁 Làm sạch công thức =VALUE(...) và tạo file sạch
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

                st.success("🎉 Đã xử lý xong!")
                st.download_button("📦 Tải File Zip Hoàn Chỉnh", data=cleaned_zip.getvalue(), file_name=f"{prefix}.zip")

            st.markdown("### 📄 Nhật ký xử lý")
            st.markdown("\n".join([f"- {line}" for line in logs]))

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

    if st.button("🚫 Xoá dòng trùng theo Tên + Ngày"):
        if base_file and zip_compare_file:
            try:
                base_df = pd.read_excel(base_file)
                base_df.columns = normalize_columns(base_df.columns)

                required_cols = {"Tên Đối Tượng", "Ngày Hạch Toán", "Phát Sinh Nợ"}
                missing_cols = required_cols - set(base_df.columns)

                if missing_cols:
                    st.error(f"""❌ File gốc **{base_file.name}** thiếu cột: {', '.join(missing_cols)}
🔍 Các cột hiện có: {', '.join(base_df.columns)}""")
                    st.stop()

                base_df["Tên chuẩn"] = base_df["Tên Đối Tượng"].apply(normalize_name)
                base_df["Ngày chuẩn"] = base_df["Ngày Hạch Toán"].apply(normalize_date)
                base_df = base_df[base_df["Tên chuẩn"].notna() & base_df["Ngày chuẩn"].notna()]
                base_lookup = base_df.set_index(["Tên chuẩn", "Ngày chuẩn"])["Phát Sinh Nợ"].to_dict()

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
                                        df["STT Gốc"] = df.index

                                        df["Trạng thái"] = df.apply(
                                            lambda row: "Trùng hoàn toàn" if (row["Tên chuẩn"], row["Ngày chuẩn"]) in base_pairs else "Không trùng",
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
                                                lambda row: base_lookup.get((row["Tên chuẩn"], row["Ngày chuẩn"])), axis=1
                                            )
                                            matched_rows_summary.append(
                                                temp_matched[[
                                                    "Loại", "Sheet", "STT Gốc", "Tên Đối Tượng",
                                                    "Ngày Hạch Toán (*)", "Số Tiền", "Phát Sinh Nợ (File Gốc)"
                                                ]]
                                            )
                                            logs.append(f"- 📄 `{file_name}` | Sheet: `{sheet}` 👉 Đã xoá {removed} dòng")

                                        df = df[df["Trạng thái"] != "Trùng hoàn toàn"]
                                        df.drop(columns=["Tên chuẩn", "Ngày chuẩn", "Trạng thái"], inplace=True)

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
    st.subheader("📊 Dòng trùng đã xoá (Tên + Ngày):")
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

            with zipfile.ZipFile(zip_input, "r") as zipf:
                for filename in zipf.namelist():
                    if not filename.endswith(".xlsx"):
                        continue

                    # Đọc nội dung file
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

            # Gộp và ghi ra file
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                for key, df_list in group_data.items():
                    if not df_list:
                        continue
                    merged_df = pd.concat(df_list, ignore_index=True)
                    merged_df.to_excel(writer, sheet_name=key, index=False)

                    # Format
                    workbook = writer.book
                    worksheet = writer.sheets[key]
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1})
                    for col_num, value in enumerate(merged_df.columns):
                        worksheet.write(0, col_num, value, header_format)
                        max_width = max(len(str(value)), *(merged_df.iloc[:, col_num].astype(str).map(len)))
                        worksheet.set_column(col_num, col_num, max_width + 2)
                    worksheet.set_tab_color("#FFD966")

            st.success("🎉 Đã gộp xong dữ liệu toàn tháng!")
            st.download_button("📥 Tải File Tổng Hợp", data=output.getvalue(), file_name="TongHop_Thang.xlsx")

        except Exception as e:
            st.error("❌ Lỗi khi xử lý file Zip:")
            st.code(traceback.format_exc(), language="python")
