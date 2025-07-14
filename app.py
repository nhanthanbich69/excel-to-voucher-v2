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

with tab2:
    st.header("🔍 So sánh với File Gốc và Xoá dòng trùng")

    base_file = st.file_uploader("📂 File Gốc (Excel)", type=["xlsx"], key="base_file")
    zip_compare_file = st.file_uploader("📦 File ZIP từ Tab 1", type=["zip"], key="zip_compare")

    def normalize_text(val):
        try:
            return str(val).strip().lower().replace("  ", " ")
        except:
            return ""

    def normalize_date(date_val):
        try:
            if pd.isna(date_val) or str(date_val).strip() in ["", "-", "NaT", "NaN"]:
                return None
            if isinstance(date_val, str):
                date_val = pd.to_datetime(date_val, dayfirst=True, errors="coerce")
            return date_val.strftime("%d/%m/%Y") if pd.notna(date_val) else None
        except:
            return None

    if st.button("🚫 Xoá dòng trùng (Tên + Ngày + Số Tiền)"):
        if not base_file or not zip_compare_file:
            st.warning("⚠️ Cần chọn đủ cả 2 file để bắt đầu.")
            st.stop()

        try:
            # Load file gốc
            base_df = pd.read_excel(base_file)
            base_df.columns = [str(c).strip().lower() for c in base_df.columns]

            if not all(col in base_df.columns for col in ["tên đối tượng", "ngày hạch toán (*)", "số tiền"]):
                st.error("❌ File gốc thiếu 1 trong 3 cột: Tên đối tượng, Ngày hạch toán (*), Số tiền")
                st.stop()

            base_df["__name__"] = base_df["tên đối tượng"].apply(normalize_text)
            base_df["__date__"] = base_df["ngày hạch toán (*)"].apply(normalize_date)
            base_df["__money__"] = pd.to_numeric(base_df["số tiền"], errors="coerce")
            base_keys = set(zip(base_df["__name__"], base_df["__date__"], base_df["__money__"]))

            zip_in = zipfile.ZipFile(zip_compare_file, 'r')
            zip_names = [f for f in zip_in.namelist() if f.endswith(".xlsx")]
            zip_buffer = BytesIO()

            progress = st.progress(0.0, text="🔄 Đang xử lý...")
            logs, removed_total, summary_rows = [], 0, []

            with zipfile.ZipFile(zip_buffer, "w") as zout:
                for idx, fname in enumerate(zip_names):
                    with zip_in.open(fname) as f:
                        xls = pd.ExcelFile(f)
                        out_buffer = BytesIO()
                        with pd.ExcelWriter(out_buffer, engine="xlsxwriter") as writer:
                            for sheet in xls.sheet_names:
                                df = pd.read_excel(xls, sheet_name=sheet)
                                df.columns = [str(c).strip().lower() for c in df.columns]

                                if not all(col in df.columns for col in ["tên đối tượng", "ngày hạch toán (*)", "số tiền"]):
                                    df.to_excel(writer, sheet_name=sheet, index=False)
                                    continue

                                df["__name__"] = df["tên đối tượng"].apply(normalize_text)
                                df["__date__"] = df["ngày hạch toán (*)"].apply(normalize_date)
                                df["__money__"] = pd.to_numeric(df["số tiền"], errors="coerce")

                                df["__key__"] = list(zip(df["__name__"], df["__date__"], df["__money__"]))
                                df["__dup__"] = df["__key__"].apply(lambda x: x in base_keys)

                                removed_rows = df[df["__dup__"]].copy()
                                removed_total += len(removed_rows)

                                if not removed_rows.empty:
                                    removed_rows["File"] = fname
                                    removed_rows["Sheet"] = sheet
                                    summary_rows.append(removed_rows[
                                        ["tên đối tượng", "ngày hạch toán (*)", "số tiền", "File", "Sheet"]
                                    ])
                                    logs.append(f"- `{fname}` | Sheet `{sheet}`: ❌ Xoá {len(removed_rows)} dòng")

                                df = df[~df["__dup__"]].drop(columns=["__name__", "__date__", "__money__", "__key__", "__dup__"])
                                df.to_excel(writer, sheet_name=sheet, index=False)

                                # Formatting
                                workbook = writer.book
                                worksheet = writer.sheets[sheet]
                                fmt = workbook.add_format({'bold': True, 'bg_color': '#FFE699', 'border': 1})
                                for i, col in enumerate(df.columns):
                                    worksheet.write(0, i, col, fmt)
                                    worksheet.set_column(i, i, min(25, max(10, df[col].astype(str).str.len().max() + 2)))
                                worksheet.set_tab_color('#FFC000')

                        out_buffer.seek(0)
                        zout.writestr(fname, out_buffer.read())

                    progress.progress((idx + 1) / len(zip_names), text=f"✅ {idx + 1}/{len(zip_names)} files done")

            st.success(f"🎉 Đã xoá tổng cộng {removed_total} dòng trùng trong {len(zip_names)} file.")

            st.session_state["tab2_zip"] = zip_buffer.getvalue()
            st.session_state["tab2_log"] = logs
            st.session_state["tab2_removed"] = pd.concat(summary_rows) if summary_rows else pd.DataFrame()

        except Exception as e:
            st.error("❌ Lỗi trong quá trình xử lý:")
            st.code(traceback.format_exc(), language="python")

    # Log
    if "tab2_log" in st.session_state:
        st.subheader("📄 Nhật ký xử lý")
        for line in st.session_state["tab2_log"]:
            st.markdown(line)

    # Preview
    if "tab2_removed" in st.session_state and not st.session_state["tab2_removed"].empty:
        st.subheader("📊 Dòng đã xoá (Tên + Ngày + Số Tiền)")
        st.dataframe(st.session_state["tab2_removed"], use_container_width=True)

    # Tải file
    if "tab2_zip" in st.session_state:
        st.download_button("📥 Tải ZIP sau khi xoá trùng", data=st.session_state["tab2_zip"], file_name="ket_qua_sau_loc_trung.zip")

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
