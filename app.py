import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback
import re
import os
from tempfile import TemporaryDirectory

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
tab1, tab2 = st.tabs(["📦 Tạo file thô", "🎨 Làm đẹp file Excel"])

# ====================== TAB 1 ======================
with tab1:
    st.header("📦 Bước 1: Tạo file ZIP chưa định dạng")

    uploaded_file = st.file_uploader("📂 Chọn file Excel đầu vào", type=["xlsx"], key="tab1_excel")

    def extract_month_year_from_filename(filename):
        try:
            match = re.search(r'(\d{4})[\.\-_]?\s*(\d{2})|\s*(\d{2})[\.\-_]?\s*(\d{4})', filename)
            if match:
                year = match.group(1) or match.group(4)
                month = match.group(2) or match.group(3)
                return month, year
        except Exception as e:
            st.error(f"❌ Lỗi tách tháng năm: {str(e)}")
        return "Tự đặt tên", "Tự đặt tên"

    if uploaded_file:
        file_name = uploaded_file.name
        thang, nam = extract_month_year_from_filename(file_name)
        st.success(f"✅ Lấy được tháng: {thang} - năm: {nam}")
    else:
        thang, nam = "Tự đặt tên", "Tự đặt tên"

    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1)", key="hau_to").strip().upper()
    prefix = f"T{thang}_{nam}" if thang != "Tự đặt tên" else "TBD"

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

    def format_name(name):
        try:
            return str(name).replace("-", "").strip().title()
        except:
            return str(name)

    def gen_so_chung_tu(date_str, category):
        try:
            d, m, y = date_str.split("/")
            return f"NVK{category}{d.zfill(2)}{m.zfill(2)}{y}{chu_hau_to}"
        except:
            return f"NVK_INVALID_{chu_hau_to}"

    if st.button("🚀 Tạo file ZIP") and uploaded_file and chu_hau_to:
        try:
            xls = pd.ExcelFile(uploaded_file)
            st.success(f"📥 Đọc file `{uploaded_file.name}` với {len(xls.sheet_names)} sheet.")

            data_by_category = {k: {} for k in category_info}
            logs = []

            try:
                has_pos = int(nam) <= 2022
            except:
                has_pos = True

            for sheet_name in xls.sheet_names:
                if not sheet_name.replace(".", "", 1).isdigit() and not sheet_name.replace(",", "", 1).isdigit():
                    logs.append(f"⏩ Bỏ sheet: {sheet_name}")
                    continue

                df = xls.parse(sheet_name)
                df.columns = [str(col).strip().upper() for col in df.columns]

                if "KHOA/BỘ PHẬN" not in df.columns or "TIỀN MẶT" not in df.columns:
                    logs.append(f"⚠️ Sheet `{sheet_name}` thiếu cột bắt buộc.")
                    continue

                date_column = 'NGÀY QUỸ' if 'NGÀY QUỸ' in df.columns else 'NGÀY KHÁM'

                df["TIỀN MẶT"] = pd.to_numeric(df["TIỀN MẶT"], errors="coerce")
                df = df[df["TIỀN MẶT"].notna() & (df["TIỀN MẶT"] != 0)]
                df = df[df["NGÀY KHÁM"].notna() & (df["NGÀY KHÁM"] != "-")]

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

                        out_df = pd.DataFrame()
                        out_df["Ngày hạch toán (*)"] = pd.to_datetime(df_mode[date_column], errors="coerce").dt.strftime("%d/%m/%Y")
                        out_df["Ngày chứng từ (*)"] = pd.to_datetime(df_mode["NGÀY KHÁM"], errors="coerce").dt.strftime("%d/%m/%Y")
                        out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(lambda x: gen_so_chung_tu(x, category))
                        out_df["Mã đối tượng"] = category_info[category]["ma"]
                        out_df["Tên đối tượng"] = df_mode["HỌ VÀ TÊN"].apply(format_name)
                        out_df["Nộp vào TK"] = "1290153594"
                        out_df["Mở tại ngân hàng"] = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam - Hoàng Mai"
                        out_df["Lý do thu"] = ""

                        ten_dv = category_info[category]['ten'].split('-')[-1].strip().lower()
                        pos_phrase = " qua pos" if has_pos else ""
                        out_df["Diễn giải lý do thu"] = (
                            ("Thu tiền" if is_pt else "Chi tiền") +
                            f" {ten_dv}{pos_phrase} ngày " + out_df["Ngày chứng từ (*)"]
                        )
                        out_df["TK Nợ (*)"] = "1368" if has_pos else "1121"
                        out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do thu"] + " " + df_mode["HỌ VÀ TÊN"].apply(format_name)
                        out_df["TK Có (*)"] = "131"
                        out_df["Số tiền"] = df_mode["TIỀN MẶT"].abs().apply(lambda x: f"=VALUE({x})")
                        out_df = out_df.astype(str)[output_columns]

                        data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                        logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

            if all(not v for v in data_by_category.values()):
                st.warning("⚠️ Không có dữ liệu hợp lệ.")
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
                                            sheet_name = mode if idx == 0 else f"{mode} {idx+1}"
                                            chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                            output.seek(0)
                            zip_path = f"{prefix}_{category}/{day}.xlsx"
                            zip_file.writestr(zip_path, output.read())

                st.success("🎉 Tạo file ZIP thành công!")
                st.download_button("⬇️ Tải file ZIP", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")

            st.markdown("### 📄 Nhật ký xử lý")
            st.markdown("\n".join([f"- {line}" for line in logs]))

        except Exception as e:
            st.error("❌ Lỗi khi xử lý:")
            st.code(traceback.format_exc())

# ====================== TAB 2 ======================
with tab2:
    st.header("🎨 Bước 2: Làm đẹp file Excel từ ZIP")

    zip_uploaded = st.file_uploader("📂 Chọn file ZIP đầu ra từ Tab 1", type=["zip"], key="tab2_zip")

    if zip_uploaded:
        try:
            output_zip = BytesIO()
            with TemporaryDirectory() as tmpdir:
                with zipfile.ZipFile(zip_uploaded, "r") as zip_ref:
                    zip_ref.extractall(tmpdir)

                with zipfile.ZipFile(output_zip, "w") as new_zip:
                    for root, _, files in os.walk(tmpdir):
                        for file in files:
                            if not file.endswith(".xlsx"):
                                continue
                            file_path = os.path.join(root, file)
                            rel_path = os.path.relpath(file_path, tmpdir)
                            styled_output = BytesIO()

                            xls = pd.ExcelFile(file_path)
                            with pd.ExcelWriter(styled_output, engine="xlsxwriter") as writer:
                                for sheet_name in xls.sheet_names:
                                    df = xls.parse(sheet_name)
                                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                                    workbook = writer.book
                                    worksheet = writer.sheets[sheet_name]

                                    header_format = workbook.add_format({
                                        'bold': True, 'bg_color': '#D9E1F2', 'border': 1
                                    })

                                    for col_num, col_name in enumerate(df.columns):
                                        worksheet.write(0, col_num, col_name, header_format)

                                    for i, col in enumerate(df.columns):
                                        max_width = max([len(str(col))] + [len(str(v)) for v in df[col]])
                                        worksheet.set_column(i, i, max_width + 2)

                                    worksheet.set_tab_color('#92D050')

                            styled_output.seek(0)
                            new_zip.writestr(rel_path, styled_output.read())

            st.success("✅ Đã làm đẹp toàn bộ file.")
            st.download_button("⬇️ Tải ZIP đã làm đẹp", data=output_zip.getvalue(), file_name="formatted_output.zip")

        except Exception as e:
            st.error("❌ Lỗi làm đẹp file:")
            st.code(traceback.format_exc())
