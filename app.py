import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback
import re
from openpyxl import load_workbook

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán Chuẩn từ Excel")

tab1, _, _ = st.tabs([
    "🧾 Tạo File Hạch Toán", 
    "🔍 So sánh và Xoá dòng trùng",
    "📊 File tuỳ chỉnh (Check thủ công)"
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
        except: pass
        return "Tự đặt tên nhé", "Tự đặt tên nhé"

    def to_ddmmyyyy(date_val):
        try:
            if pd.isnull(date_val):
                return ""
            if isinstance(date_val, pd.Timestamp):
                return date_val.strftime("%d/%m/%Y")
            if isinstance(date_val, (float, int)):
                return pd.to_datetime(date_val, origin='1899-12-30', unit='D').strftime("%d/%m/%Y")
            parsed = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
            if pd.isnull(parsed):
                parsed = pd.to_datetime(date_val, errors='coerce')
            return parsed.strftime("%d/%m/%Y") if not pd.isnull(parsed) else ""
        except: return ""

    if uploaded_file:
        file_name = uploaded_file.name
        thang, nam = extract_month_year_from_filename(file_name)
        if thang != "Tự đặt tên nhé":
            st.success(f"🗓 Lấy tháng: {thang} | năm: {nam} từ tên file")
        else:
            st.warning("⚠️ Không nhận diện được tháng/năm từ tên file.")
    else:
        thang, nam = "Tự đặt tên nhé", "Tự đặt tên nhé"

    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: DA, B1, NV123)").strip().upper()
    prefix = f"T{thang}_{nam}" if thang != "Tự đặt tên nhé" and nam != "Tự đặt tên nhé" else "TBD"

    def classify_department(value, content_value=None):
        val = str(value).upper()
        if "VACCINE" in val: return "VACCINE"
        if "THUỐC" in val: return "THUOC"
        if "THẺ" in val: return "THE"
        if content_value:
            content_val = str(content_value).upper()
            if "VACCINE" in content_val: return "VACCINE"
            if "THUỐC" in content_val: return "THUOC"
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

    def format_name(name):
        try:
            clean = re.sub(r'\s+', ' ', str(name).strip().splitlines()[0])
            return clean.replace("-", "").title()
        except: return str(name)

    if st.button("🚀 Tạo File Zip") and uploaded_file and chu_hau_to:
        try:
            xls = pd.ExcelFile(uploaded_file)
            st.success(f"📥 Đọc {len(xls.sheet_names)} sheet từ {uploaded_file.name}")
            data_by_category = {k: {} for k in category_info}
            logs = []

            has_pos = True
            try: has_pos = int(nam) <= 2022
            except: pass

            for sheet_name in xls.sheet_names:
                if not sheet_name.replace(".", "", 1).isdigit():
                    logs.append(f"⏩ Bỏ qua sheet không hợp lệ: {sheet_name}")
                    continue

                df = xls.parse(sheet_name)
                df.columns = [str(c).strip().upper() for c in df.columns]
                if "KHOA/BỘ PHẬN" not in df or "TRẢ THẺ" not in df:
                    logs.append(f"⚠️ Sheet {sheet_name} thiếu cột cần thiết.")
                    continue

                date_col = "NGÀY QUỸ" if "NGÀY QUỸ" in df else "NGÀY KHÁM"
                df = df[df[date_col].notna() & (df["TRẢ THẺ"].notna())]
                df["TRẢ THẺ"] = pd.to_numeric(df["TRẢ THẺ"], errors="coerce")
                df = df[df["TRẢ THẺ"] != 0]
                df["CATEGORY"] = df.apply(lambda r: classify_department(r["KHOA/BỘ PHẬN"], r.get("NỘI DUNG THU")), axis=1)

                for category in data_by_category:
                    cat_df = df[df["CATEGORY"] == category]
                    if cat_df.empty: continue

                    for mode in ["PT", "PC"]:
                        is_pt = mode == "PT"
                        df_mode = cat_df[cat_df["TRẢ THẺ"] > 0] if is_pt else cat_df[cat_df["TRẢ THẺ"] < 0]
                        if df_mode.empty: continue
                        df_mode = df_mode.reset_index(drop=True)

                        out_df = pd.DataFrame()
                        out_df["Ngày hạch toán (*)"] = df_mode[date_col].apply(to_ddmmyyyy)
                        out_df["Ngày chứng từ (*)"] = out_df["Ngày hạch toán (*)"]
                        out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(
                            lambda x: f'="NVK/{mode}{int(thang)}_"&TEXT(A2,"ddmmyy")&"_{chu_hau_to}"'
                        )
                        out_df["Mã đối tượng"] = category_info[category]["ma"]
                        out_df["Tên đối tượng"] = df_mode["HỌ VÀ TÊN"].apply(format_name)
                        out_df["Nộp vào TK"] = "1290153594"
                        out_df["Mở tại ngân hàng"] = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam - Hoàng Mai"
                        out_df["Lý do thu"] = ""
                        dv = category_info[category]["ten"].split("-")[-1].strip().lower()
                        pos = " qua pos" if has_pos else ""
                        out_df["Diễn giải lý do thu"] = ("Thu tiền" if is_pt else "Chi tiền") + f" {dv}{pos} ngày " + out_df["Ngày chứng từ (*)"]
                        out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do thu"] + " " + out_df["Tên đối tượng"]
                        out_df["TK Nợ (*)"] = "1368" if has_pos else "1121"
                        out_df["TK Có (*)"] = "131"
                        out_df["Số tiền"] = df_mode["TRẢ THẺ"].abs().apply(lambda x: f"=VALUE({x})")
                        out_df = out_df[output_columns].astype(str)

                        data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                        logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

            if all(not d for d in data_by_category.values()):
                st.warning("⚠️ Không có dữ liệu hợp lệ.")
                st.stop()

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for category, days in data_by_category.items():
                    for day, modes in days.items():
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            for mode in ["PT", "PC"]:
                                if mode not in modes: continue
                                full_df = modes[mode]
                                chunks = [full_df[i:i+500] for i in range(0, len(full_df), 500)]
                                for idx, chunk in enumerate(chunks):
                                    sheet_name = mode if idx == 0 else f"{mode} {idx+1}"
                                    chunk.to_excel(writer, index=False, sheet_name=sheet_name)
                                    ws = writer.sheets[sheet_name]
                                    fmt = writer.book.add_format({'bold': True, 'bg_color': '#D9E1F2'})
                                    for col_num, col_name in enumerate(chunk.columns):
                                        ws.write(0, col_num, col_name, fmt)
                                        col_width = max(len(str(col_name)), max(chunk[col_name].astype(str).apply(len)))
                                        ws.set_column(col_num, col_num, col_width + 2)
                                    ws.set_tab_color('#92D050')
                        output.seek(0)
                        zip_path = f"{prefix}_{category}/{day.strip()}.xlsx"
                        zip_file.writestr(zip_path, output.read())

            st.success("🎉 Hoàn tất tạo file!")
            st.download_button("📦 Tải File Zip", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")

            st.markdown("### 📄 Nhật ký xử lý")
            st.markdown("\n".join(f"- {log}" for log in logs))

        except Exception as e:
            st.error("❌ Lỗi xử lý:")
            st.code(traceback.format_exc(), language="python")
