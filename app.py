import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán Chuẩn từ Excel (Định dạng mới)")

uploaded_file = st.file_uploader("📂 Chọn file Excel (.xlsx)", type=["xlsx"])

col1, col2, col3 = st.columns(3)
with col1:
    thang = st.selectbox("🗓️ Chọn tháng", [str(i).zfill(2) for i in range(1, 13)])
with col2:
    nam = st.selectbox("📆 Chọn năm", [str(y) for y in range(2020, 2031)])
with col3:
    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1, NV123)").strip().upper()

prefix = f"T{thang}_{nam}"

# Cập nhật hàm phân loại dựa trên "KHOA/BỘ PHẬN" và "NỘI DUNG THU"
def classify_department(value, content_value=None):
    if isinstance(value, str):
        val = value.upper()
        if "VACCINE" in val or "VACXIN" in val:  # Kiểm tra "VACCINE" hoặc "VACXIN"
            return "VACCINE"
        elif "THUỐC" in val:  # Kiểm tra "THUỐC"
            return "THUOC"
        elif "THẺ" in val:  # Kiểm tra "THẺ"
            return "BAN THE"
    # Kiểm tra "NỘI DUNG THU" nếu có cột này
    if content_value and isinstance(content_value, str):
        content_val = content_value.upper()
        if "VACCINE" in content_val or "VACXIN" in content_val:
            return "VACCINE"
        elif "THUỐC" in content_val:
            return "THUOC"
        elif "THẺ" in content_val:
            return "BAN THE"
    return "KCB"  # Nếu không phải là "VACCINE", "THUỐC" hay "THẺ", mặc định là "KCB"

category_info = {
    "KCB":    {"ma": "KHACHLE01", "ten": "Khách hàng lẻ - Khám chữa bệnh"},
    "THUOC":  {"ma": "KHACHLE02", "ten": "Khách hàng lẻ - Bán thuốc"},
    "VACCINE": {"ma": "KHACHLE03", "ten": "Khách hàng lẻ - Vacxin"},
    "BAN THE": {"ma": "KHACHLE04", "ten": "Khách hàng lẻ - Bán thẻ"}  # Phân loại bán thẻ
}

# Danh sách cột mới theo đúng mẫu (33 cột)
output_columns = [
    "Hiển thị trên sổ", "Ngày chứng từ (*)", "Ngày hạch toán (*)", "Số chứng từ (*)",
    "Diễn giải", "Hạn thanh toán", "Diễn giải (Hạch toán)", "TK Nợ (*)", "TK Có (*)", "Số tiền",
    "Đối tượng Nợ", "Đối tượng Có", "TK ngân hàng", "Khoản mục CP", "Đơn vị", "Đối tượng THCP", "Công trình",
    "Hợp đồng bán", "CP không hợp lý", "Mã thống kê", "Diễn giải (Thuế)", "TK thuế GTGT", "Tiền thuế GTGT",
    "% thuế GTGT", "Giá trị HHDV chưa thuế", "Mẫu số HĐ", "Ngày hóa đơn", "Ký hiệu HĐ", "Số hóa đơn",
    "Nhóm HHDV mua vào", "Mã đối tượng thuế", "Tên đối tượng thuế", "Mã số thuế đối tượng thuế"
]

# Hàm xử lý tên theo yêu cầu
def format_name(name):
    # Xoá dấu "-" và chuyển thành Proper Case
    formatted_name = name.replace("-", "").strip().title()
    return formatted_name

if st.button("🚀 Tạo File Zip") and uploaded_file and chu_hau_to:
    try:
        xls = pd.ExcelFile(uploaded_file)
        st.success(f"📥 Đọc thành công file `{uploaded_file.name}` với {len(xls.sheet_names)} sheet.")

        data_by_category = {k: {} for k in category_info}
        logs = []

        for sheet_name in xls.sheet_names:
            if not sheet_name.replace(".", "", 1).isdigit() and not sheet_name.replace(",", "", 1).isdigit():
                logs.append(f"⏩ Bỏ qua sheet không hợp lệ: {sheet_name}")
                continue

            df = xls.parse(sheet_name)
            df.columns = [str(col).strip().upper() for col in df.columns]

            if "KHOA/BỘ PHẬN" not in df.columns or "TIỀN MẶT" not in df.columns:
                logs.append(f"⚠️ Sheet `{sheet_name}` thiếu cột cần thiết.")
                continue

            df["TIỀN MẶT"] = pd.to_numeric(df["TIỀN MẶT"], errors="coerce")
            df = df[df["TIỀN MẶT"].notna() & (df["TIỀN MẶT"] != 0)]

            # Bỏ qua các dòng tổng hợp (subtotal) nếu NGÀY KHÁM không có dữ liệu
            df = df[df["NGÀY KHÁM"].notna() & (df["NGÀY KHÁM"] != "-")]

            # Kiểm tra cả "KHOA/BỘ PHẬN" và "NỘI DUNG THU" (nếu có)
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
                    # Đảm bảo định dạng ngày là mm/dd/yyyy
                    out_df["Ngày hạch toán (*)"] = pd.to_datetime(df_mode["NGÀY QUỸ"], errors="coerce").dt.strftime("%m/%d/%Y")
                    out_df["Ngày chứng từ (*)"] = pd.to_datetime(df_mode["NGÀY KHÁM"], errors="coerce").dt.strftime("%m/%d/%Y")

                    def gen_so_chung_tu(date_str):
                        try:
                            d, m, y = date_str.split("/")
                            return f"{mode}{y}{m.zfill(2)}{d.zfill(2)}_{chu_hau_to}"
                        except:
                            return f"{mode}_INVALID_{chu_hau_to}"

                    out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(gen_so_chung_tu)
                    out_df["Diễn giải"] = ("Thu tiền" if is_pt else "Chi tiền") + f" {category_info[category]['ten'].split('-')[-1].strip().lower()} ngày " + out_df["Ngày chứng từ (*)"]
                    out_df["Diễn giải (Hạch toán)"] = out_df["Diễn giải"] + " - " + df_mode["HỌ VÀ TÊN"].apply(format_name)
                    out_df["TK Nợ (*)"] = "13686A"
                    out_df["TK Có (*)"] = "131"
                    out_df["Số tiền"] = df_mode["TIỀN MẶT"].abs()
                    out_df["Đối tượng Nợ"] = "NCC00002"
                    out_df["Đối tượng Có"] = "KHACHLE01"
                    out_df["TK ngân hàng"] = ""
                    out_df["Hạn thanh toán"] = ""
                    out_df["Khoản mục CP"] = ""
                    out_df["Đơn vị"] = ""
                    out_df["Đối tượng THCP"] = ""
                    out_df["Công trình"] = "003"
                    out_df["Hợp đồng bán"] = ""
                    out_df["CP không hợp lý"] = ""
                    out_df["Mã thống kê"] = ""
                    out_df["Diễn giải (Thuế)"] = ""
                    out_df["TK thuế GTGT"] = ""
                    out_df["Tiền thuế GTGT"] = ""
                    out_df["% thuế GTGT"] = ""
                    out_df["Giá trị HHDV chưa thuế"] = ""
                    out_df["Mẫu số HĐ"] = ""
                    out_df["Ngày hóa đơn"] = ""
                    out_df["Ký hiệu HĐ"] = ""
                    out_df["Số hóa đơn"] = ""
                    out_df["Nhóm HHDV mua vào"] = ""
                    out_df["Mã đối tượng thuế"] = ""
                    out_df["Tên đối tượng thuế"] = ""
                    out_df["Mã số thuế đối tượng thuế"] = ""
                    out_df["Hiển thị trên sổ"] = ""

                    # Chuyển mọi cột về dạng text
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
                        output.seek(0)
                        zip_path = f"{prefix}_{category}/{day.replace(',', '.').strip()}.xlsx"
                        zip_file.writestr(zip_path, output.read())

            st.success("🎉 Đã xử lý xong!")
            st.download_button("📦 Tải File Zip", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")

        st.markdown("### 📄 Nhật ký xử lý")
        st.markdown("\n".join([f"- {line}" for line in logs]))

    except Exception as e:
        st.error("❌ Đã xảy ra lỗi:")
        st.code(traceback.format_exc(), language="python")
