import streamlit as st
import re

def parse_misa_text(misa_text):
    # Tìm tên khách hàng
    match_name = re.search(r"Tên khách hàng:\s*(.+)", misa_text)
    customer = match_name.group(1).strip() if match_name else "Không rõ"

    # Lấy các dòng có phát sinh (số tiền)
    lines = misa_text.strip().splitlines()
    total_cost = 0
    for line in lines:
        try:
            value = float(line.strip().replace('.', '').replace(',', '.'))
            if value > 0:
                total_cost = value
        except:
            pass
    return customer, total_cost

def parse_excel_text(excel_text):
    lines = excel_text.strip().splitlines()
    total_payment = 0
    for line in lines:
        if line.strip() == "":
            continue
        fields = re.split(r'\t+|\s{2,}', line)
        for field in fields:
            try:
                val = float(field.strip().replace(',', '').replace('–', '-'))
                total_payment += val
            except:
                continue
    return total_payment

st.title("🔍 Kiểm tra khớp số liệu MISA vs Excel viện phí")

misa_input = st.text_area("📋 Dán nội dung từ MISA", height=300)
excel_input = st.text_area("📋 Dán nội dung từ bảng Excel viện phí", height=300)

if st.button("🧠 Phân tích"):
    if misa_input and excel_input:
        customer, total_misa = parse_misa_text(misa_input)
        total_excel = parse_excel_text(excel_input)
        chenh_lech = total_excel - total_misa

        st.markdown(f"### 🧑‍⚕️ Khách hàng: **{customer}**")
        st.write(f"📌 Tổng chi phí từ MISA: `{total_misa:,.0f}` đ")
        st.write(f"📌 Tổng thanh toán theo Excel: `{total_excel:,.0f}` đ")
        if abs(chenh_lech) < 1000:
            st.success("✅ Số liệu ĐÃ KHỚP!")
        elif chenh_lech > 0:
            st.warning(f"⚠️ DƯ **{chenh_lech:,.0f}** đ → Khách hàng thanh toán nhiều hơn!")
        else:
            st.error(f"❌ THIẾU **{-chenh_lech:,.0f}** đ → Khách hàng thanh toán chưa đủ!")
    else:
        st.error("Vui lòng dán đủ dữ liệu từ MISA và Excel!")
