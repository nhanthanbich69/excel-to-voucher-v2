# 📋 Excel → Hạch Toán Voucher (Streamlit)
Ứng dụng Streamlit giúp xử lý file Excel đầu vào và sinh các file Excel hạch toán, đóng gói thành `.zip`.
## ⚙️ Chạy ứng dụng
1. Cài thư viện:
```bash
pip install -r requirements.txt
Chạy:
bash
Copy
Edit
streamlit run app.py
📤 Đầu ra
File zip gồm các folder T07_KCB, T07_THUOC, T07_VACCINE
Mỗi folder chứa file Excel theo từng ngày
Mỗi file có sheet PT và/hoặc PC (nếu có dữ liệu)

✅ Ghi chú
Dòng TIỀN MẶT = 0 hoặc - sẽ bị loại
Nếu không có dữ liệu hợp lệ → báo warning, không crash
