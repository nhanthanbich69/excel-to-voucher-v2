import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from io import BytesIO
import traceback
import re

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán Chuẩn từ Excel")
tab1, tab2 = st.tabs(["🧾 Tạo File Hạch Toán", "🔍 So sánh và Xoá dòng trùng"])

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
    
    # ✅ Chuẩn hóa ngày sang DD/MM/YYYY bất kể là datetime, float, string kiểu Mỹ/VN
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
            st.success(f"📥 Đọc thành công file {uploaded_file.name} với {len(xls.sheet_names)} sheet.")
    
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
    
                        df_mode = df_mode.reset_index(drop=True)
    
                        out_df = pd.DataFrame()
                        out_df["Ngày hạch toán (*)"] = df_mode[date_column].apply(to_ddmmyyyy)
                        out_df["Ngày chứng từ (*)"] = out_df["Ngày hạch toán (*)"]
                        out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(lambda x: gen_so_chung_tu(x, category))
                        out_df["Mã đối tượng"] = "KHACHLE01"
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
    
                st.success("🎉 Đã xử lý xong!")
                st.download_button("📦 Tải File Zip", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")
    
            st.markdown("### 📄 Nhật ký xử lý")
            st.markdown("\n".join([f"- {line}" for line in logs]))
    
        except Exception as e:
            st.error("❌ Đã xảy ra lỗi:")
            st.code(traceback.format_exc(), language="python")

# ======= TAB 2: SO SÁNH XOÁ TRÙNG =======
with tab2:
    st.header("🔍 So sánh với File Gốc và Xoá dòng trùng")

    # ---------- UPLOAD ----------
    base_file        = st.file_uploader("📂 File Gốc (Base - Excel)", type=["xlsx"], key="base_file")
    zip_compare_file = st.file_uploader("📦 File ZIP đầu ra của hệ thống", type=["zip"],  key="zip_compare")

    # ---------- UTIL ----------
    def normalize_name(name: str):
        try:
            return re.sub(r'\s+', ' ', str(name).strip().lower())
        except Exception:
            return str(name)

    def normalize_money(val):
        try:
            if pd.isna(val):
                return None
            s = str(val).strip().replace("=VALUE(", "").replace(")", "")
            s = s.replace(".", "").replace(",", "")
            s = re.sub(r"[^\d\-]", "", s)
            return None if s == "" else round(float(s), 0)
        except Exception:
            return None

    # MAP alias → canonical
    ALIAS = {
        "TÊN ĐỐI TƯỢNG" : "Tên Đối Tượng",
        "TEN DOI TUONG" : "Tên Đối Tượng",
        "PHÁT SINH NỢ"  : "Phát Sinh Nợ",
        "PHAT SINH NO"  : "Phát Sinh Nợ",
        "SỐ TIỀN"       : "Số Tiền",
        "SO TIEN"       : "Số Tiền"
    }
    def canonical(col):
        up = str(col).strip().upper()
        return ALIAS.get(up, col.strip().title())

    def canon_cols(df: pd.DataFrame):
        df.columns = [canonical(c) for c in df.columns]
        return df

    def extract_date_from_filename(fn):
        m = re.search(r'(\d{2})\.(\d{2})\.xlsx', fn)
        return f"{m.group(1)}/{m.group(2)}" if m else ""

    def extract_type_from_path(p):
        p = p.upper()
        if "KCB"     in p: return "Khám chữa bệnh"
        if "THUOC"   in p: return "Thuốc"
        if "VACCINE" in p: return "Vaccine"
        if "THE"     in p: return "Thẻ"
        return "Khác"

    # ---------- BUTTON ----------
    if st.button("🚫 Xoá dòng trùng trong ZIP"):
        if not (base_file and zip_compare_file):
            st.warning("⚠️ Hãy tải cả 2 file trước khi bấm.")
        else:
            try:
                # ==== BASE FILE ====
                base_df = canon_cols(pd.read_excel(base_file))
                need = {"Tên Đối Tượng", "Phát Sinh Nợ"}
                miss = need - set(base_df.columns)
                if miss:
                    st.error(f"❌ File gốc thiếu cột: {', '.join(miss)}")
                    st.stop()

                base_df["Tên chuẩn"]  = base_df["Tên Đối Tượng"].apply(normalize_name)
                base_df["Tiền chuẩn"] = base_df["Phát Sinh Nợ"].apply(normalize_money)
                base_pairs = set(zip(base_df["Tên chuẩn"], base_df["Tiền chuẩn"]))

                # ==== ZIP ====
                z_in   = zipfile.ZipFile(zip_compare_file, "r")
                xlsx_files = [fn for fn in z_in.namelist() if fn.lower().endswith(".xlsx")]
                total  = len(xlsx_files)
                z_out_buf = BytesIO()

                logs, removed_total, summary = [], 0, []

                with zipfile.ZipFile(z_out_buf, "w") as z_out:
                    for i, fn in enumerate(xlsx_files, 1):
                        with z_in.open(fn) as raw:
                            xls = pd.ExcelFile(raw)
                            mem = BytesIO()
                            with pd.ExcelWriter(mem, engine="xlsxwriter") as writer:
                                for sh in xls.sheet_names:
                                    df = canon_cols(pd.read_excel(xls, sheet_name=sh))

                                    if {"Tên Đối Tượng", "Số Tiền"} <= set(df.columns):
                                        df["Tên chuẩn"]  = df["Tên Đối Tượng"].apply(normalize_name)
                                        df["Tiền chuẩn"] = df["Số Tiền"].apply(normalize_money)
                                        df["STT Gốc"]    = df.index

                                        dup = df[df[["Tên chuẩn","Tiền chuẩn"]].apply(tuple,axis=1).isin(base_pairs)]
                                        removed_total += len(dup)

                                        if not dup.empty:
                                            dup = dup.assign(
                                                Loại  = extract_type_from_path(fn),
                                                Ngày  = extract_date_from_filename(fn),
                                                Sheet = sh
                                            )
                                            summary.append(
                                                dup[["Loại","Ngày","Sheet","STT Gốc","Tên Đối Tượng","Số Tiền"]]
                                                .merge(base_df[["Tên chuẩn","Phát Sinh Nợ"]],
                                                       left_on="Tên chuẩn", right_on="Tên chuẩn", how="left")
                                                .rename(columns={"Phát Sinh Nợ":"Phát Sinh Nợ"})
                                            )
                                            logs.append(f"- `{fn}` | `{sh}` xoá {len(dup)} dòng")

                                        df = df.drop(dup.index)
                                        df.drop(columns=["Tên chuẩn","Tiền chuẩn"], inplace=True)

                                    # --- save sheet ---
                                    df.to_excel(writer, sheet_name=sh, index=False)
                                    # quick header style
                                    wb = writer.book; ws = writer.sheets[sh]
                                    hdr = wb.add_format({"bold":True,"bg_color":"#FFE699","border":1})
                                    [ws.write(0,c, col,hdr) for c,col in enumerate(df.columns)]

                            mem.seek(0)
                            z_out.writestr(fn, mem.read())

                        st.progress(i/total, text=f"🗃️ Xử lý {i}/{total} file")

                # store session
                st.session_state["zip_ready"]            = True
                st.session_state["zip_buffer"]           = z_out_buf.getvalue()
                st.session_state["matched_rows_summary"] = summary
                st.session_state["logs"]                 = logs

                st.success(f"✅ Đã xoá {removed_total} dòng trùng trong {total} file Excel.")

            except Exception as e:
                st.error("❌ Lỗi khi xử lý:")
                st.code(traceback.format_exc(), language="python")

    # ---------- SHOW RESULT ----------
    if st.session_state.get("zip_ready"):
        st.download_button("📥 Tải ZIP đã xoá trùng",
                           data=st.session_state["zip_buffer"],
                           file_name="sau_xoa_trung.zip")

        if st.session_state["logs"]:
            st.markdown("### 📋 Tóm tắt")
            st.markdown("\n".join(st.session_state["logs"]))

        if st.session_state["matched_rows_summary"]:
            st.markdown("### 🧾 Danh sách chi tiết các dòng đã xoá")
            big_df = pd.concat(st.session_state["matched_rows_summary"], ignore_index=True)
            big_df.sort_values(["Ngày","Loại","Sheet"], inplace=True)

            # format tiền đẹp
            for col in ["Số Tiền","Phát Sinh Nợ"]:
                if col in big_df.columns:
                    big_df[col] = big_df[col].apply(
                        lambda x: f"{int(x):,}" if isinstance(x,(int,float)) and not pd.isna(x) else ""
                    )

            # ---- FILTER UI ----
            c1,c2,c3,c4 = st.columns(4)
            with c1:
                f_loai  = st.selectbox("🗂️ Loại",  ["(Tất cả)"]+sorted(big_df["Loại"].unique()))
            with c2:
                f_ngay  = st.selectbox("📅 Ngày",   ["(Tất cả)"]+sorted(big_df["Ngày"].unique()))
            with c3:
                f_sheet = st.selectbox("📄 Sheet", ["(Tất cả)"]+sorted(big_df["Sheet"].unique()))
            with c4:
                f_ten   = st.selectbox("👤 Tên",   ["(Tất cả)"]+sorted(big_df["Tên Đối Tượng"].unique()))

            filt = big_df.copy()
            if f_loai  != "(Tất cả)": filt = filt[filt["Loại"]==f_loai]
            if f_ngay  != "(Tất cả)": filt = filt[filt["Ngày"]==f_ngay]
            if f_sheet != "(Tất cả)": filt = filt[filt["Sheet"]==f_sheet]
            if f_ten   != "(Tất cả)": filt = filt[filt["Tên Đối Tượng"]==f_ten]

            st.dataframe(
                filt[["Loại","Ngày","Sheet","STT Gốc","Tên Đối Tượng","Số Tiền","Phát Sinh Nợ"]],
                use_container_width=True
            )
