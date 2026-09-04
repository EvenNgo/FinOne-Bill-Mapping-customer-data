import io
import os
import re
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="FinOne Data Mapper Pro", page_icon="📊", layout="wide"
)
st.title("📊 Chuyển Đổi Dữ Liệu Khách Hàng")

# ==============================================================================
# 1. BỘ TỪ KHÓA MEGA & TỪ ĐIỂN HỌ NGƯỜI VIỆT (AI DATA BRAIN)
# ==============================================================================
KEYWORDS = {
    "stt": ["tt", "stt", "số tt", "số thứ tự", "no.", "no", "thứ tự"],
    "name": ["họ và tên", "họ tên", "tên khách hàng", "người đại diện", "chủ hộ", "học sinh", "tên học sinh", "người nộp", "bên a", "kh", "khách", "tên cháu", "tên hv", "học viên", "người mua", "tên đv", "tên"],
    "last_name": ["họ đệm", "họ và đệm", "họ và tên đệm", "họ lót", "họ"],
    "first_name": ["tên gọi", "tên"],
    "phone": ["điện thoại", "sđt", "sdt", "phone", "tel", "di động", "mobile", "liên hệ", "đt cha", "đt mẹ", "hotline", "số đt"],
    "id": ["mã", "id", "định danh", "cccd", "cmnd", "cmt", "mã kh", "số định danh", "mã học sinh", "mã hv", "passport"],
    "group": ["khu vực", "nhóm", "phường", "xã", "tổ", "cụm", "khối", "lớp"]
}

# Từ điển Họ phổ biến để AI nhận diện nếu không có tiêu đề
COMMON_SURNAMES = {"nguyễn", "trần", "lê", "phạm", "hoàng", "huỳnh", "phan", "vũ", "võ", "đặng", "bùi", "đỗ", "hồ", "ngô", "dương", "lý", "đoàn", "chu", "trịnh"}

# ==============================================================================
# 2. HÀM CORE & CACHE TỐC ĐỘ CAO (TỪ CODE CŨ)
# ==============================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names

@st.cache_data(show_spinner=False)
def load_sheet(file_bytes, sheet_name, header=None, nrows=None):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header, nrows=nrows)

def auto_detect_header_row_smart(raw_df):
    best_row, max_matches = 0, 0
    for r_idx in range(min(15, len(raw_df))):
        row_vals = [str(v).lower().strip() for v in raw_df.iloc[r_idx].values if pd.notna(v)]
        row_text = " ".join(row_vals)
        matches = sum(1 for kws in KEYWORDS.values() if any(k in row_text for k in kws))
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx
    return best_row, max_matches

# ==============================================================================
# 3. ĐỘNG CƠ AI CẢM BIẾN DỮ LIỆU (SUPER SENSOR)
# ==============================================================================
def resolve_column_super_sensor(df, target_col, category):
    cols = df.columns.tolist()
    
    # ƯU TIÊN 1: Map đích danh
    if target_col and target_col not in ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"]:
        for c in cols:
            if str(c).strip().lower() == str(target_col).strip().lower(): return c
                
    # ƯU TIÊN 2: Quét Keyword
    if category in KEYWORDS:
        for c in cols:
            c_str = str(c).lower().strip()
            if any(kw == c_str or kw in c_str for kw in KEYWORDS[category]): return c

    # ƯU TIÊN 3: CẢM BIẾN DỮ LIỆU
    best_col = None
    best_score = 0
    
    for c in cols:
        sample_raw = df[c].dropna().astype(str).str.strip().tolist()[:15]
        sample = [v for v in sample_raw if v.lower() not in ["none", "nan", "null", "", "0", "0.0"]]
        if len(sample) < 3: continue
        
        score = 0
        if category == "name":
            for val in sample:
                if re.search(r'\d', val): score -= 5
                words = val.split()
                if 2 <= len(words) <= 5: score += 2
                if len(words) > 6: score -= 3
                if words and words[0].lower() in COMMON_SURNAMES: score += 5
                if val.istitle(): score += 1
                
        elif category == "last_name":
            for val in sample:
                if re.search(r'\d', val): score -= 5
                words = val.split()
                if 1 <= len(words) <= 3: score += 2
                if words and words[0].lower() in COMMON_SURNAMES: score += 5
                
        elif category == "first_name":
            for val in sample:
                if re.search(r'\d', val): score -= 5
                words = val.split()
                if len(words) == 1: score += 4
                if val.istitle(): score += 1
                
        elif category == "phone":
            for val in sample:
                # Tách riêng các cụm số (Để trị case "0901471608 0909754524")
                tokens = re.findall(r'\d+', val)
                found = False
                for t in tokens:
                    if len(t) in [9, 10, 11] and t.startswith(("0", "84", "3", "5", "7", "8", "9")):
                        found = True; break
                # Dự phòng (Trị case có dấu gạch ngang "090-123-4567")
                if not found:
                    d_all = re.sub(r'\D', '', val.split(".")[0])
                    if len(d_all) in [9, 10, 11] and d_all.startswith(("0", "84", "3", "5", "7", "8", "9")):
                        found = True
                        
                if found: score += 5
                else: score -= 2
                
        elif category == "id":
            for val in sample:
                d = re.sub(r'\D', '', val.split(".")[0])
                # Trị case 9 (CMND), 12 (CCCD) và 11 (CCCD bị rụng số 0 ở đầu)
                if len(d) in [9, 11, 12]: score += 5
                else: score -= 2
                
        avg_score = score / len(sample)
        if avg_score > best_score and avg_score > 0.5:
            best_score = avg_score
            best_col = c
            
    return best_col

# ==============================================================================
# 4. CÁC HÀM LÀM SẠCH VÀ CHUẨN HÓA CỦA CODE CŨ
# ==============================================================================
def clean_phone(phone_val):
    if pd.isna(phone_val): return ""
    val_str = str(phone_val).strip()
    if val_str.lower() in ["0", "none", "nan", "null", ""]: return ""
    
    # Biến các ký tự rác thành khoảng trắng
    cleaned = re.sub(r"[/,;\-–—|và+]+", " ", val_str)
    tokens = cleaned.split()
    
    # 1. Quét từng cụm để bóc SĐT chuẩn xác đầu tiên
    for tok in tokens:
        digits = re.sub(r"\D", "", tok)
        if digits.startswith("84") and len(digits) == 11: digits = "0" + digits[2:]
        elif len(digits) == 9 and not digits.startswith("0"): digits = "0" + digits
        
        if len(digits) == 10 and digits.startswith("0"):
            return digits # Lấy được là thoát luôn
            
    # 2. Back-up: Chữa cháy cho case 090 123 4567 (bị phân cách bằng dấu cách)
    digits_all = re.sub(r"\D", "", val_str.split(".")[0])
    if digits_all.startswith("84") and len(digits_all) == 11: digits_all = "0" + digits_all[2:]
    elif len(digits_all) == 9 and not digits_all.startswith("0"): digits_all = "0" + digits_all
    
    if len(digits_all) == 10 and digits_all.startswith("0"):
        return digits_all
        
    return ""

def clean_id_val(raw_val):
    if pd.isna(raw_val): return ""
    val_str = str(raw_val).split(".")[0].strip()
    if val_str.lower() in ["none", "nan", "null", ""]: return ""
    val_clean = re.sub(r"\s+", "", val_str)
    
    # AI tự động nhận biết CCCD bị mất số 0 và chắp vá lại
    if len(val_clean) == 11:
        val_clean = "0" + val_clean
        
    return val_clean

def is_valid_human_name(name):
    if not name or len(str(name).strip()) < 2: return False
    name_clean = str(name).strip()
    if re.match(r"^[\d\s\.\,\-]+$", name_clean): return False
    if name_clean.lower() in ["họ và tên", "họ tên", "tên", "người liên hệ", "tổng cộng", "stt", "nan", "none"]: return False
    return True

# ==============================================================================
# 5. XỬ LÝ LÕI TỪNG SHEET (PHÂN TÁCH V_ROWS VÀ R_ROWS)
# ==============================================================================
def process_sheet_data(s_name, file_bytes, config, file_name=""):
    raw_df = load_sheet(file_bytes, sheet_name=s_name, header=None, nrows=15)
    if raw_df.empty: return [], [], "Trống", True

    auto_h, matches = auto_detect_header_row_smart(raw_df)
    
    # Nếu không thấy bất kỳ Keyword nào (File rác, mất Header) -> Đọc raw data
    if matches == 0 and not config["apply_fixed_header"]:
        df = load_sheet(file_bytes, sheet_name=s_name, header=None).dropna(how="all")
        df.columns = [f"Col_{i}" for i in range(len(df.columns))]
        hdr_text = "AI Cảm Biến tự dò (Không Header)"
    else:
        use_h_idx = config["header_row_idx"] if config["apply_fixed_header"] else auto_h
        df = load_sheet(file_bytes, sheet_name=s_name, header=use_h_idx).dropna(how="all")
        hdr_text = f"Dòng {use_h_idx + 1}"

    # Chạy hệ thống AI Cảm Biến rà soát từng cột
    s_name_col = resolve_column_super_sensor(df, config["map_name"] if config["name_mode"] == "Họ và Tên gộp chung 1 cột" else None, "name")
    s_last_col = resolve_column_super_sensor(df, config["map_last_name"], "last_name")
    s_first_col = resolve_column_super_sensor(df, config["map_first_name"], "first_name")
    s_phone_col = resolve_column_super_sensor(df, config["map_phone"], "phone")
    s_id_col = resolve_column_super_sensor(df, config["map_id"], "id")

    # TỰ ĐỘNG CHỮA CHÁY (Self-Heal): 
    # Nếu chọn chế độ 1 Cột Tên, nhưng AI không thấy Cột Tên gộp nào, nó tự động tìm cột Họ đệm + Cột Tên để ghép lại (Giải quyết trọn vẹn file Trẻ 1.xlsx)
    if config["name_mode"] == "Họ và Tên gộp chung 1 cột" and not s_name_col:
        s_last_col = resolve_column_super_sensor(df, None, "last_name")
        s_first_col = resolve_column_super_sensor(df, None, "first_name")
        
    is_missing_name = False
    if not s_name_col and not (s_last_col and s_first_col):
        is_missing_name = True

    if is_missing_name: 
        return [], [], hdr_text, True

    v_rows, r_rows = [], []
    for _, r in df.iterrows():
        # Xử lý gộp tên
        if s_name_col:
            full_name = str(r.get(s_name_col, "")).strip()
        else:
            p_last = str(r.get(s_last_col, "")).strip() if pd.notna(r.get(s_last_col)) else ""
            p_first = str(r.get(s_first_col, "")).strip() if pd.notna(r.get(s_first_col)) else ""
            if p_last.lower() in ['nan', 'none']: p_last = ""
            if p_first.lower() in ['nan', 'none']: p_first = ""
            full_name = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()

        # Dòng bị lỗi/trống tên -> Đưa vào báo cáo lỗi
        if not is_valid_human_name(full_name):
            r_rows.append({"File Nguồn": file_name, "Sheet": s_name, "Dữ liệu Tên": full_name, "Lý do Loại": "Tên rỗng hoặc chứa số/ký tự sai định dạng"})
            continue
            
        phone_raw = config["fix_phone"] if config["map_phone"] == ">> Nhập giá trị cố định <<" else r.get(s_phone_col, "")
        id_raw = config["fix_id"] if config["map_id"] == ">> Nhập giá trị cố định <<" else r.get(s_id_col, "")
        
        # Xử lý Chiến lược Nhóm KH
        if "Tên từng File" in config["group_strategy"]: grp_val = file_name
        elif "Tên từng Sheet" in config["group_strategy"]: grp_val = s_name
        elif "Một cột trong bảng" in config["group_strategy"]:
            grp_col = resolve_column_super_sensor(df, config["map_group"], "group")
            grp_val = str(r.get(grp_col, config["fix_group"])) if grp_col and pd.notna(r.get(grp_col)) else config["fix_group"]
        else:
            grp_val = config["fix_group"]
            
        if str(grp_val).lower() in ['nan', 'none', '']: grp_val = config["fix_group"]

        v_rows.append({
            "Cơ sở (*)": config["val_coso"],
            "Nhóm KH (*)": grp_val,
            "Tên KH (*)": full_name,
            "Mã định danh": clean_id_val(id_raw),
            "Điện thoại (*)": clean_phone(phone_raw),
            "Email": "", "Loại KH": "", "Địa chỉ/Ghi chú": "", "Tên doanh nghiệp": "",
        })
        
    return v_rows, r_rows, hdr_text, False

# ==============================================================================
# 6. GIAO DIỆN STREAMLIT (NỀN TẢNG CODE CŨ)
# ==============================================================================
uploaded_files = st.file_uploader("1. Tải lên 1 hoặc nhiều file Excel (.xlsx, .xls):", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã nạp thành công **{len(uploaded_files)} file**")

    col_f1, col_f2 = st.columns([1, 2])
    with col_f1: sample_file_name = st.selectbox("🎯 Chọn File mẫu:", options=list(file_map.keys()))
    sample_file_bytes = file_map[sample_file_name]
    sample_sheets = get_sheet_names(sample_file_bytes)

    default_sheets = [s for s in sample_sheets if s.lower() not in ["dulieu", "thongbaoloi"]] or sample_sheets
    with col_f2: selected_sheets = st.multiselect("📋 Danh sách Sheet xử lý:", options=sample_sheets, default=default_sheets)

    st.markdown("---")
    st.write("🔗 **BƯỚC 1: KIỂM SOÁT TIÊU ĐỀ & GHÉP CỘT (AI HỖ TRỢ)**")

    preview_ref_sheet = selected_sheets[0] if selected_sheets else sample_sheets[0]
    raw_preview = load_sheet(sample_file_bytes, sheet_name=preview_ref_sheet, header=None, nrows=15)
    detected_h_row, _ = auto_detect_header_row_smart(raw_preview)

    c_h1, c_h2 = st.columns([1, 2])
    with c_h1: header_choice = st.number_input("📌 Vị trí dòng tiêu đề (Header):", min_value=1, max_value=15, value=int(detected_h_row + 1))
    with c_h2: apply_fixed_header = st.checkbox("Khóa cứng dòng tiêu đề này cho mọi file (Bỏ tích để AI tự do quét từng file)", value=False)

    curr_header_idx = int(header_choice) - 1
    sample_cols_df = load_sheet(sample_file_bytes, sheet_name=preview_ref_sheet, header=curr_header_idx, nrows=5)
    valid_cols = [str(c).strip() for c in sample_cols_df.columns if not str(c).startswith("Unnamed:") and pd.notna(c)]
    dropdown_opts = ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"] + valid_cols

    val_coso = st.text_input("🏢 Tên Cơ sở (Áp dụng tất cả):", placeholder="Ví dụ: Cơ sở A...")

    col_map1, col_map2 = st.columns(2)
    with col_map1:
        grp_strategy = st.radio("🏢 Nhóm Khách hàng:", ["Tên từng File", "Tên từng Sheet", "Nhập tên cố định", "Một cột trong bảng"], index=0)
        f_group, m_group = "", None
        if grp_strategy == "Nhập tên cố định": f_group = st.text_input("✍️ Nhập Nhóm cố định:")
        elif grp_strategy == "Một cột trong bảng": 
            m_group = st.selectbox("Cột Nhóm:", dropdown_opts)
            f_group = st.text_input("Dự phòng nếu ô trống:")

        name_mode = st.radio("👤 Định dạng cột Họ Tên:", ["Tách riêng 2 cột (Họ đệm + Tên)", "Họ và Tên gộp chung 1 cột"], index=1)
        m_name, m_last, m_first = None, None, None
        if name_mode == "Họ và Tên gộp chung 1 cột": m_name = st.selectbox("Cột Họ và Tên:", dropdown_opts)
        else:
            m_last = st.selectbox("Cột Họ đệm:", dropdown_opts)
            m_first = st.selectbox("Cột Tên:", dropdown_opts)

    with col_map2:
        m_phone = st.selectbox("Cột Điện thoại:", dropdown_opts)
        f_phone = st.text_input("✍️ SĐT cố định:") if m_phone == ">> Nhập giá trị cố định <<" else ""
        m_id = st.selectbox("Cột Mã định danh:", dropdown_opts)
        f_id = st.text_input("✍️ Mã cố định:") if m_id == ">> Nhập giá trị cố định <<" else ""
        
        st.info("🧠 **Sensor Đang Hoạt Động:** Nếu các cột bạn chọn ở trên không tồn tại ở các file khác, hệ thống sẽ tự động quét Data (chữ, số) để tìm và ghép cột Tên & Điện thoại chính xác.")

    config = {
        "val_coso": val_coso.strip(), "header_row_idx": curr_header_idx, "apply_fixed_header": apply_fixed_header,
        "group_strategy": grp_strategy, "fix_group": f_group, "map_group": m_group,
        "name_mode": name_mode, "map_name": m_name, "map_last_name": m_last, "map_first_name": m_first,
        "map_phone": m_phone, "fix_phone": f_phone, "map_id": m_id, "fix_id": f_id,
    }

    # ==============================================================================
    # BƯỚC 3: XUẤT FILE HOÀN CHỈNH & ĐỐI SOÁT BÁO LỖI (TỪ CODE CŨ)
    # ==============================================================================
    st.markdown("---")
    st.subheader("BƯỚC 3: XÁC NHẬN VÀ XUẤT TOÀN BỘ FILE")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_FinOne.xlsx")

    if st.button("🚀 GỘP VÀ XUẤT FILE FINONE", type="primary"):
        if not val_coso.strip():
            st.error("🚨 LỖI: Bắt buộc điền 'Tên Cơ sở'.")
            st.stop()

        wb = openpyxl.load_workbook("mau-nhap-lieu-khach-hang.xlsx")
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        total_valid, total_rejected = 0, 0
        summary_stats, all_rejected_rows, bulk_write_data = [], [], []
        sheets_missing_cols = []

        progress_bar = st.progress(0)
        
        for f_idx, (fname, fbytes) in enumerate(file_map.items()):
            base_n = os.path.splitext(fname)[0]
            target_sheets = [s for s in get_sheet_names(fbytes) if s.lower() not in ["dulieu", "thongbaoloi"]]
            if not target_sheets: target_sheets = get_sheet_names(fbytes)

            for s_name in target_sheets:
                v_rows, r_rows, hdr_text, is_missing = process_sheet_data(s_name, fbytes, config, base_n)

                if is_missing: sheets_missing_cols.append(f"{fname} - {s_name}")
                
                bulk_write_data.extend(v_rows)
                total_valid += len(v_rows)
                
                all_rejected_rows.extend(r_rows)
                total_rejected += len(r_rows)

                summary_stats.append({
                    "Tên File": fname, "Tên Sheet": s_name, "Dòng Header": hdr_text,
                    "Số dòng Hợp lệ": len(v_rows), "Số dòng Bị loại": len(r_rows),
                })
            progress_bar.progress((f_idx + 1) / len(file_map))

        # =========================================================
        # CƠ CHẾ GHI CHUẨN CỦA CODE CŨ (CHỐNG MẤT SỐ 0 TUYỆT ĐỐI)
        # =========================================================
        col_mapping = {"Cơ sở (*)": 2, "Nhóm KH (*)": 3, "Tên KH (*)": 4, "Mã định danh": 5, "Điện thoại (*)": 6, "Email": 7, "Loại KH": 8, "Địa chỉ/Ghi chú": 9, "Tên doanh nghiệp": 10}
        
        current_row = 6
        for row_data in bulk_write_data:
            for key, col_idx in col_mapping.items():
                cell_val = row_data.get(key, "")
                safe_val = "" if str(cell_val).strip().lower() in ["", "none", "nan", "null", "0", "0.0"] or pd.isna(cell_val) else str(cell_val).strip()
                
                cell = ws.cell(row=current_row, column=col_idx)
                
                # CHỐT CHẶN BẢO VỆ SỐ ĐIỆN THOẠI CỦA CODE CŨ
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'  # Ép thư viện Python ghi String
                    cell.number_format = '@'  # Ép Excel hiển thị Text
                    cell.value = safe_val
                else:
                    cell.value = safe_val
            current_row += 1

        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        st.success(f"🎉 Đã gộp thành công **{total_valid} khách hàng**!")

        if sheets_missing_cols:
            st.error("🚨 **CẢNH BÁO: AI Cảm biến KHÔNG THỂ tìm thấy Tên khách hàng ở các sheet sau (Dữ liệu bị bỏ qua):**\n\n" + "\n".join(f"- {s}" for s in sheets_missing_cols))

        st.info(f"✅ **Đối soát quân số:** Tổng dòng quét = **{total_valid + total_rejected}** | Hợp lệ = **{total_valid}** | Bị loại = **{total_rejected}**")

        tab_final1, tab_final2 = st.tabs(["📥 Tải File Hoàn Chỉnh (FinOne)", "❌ Báo Cáo Tổng Hợp Dòng Bị Loại"])

        with tab_final1:
            st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)
            st.download_button("📥 TẢI FILE EXCEL KẾT QUẢ", output_buffer, output_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with tab_final2:
            if all_rejected_rows:
                df_all_rej = pd.DataFrame(all_rejected_rows)
                st.dataframe(df_all_rej, use_container_width=True)
                rej_buf = io.BytesIO()
                df_all_rej.to_excel(rej_buf, index=False, sheet_name="Dong_Bi_Loai", engine="openpyxl")
                rej_buf.seek(0)
                st.download_button("📥 TẢI FILE ĐỐI SOÁT DÒNG BỊ LOẠI (.XLSX)", rej_buf, "Tong_Hop_Dong_Bi_Loai.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("🎉 Toàn bộ dữ liệu đều hợp lệ 100%!")
