import io
import os
import re
import zipfile
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="FinOne Data Mapper", page_icon="📊", layout="wide"
)
st.title("📊 Chuyển Đổi Dữ Liệu Khách Hàng - FinOne Bill")

# ==============================================================================
# 1. BỘ TỪ KHÓA MEGA (ĐÃ MỞ RỘNG)
# ==============================================================================
KEYWORDS = {
    "stt": ["tt", "stt", "số tt", "số thứ tự", "no.", "no", "thứ tự"],
    "name": ["họ và tên", "họ tên", "tên khách hàng", "người đại diện", "chủ hộ", "học sinh", "tên học sinh", "người nộp", "bên a", "kh", "khách", "tên cháu", "tên hv", "học viên", "người mua", "tên đv"],
    "last_name": ["họ đệm", "họ và đệm", "họ và tên đệm", "họ lót", "họ"],
    "first_name": ["tên gọi", "tên"],
    "phone": ["điện thoại", "sđt", "sdt", "phone", "tel", "di động", "mobile", "liên hệ", "đt cha", "đt mẹ", "hotline", "số đt"],
    "id": ["mã", "id", "định danh", "cccd", "cmnd", "cmt", "mã kh", "số định danh", "mã học sinh", "mã hv", "passport"],
}

# ==============================================================================
# 2. HỆ THỐNG CACHE VÀ LÀM SẠCH BỘ NHỚ RAM
# ==============================================================================
@st.cache_data(show_spinner=False)
def sanitize_excel_bytes(file_bytes):
    try:
        in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), "r")
        out_buf = io.BytesIO()
        out_zip = zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED)
        for item in in_zip.infolist():
            data = in_zip.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                data = re.sub(b"<(?:\w+:)?dataValidations[^>]*>.*?</(?:\w+:)?dataValidations>", b"", data, flags=re.DOTALL)
                data = re.sub(b"<(?:\w+:)?dataValidation[^>]*>.*?</(?:\w+:)?dataValidation>", b"", data, flags=re.DOTALL)
            out_zip.writestr(item, data)
        out_zip.close()
        out_buf.seek(0)
        return out_buf.getvalue()
    except Exception:
        return file_bytes

@st.cache_data(show_spinner=False)
def get_safe_sheet_names(file_bytes):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    return pd.ExcelFile(io.BytesIO(safe_bytes)).sheet_names

@st.cache_data(show_spinner=False)
def load_preview_df(file_bytes, sheet_name, header=None, nrows=20):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    return pd.read_excel(io.BytesIO(safe_bytes), sheet_name=sheet_name, header=header, nrows=nrows)

# ==============================================================================
# 3. ĐỘNG CƠ NHẬN DIỆN THÔNG MINH (HEADER & DATA)
# ==============================================================================
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

def detect_column_by_data_regex(df, col_type):
    """Cảm biến dữ liệu: Đọc thử 15 dòng đầu để đoán cột nếu không tìm thấy bằng Tiêu đề"""
    for col in df.columns:
        valid_count, match_count = 0, 0
        sample_data = df[col].dropna().astype(str).head(15).tolist()
        
        for val in sample_data:
            val_clean = str(val).strip().split(".")[0]
            if not val_clean or val_clean.lower() in ["none", "nan", "null", "0"]: 
                continue
            valid_count += 1
            
            if col_type == "phone":
                digits = re.sub(r"\D", "", val_clean)
                if len(digits) in [9, 10, 11]:
                    match_count += 1
            elif col_type == "id":
                val_no_space = re.sub(r"\s+", "", val_clean)
                if len(val_no_space) in [9, 12] and val_no_space.isdigit():
                    match_count += 1
                    
        if valid_count >= 3 and (match_count / valid_count) >= 0.5:
            return col
    return None

def resolve_column_dynamic(df, target_col_name, keyword_category):
    cols_list = df.columns.tolist()
    
    # 1. Ưu tiên 1: Tên cột được người dùng map đích danh từ bảng mẫu
    if target_col_name and target_col_name not in ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"]:
        for c in cols_list:
            if str(c).strip().lower() == str(target_col_name).strip().lower():
                return c
                
    # 2. Ưu tiên 2: Tìm bằng bộ Keywords mở rộng trên File đang xét
    if keyword_category in KEYWORDS:
        for col in cols_list:
            col_str = str(col).lower().strip()
            if any(kw == col_str or kw in col_str for kw in KEYWORDS[keyword_category]):
                return col
                
    # 3. Ưu tiên 3: Quét trực tiếp Data (Kế hoạch dự phòng cho file mất/sai header)
    if keyword_category in ["phone", "id"]:
        return detect_column_by_data_regex(df, keyword_category)
        
    return None

# ==============================================================================
# 4. CÁC HÀM LÀM SẠCH DỮ LIỆU CÁ NHÂN
# ==============================================================================
def extract_first_phone(phone_raw):
    if not phone_raw or pd.isna(phone_raw): return ""
    val_str = str(phone_raw).strip().split(".")[0]
    if val_str.lower() in ["none", "nan", "null", "0", "0.0", ""]: return ""

    cleaned = re.sub(r"[/,;\-–—|và+]+", " ", val_str)
    for tok in cleaned.split():
        d = re.sub(r"\D", "", tok)
        if d.startswith("84") and len(d) == 11: d = "0" + d[2:]
        elif len(d) == 9 and not d.startswith("0"): d = "0" + d
        if len(d) == 10 and d.startswith("0"): return d

    match = re.search(r"(0\d{9})", re.sub(r"\D", "", val_str))
    return match.group(1) if match else ""

def clean_id_val(raw_val):
    if pd.isna(raw_val): return ""
    val_str = str(raw_val).split(".")[0].strip()
    if val_str.lower() in ["none", "nan", "null", ""]: return ""
    val_clean = re.sub(r"\s+", "", val_str)
    if len(val_clean) == 11 and val_clean.startswith("79"): val_clean = "0" + val_clean
    return val_clean

def is_valid_human_name(name):
    if not name or len(str(name).strip()) < 2: return False
    name_clean = str(name).strip()
    if re.match(r"^[\d\s\.\,\-]+$", name_clean): return False
    if name_clean.lower() in ["họ và tên", "họ tên", "tên", "người liên hệ", "tổng cộng", "stt", "nan", "none"]: return False
    return True

# ==============================================================================
# 5. XỬ LÝ LÕI TỪNG SHEET ĐỘC LẬP
# ==============================================================================
def process_single_sheet(s_name, file_bytes, config, file_base_name=""):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    raw_preview = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=None, nrows=15)
    
    if raw_preview.empty:
        return [], [], "Trống", True

    auto_h, matches = auto_detect_header_row_smart(raw_preview)
    use_h_idx = config["header_row_idx"] if config["apply_fixed_header"] else auto_h
    
    # Nếu matches quá thấp (file ko cấu trúc rõ), ép chạy header giả (dùng index số)
    if matches < 1 and not config["apply_fixed_header"]:
        df_sheet = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=None).dropna(how="all")
        df_sheet.columns = [f"Col_{i}" for i in range(len(df_sheet.columns))]
        hdr_desc = "Không có Header chuẩn (Tự động map bằng Data)"
    else:
        df_sheet = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=use_h_idx).dropna(how="all")
        hdr_desc = f"Dòng {use_h_idx + 1}"

    # Dò tìm động trên sheet hiện tại (Dynamic Mapping)
    s_name_col = resolve_column_dynamic(df_sheet, config["map_name"], "name")
    s_last_col = resolve_column_dynamic(df_sheet, config["map_last_name"], "last_name")
    s_first_col = resolve_column_dynamic(df_sheet, config["map_first_name"], "first_name")
    s_phone_col = resolve_column_dynamic(df_sheet, config["map_phone"], "phone")
    s_id_col = resolve_column_dynamic(df_sheet, config["map_id"], "id")

    v_rows, r_rows = [], []
    is_missing_name = False

    # Kiểm tra kịch bản mất trắng cột Tên
    if config["name_mode"] == "Tách riêng 2 cột (Họ đệm + Tên)":
        if not s_last_col and not s_first_col: is_missing_name = True
    else:
        if not s_name_col: is_missing_name = True

    if is_missing_name:
        return [], [], hdr_desc, True

    for row_idx, r in df_sheet.iterrows():
        # Xử lý Tên
        if config["name_mode"] == "Tách riêng 2 cột (Họ đệm + Tên)":
            p_last = str(r.get(s_last_col, "")).strip() if s_last_col else ""
            p_first = str(r.get(s_first_col, "")).strip() if s_first_col else ""
            if p_last.lower() in ["nan", "none"]: p_last = ""
            if p_first.lower() in ["nan", "none"]: p_first = ""
            full_name = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()
        else:
            full_name = str(r.get(s_name_col, "")).strip() if s_name_col else ""

        # Ghi nhận dòng lỗi
        if not is_valid_human_name(full_name):
            r_rows.append({"File": file_base_name, "Sheet": s_name, "Lý do": "Tên không hợp lệ / Dòng trống", "Dữ liệu Tên": full_name})
            continue

        phone_raw = config["fix_phone"] if config["map_phone"] == ">> Nhập giá trị cố định <<" else r.get(s_phone_col, "")
        id_raw = config["fix_id"] if config["map_id"] == ">> Nhập giá trị cố định <<" else r.get(s_id_col, "")
        grp_val = file_base_name if "Tên file" in config["group_strategy"] else (config["fix_group"] if "cố định" in config["group_strategy"] else s_name)

        v_rows.append({
            "Cơ sở (*)": config["val_coso"],
            "Nhóm KH (*)": grp_val,
            "Tên KH (*)": full_name,
            "Mã định danh": clean_id_val(id_raw),
            "Điện thoại (*)": extract_first_phone(phone_raw),
            "Email": "", "Loại KH": "", "Địa chỉ/Ghi chú": "", "Tên doanh nghiệp": "",
        })

    return v_rows, r_rows, hdr_desc, False

# ==============================================================================
# 6. GIAO DIỆN STREAMLIT 
# ==============================================================================
uploaded_files = st.file_uploader("1. Tải lên 1 hoặc nhiều file Excel (.xlsx, .xls):", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã nạp thành công **{len(uploaded_files)} file**")

    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        sample_file_name = st.selectbox("🎯 Chọn File làm mẫu cấu hình:", options=list(file_map.keys()))
    sample_file_bytes = file_map[sample_file_name]
    sample_sheets = get_safe_sheet_names(sample_file_bytes)

    default_sheets = [s for s in sample_sheets if s.lower() not in ["dulieu", "thongbaoloi"]]
    if not default_sheets: default_sheets = sample_sheets

    with col_f2:
        selected_sheets = st.multiselect("📋 Danh sách Sheet sẽ xử lý trong file này:", options=sample_sheets, default=default_sheets)

    st.markdown("---")
    st.subheader("BƯỚC 1: KIỂM SOÁT TIÊU ĐỀ & ĐỊNH HƯỚNG CỘT")

    preview_ref_sheet = selected_sheets[0] if selected_sheets else sample_sheets[0]
    raw_preview = load_preview_df(sample_file_bytes, preview_ref_sheet, header=None, nrows=15)
    detected_h_row, _ = auto_detect_header_row_smart(raw_preview)

    c_h1, c_h2 = st.columns([1, 2])
    with c_h1:
        header_choice = st.number_input("📌 Vị trí dòng tiêu đề (Header):", min_value=1, max_value=15, value=int(detected_h_row + 1))
    with c_h2:
        apply_fixed_header = st.checkbox("Khóa cứng dòng tiêu đề này cho tất cả file (Bỏ tích để AI tự dò từng file)", value=False)

    curr_header_idx = int(header_choice) - 1
    sample_cols_df = load_preview_df(sample_file_bytes, preview_ref_sheet, header=curr_header_idx, nrows=5)
    valid_cols = [str(c).strip() for c in sample_cols_df.columns if not str(c).startswith("Unnamed:") and pd.notna(c)]
    dropdown_opts = ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"] + valid_cols

    def get_auto_index(cat_key):
        for idx, col in enumerate(dropdown_opts[2:]):
            if any(kw in col.lower() for kw in KEYWORDS.get(cat_key, [])):
                return idx + 2
        return 0

    st.markdown("---")
    st.subheader("BƯỚC 2: KHAI BÁO THÔNG TIN CHUNG & GHÉP CỘT")

    val_coso = st.text_input("🏢 Tên Cơ sở (*) (Bắt buộc theo chuẩn FinOne):", placeholder="Ví dụ: MNTT Bông Sen Hồng...")

    c_map1, c_map2 = st.columns(2)
    with c_map1:
        grp_strategy = st.radio("🏢 Nhóm KH (Lớp/Khu vực) xác định theo:", ["Tên từng File", "Tên từng Sheet", "Nhập tên cố định"])
        f_group = st.text_input("✍️ Nhập tên Nhóm KH cố định:") if "cố định" in grp_strategy else ""
        
        name_mode = st.radio("👤 Định dạng cột Họ và Tên:", ["Tách riêng 2 cột (Họ đệm + Tên)", "Họ và Tên gộp chung 1 cột"], index=1)
        m_name, m_last_name, m_first_name = None, None, None
        if name_mode == "Tách riêng 2 cột (Họ đệm + Tên)":
            c_a, c_b = st.columns(2)
            with c_a: m_last_name = st.selectbox("Cột Họ đệm (*):", dropdown_opts, index=get_auto_index("last_name"))
            with c_b: m_first_name = st.selectbox("Cột Tên (*):", dropdown_opts, index=get_auto_index("first_name"))
        else:
            m_name = st.selectbox("Cột Họ và Tên (*):", dropdown_opts, index=get_auto_index("name"))

    with c_map2:
        st.info("💡 **Gợi ý:** Nếu File khác không có tên cột giống hệt, AI sẽ tự dò theo Data/Từ khóa.")
        m_phone = st.selectbox("Cột Điện thoại liên lạc (*):", dropdown_opts, index=get_auto_index("phone"))
        f_phone = st.text_input("✍️ Nhập SĐT cố định:") if m_phone == ">> Nhập giá trị cố định <<" else ""

        m_id = st.selectbox("Cột Mã định danh / CCCD:", dropdown_opts, index=get_auto_index("id"))
        f_id = st.text_input("✍️ Nhập Mã định danh cố định:") if m_id == ">> Nhập giá trị cố định <<" else ""

    config = {
        "val_coso": val_coso.strip(), "header_row_idx": curr_header_idx, "apply_fixed_header": apply_fixed_header,
        "group_strategy": grp_strategy, "fix_group": f_group,
        "name_mode": name_mode, "map_name": m_name, "map_last_name": m_last_name, "map_first_name": m_first_name,
        "map_phone": m_phone, "fix_phone": f_phone, "map_id": m_id, "fix_id": f_id,
    }

    # ==============================================================================
    # BƯỚC 3 & 4: XUẤT FILE HOÀN CHỈNH & BÁO CÁO LỖI
    # ==============================================================================
    st.markdown("---")
    st.subheader("🚀 BƯỚC 3: XÁC NHẬN & GỘP TOÀN BỘ FILE")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_FinOne.xlsx")

    if st.button("🚀 THỰC THI CHUYỂN ĐỔI", type="primary"):
        if not val_coso.strip():
            st.error("🚨 LỖI: Cần điền 'Tên Cơ sở (*)' để import vào FinOne!")
            st.stop()

        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}]. Vui lòng để chung thư mục code.")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        progress_bar = st.progress(0)
        status_text = st.empty()

        all_valid_records, all_rejected_rows, summary_stats = [], [], []
        sheets_missing_cols = []
        total_files = len(file_map)

        for f_idx, (fname, fbytes) in enumerate(file_map.items()):
            base_n = os.path.splitext(fname)[0]
            sheets = get_safe_sheet_names(fbytes)
            target_sheets = [s for s in sheets if s.lower() not in ["dulieu", "thongbaoloi"]] or sheets

            for s_name in target_sheets:
                status_text.text(f"Đang phân tích AI: '{fname}' ➔ '{s_name}'...")
                v_rows, r_rows, hdr_desc, is_missing = process_single_sheet(s_name, fbytes, config, base_n)
                
                if is_missing:
                    sheets_missing_cols.append(f"File: {fname} - Sheet: {s_name}")
                
                all_valid_records.extend(v_rows)
                all_rejected_rows.extend(r_rows)

                summary_stats.append({
                    "File": fname, "Sheet": s_name, "Header": hdr_desc,
                    "Hợp lệ": len(v_rows), "Bị loại": len(r_rows)
                })

            progress_bar.progress((f_idx + 1) / total_files)

        status_text.text("✅ Đang nạp và khóa định dạng Text an toàn vào mẫu FinOne...")

        col_mapping = {"Cơ sở (*)": 2, "Nhóm KH (*)": 3, "Tên KH (*)": 4, "Mã định danh": 5, "Điện thoại (*)": 6, "Email": 7, "Loại KH": 8, "Địa chỉ/Ghi chú": 9, "Tên doanh nghiệp": 10}
        
        current_row = 6
        for r_dict in all_valid_records:
            for key, col_idx in col_mapping.items():
                cell = ws.cell(row=current_row, column=col_idx)
                val = r_dict.get(key, "")
                
                # KHÓA CỨNG: Ép kiểu String (s) để bảo vệ tuyệt đối số 0 đầu
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'
                    cell.number_format = '@'
                    cell.value = val if val else ""
                else:
                    cell.value = val
            current_row += 1

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        # Giao diện Báo Cáo
        st.success(f"🎉 Xử lý hoàn tất! Đã gom thành công **{len(all_valid_records)} KH**")
        
        if sheets_missing_cols:
            st.error("🚨 **CẢNH BÁO: Các Sheet sau KHÔNG TÌM THẤY CỘT TÊN (toàn bộ dữ liệu bị bỏ qua):**\n" + "\n".join(f"- {s}" for s in sheets_missing_cols))

        tab1, tab2 = st.tabs(["📥 KẾT QUẢ FINONE (Hợp lệ)", "❌ ĐỐI SOÁT (Dòng bị loại)"])

        with tab1:
            st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)
            st.download_button("📥 TẢI FILE EXCEL CHO FINONE", out_buf, output_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with tab2:
            st.metric("Tổng số dòng bị loại (Không có Tên/Sai định dạng)", len(all_rejected_rows))
            if all_rejected_rows:
                df_rej = pd.DataFrame(all_rejected_rows)
                st.dataframe(df_rej, use_container_width=True)
                
                rej_buf = io.BytesIO()
                df_rej.to_excel(rej_buf, index=False, engine="openpyxl")
                rej_buf.seek(0)
                st.download_button("📥 TẢI BÁO CÁO DÒNG LỖI", rej_buf, "Danh_Sach_Loi.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("Tuyệt vời! Không có dòng dữ liệu nào bị loại.")
