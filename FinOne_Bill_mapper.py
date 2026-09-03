import io
import os
import re
import zipfile
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="FinOne Data Mapper Pro", page_icon="📊", layout="wide"
)
st.title("📊 Chuyển Đổi Dữ Liệu Khách Hàng - FinOne Bill")

# ==============================================================================
# 1. BỘ TỪ KHÓA TỰ ĐỘNG DÒ TÌM
# ==============================================================================
KEYWORDS = {
    "stt": ["tt", "stt", "số tt", "số thứ tự", "no.", "no"],
    "name": ["họ và tên", "họ tên", "tên khách hàng", "người đại diện", "chủ hộ", "học sinh", "tên", "họ"],
    "last_name": ["họ đệm", "họ và đệm", "họ và tên đệm", "họ lót", "họ"],
    "first_name": ["tên gọi", "tên"],
    "phone": ["điện thoại", "sđt", "sdt", "phone", "tel", "di động", "mobile", "liên hệ"],
    "id": ["mã", "id", "định danh", "cccd", "cmnd", "cmt", "mã kh", "số định danh"],
    "email": ["email", "mail", "hòm thư"],
    "address": ["địa chỉ", "address", "nơi ở", "thường trú", "tạm trú"],
    "company": ["công ty", "doanh nghiệp", "tổ chức", "đơn vị", "cơ quan"],
    "group": ["khu vực", "nhóm", "phường", "xã", "tổ", "cụm", "khối", "lớp"],
}

# ==============================================================================
# 2. HÀM CORE & LÀM SẠCH FILE LỖI TRỰC TIẾP TRÊN RAM (CHỐNG CRASH OPENPYXL)
# ==============================================================================
def sanitize_excel_bytes(file_bytes):
    """Bóc tách và gọt bỏ các thẻ Data Validation bị hỏng trực tiếp trong RAM."""
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
def get_sheet_names(file_bytes):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    return pd.ExcelFile(io.BytesIO(safe_bytes)).sheet_names

@st.cache_data(show_spinner=False)
def load_sheet(file_bytes, sheet_name, header=None, nrows=None):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    return pd.read_excel(io.BytesIO(safe_bytes), sheet_name=sheet_name, header=header, nrows=nrows)

def find_column_by_keywords(columns, keyword_list):
    for col in columns:
        col_str = str(col).lower().strip()
        for kw in keyword_list:
            if kw == col_str or kw in col_str:
                return col
    return None

def resolve_column_for_sheet(columns_list, selected_col_name, keyword_category):
    if selected_col_name in ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"] or not selected_col_name:
        return None
    for c in columns_list:
        if str(c).strip().lower() == str(selected_col_name).strip().lower():
            return c
    if keyword_category in KEYWORDS:
        matched = find_column_by_keywords(columns_list, KEYWORDS[keyword_category])
        if matched:
            return matched
    return None

def auto_detect_header_row_smart(raw_df):
    best_row = 0
    max_matches = 0
    for r_idx in range(min(15, len(raw_df))):
        row_vals = [str(v).lower().strip() for v in raw_df.iloc[r_idx].values if pd.notna(v)]
        row_text = " ".join(row_vals)
        matches = sum(1 for kws in KEYWORDS.values() if any(k in row_text for k in kws))
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx
    return best_row

def clean_phone(phone_val):
    if pd.isna(phone_val): return ""
    val_str = str(phone_val).strip()
    if val_str.endswith(".0"): 
        val_str = val_str[:-2]
    if val_str.lower() in ["none", "nan", "null", ""]: 
        return ""
    if val_str == "0": 
        return "0"
        
    digits = re.sub(r"\D", "", val_str)
    if not digits: return val_str
    
    if digits.startswith("84") and len(digits) == 11: 
        digits = "0" + digits[2:]
    elif len(digits) == 9 and not digits.startswith("0"): 
        digits = "0" + digits
    return digits

def safe_str(val, default=""):
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return default
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null") and default == "":
        return ""
    return default if s.lower() in ("nan", "none", "null") else s

# ==============================================================================
# 3. TRÁI TIM NGHIỆP VỤ: XỬ LÝ & LỌC DỮ LIỆU
# ==============================================================================
def process_sheet_data(s_name, file_bytes, config, file_base_name=""):
    if s_name == config["preview_sheet_ref"]:
        h_idx = config["current_header_idx"]
    else:
        raw_preview = load_sheet(file_bytes, sheet_name=s_name, header=None, nrows=15)
        h_idx = auto_detect_header_row_smart(raw_preview)

    df_sheet = load_sheet(file_bytes, sheet_name=s_name, header=h_idx).dropna(how="all")
    if df_sheet.empty:
        return [], [], f"Dòng {h_idx+1}", "Trống", False

    cols_list = df_sheet.columns.tolist()

    s_name_col, s_last_name_col, s_first_name_col = None, None, None
    if config.get("name_mode") == "Tách riêng 2 cột (Họ đệm + Tên)":
        s_last_name_col = resolve_column_for_sheet(cols_list, config.get("map_last_name"), "last_name")
        s_first_name_col = resolve_column_for_sheet(cols_list, config.get("map_first_name"), "first_name")
        if not s_first_name_col and not s_last_name_col:
            rejected_all = []
            for idx, r in df_sheet.iterrows():
                excel_row_num = h_idx + idx + 2
                rej_dict = {
                    "File Nguồn": file_base_name, "Tên Sheet": s_name, "Dòng Excel số": excel_row_num,
                    "Lý do loại bỏ": "KHÔNG XÁC ĐỊNH ĐƯỢC CỘT HỌ / TÊN",
                }
                for k, v in r.items():
                    if not str(k).startswith("Unnamed:") and pd.notna(k):
                        rej_dict[str(k)] = "" if pd.isna(v) else str(v)
                rejected_all.append(rej_dict)
            return [], rejected_all, f"Dòng {h_idx+1}", "❌ KHÔNG TÌM THẤY", True
        col_display_name = f"{s_last_name_col} + {s_first_name_col}"
    else:
        s_name_col = resolve_column_for_sheet(cols_list, config["map_name"], "name")
        if not s_name_col:
            rejected_all = []
            for idx, r in df_sheet.iterrows():
                excel_row_num = h_idx + idx + 2
                rej_dict = {
                    "File Nguồn": file_base_name, "Tên Sheet": s_name, "Dòng Excel số": excel_row_num,
                    "Lý do loại bỏ": "KHÔNG XÁC ĐỊNH ĐƯỢC CỘT TÊN",
                }
                for k, v in r.items():
                    if not str(k).startswith("Unnamed:") and pd.notna(k):
                        rej_dict[str(k)] = "" if pd.isna(v) else str(v)
                rejected_all.append(rej_dict)
            return [], rejected_all, f"Dòng {h_idx+1}", "❌ KHÔNG TÌM THẤY", True
        col_display_name = str(s_name_col)

    s_phone_col = resolve_column_for_sheet(cols_list, config["map_phone"], "phone")
    s_id_col = resolve_column_for_sheet(cols_list, config["map_id"], "id")
    s_email_col = resolve_column_for_sheet(cols_list, config["map_email"], "email")
    s_addr_col = resolve_column_for_sheet(cols_list, config["map_address"], "address")
    s_comp_col = resolve_column_for_sheet(cols_list, config["map_company"], "company")
    s_stt_col = resolve_column_for_sheet(cols_list, config["col_stt_selected"], "stt")
    
    s_group_col = None
    if "Một cột trong bảng" in config["group_strategy"]:
        s_group_col = resolve_column_for_sheet(cols_list, config.get("group_col"), "group")

    valid_rows, rejected_rows = [], []
    garbage_phrases = ["tổng cộng", "tổng số", "tổng tiền", "tổng:", "cộng:", "người lập", "ký tên", "kế toán", "xác nhận", "trưởng phòng", "chữ ký", "đvt:", "đơn vị tính", "bằng chữ:"]

    for idx, r in df_sheet.iterrows():
        excel_row_num = h_idx + idx + 2
        is_val, reason = True, ""

        if config.get("name_mode") == "Tách riêng 2 cột (Họ đệm + Tên)":
            part_last = str(r.get(s_last_name_col, "")).strip() if s_last_name_col and pd.notna(r.get(s_last_name_col)) else ""
            part_first = str(r.get(s_first_name_col, "")).strip() if s_first_name_col and pd.notna(r.get(s_first_name_col)) else ""
            if part_last.lower() in ["nan", "none", "null"]: part_last = ""
            if part_first.lower() in ["nan", "none", "null"]: part_first = ""
            name_str = re.sub(r"\s+", " ", f"{part_last} {part_first}").strip()
        else:
            name_val = r.get(s_name_col, "")
            name_str = str(name_val).strip()

        lower_name = name_str.lower()
        v_phone = config["fix_phone"] if config["map_phone"] == ">> Nhập giá trị cố định <<" else r.get(s_phone_col, "")
        v_id = config["fix_id"] if config["map_id"] == ">> Nhập giá trị cố định <<" else r.get(s_id_col, "")
        v_email = config["fix_email"] if config["map_email"] == ">> Nhập giá trị cố định <<" else r.get(s_email_col, "")
        v_addr = config["fix_address"] if config["map_address"] == ">> Nhập giá trị cố định <<" else r.get(s_addr_col, "")
        v_comp = config["fix_company"] if config["map_company"] == ">> Nhập giá trị cố định <<" else r.get(s_comp_col, "")

        if not name_str or lower_name in ["nan", "none", "null", ""]:
            is_val = False
            reason = "[Loại 3] - Thiếu Tên khách hàng"
        elif any(phrase in lower_name for phrase in garbage_phrases):
            is_val = False
            reason = f"[Loại 2] - Dòng tiêu đề rác / chữ ký ('{name_str}')"
        elif re.search(r"ngày\s*[\.\d]*\s*tháng\s*[\.\d]*\s*năm", lower_name):
            is_val = False
            reason = f"[Loại 2] - Dòng ngày tháng ('{name_str}')"
        elif "Option 1" in config["filter_mode"]:
            if s_stt_col and s_stt_col != "-- Bỏ trống --":
                stt_val = r.get(s_stt_col)
                if pd.isna(stt_val) or str(stt_val).strip() == "":
                    is_val = False
                    reason = "[Loại 1] - Cột STT trống (Tiêu đề nhóm)"
                else:
                    try:
                        if float(str(stt_val).strip()) <= 0:
                            is_val = False
                            reason = "[Loại 1] - STT không hợp lệ (<= 0)"
                    except ValueError:
                        is_val = False
                        reason = f"[Loại 1] - STT không phải số ('{stt_val}')"
        elif "Option 2" in config["filter_mode"] and config["required_second_field"] != "Không (Chỉ cần Họ tên)":
            if "Điện thoại" in config["required_second_field"]:
                if not v_phone or str(v_phone).strip() == "" or str(v_phone).strip().lower() in ["nan", "none"]:
                    is_val = False
                    reason = "Thiếu thông tin bắt buộc: [Điện thoại liên lạc]"
                elif not clean_phone(v_phone):
                    is_val = False
                    reason = f"SĐT không hợp lệ: '{v_phone}'"

        if not is_val:
            if not pd.isna(r.values).all():
                rej_dict = {"File Nguồn": file_base_name, "Tên Sheet": s_name, "Dòng Excel số": excel_row_num, "Lý do loại bỏ": reason}
                for k, v in r.items():
                    if not str(k).startswith("Unnamed:") and pd.notna(k):
                        rej_dict[str(k)] = "" if pd.isna(v) else str(v)
                rejected_rows.append(rej_dict)
        else:
            # Xác định Tên Nhóm KH theo cấu hình linh hoạt
            if "Tên từng File" in config["group_strategy"]:
                grp_val = file_base_name
            elif "Tên từng Sheet" in config["group_strategy"]:
                grp_val = s_name
            elif "Một cột trong bảng" in config["group_strategy"]:
                grp_val = safe_str(r.get(s_group_col), config["fixed_group_val"]) if s_group_col else config["fixed_group_val"]
            else:
                grp_val = config["fixed_group_val"]

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp_val,
                "Tên KH (*)": name_str,
                "Mã định danh": safe_str(v_id),
                "Điện thoại (*)": clean_phone(v_phone),
                "Email": safe_str(v_email),
                "Loại KH": "", 
                "Địa chỉ/Ghi chú": safe_str(v_addr),
                "Tên doanh nghiệp": safe_str(v_comp),
            })

    return valid_rows, rejected_rows, f"Dòng {h_idx+1}", col_display_name, False

# ==============================================================================
# 4. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
uploaded_files = st.file_uploader(
    "1. Tải lên 1 hoặc nhiều file Excel (.xlsx, .xls):",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"📂 Đã nạp **{len(uploaded_files)} file** thành công!")
    
    # Tạo danh bạ file để người dùng tra cứu
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    sample_file_name = list(file_map.keys())[0]
    sample_file_bytes = file_map[sample_file_name]

    all_sheet_names = get_sheet_names(sample_file_bytes)

    st.markdown("---")
    st.subheader("BƯỚC 1: HỆ THỐNG TỰ ĐỘNG DÒ TÌM & GỢI Ý CẤU HÌNH")
    st.caption(f"💡 Đang dùng file **`{sample_file_name}`** làm mẫu cấu hình chuẩn.")

    preview_sheet_ref = all_sheet_names[0]
    raw_preview_df = load_sheet(sample_file_bytes, sheet_name=preview_sheet_ref, header=None, nrows=15)
    detected_header_idx = auto_detect_header_row_smart(raw_preview_df)

    col_hdr1, col_hdr2 = st.columns([1, 3])
    with col_hdr1:
        header_row_excel = st.number_input("📌 Dòng tiêu đề trên Sheet mẫu:", min_value=1, max_value=15, value=int(detected_header_idx + 1))
    with col_hdr2:
        st.info(f"💡 Tất cả các file và sheet được tải lên sẽ tự động đồng bộ theo dòng tiêu đề số **{header_row_excel}**.")

    current_header_idx = int(header_row_excel) - 1
    df_sample = load_sheet(sample_file_bytes, sheet_name=preview_sheet_ref, header=current_header_idx)
    valid_cols = [str(c).strip() for c in df_sample.columns if not str(c).startswith("Unnamed:") and pd.notna(c)]
    
    dropdown_opts = ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"] + valid_cols

    def get_auto_index(category_key):
        matched = find_column_by_keywords(valid_cols, KEYWORDS[category_key])
        return dropdown_opts.index(matched) if matched else 0

    st.markdown("---")
    st.subheader("BƯỚC 2: CHỌN CHẾ ĐỘ LỌC DỮ LIỆU & GHÉP CỘT")

    filter_mode = st.selectbox(
        "Chọn chế độ kiểm tra dữ liệu:",
        [
            "Option 1: Lọc theo cột STT/TT (Khuyên dùng khi file có cột STT 1, 2, 3...)",
            "Option 2: Bắt buộc Họ tên + 1 Trường tùy chọn (Khuyên dùng khi KHÔNG CÓ STT)",
            "Option 3: Lấy tất cả dòng có Tên (Chỉ loại bỏ dòng chữ ký / trống)",
        ],
        index=0 if get_auto_index("stt") != 0 else 1,
    )

    col_stt_selected, required_second_field = None, None
    if "Option 1" in filter_mode:
        col_stt_selected = st.selectbox("👉 Chọn cột Số Thứ Tự (STT/TT):", dropdown_opts, index=get_auto_index("stt"))
    elif "Option 2" in filter_mode:
        col_sub1, col_sub2 = st.columns(2)
        with col_sub1: st.text_input("Trường bắt buộc 1:", value="Tên khách hàng (*)", disabled=True)
        with col_sub2:
            required_second_field = st.selectbox("👉 Trường bắt buộc thứ 2 đi kèm:", ["Điện thoại liên lạc (*)", "Mã định danh", "Địa chỉ/Ghi chú", "Email liên lạc", "Không (Chỉ cần Họ tên)"], index=0)

    st.markdown("---")
    val_coso = st.text_input("🏢 Tên Cơ sở (Áp dụng chung cho tất cả khách hàng trong tất cả các file):", value="", placeholder="Ví dụ: Trường Mầm non Ánh Dương...")

    col_map1, col_map2 = st.columns(2)
    with col_map1:
        group_strategy = st.radio(
            "🏢 Nhóm KH xác định theo:", 
            [
                "Tên từng File (Bỏ đuôi .xlsx - Thích hợp khi mỗi lớp/tòa là 1 file)", 
                "Tên từng Sheet (Thích hợp khi gộp chung trong 1 file)", 
                "Một cột trong bảng", 
                "Nhập tên cố định"
            ],
            index=0 if len(uploaded_files) > 1 else 1
        )
        group_col = st.selectbox("Cột Nhóm:", dropdown_opts) if "Một cột" in group_strategy else None
        fixed_group_val = st.text_input("Tên nhóm chung:", value="Khách hàng chung") if "cố định" in group_strategy else ""
        
        name_mode = st.radio(
            "👤 Định dạng cột Họ và Tên:",
            ["Họ và Tên gộp chung 1 cột", "Tách riêng 2 cột (Họ đệm + Tên)"],
            index=1 if (find_column_by_keywords(valid_cols, KEYWORDS["last_name"]) and find_column_by_keywords(valid_cols, KEYWORDS["first_name"])) else 0
        )

        m_name, m_last_name, m_first_name = None, None, None
        if name_mode == "Tách riêng 2 cột (Họ đệm + Tên)":
            c_name_a, c_name_b = st.columns(2)
            with c_name_a:
                m_last_name = st.selectbox("1a. Cột Họ / Họ đệm (*):", dropdown_opts, index=get_auto_index("last_name"))
            with c_name_b:
                m_first_name = st.selectbox("1b. Cột Tên (*):", dropdown_opts, index=get_auto_index("first_name"))
        else:
            m_name = st.selectbox("1. Cột Tên khách hàng (*):", dropdown_opts, index=get_auto_index("name"))
        
        m_phone = st.selectbox("2. Cột Điện thoại (*):", dropdown_opts, index=get_auto_index("phone"))
        f_phone = st.text_input("✍️ Nhập số điện thoại cố định:") if m_phone == ">> Nhập giá trị cố định <<" else ""

    with col_map2:
        m_id = st.selectbox("3. Cột Mã định danh:", dropdown_opts, index=get_auto_index("id"))
        f_id = st.text_input("✍️ Nhập Mã định danh cố định:") if m_id == ">> Nhập giá trị cố định <<" else ""
        
        m_email = st.selectbox("4. Cột Email liên lạc:", dropdown_opts, index=get_auto_index("email"))
        f_email = st.text_input("✍️ Nhập Email cố định:") if m_email == ">> Nhập giá trị cố định <<" else ""
        
        m_address = st.selectbox("5. Cột Địa chỉ/Ghi chú:", dropdown_opts, index=get_auto_index("address"))
        f_address = st.text_input("✍️ Nhập Địa chỉ cố định:") if m_address == ">> Nhập giá trị cố định <<" else ""
        
        m_company = st.selectbox("6. Cột Tên doanh nghiệp:", dropdown_opts, index=get_auto_index("company"))
        f_company = st.text_input("✍️ Nhập Tên doanh nghiệp cố định:") if m_company == ">> Nhập giá trị cố định <<" else ""

    config = {
        "preview_sheet_ref": preview_sheet_ref,
        "current_header_idx": current_header_idx,
        "filter_mode": filter_mode,
        "col_stt_selected": col_stt_selected,
        "required_second_field": required_second_field,
        "val_coso": val_coso,
        "group_strategy": group_strategy,
        "group_col": group_col,
        "fixed_group_val": fixed_group_val,
        "name_mode": name_mode,
        "map_name": m_name,
        "map_last_name": m_last_name,
        "map_first_name": m_first_name,
        "map_phone": m_phone, "fix_phone": f_phone,
        "map_id": m_id, "fix_id": f_id,
        "map_email": m_email, "fix_email": f_email,
        "map_address": m_address, "fix_address": f_address,
        "map_company": m_company, "fix_company": f_company,
    }

    # ==============================================================================
    # LIVE PREVIEW CHI TIẾT
    # ==============================================================================
    st.markdown("---")
    st.subheader("🔍 LIVE PREVIEW - CHI TIẾT THEO TỪNG FILE")
    
    sel_pv_file = st.selectbox("👉 Chọn File muốn xem trước kết quả:", options=list(file_map.keys()))
    sel_pv_bytes = file_map[sel_pv_file]
    sel_pv_sheets = get_sheet_names(sel_pv_bytes)
    sel_pv_sheet = st.selectbox("👉 Chọn Sheet:", options=sel_pv_sheets)

    base_fname = os.path.splitext(sel_pv_file)[0]
    v_rows, r_rows, _, col_name_found, is_missing = process_sheet_data(sel_pv_sheet, sel_pv_bytes, config, base_fname)

    if is_missing:
        st.error(f"🚨 KHÔNG XÁC ĐỊNH ĐƯỢC CỘT TÊN trên file '{sel_pv_file}' sheet '{sel_pv_sheet}'!")

    col_pv1, col_pv2 = st.columns(2)
    col_pv1.metric("✅ HỢP LỆ (Thành công)", f"{len(v_rows)} dòng")
    col_pv2.metric("❌ TỪ CHỐI (Bị loại)", f"{len(r_rows)} dòng")

    t1, t2 = st.tabs(["✅ Danh sách Hợp Lệ (Mẫu 5 dòng)", f"❌ Chi Tiết {len(r_rows)} Dòng Bị Từ Chối"])
    with t1:
        if v_rows: st.dataframe(pd.DataFrame(v_rows[:5]), use_container_width=True)
        else: st.warning("Chưa có dòng hợp lệ.")
    with t2:
        if r_rows: st.dataframe(pd.DataFrame(r_rows), use_container_width=True)
        else: st.success("Không có dòng bị loại!")

    # ==============================================================================
    # XUẤT FILE HOÀN CHỈNH GỘP TẤT CẢ FILE
    # ==============================================================================
    st.markdown("---")
    st.subheader("BƯỚC 3: XÁC NHẬN VÀ XUẤT TOÀN BỘ FILE")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_FinOne.xlsx")

    if st.button("🚀 XÁC NHẬN & GỘP TẤT CẢ FILE VÀO MẪU FINONE", type="primary"):
        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}]!")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        total_valid, total_rejected = 0, 0
        summary_stats, all_rejected_rows, bulk_write_data = [], [], []

        # Quét lần lượt qua tất cả các File được tải lên
        for fname, fbytes in file_map.items():
            base_n = os.path.splitext(fname)[0]
            try:
                sheet_list = get_sheet_names(fbytes)
            except Exception:
                sheet_list = ["Sheet1"]

            for s_name in sheet_list:
                v_rows, r_rows, hdr_text, col_text, is_missing = process_sheet_data(s_name, fbytes, config, base_n)
                for row_dict in v_rows:
                    bulk_write_data.append(row_dict)
                    total_valid += 1
                total_rejected += len(r_rows)
                all_rejected_rows.extend(r_rows)
                summary_stats.append({
                    "Tên File": fname,
                    "Tên Sheet": s_name, 
                    "Dòng Header": hdr_text, 
                    "Cột Tên": col_text,
                    "Số dòng Hợp lệ": len(v_rows), 
                    "Số dòng Bị loại": len(r_rows),
                })

        col_mapping = {
            "Cơ sở (*)": 2,          # Cột B
            "Nhóm KH (*)": 3,        # Cột C
            "Tên KH (*)": 4,         # Cột D
            "Mã định danh": 5,       # Cột E
            "Điện thoại (*)": 6,     # Cột F
            "Email": 7,              # Cột G
            "Loại KH": 8,            # Cột H
            "Địa chỉ/Ghi chú": 9,    # Cột I
            "Tên doanh nghiệp": 10   # Cột J
        }

        current_row = 6
        for row_data in bulk_write_data:
            for key, col_idx in col_mapping.items():
                cell_val = row_data.get(key, "")
                safe_val = "" if str(cell_val).strip().lower() in ["", "none", "nan", "null"] or (isinstance(cell_val, float) and pd.isna(cell_val)) else str(cell_val).strip()
                
                cell = ws.cell(row=current_row, column=col_idx)
                
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'
                    cell.number_format = '@'
                    cell.value = safe_val
                else:
                    cell.value = safe_val
                    
            current_row += 1

        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        st.success(f"🎉 Đã gộp thành công **{total_valid} khách hàng** từ **{len(uploaded_files)} File**!")
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
