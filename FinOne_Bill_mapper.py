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
# 1. BỘ TỪ KHÓA DÒ TÌM CỘT TỰ ĐỘNG
# ==============================================================================
KEYWORDS = {
    "stt": ["tt", "stt", "số tt", "số thứ tự", "no.", "no"],
    "name": ["họ và tên", "họ tên", "tên khách hàng", "người đại diện", "chủ hộ", "học sinh", "tên học sinh", "người nộp"],
    "last_name": ["họ đệm", "họ và đệm", "họ và tên đệm", "họ lót", "họ"],
    "first_name": ["tên gọi", "tên"],
    "phone": ["điện thoại", "sđt", "sdt", "phone", "tel", "di động", "mobile", "liên hệ", "đt cha", "đt mẹ"],
    "id": ["mã", "id", "định danh", "cccd", "cmnd", "cmt", "mã kh", "số định danh", "mã học sinh"],
    "email": ["email", "mail", "hòm thư"],
    "address": ["địa chỉ", "address", "nơi ở", "thường trú", "tạm trú", "tổ", "phường"],
    "company": ["công ty", "doanh nghiệp", "tổ chức", "đơn vị", "cơ quan"],
    "group": ["khu vực", "nhóm", "phường", "xã", "tổ", "cụm", "khối", "lớp"],
}

# ==============================================================================
# 2. CACHE TỐC ĐỘ CAO & LÀM SẠCH FILE TRỰC TIẾP TRÊN RAM
# ==============================================================================
@st.cache_data(show_spinner=False)
def sanitize_excel_bytes(file_bytes):
    """Làm sạch XML một lần duy nhất và đưa vào Cache RAM."""
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
# 3. CÁC HÀM XỬ LÝ CHUẨN HÓA DỮ LIỆU
# ==============================================================================
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
    best_row, max_matches = 0, 0
    for r_idx in range(min(15, len(raw_df))):
        row_vals = [str(v).lower().strip() for v in raw_df.iloc[r_idx].values if pd.notna(v)]
        row_text = " ".join(row_vals)
        matches = sum(1 for kws in KEYWORDS.values() if any(k in row_text for k in kws))
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx
    return best_row, max_matches

def extract_first_phone(phone_raw):
    if not phone_raw or pd.isna(phone_raw):
        return ""
    val_str = str(phone_raw).strip().split(".")[0]
    if val_str.lower() in ["none", "nan", "null", "0", "0.0", ""]:
        return ""

    cleaned = re.sub(r"[/,;\-–—|và+]+", " ", val_str)
    tokens = cleaned.split()

    for tok in tokens:
        d = re.sub(r"\D", "", tok)
        if d.startswith("84") and len(d) == 11:
            d = "0" + d[2:]
        elif len(d) == 9 and not d.startswith("0"):
            d = "0" + d
        if len(d) == 10 and d.startswith("0"):
            return d

    raw_digits = re.sub(r"\D", "", val_str)
    match = re.search(r"(0\d{9})", raw_digits)
    if match:
        return match.group(1)
    return ""

def clean_id_val(raw_val):
    if pd.isna(raw_val) or str(raw_val).strip().lower() in ["none", "nan", "null", ""]:
        return ""
    val_str = str(raw_val).split(".")[0].strip()
    val_clean = re.sub(r"\s+", "", val_str)
    if len(val_clean) == 11 and val_clean.startswith("79"):
        val_clean = "0" + val_clean
    return val_clean

def is_valid_human_name(name):
    if not name or len(name) < 2:
        return False
    name_clean = str(name).strip()
    if re.match(r"^[\d\s\.\,\-]+$", name_clean):
        return False
    if name_clean.lower() in ["họ và tên", "họ tên", "tên", "người liên hệ", "tổng cộng", "stt", "nan", "none"]:
        return False
    return True

# ==============================================================================
# 4. BỘ PHÂN TÍCH SHEET LINH HOẠT THEO CẤU HÌNH KIỂM SOÁT
# ==============================================================================
def process_single_sheet(s_name, file_bytes, config, file_base_name=""):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    raw_preview = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=None, nrows=15)
    if raw_preview.empty:
        return [], "Trống", False

    # Kiểm tra Header
    auto_h, matches = auto_detect_header_row_smart(raw_preview)
    use_h_idx = config["header_row_idx"] if config["apply_fixed_header"] else auto_h

    # Kiểm tra xem sheet này có header hợp lệ không hay là file phi cấu trúc (như Trẻ 1)
    is_headerless = (matches < 2 and not config["apply_fixed_header"])

    valid_rows = []
    if not is_headerless:
        df_sheet = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=use_h_idx).dropna(how="all")
        cols_list = df_sheet.columns.tolist()

        s_name_col = resolve_column_for_sheet(cols_list, config["map_name"], "name")
        s_last_col = resolve_column_for_sheet(cols_list, config["map_last_name"], "last_name")
        s_first_col = resolve_column_for_sheet(cols_list, config["map_first_name"], "first_name")
        s_phone_col = resolve_column_for_sheet(cols_list, config["map_phone"], "phone")
        s_id_col = resolve_column_for_sheet(cols_list, config["map_id"], "id")

        for _, r in df_sheet.iterrows():
            if config["name_mode"] == "Tách riêng 2 cột (Họ đệm + Tên)":
                p_last = str(r.get(s_last_col, "")).strip() if s_last_col and pd.notna(r.get(s_last_col)) else ""
                p_first = str(r.get(s_first_col, "")).strip() if s_first_col and pd.notna(r.get(s_first_col)) else ""
                if p_last.lower() in ["nan", "none"]: p_last = ""
                if p_first.lower() in ["nan", "none"]: p_first = ""
                full_name = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()
            else:
                full_name = str(r.get(s_name_col, "")).strip() if s_name_col else ""

            if not is_valid_human_name(full_name):
                continue

            phone_raw = config["fix_phone"] if config["map_phone"] == ">> Nhập giá trị cố định <<" else r.get(s_phone_col, "")
            id_raw = config["fix_id"] if config["map_id"] == ">> Nhập giá trị cố định <<" else r.get(s_id_col, "")

            grp_val = file_base_name if "Tên file" in config["group_strategy"] else s_name

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp_val,
                "Tên KH (*)": full_name,
                "Mã định danh": clean_id_val(id_raw),
                "Điện thoại (*)": extract_first_phone(phone_raw),
                "Email": "",
                "Loại KH": "",
                "Địa chỉ/Ghi chú": "",
                "Tên doanh nghiệp": "",
            })
        hdr_desc = f"Dòng {use_h_idx + 1}"
    else:
        # Fallback cho file không header (như Trẻ 1)
        df_all = pd.read_excel(io.BytesIO(safe_bytes), sheet_name=s_name, header=None).dropna(how="all")
        for _, r in df_all.iterrows():
            p_last = str(r[1]).strip() if len(r) > 1 and pd.notna(r[1]) else ""
            p_first = str(r[2]).strip() if len(r) > 2 and pd.notna(r[2]) else ""
            full_name = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()
            if not is_valid_human_name(full_name):
                continue

            v_id = clean_id_val(r[3]) if len(r) > 3 else ""
            phone_raw = r[33] if len(r) > 33 else ""
            grp_val = file_base_name if "Tên file" in config["group_strategy"] else s_name

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp_val,
                "Tên KH (*)": full_name,
                "Mã định danh": v_id,
                "Điện thoại (*)": extract_first_phone(phone_raw),
                "Email": "",
                "Loại KH": "",
                "Địa chỉ/Ghi chú": "",
                "Tên doanh nghiệp": "",
            })
        hdr_desc = "Tự động đọc trực tiếp (Không có Header)"

    return valid_rows, hdr_desc, False

# ==============================================================================
# 5. GIAO DIỆN CHÍNH STREAMLIT - PHẦN KIỂM SOÁT
# ==============================================================================
uploaded_files = st.file_uploader(
    "1. Tải lên 1 hoặc nhiều file Excel (.xlsx, .xls):",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã nạp thành công **{len(uploaded_files)} file**")

    # Chọn File đại diện để xem xét và cấu hình
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        sample_file_name = st.selectbox("🎯 Chọn File làm mẫu cấu hình:", options=list(file_map.keys()))
    sample_file_bytes = file_map[sample_file_name]
    sample_sheets = get_safe_sheet_names(sample_file_bytes)

    # Bộ lọc sheet: Tự động loại bỏ các sheet phụ nếu là file mầm non
    default_sheets = [s for s in sample_sheets if s.lower() not in ["dulieu", "thongbaoloi"]]
    if not default_sheets:
        default_sheets = sample_sheets

    with col_f2:
        selected_sheets = st.multiselect(
            "📋 Danh sách Sheet sẽ xử lý trong file này:",
            options=sample_sheets,
            default=default_sheets
        )

    st.markdown("---")
    st.subheader("BƯỚC 1: KIỂM SOÁT TIÊU ĐỀ & CHỌN CỘT")

    # Đọc nhanh 15 dòng của sheet mẫu để dò tìm
    preview_ref_sheet = selected_sheets[0] if selected_sheets else sample_sheets[0]
    raw_preview = load_preview_df(sample_file_bytes, preview_ref_sheet, header=None, nrows=15)
    detected_h_row, match_count = auto_detect_header_row_smart(raw_preview)

    c_h1, c_h2 = st.columns([1, 2])
    with c_h1:
        header_choice = st.number_input(
            "📌 Vị trí dòng tiêu đề (Header):",
            min_value=1,
            max_value=15,
            value=int(detected_h_row + 1)
        )
    with c_h2:
        apply_fixed_header = st.checkbox(
            "Áp dụng cố định dòng tiêu đề này cho tất cả file/sheet khác",
            value=False,
            help="Nếu bỏ tích, hệ thống sẽ tự động quét linh hoạt theo từng sheet"
        )

    curr_header_idx = int(header_choice) - 1
    sample_cols_df = load_preview_df(sample_file_bytes, preview_ref_sheet, header=curr_header_idx, nrows=5)
    valid_cols = [str(c).strip() for c in sample_cols_df.columns if not str(c).startswith("Unnamed:") and pd.notna(c)]
    dropdown_opts = ["-- Bỏ trống --", ">> Nhập giá trị cố định <<"] + valid_cols

    def get_auto_index(category_key):
        matched = find_column_by_keywords(valid_cols, KEYWORDS[category_key])
        return dropdown_opts.index(matched) if matched else 0

    st.markdown("---")
    st.subheader("BƯỚC 2: GHÉP CỘT & ĐIỀU CHỈNH THÔNG TIN")

    val_coso = st.text_input(
        "🏢 Tên Cơ sở (*) (Bắt buộc theo chuẩn FinOne):",
        value="",
        placeholder="Ví dụ: MNTT Bông Sen Hồng, Tòa nhà Golden..."
    )

    c_map1, c_map2 = st.columns(2)
    with c_map1:
        grp_strategy = st.radio(
            "🏢 Nhóm KH (Lớp/Khu vực) xác định theo:",
            [
                "Tên từng File (Thích hợp khi mỗi lớp là 1 file)",
                "Tên từng Sheet (Thích hợp khi file có nhiều sheet lớp)",
                "Nhập tên cố định"
            ],
            index=0 if len(uploaded_files) > 1 else 1
        )
        name_mode = st.radio(
            "👤 Định dạng cột Họ và Tên:",
            ["Tách riêng 2 cột (Họ đệm + Tên)", "Họ và Tên gộp chung 1 cột"],
            index=1 if (get_auto_index("name") != 0 or get_auto_index("last_name") == 0) else 0
        )

        m_name, m_last_name, m_first_name = None, None, None
        if name_mode == "Tách riêng 2 cột (Họ đệm + Tên)":
            c_a, c_b = st.columns(2)
            with c_a: m_last_name = st.selectbox("Cột Họ / Họ đệm (*):", dropdown_opts, index=get_auto_index("last_name"))
            with c_b: m_first_name = st.selectbox("Cột Tên (*):", dropdown_opts, index=get_auto_index("first_name"))
        else:
            m_name = st.selectbox("Cột Họ và Tên (*):", dropdown_opts, index=get_auto_index("name"))

    with c_map2:
        m_phone = st.selectbox("Cột Điện thoại liên lạc (*):", dropdown_opts, index=get_auto_index("phone"))
        f_phone = st.text_input("✍️ Nhập SĐT cố định:") if m_phone == ">> Nhập giá trị cố định <<" else ""

        m_id = st.selectbox("Cột Mã định danh / CCCD:", dropdown_opts, index=get_auto_index("id"))
        f_id = st.text_input("✍️ Nhập Mã định danh cố định:") if m_id == ">> Nhập giá trị cố định <<" else ""

    config = {
        "val_coso": val_coso.strip(),
        "header_row_idx": curr_header_idx,
        "apply_fixed_header": apply_fixed_header,
        "group_strategy": grp_strategy,
        "name_mode": name_mode,
        "map_name": m_name,
        "map_last_name": m_last_name,
        "map_first_name": m_first_name,
        "map_phone": m_phone, "fix_phone": f_phone,
        "map_id": m_id, "fix_id": f_id,
    }

    # ==============================================================================
    # BƯỚC 3: LIVE PREVIEW KIỂM SOÁT TRỰC QUAN
    # ==============================================================================
    st.markdown("---")
    st.subheader("🔍 BƯỚC 3: XEM TRƯỚC DỮ LIỆU (LIVE PREVIEW)")

    pv_col1, pv_col2 = st.columns(2)
    with pv_col1:
        pv_file = st.selectbox("Chọn File xem trước:", options=list(file_map.keys()))
    with pv_col2:
        pv_sheets = get_safe_sheet_names(file_map[pv_file])
        pv_sheet = st.selectbox("Chọn Sheet xem trước:", options=pv_sheets)

    base_fname = os.path.splitext(pv_file)[0]
    pv_rows, pv_hdr_desc, _ = process_single_sheet(pv_sheet, file_map[pv_file], config, base_fname)

    if pv_rows:
        st.caption(f"🔎 Nhận diện: **{pv_hdr_desc}** | Tìm thấy **{len(pv_rows)} dòng dữ liệu**")
        df_pv_show = pd.DataFrame(pv_rows[:5])
        st.dataframe(df_pv_show[["Cơ sở (*)", "Nhóm KH (*)", "Tên KH (*)", "Mã định danh", "Điện thoại (*)"]], use_container_width=True)
    else:
        st.warning(f"⚠️ Sheet '{pv_sheet}' trong file '{pv_file}' hiện chưa trích xuất được dòng nào. Hãy kiểm tra lại cấu hình cột hoặc dòng tiêu đề.")

    # ==============================================================================
    # BƯỚC 4: GỘP TOÀN BỘ VÀ XUẤT FILE FINONE
    # ==============================================================================
    st.markdown("---")
    st.subheader("BƯỚC 4: XUẤT FILE HOÀN CHỈNH")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_FinOne.xlsx")

    if st.button("🚀 XÁC NHẬN & GỘP TẤT CẢ FILE", type="primary"):
        if not val_coso.strip():
            st.error("🚨 LỖI: Bạn chưa điền 'Tên Cơ sở (*)' ở Bước 2. FinOne bắt buộc phải có tên cơ sở để import!")
            st.stop()

        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}] trong thư mục chạy ứng dụng!")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        progress_bar = st.progress(0)
        status_text = st.empty()

        all_records = []
        summary_stats = []
        total_files = len(file_map)

        for f_idx, (fname, fbytes) in enumerate(file_map.items()):
            base_n = os.path.splitext(fname)[0]
            sheets = get_safe_sheet_names(fbytes)

            # Lọc bỏ sheet rác
            target_sheets = [s for s in sheets if s.lower() not in ["dulieu", "thongbaoloi"]]
            if not target_sheets:
                target_sheets = sheets

            for s_name in target_sheets:
                status_text.text(f"Đang xử lý: File '{fname}' ➔ Sheet '{s_name}'...")
                rows, hdr_desc, _ = process_single_sheet(s_name, fbytes, config, base_n)
                all_records.extend(rows)

                summary_stats.append({
                    "File Nguồn": fname,
                    "Sheet": s_name,
                    "Dòng Header": hdr_desc,
                    "Số dòng hợp lệ": len(rows),
                    "Người đại diện đầu": rows[0]["Tên KH (*)"] if rows else "--",
                    "Người đại diện cuối": rows[-1]["Tên KH (*)"] if rows else "--"
                })

            progress_bar.progress((f_idx + 1) / total_files)

        status_text.text("✅ Đang nạp dữ liệu vào mẫu FinOne...")

        # Ghi trực tiếp vào từng cột chuẩn từ B đến J
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
        for r_dict in all_records:
            for key, col_idx in col_mapping.items():
                cell_val = r_dict.get(key, "")
                cell = ws.cell(row=current_row, column=col_idx)
                
                # Định dạng Text chuẩn (@) không ép cờ 's' gây cảnh báo xanh lá
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.number_format = '@'
                    cell.value = str(cell_val) if cell_val else ""
                else:
                    cell.value = cell_val
            current_row += 1

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        st.success(f"🎉 Gộp thành công tổng cộng **{len(all_records)} khách hàng**!")
        st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)

        st.download_button(
            "📥 TẢI FILE EXCEL KẾT QUẢ CHO FINONE",
            out_buf,
            output_filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
