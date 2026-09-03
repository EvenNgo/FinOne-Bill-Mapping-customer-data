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
# 2. XỬ LÝ LỖI CORRUPT XML TRỰC TIẾP TRÊN RAM (CHỐNG LỖI MULTICELLRANGE)
# ==============================================================================
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

# ==============================================================================
# 3. HÀM LẤY DUY NHẤT 1 SỐ ĐIỆN THOẠI ĐẦU TIÊN
# ==============================================================================
def extract_first_phone(phone_raw):
    """
    Tự động tìm và chỉ lấy DUY NHẤT số điện thoại đầu tiên trong ô (dù có 2-3 số).
    Đảm bảo đủ 10 chữ số, giữ số 0 ở đầu.
    """
    if not phone_raw or pd.isna(phone_raw):
        return ""
    val_str = str(phone_raw).strip()
    if val_str.lower() in ["none", "nan", "null", "0", "0.0", ""]:
        return ""

    # Chuyển các ký tự phân cách thành khoảng trắng
    cleaned = re.sub(r"[/,;\-–—|và+]+", " ", val_str)
    tokens = cleaned.split()

    for tok in tokens:
        d = re.sub(r"\D", "", tok)
        if d.startswith("84") and len(d) == 11:
            d = "0" + d[2:]
        elif len(d) == 9 and not d.startswith("0"):
            d = "0" + d

        if len(d) == 10 and d.startswith("0"):
            return d  # Trả về ngay số đầu tiên tìm thấy

    # Dự phòng: nếu dính liền 20 số không dấu cách (vd: 09150666200981833877)
    raw_digits = re.sub(r"\D", "", val_str)
    match = re.search(r"(0\d{9})", raw_digits)
    if match:
        return match.group(1)

    return ""

def is_valid_human_name(name):
    """Loại bỏ dòng rác đánh số cột (ví dụ dòng 3 chứa số '2 3')."""
    if not name or len(name) < 2:
        return False
    if re.match(r"^[\d\s\.\,\-]+$", name):
        return False
    return True

# ==============================================================================
# 4. BỘ XỬ LÝ DỮ LIỆU TỰ THÍCH ỨNG (CÓ / KHÔNG CÓ HEADER)
# ==============================================================================
def process_sheet_data_adaptive(s_name, file_bytes, config, file_base_name=""):
    raw_df = load_sheet(file_bytes, sheet_name=s_name, header=None)
    if raw_df.empty:
        return [], [], "Trống", "Trống", False

    has_header = False
    detected_h_idx = 0
    for r_idx in range(min(5, len(raw_df))):
        row_text = " ".join([str(v).lower() for v in raw_df.iloc[r_idx].values if pd.notna(v)])
        if ("họ đệm" in row_text and "tên" in row_text) or ("họ và tên" in row_text):
            has_header = True
            detected_h_idx = r_idx
            break

    valid_rows = []

    # TRƯỜNG HỢP 1: File chuẩn có Header (như Chồi 1, Lá 1, Mầm 1...)
    if has_header:
        df_sheet = load_sheet(file_bytes, sheet_name=s_name, header=detected_h_idx)
        cols_list = df_sheet.columns.tolist()

        s_last_col = resolve_column_for_sheet(cols_list, config.get("map_last_name"), "last_name")
        s_first_col = resolve_column_for_sheet(cols_list, config.get("map_first_name"), "first_name")
        s_name_col = resolve_column_for_sheet(cols_list, config.get("map_name"), "name")
        s_phone_col = resolve_column_for_sheet(cols_list, config["map_phone"], "phone")
        s_id_col = resolve_column_for_sheet(cols_list, config["map_id"], "id")

        for idx, r in df_sheet.iterrows():
            if config.get("name_mode") == "Tách riêng 2 cột (Họ đệm + Tên)":
                p_last = str(r.get(s_last_col, "")).strip() if s_last_col and pd.notna(r.get(s_last_col)) else ""
                p_first = str(r.get(s_first_col, "")).strip() if s_first_col and pd.notna(r.get(s_first_col)) else ""
                if p_last.lower() in ["nan", "none"]: p_last = ""
                if p_first.lower() in ["nan", "none"]: p_first = ""
                name_str = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()
            else:
                name_str = str(r.get(s_name_col, "")).strip()

            if not is_valid_human_name(name_str):
                continue

            grp = file_base_name if "Tên từng File" in config["group_strategy"] else s_name
            
            v_id = str(r.get(s_id_col, "")).split(".")[0].strip() if s_id_col and pd.notna(r.get(s_id_col)) else ""
            if len(v_id) == 11 and v_id.startswith("79"):
                v_id = "0" + v_id

            # Chỉ lấy số điện thoại đầu tiên
            phone_raw = r.get(s_phone_col, "") if s_phone_col else ""
            main_phone = extract_first_phone(phone_raw)

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp,
                "Tên KH (*)": name_str,
                "Mã định danh": v_id,
                "Điện thoại (*)": main_phone,
                "Email": "",
                "Loại KH": "",
                "Địa chỉ/Ghi chú": "",  # Để trống hoàn toàn, không ghi chú SĐT phụ
                "Tên doanh nghiệp": "",
            })
        hdr_desc = f"Dòng {detected_h_idx+1}"

    # TRƯỜNG HỢP 2: File KHÔNG có Header (như Trẻ 1.xlsx)
    else:
        for idx, r in raw_df.iterrows():
            p_last = str(r[1]).strip() if pd.notna(r[1]) else ""
            p_first = str(r[2]).strip() if pd.notna(r[2]) else ""
            name_str = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()

            if not is_valid_human_name(name_str):
                continue

            grp = file_base_name if "Tên từng File" in config["group_strategy"] else s_name
            
            v_id = str(r[3]).split(".")[0].strip() if pd.notna(r[3]) else ""
            if len(v_id) == 11 and v_id.startswith("79"):
                v_id = "0" + v_id

            # Cột 33 trong file Trẻ 1 là cột điện thoại
            phone_raw = r[33] if len(r) > 33 else ""
            main_phone = extract_first_phone(phone_raw)

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp,
                "Tên KH (*)": name_str,
                "Mã định danh": v_id,
                "Điện thoại (*)": main_phone,
                "Email": "",
                "Loại KH": "",
                "Địa chỉ/Ghi chú": "",  # Để trống hoàn toàn
                "Tên doanh nghiệp": "",
            })
        hdr_desc = "Tự động nhận diện (Không có Header)"

    return valid_rows, [], hdr_desc, "OK", False

# ==============================================================================
# 5. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
uploaded_files = st.file_uploader(
    "1. Kéo thả 1 hoặc toàn bộ các file Excel của trường vào đây:",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã nạp **{len(uploaded_files)} file** thành công!")

    # Lấy mẫu cấu hình từ file đầu tiên
    sample_fn = list(file_map.keys())[0]
    sample_bytes = file_map[sample_fn]
    sheet_opts = get_sheet_names(sample_bytes)
    def_sheet = "MauNhapLieu" if "MauNhapLieu" in sheet_opts else sheet_opts[0]

    raw_sample = load_sheet(sample_bytes, sheet_name=def_sheet, header=None, nrows=10)
    
    detected_h = 1
    for r_i in range(min(5, len(raw_sample))):
        txt = " ".join([str(v).lower() for v in raw_sample.iloc[r_i].values if pd.notna(v)])
        if "họ đệm" in txt or "họ và tên" in txt:
            detected_h = r_i
            break

    df_sample = load_sheet(sample_bytes, sheet_name=def_sheet, header=detected_h)
    valid_cols = [str(c).strip() for c in df_sample.columns if not str(c).startswith("Unnamed:") and pd.notna(c)]
    dropdown_opts = ["-- Bỏ trống --"] + valid_cols

    def get_auto_index(cat_key):
        m = find_column_by_keywords(valid_cols, KEYWORDS[cat_key])
        return dropdown_opts.index(m) if m else 0

    st.markdown("---")
    st.subheader("BƯỚC 2: XÁC NHẬN THÔNG TIN GHÉP CỘT")

    # ĐỂ TRỐNG Ô CƠ SỞ ĐỂ NGƯỜI DÙNG TỰ NHẬP
    val_coso = st.text_input(
        "🏢 Tên Cơ sở (*) (Nhập tên trường/cơ sở áp dụng chung cho tất cả khách hàng):", 
        value="", 
        placeholder="Ví dụ: Trường Mầm Non Hoa Sen, Tòa nhà Bitexco..."
    )

    col1, col2 = st.columns(2)
    with col1:
        grp_strategy = st.radio(
            "🏢 Cách đặt tên Nhóm KH (Lớp học):",
            [
                "Tên từng File (Bỏ đuôi .xlsx - File 'Chồi 1.xlsx' thành lớp 'Chồi 1')",
                "Tên từng Sheet",
                "Nhập tên cố định"
            ],
            index=0 if len(uploaded_files) > 1 else 1
        )
        name_mode = st.radio(
            "👤 Cấu trúc Họ và Tên:",
            ["Tách riêng 2 cột (Họ đệm + Tên)", "Họ và Tên gộp chung 1 cột"],
            index=0
        )
        if name_mode == "Tách riêng 2 cột (Họ đệm + Tên)":
            c_a, c_b = st.columns(2)
            with c_a: m_last = st.selectbox("Cột Họ đệm (*):", dropdown_opts, index=get_auto_index("last_name"))
            with c_b: m_first = st.selectbox("Cột Tên (*):", dropdown_opts, index=get_auto_index("first_name"))
            m_name = None
        else:
            m_name = st.selectbox("Cột Họ và Tên (*):", dropdown_opts, index=get_auto_index("name"))
            m_last, m_first = None, None

    with col2:
        m_phone = st.selectbox("Cột Điện thoại liên lạc (*):", dropdown_opts, index=get_auto_index("phone"))
        m_id = st.selectbox("Cột Mã định danh:", dropdown_opts, index=get_auto_index("id"))

    config = {
        "val_coso": val_coso.strip(),
        "group_strategy": grp_strategy,
        "name_mode": name_mode,
        "map_last_name": m_last,
        "map_first_name": m_first,
        "map_name": m_name,
        "map_phone": m_phone,
        "map_id": m_id,
    }

    st.markdown("---")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_Khach_Hang.xlsx")

    if st.button("🚀 GỘP TOÀN BỘ FILE VÀO MẪU FINONE", type="primary"):
        if not val_coso.strip():
            st.warning("⚠️ Bạn chưa nhập 'Tên Cơ sở (*)', vui lòng điền tên cơ sở ở Bước 2 trước khi gộp file!")
            st.stop()

        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}] trong thư mục ứng dụng!")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        total_valid = 0
        summary_stats = []
        bulk_write_data = []

        for fn, fbytes in file_map.items():
            base_n = os.path.splitext(fn)[0]
            sheets = get_sheet_names(fbytes)
            target_s = "MauNhapLieu" if "MauNhapLieu" in sheets else sheets[0]

            v_rows, _, hdr_desc, _, _ = process_sheet_data_adaptive(target_s, fbytes, config, base_n)
            bulk_write_data.extend(v_rows)
            total_valid += len(v_rows)
            summary_stats.append({
                "Tên File": fn,
                "Cấu trúc phát hiện": hdr_desc,
                "Số học sinh": len(v_rows),
                "Học sinh đầu tiên": v_rows[0]["Tên KH (*)"] if v_rows else "",
                "Học sinh cuối cùng": v_rows[-1]["Tên KH (*)"] if v_rows else ""
            })

        # Mỏ neo chuẩn mẫu FinOne (Cột B đến Cột J)
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
        for r_dict in bulk_write_data:
            for key, col_idx in col_mapping.items():
                cell_val = r_dict.get(key, "")
                cell = ws.cell(row=current_row, column=col_idx)
                
                # Ép kiểu Text để Excel không bao giờ nuốt mất số 0 ở đầu SĐT và Mã định danh
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'
                    cell.number_format = '@'
                cell.value = cell_val
            current_row += 1

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        st.success(f"🎉 Đã gộp thành công toàn bộ **{total_valid} học sinh**!")
        st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)

        st.download_button(
            "📥 TẢI FILE EXCEL HOÀN CHỈNH CHO FINONE",
            out_buf,
            output_filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
