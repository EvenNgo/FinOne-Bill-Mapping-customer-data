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

def auto_detect_header_row_smart(raw_df):
    best_row, max_matches = 0, 0
    for r_idx in range(min(10, len(raw_df))):
        row_vals = [str(v).lower().strip() for v in raw_df.iloc[r_idx].values if pd.notna(v)]
        row_text = " ".join(row_vals)
        matches = sum(1 for kws in KEYWORDS.values() if any(k in row_text for k in kws))
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx
    return best_row, max_matches

def clean_phone(phone_val):
    if pd.isna(phone_val): return ""
    val_str = str(phone_val).strip().split(".")[0]
    if val_str.lower() in ["none", "nan", "null", ""]: return ""
    if val_str == "0": return "0"
    digits = re.sub(r"\D", "", val_str)
    if not digits: return val_str
    if digits.startswith("84") and len(digits) == 11: digits = "0" + digits[2:]
    elif len(digits) == 9 and not digits.startswith("0"): digits = "0" + digits
    return digits

def safe_str(val, default=""):
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return default
    s = str(val).strip()
    return default if s.lower() in ("nan", "none", "null") else s

def is_valid_human_name(name):
    if not name or len(name) < 2: return False
    if re.match(r"^[\d\s\.\,\-]+$", name): return False
    return True

# ==============================================================================
# BỘ XỬ LÝ DỮ LIỆU TỰ ĐỘNG THÍCH ỨNG (CÓ/KHÔNG CÓ HEADER)
# ==============================================================================
def process_sheet_data_adaptive(s_name, file_bytes, config, file_base_name=""):
    raw_df = load_sheet(file_bytes, sheet_name=s_name, header=None)
    if raw_df.empty:
        return [], [], "Trống", "Trống", False

    # 1. Kiểm tra file này có Dòng tiêu đề hay không
    has_header = False
    detected_h_idx = 0
    for r_idx in range(min(5, len(raw_df))):
        row_text = " ".join([str(v).lower() for v in raw_df.iloc[r_idx].values if pd.notna(v)])
        if ("họ đệm" in row_text and "tên" in row_text) or ("họ và tên" in row_text):
            has_header = True
            detected_h_idx = r_idx
            break

    valid_rows, rejected_rows = [], []

    # TRƯỜNG HỢP 1: File có Header chuẩn
    if has_header:
        df_sheet = load_sheet(file_bytes, sheet_name=s_name, header=detected_h_idx)
        cols_list = df_sheet.columns.tolist()

        s_last_col = resolve_column_for_sheet(cols_list, config.get("map_last_name"), "last_name")
        s_first_col = resolve_column_for_sheet(cols_list, config.get("map_first_name"), "first_name")
        s_name_col = resolve_column_for_sheet(cols_list, config.get("map_name"), "name")
        s_phone_col = resolve_column_for_sheet(cols_list, config["map_phone"], "phone")
        s_id_col = resolve_column_for_sheet(cols_list, config["map_id"], "id")
        s_group_col = resolve_column_for_sheet(cols_list, config.get("group_col"), "group") if "Một cột" in config["group_strategy"] else None

        for idx, r in df_sheet.iterrows():
            excel_row = detected_h_idx + idx + 2
            if config.get("name_mode") == "Tách riêng 2 cột (Họ đệm + Tên)":
                p_last = str(r.get(s_last_col, "")).strip() if s_last_col and pd.notna(r.get(s_last_col)) else ""
                p_first = str(r.get(s_first_col, "")).strip() if s_first_col and pd.notna(r.get(s_first_col)) else ""
                name_str = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()
            else:
                name_str = str(r.get(s_name_col, "")).strip()

            if not is_valid_human_name(name_str):
                continue

            grp = file_base_name if "Tên từng File" in config["group_strategy"] else s_name
            v_id = str(r.get(s_id_col, "")).split(".")[0].strip() if s_id_col and pd.notna(r.get(s_id_col)) else ""
            if len(v_id) == 11 and v_id.startswith("79"): v_id = "0" + v_id

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp,
                "Tên KH (*)": name_str,
                "Mã định danh": v_id,
                "Điện thoại (*)": clean_phone(r.get(s_phone_col, "")) if s_phone_col else "",
                "Email": "", "Loại KH": "", "Địa chỉ/Ghi chú": "", "Tên doanh nghiệp": "",
            })
        hdr_desc = f"Dòng {detected_h_idx+1}"

    # TRƯỜNG HỢP 2: File KHÔNG có Header (như Trẻ 1.xlsx)
    else:
        for idx, r in raw_df.iterrows():
            excel_row = idx + 1
            # Vị trí cố định: Cột 0: STT, Cột 1: Họ đệm, Cột 2: Tên, Cột 3: Mã định danh
            p_last = str(r[1]).strip() if pd.notna(r[1]) else ""
            p_first = str(r[2]).strip() if pd.notna(r[2]) else ""
            name_str = re.sub(r"\s+", " ", f"{p_last} {p_first}").strip()

            if not is_valid_human_name(name_str):
                continue

            grp = file_base_name if "Tên từng File" in config["group_strategy"] else s_name
            v_id = str(r[3]).split(".")[0].strip() if pd.notna(r[3]) else ""
            if len(v_id) == 11 and v_id.startswith("79"): v_id = "0" + v_id

            valid_rows.append({
                "Cơ sở (*)": config["val_coso"],
                "Nhóm KH (*)": grp,
                "Tên KH (*)": name_str,
                "Mã định danh": v_id,
                "Điện thoại (*)": "",
                "Email": "", "Loại KH": "", "Địa chỉ/Ghi chú": "", "Tên doanh nghiệp": "",
            })
        hdr_desc = "Tự động kích hoạt (Không có Header)"

    return valid_rows, rejected_rows, hdr_desc, "Tự động nhận diện", False

# ==============================================================================
# GIAO DIỆN CHÍNH
# ==============================================================================
uploaded_files = st.file_uploader(
    "1. Tải lên danh sách các file Excel (.xlsx):",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã tải lên **{len(uploaded_files)} file** thành công!")

    sample_fn = list(file_map.keys())[0]
    sample_bytes = file_map[sample_fn]
    sheet_opts = get_sheet_names(sample_bytes)
    
    # Tự động chọn sheet 'MauNhapLieu' nếu có
    def_sheet = "MauNhapLieu" if "MauNhapLieu" in sheet_opts else sheet_opts[0]

    st.markdown("---")
    st.subheader("CẤU HÌNH GỘP DỮ LIỆU")

    c1, c2 = st.columns(2)
    with c1:
        val_coso = st.text_input("🏢 Tên Cơ sở:", value="Trường Mầm Non Phú Lương")
        grp_strategy = st.radio("🏢 Nhóm KH xác định theo:", ["Tên từng File (Bỏ đuôi .xlsx)", "Tên từng Sheet", "Nhập tên cố định"])
    with c2:
        name_mode = st.radio("👤 Định dạng Họ và Tên:", ["Tách riêng 2 cột (Họ đệm + Tên)", "Họ và Tên gộp chung 1 cột"])

    config = {
        "val_coso": val_coso,
        "group_strategy": grp_strategy,
        "name_mode": name_mode,
        "map_last_name": "Họ đệm",
        "map_first_name": "Tên",
        "map_name": "Họ và tên",
        "map_phone": "-- Bỏ trống --", "fix_phone": "",
        "map_id": "Mã định danh", "fix_id": "",
    }

    st.markdown("---")
    if st.button("🚀 GỘP TẤT CẢ FILE VÀO MẪU FINONE", type="primary"):
        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}]!")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        total_valid = 0
        summary_stats = []
        bulk_write_data = []

        for fn, fbytes in file_map.items():
            base_n = os.path.splitext(fn)[0]
            sheets = get_sheet_names(fbytes)
            # Chỉ đọc sheet MauNhapLieu, bỏ qua DuLieu và ThongBaoLoi
            target_s = "MauNhapLieu" if "MauNhapLieu" in sheets else sheets[0]

            v_rows, _, hdr_desc, _, _ = process_sheet_data_adaptive(target_s, fbytes, config, base_n)
            bulk_write_data.extend(v_rows)
            total_valid += len(v_rows)
            summary_stats.append({
                "Tên File": fn,
                "Trạng thái nhận diện": hdr_desc,
                "Số học sinh": len(v_rows),
                "Học sinh đầu": v_rows[0]["Tên KH (*)"] if v_rows else "",
                "Học sinh cuối": v_rows[-1]["Tên KH (*)"] if v_rows else ""
            })

        # Ghi dữ liệu vào File mẫu
        col_mapping = {"Cơ sở (*)": 2, "Nhóm KH (*)": 3, "Tên KH (*)": 4, "Mã định danh": 5, "Điện thoại (*)": 6}
        cur_row = 6
        for r_dict in bulk_write_data:
            for k, c_idx in col_mapping.items():
                val = r_dict.get(k, "")
                cell = ws.cell(row=cur_row, column=c_idx)
                if k in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'
                    cell.number_format = '@'
                cell.value = val
            cur_row += 1

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        st.success(f"🎉 Gộp thành công **{total_valid} học sinh** từ **{len(uploaded_files)} file**!")
        st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)

        st.download_button(
            "📥 TẢI FILE EXCEL HOÀN CHỈNH CHO FINONE",
            out_buf,
            "Ket_Qua_Gop_Mam_Non_FinOne.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
