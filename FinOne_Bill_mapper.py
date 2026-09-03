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
# 1. TỪ KHÓA TỰ ĐỘNG NHẬN DIỆN CỘT
# ==============================================================================
KEYWORDS = {
    "stt": ["tt", "stt", "số tt", "số thứ tự", "no.", "no"],
    "full_name": ["họ và tên", "họ tên", "tên khách hàng", "người đại diện", "chủ hộ", "học sinh", "tên học sinh", "người nộp"],
    "last_name": ["họ đệm", "họ và đệm", "họ và tên đệm", "họ lót", "họ"],
    "first_name": ["tên gọi", "tên"],
    "phone": ["điện thoại", "sđt", "sdt", "phone", "tel", "di động", "mobile", "liên hệ", "đt cha", "đt mẹ", "sđt ba", "sđt mẹ"],
    "id": ["mã", "id", "định danh", "cccd", "cmnd", "cmt", "mã kh", "số định danh", "mã định danh", "mã học sinh", "mã hs"],
}

# ==============================================================================
# 2. XỬ LÝ LỖI CORRUPT XML TRÊN RAM
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
def load_sheet_raw(file_bytes, sheet_name, nrows=None):
    safe_bytes = sanitize_excel_bytes(file_bytes)
    return pd.read_excel(io.BytesIO(safe_bytes), sheet_name=sheet_name, header=None, nrows=nrows)

# ==============================================================================
# 3. TRÍCH XUẤT VÀ CHUẨN HÓA DỮ LIỆU
# ==============================================================================
def extract_first_phone(phone_raw):
    if not phone_raw or pd.isna(phone_raw):
        return ""
    val_str = str(phone_raw).strip()
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
    if name_clean.lower() in ["họ và tên", "họ tên", "tên", "người liên hệ", "tổng cộng", "stt"]:
        return False
    return True

# ==============================================================================
# 4. NHẬN DIỆN CẤU TRÚC BẢNG DỮ LIỆU TỰ ĐỘNG
# ==============================================================================
def auto_detect_table_structure(df_raw):
    best_row = 0
    max_score = 0
    detected_cols = {}

    for r_idx in range(min(20, len(df_raw))):
        row_vals = [str(v).lower().strip() for v in df_raw.iloc[r_idx].values if pd.notna(v)]
        score = 0
        cols = {}
        for c_idx, cell in enumerate(row_vals):
            for role, kws in KEYWORDS.items():
                if any(kw == cell or kw in cell for kw in kws):
                    if role not in cols:
                        cols[role] = c_idx
                        score += 2 if role in ["full_name", "first_name", "phone"] else 1
        if score > max_score:
            max_score = score
            best_row = r_idx
            detected_cols = cols

    # Trường hợp không thấy dòng header rõ ràng (Dò theo pattern dữ liệu)
    if max_score < 2:
        phone_col = None
        for col_idx in range(df_raw.shape[1]):
            matches = 0
            for r in range(min(15, len(df_raw))):
                val = str(df_raw.iat[r, col_idx])
                if extract_first_phone(val):
                    matches += 1
            if matches >= 2:
                phone_col = col_idx
                break
        return 0, {"phone": phone_col}

    return best_row, detected_cols

def parse_sheet(df_raw, header_row, col_map, config, group_name):
    records = []
    start_row = header_row + 1 if header_row is not None else 0

    has_split_name = ("last_name" in col_map and "first_name" in col_map)
    full_name_col = col_map.get("full_name")
    last_name_col = col_map.get("last_name")
    first_name_col = col_map.get("first_name")
    phone_col = col_map.get("phone")
    id_col = col_map.get("id")

    for r in range(start_row, len(df_raw)):
        # Trích xuất Họ Tên
        if has_split_name:
            last = str(df_raw.iat[r, last_name_col]).strip() if pd.notna(df_raw.iat[r, last_name_col]) else ""
            first = str(df_raw.iat[r, first_name_col]).strip() if pd.notna(df_raw.iat[r, first_name_col]) else ""
            if last.lower() in ["nan", "none"]: last = ""
            if first.lower() in ["nan", "none"]: first = ""
            name_str = re.sub(r"\s+", " ", f"{last} {first}").strip()
        elif full_name_col is not None:
            val = df_raw.iat[r, full_name_col]
            name_str = str(val).strip() if pd.notna(val) else ""
        else:
            name_str = ""

        if not is_valid_human_name(name_str):
            continue

        # Trích xuất SĐT & ID
        phone_val = extract_first_phone(df_raw.iat[r, phone_col]) if phone_col is not None and phone_col < df_raw.shape[1] else ""
        id_val = clean_id_val(df_raw.iat[r, id_col]) if id_col is not None and id_col < df_raw.shape[1] else ""

        records.append({
            "Cơ sở (*)": config["val_coso"],
            "Nhóm KH (*)": group_name,
            "Tên KH (*)": name_str,
            "Mã định danh": id_val,
            "Điện thoại (*)": phone_val,
            "Email": "",
            "Loại KH": "",
            "Địa chỉ/Ghi chú": "",
            "Tên doanh nghiệp": "",
        })
    return records

# ==============================================================================
# 5. GIAO DIỆN STREAMLIT
# ==============================================================================
uploaded_files = st.file_uploader(
    "1. Tải lên danh sách các file Excel:",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    file_map = {f.name: f.getvalue() for f in uploaded_files}
    st.success(f"📂 Đã nạp thành công **{len(uploaded_files)} file**")

    st.markdown("---")
    st.subheader("2. CẤU HÌNH NHẬP LIỆU")

    val_coso = st.text_input(
        "🏢 Tên Cơ sở (*):", 
        value="", 
        placeholder="Nhập tên trường, trung tâm hoặc đơn vị nạp tiền..."
    )

    col1, col2 = st.columns(2)
    with col1:
        grp_strategy = st.radio(
            "Tên 'Nhóm KH' lấy theo:",
            ["Tên file (loại bỏ đuôi .xlsx)", "Tên Sheet trong file"],
            index=0 if len(uploaded_files) > 1 else 1
        )
    with col2:
        output_filename = st.text_input("Tên file kết quả:", value="Ket_Qua_Nhap_Lieu_Khach_Hang.xlsx")

    config = {
        "val_coso": val_coso.strip(),
        "group_strategy": grp_strategy
    }

    if st.button("BẮT ĐẦU XỬ LÝ & GỘP DỮ LIỆU", type="primary"):
        if not val_coso.strip():
            st.warning("⚠️ Vui lòng nhập 'Tên Cơ sở (*)' trước khi chuyển đổi.")
            st.stop()

        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}]. Hãy đặt file mẫu trong cùng thư mục chạy ứng dụng.")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        all_records = []
        summary_stats = []

        for fn, fbytes in file_map.items():
            base_n = os.path.splitext(fn)[0]
            sheets = get_sheet_names(fbytes)

            for s_name in sheets:
                df_raw = load_sheet_raw(fbytes, s_name)
                if df_raw.empty:
                    continue

                best_hdr, detected_cols = auto_detect_table_structure(df_raw)
                
                # Bỏ qua sheet nếu không quét được cột tên và không có dữ liệu
                if "full_name" not in detected_cols and "first_name" not in detected_cols and len(detected_cols) <= 1:
                    continue

                grp = base_n if "Tên file" in grp_strategy else s_name
                sheet_records = parse_sheet(df_raw, best_hdr, detected_cols, config, grp)

                if sheet_records:
                    all_records.extend(sheet_records)
                    summary_stats.append({
                        "File": fn,
                        "Sheet": s_name,
                        "Dòng Header": best_hdr + 1,
                        "Số bản ghi hợp lệ": len(sheet_records),
                        "Đại diện": sheet_records[0]["Tên KH (*)"]
                    })

        if not all_records:
            st.error("Không tìm thấy dòng dữ liệu nào khớp định dạng học sinh/khách hàng. Vui lòng kiểm tra lại cấu trúc file.")
            st.stop()

        # Ánh xạ ghi vào các cột của FinOne (Cột B đến J)
        col_mapping = {
            "Cơ sở (*)": 2,
            "Nhóm KH (*)": 3,
            "Tên KH (*)": 4,
            "Mã định danh": 5,
            "Điện thoại (*)": 6,
            "Email": 7,
            "Loại KH": 8,
            "Địa chỉ/Ghi chú": 9,
            "Tên doanh nghiệp": 10
        }

        current_row = 6
        for record in all_records:
            for key, col_idx in col_mapping.items():
                cell = ws.cell(row=current_row, column=col_idx)
                val = record.get(key, "")
                if key in ["Điện thoại (*)", "Mã định danh"]:
                    cell.data_type = 's'
                    cell.number_format = '@'
                cell.value = val
            current_row += 1

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        st.success(f"🎉 Đã xử lý hoàn tất **{len(all_records)} khách hàng**!")
        st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)

        st.download_button(
            "📥 TẢI FILE EXCEL HOÀN CHỈNH CHO FINONE",
            out_buf,
            output_filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
