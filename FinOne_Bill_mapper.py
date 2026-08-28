import io
import os
import re
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
    "name": [
        "tên",
        "họ",
        "name",
        "khách hàng",
        "người đại diện",
        "chủ hộ",
        "họ và tên",
    ],
    "phone": [
        "điện thoại",
        "sđt",
        "sdt",
        "phone",
        "tel",
        "di động",
        "mobile",
        "liên hệ",
    ],
    "id": [
        "mã",
        "id",
        "định danh",
        "cccd",
        "cmnd",
        "cmt",
        "mã kh",
        "số định danh",
    ],
    "email": ["email", "mail", "hòm thư"],
    "address": ["địa chỉ", "address", "nơi ở", "thường trú", "tạm trú"],
    "company": ["công ty", "doanh nghiệp", "tổ chức", "đơn vị", "cơ quan"],
    "group": ["khu vực", "nhóm", "phường", "xã", "tổ", "cụm", "khối", "lớp"],
}

# ==============================================================================
# 2. HÀM CORE & CACHE THEO NHU CẦU (ON-DEMAND)
# ==============================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def load_sheet(file_bytes, sheet_name, header=None, nrows=None):
    return pd.read_excel(
        io.BytesIO(file_bytes), sheet_name=sheet_name, header=header, nrows=nrows
    )


def find_column_by_keywords(columns, keyword_list):
    for col in columns:
        col_str = str(col).lower().strip()
        for kw in keyword_list:
            if kw == col_str or kw in col_str:
                return col
    return None


def resolve_column_for_sheet(columns_list, selected_col_name, keyword_category):
    if selected_col_name == "-- Bỏ trống --" or not selected_col_name:
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
        row_vals = [
            str(v).lower().strip() for v in raw_df.iloc[r_idx].values if pd.notna(v)
        ]
        row_text = " ".join(row_vals)
        matches = sum(
            1 for kws in KEYWORDS.values() if any(k in row_text for k in kws)
        )
        if matches > max_matches:
            max_matches = matches
            best_row = r_idx
    return best_row


def clean_phone(phone_val):
    """Chuẩn hóa SĐT về đúng 10 chữ số bắt đầu bằng 0 (0xxxxxxxxx)."""
    if pd.isna(phone_val):
        return ""
    digits = re.sub(r"\D", "", str(phone_val).split(".")[0])
    if digits.startswith("84") and len(digits) == 11:
        digits = "0" + digits[2:]
    elif len(digits) == 9 and not digits.startswith("0"):
        digits = "0" + digits
    return digits if len(digits) == 10 and digits.startswith("0") else ""


def safe_str(val, default=""):
    """Ép về chuỗi an toàn, không bao giờ để lọt literal NaN/None vào ô Excel."""
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return default
    s = str(val).strip()
    return default if s.lower() in ("nan", "none", "null") else s


# ==============================================================================
# 3. TRÁI TIM NGHIỆP VỤ: XỬ LÝ & LỌC DỮ LIỆU DÙNG CHUNG
# ==============================================================================
def process_sheet_data(s_name, file_bytes, config):
    """Xử lý đồng nhất cho cả Preview và Export, bảo đảm phát hiện lệch cột."""
    if s_name == config["preview_sheet_ref"]:
        h_idx = config["current_header_idx"]
    else:
        raw_preview = load_sheet(file_bytes, sheet_name=s_name, header=None, nrows=15)
        h_idx = auto_detect_header_row_smart(raw_preview)

    df_sheet = load_sheet(file_bytes, sheet_name=s_name, header=h_idx).dropna(how="all")
    if df_sheet.empty:
        return [], [], f"Dòng {h_idx+1}", "Trống", False

    cols_list = df_sheet.columns.tolist()
    s_name_col = resolve_column_for_sheet(cols_list, config["map_name"], "name")

    # Bắt lỗi lệch cột Tên giữa các sheet
    if not s_name_col:
        rejected_all = []
        for idx, r in df_sheet.iterrows():
            excel_row_num = h_idx + idx + 2
            rej_dict = {
                "Tên Sheet": s_name,
                "Dòng Excel số": excel_row_num,
                "Lý do loại bỏ": "KHÔNG XÁC ĐỊNH ĐƯỢC CỘT TÊN (Cấu trúc cột khác sheet mẫu)",
            }
            for k, v in r.items():
                if not str(k).startswith("Unnamed:") and pd.notna(k):
                    rej_dict[str(k)] = "" if pd.isna(v) else str(v)
            rejected_all.append(rej_dict)
        return [], rejected_all, f"Dòng {h_idx+1}", "❌ KHÔNG TÌM THẤY", True

    s_phone_col = resolve_column_for_sheet(cols_list, config["map_phone"], "phone")
    s_id_col = resolve_column_for_sheet(cols_list, config["map_id"], "id")
    s_email_col = resolve_column_for_sheet(cols_list, config["map_email"], "email")
    s_addr_col = resolve_column_for_sheet(cols_list, config["map_address"], "address")
    s_comp_col = resolve_column_for_sheet(cols_list, config["map_company"], "company")
    s_stt_col = resolve_column_for_sheet(cols_list, config["col_stt_selected"], "stt")

    # FIX: cột Nhóm/Khu vực (khi chọn "Một cột trong bảng") PHẢI được resolve
    # riêng cho từng sheet giống mọi cột khác — sheet khác có thể đặt tên cột
    # nhóm khác hẳn ("Khu vực" vs "Phường/Xã" vs "Tổ dân phố"...). Trước đây
    # code dùng thẳng tên cột lấy từ sheet mẫu -> sang sheet khác không khớp
    # tên sẽ lặng lẽ trả về giá trị mặc định (rỗng) mà không cảnh báo.
    s_group_col = None
    if "Một cột trong bảng" in config["group_strategy"]:
        s_group_col = resolve_column_for_sheet(
            cols_list, config.get("group_col"), "group"
        )

    valid_rows, rejected_rows = [], []

    garbage_phrases = [
        "tổng cộng", "tổng số", "tổng tiền", "tổng:", "cộng:",
        "người lập biểu", "người lập", "ký tên", "kế toán trưởng",
        "kế toán", "xác nhận", "trưởng phòng", "chữ ký",
        "đvt:", "đơn vị tính", "bằng chữ:", "bằng chữ"
    ]

    for idx, r in df_sheet.iterrows():
        excel_row_num = h_idx + idx + 2
        is_val = True
        reason = ""

        name_val = r.get(s_name_col, "")
        name_str = str(name_val).strip()
        lower_name = name_str.lower()

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
                        stt_num = float(str(stt_val).strip())
                        if stt_num <= 0:
                            is_val = False
                            reason = "[Loại 1] - STT không hợp lệ (<= 0)"
                    except ValueError:
                        is_val = False
                        reason = f"[Loại 1] - STT không phải số ('{stt_val}')"
        elif "Option 2" in config["filter_mode"] and config["required_second_field"] != "Không (Chỉ cần Họ tên)":
            target_col = {
                "Điện thoại liên lạc (*)": s_phone_col,
                "Mã định danh": s_id_col,
                "Địa chỉ/Ghi chú": s_addr_col,
                "Email liên lạc": s_email_col,
            }.get(config["required_second_field"])

            if not target_col:
                is_val = False
                reason = f"Chưa thiết lập cột cho [{config['required_second_field']}]"
            else:
                val = r.get(target_col, "")
                if pd.isna(val) or not str(val).strip() or str(val).strip().lower() in ["nan", "none"]:
                    is_val = False
                    reason = f"Thiếu thông tin bắt buộc: [{config['required_second_field']}]"
                elif "Điện thoại" in config["required_second_field"] and not clean_phone(val):
                    is_val = False
                    reason = f"SĐT không hợp lệ hoặc không đủ 10 số: '{val}'"

        if not is_val:
            if not pd.isna(r.values).all():
                rej_dict = {
                    "Tên Sheet": s_name,
                    "Dòng Excel số": excel_row_num,
                    "Lý do loại bỏ": reason,
                }
                for k, v in r.items():
                    if not str(k).startswith("Unnamed:") and pd.notna(k):
                        rej_dict[str(k)] = "" if pd.isna(v) else str(v)
                rejected_rows.append(rej_dict)
        else:
            # FIX: dùng safe_str() để không bao giờ lọt NaN/"nan" vào ô Excel,
            # và group theo cột đã resolve riêng cho từng sheet (s_group_col).
            if "Tên từng Sheet" in config["group_strategy"]:
                grp_val = s_name
            elif "Một cột trong bảng" in config["group_strategy"]:
                grp_val = (
                    safe_str(r.get(s_group_col), config["fixed_group_val"])
                    if s_group_col
                    else config["fixed_group_val"]
                )
            else:
                grp_val = config["fixed_group_val"]

            valid_rows.append({
                "Cơ sở (*)": "",
                "Nhóm KH (*)": grp_val,
                "Tên KH (*)": name_str,
                "Mã định danh": safe_str(r.get(s_id_col)) if s_id_col else "",
                "Điện thoại (*)": clean_phone(r.get(s_phone_col, "")) if s_phone_col else "",
                "Email": safe_str(r.get(s_email_col)) if s_email_col else "",
                "Loại KH": "Cá nhân",
                "Địa chỉ/Ghi chú": safe_str(r.get(s_addr_col)) if s_addr_col else "",
                "Tên doanh nghiệp": safe_str(r.get(s_comp_col)) if s_comp_col else "",
            })

    return valid_rows, rejected_rows, f"Dòng {h_idx+1}", str(s_name_col), False


# ==============================================================================
# 4. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
uploaded_file = st.file_uploader(
    "1. Tải file Excel của khách hàng (.xlsx, .xls)", type=["xlsx", "xls"]
)

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    all_sheet_names = get_sheet_names(file_bytes)

    st.markdown("---")
    st.subheader("BƯỚC 1: HỆ THỐNG TỰ ĐỘNG DÒ TÌM & GỢI Ý CẤU HÌNH")

    selected_sheets = st.multiselect(
        "📁 Các Sheet cần xử lý (Mặc định chọn tất cả, bỏ tích nếu có sheet thừa):",
        options=all_sheet_names,
        default=all_sheet_names,
    )
    if not selected_sheets:
        st.stop()

    preview_sheet_ref = selected_sheets[0]
    raw_preview_df = load_sheet(file_bytes, sheet_name=preview_sheet_ref, header=None, nrows=15)
    detected_header_idx = auto_detect_header_row_smart(raw_preview_df)

    col_hdr1, col_hdr2 = st.columns([1, 3])
    with col_hdr1:
        header_row_excel = st.number_input(
            "📌 Dòng tiêu đề trên Sheet mẫu:",
            min_value=1,
            max_value=15,
            value=int(detected_header_idx + 1),
        )
    with col_hdr2:
        st.info(
            f"💡 Đang dùng Sheet **'{preview_sheet_ref}'** làm mẫu cấu hình. Tất cả Sheet khác sẽ tự động đồng bộ."
        )

    current_header_idx = int(header_row_excel) - 1
    df_sample = load_sheet(file_bytes, sheet_name=preview_sheet_ref, header=current_header_idx)
    valid_cols = [
        str(c).strip() for c in df_sample.columns if not str(c).startswith("Unnamed:") and pd.notna(c)
    ]
    dropdown_opts = ["-- Bỏ trống --"] + valid_cols

    def get_auto_index(category_key):
        matched = find_column_by_keywords(valid_cols, KEYWORDS[category_key])
        return dropdown_opts.index(matched) if matched else 0

    st.markdown("---")
    st.subheader("BƯỚC 2: CHỌN CHẾ ĐỘ LỌC DỮ LIỆU & GHÉP CỘT")
    st.write("🛡️ **Thiết lập chế độ lọc bỏ dòng rác / tiêu đề phụ:**")

    filter_mode = st.selectbox(
        "Chọn chế độ kiểm tra dữ liệu:",
        [
            "Option 1: Lọc theo cột STT/TT (Khuyên dùng khi file có cột STT 1, 2, 3...)",
            "Option 2: Bắt buộc Họ tên + 1 Trường tùy chọn (Khuyên dùng khi KHÔNG CÓ STT)",
            "Option 3: Lấy tất cả dòng có Tên (Chỉ loại bỏ dòng chữ ký / trống)",
        ],
        index=0 if get_auto_index("stt") != 0 else 1,
    )

    col_stt_selected = None
    required_second_field = None

    if "Option 1" in filter_mode:
        col_stt_selected = st.selectbox(
            "👉 Chọn cột Số Thứ Tự (STT/TT):",
            dropdown_opts,
            index=get_auto_index("stt"),
        )
    elif "Option 2" in filter_mode:
        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
            st.text_input("Trường bắt buộc 1:", value="Tên khách hàng (*)", disabled=True)
        with col_sub2:
            required_second_field = st.selectbox(
                "👉 Trường bắt buộc thứ 2 đi kèm:",
                [
                    "Điện thoại liên lạc (*)",
                    "Mã định danh",
                    "Địa chỉ/Ghi chú",
                    "Email liên lạc",
                    "Không (Chỉ cần Họ tên)",
                ],
                index=0,
            )

    st.markdown("---")
    st.write("🔗 **Ghép các cột dữ liệu vào mẫu FinOne Bill:**")

    col_map1, col_map2 = st.columns(2)
    with col_map1:
        group_strategy = st.radio(
            "🏢 Nhóm KH xác định theo:",
            ["Tên từng Sheet", "Một cột trong bảng", "Nhập tên cố định"],
        )
        group_col = st.selectbox("Cột Nhóm:", dropdown_opts) if "Một cột" in group_strategy else None
        fixed_group_val = st.text_input("Tên nhóm chung:", value="Khách hàng chung") if "cố định" in group_strategy else ""
        map_name = st.selectbox("1. Cột Tên khách hàng (*):", dropdown_opts, index=get_auto_index("name"))
        map_phone = st.selectbox("2. Cột Điện thoại (*):", dropdown_opts, index=get_auto_index("phone"))

    with col_map2:
        map_id = st.selectbox("3. Cột Mã định danh:", dropdown_opts, index=get_auto_index("id"))
        map_email = st.selectbox("4. Cột Email liên lạc:", dropdown_opts, index=get_auto_index("email"))
        map_address = st.selectbox("5. Cột Địa chỉ/Ghi chú:", dropdown_opts, index=get_auto_index("address"))
        map_company = st.selectbox("6. Cột Tên doanh nghiệp:", dropdown_opts, index=get_auto_index("company"))

    config = {
        "preview_sheet_ref": preview_sheet_ref,
        "current_header_idx": current_header_idx,
        "filter_mode": filter_mode,
        "col_stt_selected": col_stt_selected,
        "required_second_field": required_second_field,
        "group_strategy": group_strategy,
        "group_col": group_col,
        "fixed_group_val": fixed_group_val,
        "map_name": map_name,
        "map_phone": map_phone,
        "map_id": map_id,
        "map_email": map_email,
        "map_address": map_address,
        "map_company": map_company,
    }

    # ==============================================================================
    # LIVE PREVIEW CHI TIẾT
    # ==============================================================================
    st.markdown("---")
    st.subheader("🔍 LIVE PREVIEW - CHI TIẾT THEO TỪNG SHEET")
    selected_preview_sheet = st.selectbox("👉 Chọn Sheet muốn xem chi tiết:", options=selected_sheets)

    v_rows, r_rows, _, col_name_found, is_missing = process_sheet_data(
        selected_preview_sheet, file_bytes, config
    )

    if is_missing:
        st.error(
            f"🚨 KHÔNG XÁC ĐỊNH ĐƯỢC CỘT TÊN trên sheet '{selected_preview_sheet}'! "
            f"Vui lòng kiểm tra lại cấu trúc cột của sheet này."
        )

    if "Một cột trong bảng" in group_strategy and not is_missing:
        _cols_check = load_sheet(
            file_bytes, sheet_name=selected_preview_sheet,
            header=(current_header_idx if selected_preview_sheet == preview_sheet_ref
                    else auto_detect_header_row_smart(
                        load_sheet(file_bytes, sheet_name=selected_preview_sheet, header=None, nrows=15)
                    )),
        ).columns.tolist()
        if not resolve_column_for_sheet(_cols_check, group_col, "group"):
            st.warning(
                f"⚠️ Sheet '{selected_preview_sheet}' không tìm thấy cột Nhóm '{group_col}' — "
                f"các dòng hợp lệ sẽ dùng giá trị mặc định '{fixed_group_val or '(rỗng)'}' cho Nhóm KH."
            )

    col_pv1, col_pv2 = st.columns(2)
    col_pv1.metric("✅ HỢP LỆ (Thành công)", f"{len(v_rows)} dòng")
    col_pv2.metric("❌ TỪ CHỐI (Bị loại)", f"{len(r_rows)} dòng")

    t1, t2 = st.tabs(["✅ Danh sách Hợp Lệ (Mẫu 5 dòng)", f"❌ Chi Tiết {len(r_rows)} Dòng Bị Từ Chối"])
    with t1:
        if v_rows:
            st.dataframe(pd.DataFrame(v_rows[:5]), use_container_width=True)
        else:
            st.warning("Chưa có dòng hợp lệ.")
    with t2:
        if r_rows:
            st.dataframe(pd.DataFrame(r_rows), use_container_width=True)
        else:
            st.success("Không có dòng bị loại!")

    # ==============================================================================
    # XUẤT FILE HOÀN CHỈNH & ĐỐI SOÁT QUÂN SỐ
    # ==============================================================================
    st.markdown("---")
    st.subheader("BƯỚC 3: XÁC NHẬN VÀ XUẤT TOÀN BỘ FILE")
    output_filename = st.text_input("Tên file xuất ra:", value="Ket_Qua_Nhap_Lieu_FinOne.xlsx")

    if st.button("🚀 XÁC NHẬN & GỘP TẤT CẢ SHEET VÀO FILE MẪU", type="primary"):
        template_file = "mau-nhap-lieu-khach-hang.xlsx"
        if not os.path.exists(template_file):
            st.error(f"❌ Không tìm thấy file mẫu [{template_file}]!")
            st.stop()

        wb = openpyxl.load_workbook(template_file)
        ws = wb["Bảng nhập liệu khách hàng"] if "Bảng nhập liệu khách hàng" in wb.sheetnames else wb.active

        total_valid = 0
        total_rejected = 0
        summary_stats = []
        all_rejected_rows = []
        sheets_missing_cols = []
        sheets_missing_group_col = []
        bulk_write_data = []

        for s_name in selected_sheets:
            v_rows, r_rows, hdr_text, col_text, is_missing = process_sheet_data(s_name, file_bytes, config)

            if is_missing:
                sheets_missing_cols.append(s_name)
            elif "Một cột trong bảng" in group_strategy:
                _raw = load_sheet(file_bytes, sheet_name=s_name, header=None, nrows=15)
                _h = current_header_idx if s_name == preview_sheet_ref else auto_detect_header_row_smart(_raw)
                _cols = load_sheet(file_bytes, sheet_name=s_name, header=_h).columns.tolist()
                if not resolve_column_for_sheet(_cols, group_col, "group"):
                    sheets_missing_group_col.append(s_name)

            for row_dict in v_rows:
                bulk_write_data.append([
                    "", row_dict["Nhóm KH (*)"], row_dict["Tên KH (*)"], row_dict["Mã định danh"],
                    row_dict["Điện thoại (*)"], row_dict["Email"], row_dict["Loại KH"],
                    row_dict["Địa chỉ/Ghi chú"], row_dict["Tên doanh nghiệp"]
                ])
                total_valid += 1

            total_rejected += len(r_rows)
            all_rejected_rows.extend(r_rows)
            summary_stats.append({
                "Tên Sheet": s_name,
                "Dòng Header": hdr_text,
                "Cột Tên": col_text,
                "Số dòng Hợp lệ": len(v_rows),
                "Số dòng Bị loại": len(r_rows),
            })

        current_row = 6
        for row_vals in bulk_write_data:
            for col_idx, cell_val in enumerate(row_vals, start=2):
                # An toàn tuyệt đối: không bao giờ ghi NaN/None thô vào Excel
                safe_val = "" if (cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val))) else cell_val
                ws.cell(row=current_row, column=col_idx, value=safe_val)
            current_row += 1

        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        st.success(f"🎉 Đã gộp thành công **{total_valid} khách hàng** từ **{len(selected_sheets)} Sheet**!")

        if sheets_missing_cols:
            st.error(
                "🚨 **CẢNH BÁO: Có Sheet KHÔNG map được cột Tên khách hàng — toàn bộ dữ liệu đã bị loại:**\n\n"
                + "\n".join(f"- {s}" for s in sheets_missing_cols)
            )

        if sheets_missing_group_col:
            st.warning(
                "⚠️ **Có Sheet không tìm thấy cột Nhóm KH đã chọn — các dòng hợp lệ trong sheet này dùng "
                f"giá trị mặc định '{fixed_group_val or '(rỗng)'}' cho Nhóm KH, cần kiểm tra lại:**\n\n"
                + "\n".join(f"- {s}" for s in sheets_missing_group_col)
            )

        # Đối soát tổng quân số
        total_scanned = total_valid + total_rejected
        st.info(
            f"✅ **Đối soát quân số:** Tổng dòng quét = **{total_scanned}** | "
            f"Hợp lệ = **{total_valid}** | Bị loại = **{total_rejected}**"
        )

        tab_final1, tab_final2 = st.tabs(["📥 Tải File Hoàn Chỉnh (FinOne)", "❌ Báo Cáo Tổng Hợp Dòng Bị Loại"])

        with tab_final1:
            st.dataframe(pd.DataFrame(summary_stats), use_container_width=True)
            st.download_button(
                "📥 TẢI FILE EXCEL KẾT QUẢ",
                output_buffer,
                output_filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab_final2:
            if all_rejected_rows:
                df_all_rej = pd.DataFrame(all_rejected_rows)
                st.dataframe(df_all_rej, use_container_width=True)
                rej_buf = io.BytesIO()
                df_all_rej.to_excel(rej_buf, index=False, sheet_name="Dong_Bi_Loai", engine="openpyxl")
                rej_buf.seek(0)
                st.download_button(
                    "📥 TẢI FILE ĐỐI SOÁT DÒNG BỊ LOẠI (.XLSX)",
                    rej_buf,
                    "Tong_Hop_Dong_Bi_Loai.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.success("🎉 Toàn bộ dữ liệu đều hợp lệ 100%!")
