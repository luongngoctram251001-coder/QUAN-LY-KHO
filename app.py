import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ================= CẤU HÌNH TRANG WEB =================
st.set_page_config(page_title="Quản Lý Kho Máy & Điều Phối Online", page_icon="🏥", layout="wide")

SHEET_NAME = "Kho_Thiet_Bi_Online" # Tên file Google Sheets của bạn

# ================= KẾT NỐI GOOGLE SHEETS =================
@st.cache_resource
def init_connection():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # Đọc thông tin xác thực từ Streamlit Secrets (khi chạy online) hoặc file credentials.json (khi chạy thử trên máy)
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME)

try:
    sh = init_connection()
except Exception as e:
    st.error(f"❌ Không thể kết nối Google Sheets '{SHEET_NAME}'. Hãy kiểm tra lại quyền Share hoặc file credentials.json!\nChi tiết lỗi: {e}")
    st.stop()

def khoi_tao_cau_truc_sheet():
    cau_truc = {
        "may_moc": ["ma_may", "ten_may", "nhom", "trang_thai", "ghi_chu"],
        "don_hang": ["ma_don", "doi_tac", "ngay_xuat", "ngay_ve", "trang_thai"],
        "chi_tiet_don": ["id", "ma_don", "ma_may", "da_tra", "ngay_tra_thuc_te"],
        "nhan_su_don_hang": ["id", "ma_don", "vai_tro", "ho_ten", "don_vi_tinh", "so_luong", "don_gia", "thanh_tien"]
    }
    existing_titles = [ws.title for ws in sh.worksheets()]
    for ten_sheet, headers in cau_truc.items():
        if ten_sheet not in existing_titles:
            ws = sh.add_worksheet(title=ten_sheet, rows=1000, cols=15)
            ws.append_row(headers)
        else:
            ws = sh.worksheet(ten_sheet)
            if not ws.row_values(1):
                ws.append_row(headers)

khoi_tao_cau_truc_sheet()

# Hàm đọc dữ liệu nhanh từ Google Sheets ra DataFrame
def doc_du_lieu(ten_sheet):
    ws = sh.worksheet(ten_sheet)
    data = ws.get_all_records()
    if not data:
        headers = ws.row_values(1)
        return pd.DataFrame(columns=headers)
    df = pd.DataFrame(data)
    # Đảm bảo các cột mã luôn là kiểu chuỗi (string)
    for col in ["ma_may", "ma_don", "ngay_xuat", "ngay_ve", "ngay_tra_thuc_te"]:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

# ================= GIAO DIỆN CHÍNH =================
st.title("🏥 HỆ THỐNG QUẢN LÝ KHO MÁY, NHÂN SỰ & ĐIỀU PHỐI ONLINE")

if st.button("🔄 Làm mới dữ liệu mới nhất"):
    st.rerun()

tab0, tab1, tab2, tab3 = st.tabs([
    "0. Tổng Quan Vị Trí Máy",
    "1. Danh Mục Máy Trong Kho",
    "2. Tạo Đơn (Máy & BS/Thư Ký)",
    "3. Theo Dõi, Thu Hồi & Đối Chiếu"
])

df_may = doc_du_lieu("may_moc")
df_don = doc_du_lieu("don_hang")
df_ct = doc_du_lieu("chi_tiet_don")
df_ns = doc_du_lieu("nhan_su_don_hang")

# ================= TAB 0: TỔNG QUAN VỊ TRÍ MÁY =================
with tab0:
    st.subheader("Dashboard Trạng Thái & Vị Trí Thiết Bị Thực Tế")
    today = datetime.now().date()
    
    tong_may = len(df_may)
    tai_kho = len(df_may[df_may["trang_thai"] == "Sẵn sàng"]) if not df_may.empty else 0
    dang_di = len(df_may[df_may["trang_thai"] == "Đang đi đơn"]) if not df_may.empty else 0
    
    # Tính chi tiết vị trí từng máy
    rows_tq = []
    sap_ve_count = 0
    
    if not df_may.empty:
        for _, m in df_may.iterrows():
            ma_m = m["ma_may"]
            vi_tri = "Sẵn sàng tại Kho"
            ngay_ve_str = "-"
            trang_thai_ht = "🟢 Tại kho"
            lich_tiep = "Trống lịch"
            is_sap_ve = False

            # Tìm đơn đang giữ máy
            if not df_ct.empty and not df_don.empty:
                ct_chua_tra = df_ct[(df_ct["ma_may"] == ma_m) & (df_ct["da_tra"].astype(str) == "0")]
                if not ct_chua_tra.empty:
                    ma_dons_lien_quan = ct_chua_tra["ma_don"].tolist()
                    don_dang_di = df_don[(df_don["ma_don"].isin(ma_dons_lien_quan)) & (df_don["trang_thai"] == "Đang đi đơn")]
                    if not don_dang_di.empty:
                        d_row = don_dang_di.iloc[0]
                        vi_tri = f"Đang ở: {d_row['doi_tac']} ({d_row['ma_don']})"
                        ngay_ve_str = d_row["ngay_ve"]
                        trang_thai_ht = "🔴 Đang đi đoàn"
                        try:
                            d_ve = datetime.strptime(ngay_ve_str, "%d/%m/%Y").date()
                            if (d_ve - today).days <= 2:
                                sap_ve_count += 1
                                is_sap_ve = True
                                trang_thai_ht = "🟡 Sắp về kho"
                        except:
                            pass

                    # Tìm đơn đặt trước
                    don_dat_truoc = df_don[(df_don["ma_don"].isin(ma_dons_lien_quan)) & (df_don["trang_thai"] == "Đặt trước / Chờ máy về")]
                    if not don_dat_truoc.empty:
                        dt_row = don_dat_truoc.iloc[0]
                        lich_tiep = f"{dt_row['doi_tac']} (Từ {dt_row['ngay_xuat']})"

            rows_tq.append({
                "Mã Máy / Serial": ma_m,
                "Tên Thiết Bị": m["ten_may"],
                "Nhóm": m["nhom"],
                "Trạng Thái": trang_thai_ht,
                "Vị Trí Hiện Tại": vi_tri,
                "Dự Kiến Về Kho": ngay_ve_str,
                "Lịch Đặt Trước Tiếp Theo": lich_tiep,
                "is_sap_ve": is_sap_ve
            })

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng thiết bị", f"{tong_may} máy")
    c2.metric("Đang tại kho", f"{tai_kho} máy")
    c3.metric("Đang đi đoàn", f"{dang_di} máy")
    c4.metric("Sắp về (<= 2 ngày)", f"{sap_ve_count} máy")

    loc_tq = st.radio("Lọc nhanh:", ["Tất cả máy", "Tại kho", "Đang đi đoàn", "Sắp về kho"], horizontal=True)
    df_tq = pd.DataFrame(rows_tq)
    if not df_tq.empty:
        if loc_tq == "Tại kho":
            df_tq = df_tq[df_tq["Trạng Thái"] == "🟢 Tại kho"]
        elif loc_tq == "Đang đi đoàn":
            df_tq = df_tq[df_tq["Trạng Thái"].isin(["🔴 Đang đi đoàn", "🟡 Sắp về kho"])]
        elif loc_tq == "Sắp về kho":
            df_tq = df_tq[df_tq["is_sap_ve"] == True]
        st.dataframe(df_tq.drop(columns=["is_sap_ve"]), use_container_width=True, hide_index=True)
    else:
        st.info("Chưa có thiết bị nào trong kho. Hãy sang Tab 1 để thêm máy.")

# ================= TAB 1: DANH MỤC MÁY TRONG KHO =================
with tab1:
    st.subheader("Quản Lý Danh Mục Thiết Bị Trong Kho")
    col_add, col_imp = st.columns(2)
    
    with col_add:
        with st.form("form_them_may", clear_on_submit=True):
            st.write("**Thêm thủ công 1 thiết bị:**")
            ma_m = st.text_input("Mã máy / Số Serial *").strip().upper()
            ten_m = st.text_input("Tên thiết bị *").strip()
            nhom_m = st.selectbox("Nhóm thiết bị", ["Siêu âm", "Điện tim", "Đo loãng xương", "Điện não", "X-Quang", "Phụ kiện"])
            gc_m = st.text_input("Phụ kiện đi kèm / Ghi chú").strip()
            if st.form_submit_button("Lưu Thiết Bị Vào Kho"):
                if not ma_m or not ten_m:
                    st.warning("Vui lòng nhập đủ Mã máy và Tên thiết bị!")
                elif not df_may.empty and ma_m in df_may["ma_may"].values:
                    st.error(f"Mã máy '{ma_m}' đã tồn tại!")
                else:
                    sh.worksheet("may_moc").append_row([ma_m, ten_m, nhom_m, "Sẵn sàng", gc_m])
                    st.success(f"Đã thêm máy {ma_m} thành công!")
                    st.rerun()

    with col_imp:
        st.write("**Nhập hàng loạt từ Excel / CSV:**")
        # Tạo file Excel mẫu để tải về
        wb_mau = openpyxl.Workbook()
        ws_mau = wb_mau.active
        ws_mau.append(["ma_may", "ten_may", "nhom", "ghi_chu"])
        ws_mau.append(["CM300-01", "Máy đo loãng xương Furuno CM300", "Đo loãng xương", "Dây nguồn, gel, miếng chuẩn"])
        ws_mau.append(["ECG-1501-01", "Máy điện tim NewTech ECG 1501", "Điện tim", "Dây điện cực, kẹp chi"])
        buf_mau = io.BytesIO()
        wb_mau.save(buf_mau)
        st.download_button("📥 Tải File Excel Mẫu", data=buf_mau.getvalue(), file_name="Mau_Nhap_Kho_May.xlsx")

        file_up = st.file_uploader("Chọn file Excel (.xlsx) hoặc CSV (.csv) để nạp vào kho", type=["xlsx", "csv"])
        if file_up and st.button("🚀 Nạp Danh Sách Lên Google Sheets"):
            df_up = pd.read_excel(file_up) if file_up.name.endswith(".xlsx") else pd.read_csv(file_up)
            ws_may = sh.worksheet("may_moc")
            existing_codes = set(df_may["ma_may"].values) if not df_may.empty else set()
            added = 0
            for _, r in df_up.iterrows():
                code = str(r.iloc[0]).strip().upper()
                name = str(r.iloc[1]).strip()
                grp = str(r.iloc[2]).strip() if len(r) > 2 and pd.notna(r.iloc[2]) else "Khác"
                note = str(r.iloc[3]).strip() if len(r) > 3 and pd.notna(r.iloc[3]) else ""
                if code and name and code not in existing_codes:
                    ws_may.append_row([code, name, grp, "Sẵn sàng", note])
                    existing_codes.add(code)
                    added += 1
            st.success(f"Đã nạp thêm {added} thiết bị mới vào kho online!")
            st.rerun()

    st.markdown("---")
    st.write("**Danh sách toàn bộ thiết bị hiện có:**")
    if not df_may.empty:
        st.dataframe(df_may, use_container_width=True, hide_index=True)

# ================= TAB 2: TẠO ĐƠN (MÁY & NHÂN SỰ BS/THƯ KÝ) =================
with tab2:
    st.subheader("Tạo Đơn Hàng Mới & Điều Phối Gối Đầu")
    
    c_info1, c_info2, c_info3 = st.columns([2, 1, 1])
    doi_tac = c_info1.text_input("Tên Đối Tác / Đơn Vị *")
    ngay_xuat_dt = c_info2.date_input("Ngày xuất kho", datetime.now())
    ngay_ve_dt = c_info3.date_input("Ngày dự kiến về", datetime.now())

    col_left, col_right = st.columns(2)

    # Cột trái: Chọn máy (cả máy rảnh lẫn máy đang đi đoàn)
    with col_left:
        st.write("**1. Chọn thiết bị điều phối:**")
        options_may = []
        map_may_info = {}
        if not df_may.empty:
            for _, m in df_may.iterrows():
                ma_m = m["ma_may"]
                if m["trang_thai"] == "Sẵn sàng":
                    label = f"🟢 [SẴN SÀNG] [{ma_m}] {m['ten_may']}"
                    options_may.append(label)
                    map_may_info[label] = (ma_m, "Sẵn sàng", None, None)
                else:
                    ngay_ve_cu = "Chưa rõ"
                    dt_cu = "Đoàn khác"
                    if not df_ct.empty and not df_don.empty:
                        ct_m = df_ct[(df_ct["ma_may"] == ma_m) & (df_ct["da_tra"].astype(str) == "0")]
                        if not ct_m.empty:
                            d_m = df_don[(df_don["ma_don"].isin(ct_m["ma_don"])) & (df_don["trang_thai"] == "Đang đi đơn")]
                            if not d_m.empty:
                                ngay_ve_cu = d_m.iloc[0]["ngay_ve"]
                                dt_cu = d_m.iloc[0]["doi_tac"]
                    label = f"🟡 [Về {ngay_ve_cu}] [{ma_m}] {m['ten_may']} (Ở: {dt_cu})"
                    options_may.append(label)
                    map_may_info[label] = (ma_m, "Đang đi đơn", ngay_ve_cu, dt_cu)

        selected_may_labels = st.multiselect("Chọn một hoặc nhiều máy:", options_may)

    # Cột phải: Thuê thêm Bác sĩ / Thư ký theo đơn giá riêng
    with col_right:
        st.write("**2. Nhân sự đi kèm (Bác sĩ / Thư ký - Đơn giá riêng theo đoàn):**")
        if "ns_tam" not in st.session_state:
            st.session_state.ns_tam = []

        r1, r2 = st.columns(2)
        vt_ns = r1.selectbox("Vai trò", ["Bác sĩ siêu âm", "Bác sĩ đọc điện tim", "Bác sĩ khám tổng quát", "Thư ký siêu âm", "Thư ký điện tim", "Kỹ thuật viên"])
        ten_ns = r2.text_input("Họ tên (Tùy chọn)")
        
        r3, r4, r5 = st.columns(3)
        dvt_ns = r3.selectbox("Tính theo", ["Ngày", "Buổi"])
        sl_ns = r4.number_input("Số lượng", min_value=0.5, value=1.0, step=0.5)
        gia_ns = r5.number_input("Đơn giá (VNĐ)", min_value=0, value=1000000, step=100000)

        if st.button("➕ Thêm Nhân Sự Vào Đơn"):
            st.session_state.ns_tam.append({
                "Vai Trò": vt_ns,
                "Họ Tên": ten_ns if ten_ns.strip() else "Theo điều động",
                "ĐVT": dvt_ns,
                "SL": sl_ns,
                "Đơn Giá": gia_ns,
                "Thành Tiền": sl_ns * gia_ns
            })

        if st.session_state.ns_tam:
            st.dataframe(pd.DataFrame(st.session_state.ns_tam), use_container_width=True, hide_index=True)
            if st.button("🗑️ Xóa danh sách nhân sự tạm"):
                st.session_state.ns_tam = []
                st.rerun()

    st.markdown("---")
    if st.button("✅ XÁC NHẬN TẠO ĐƠN HÀNG ONLINE", type="primary"):
        if not doi_tac.strip() or (not selected_may_labels and not st.session_state.ns_tam):
            st.warning("Vui lòng nhập Tên đối tác và chọn ít nhất 1 máy hoặc 1 nhân sự!")
        else:
            ngay_xuat_str = ngay_xuat_dt.strftime("%d/%m/%Y")
            ngay_ve_str = ngay_ve_dt.strftime("%d/%m/%Y")
            
            # Tạo mã đơn YYMMDD-XXXX
            prefix = datetime.now().strftime("%y%m%d")
            count_today = len(df_don[df_don["ma_don"].str.startswith(prefix)]) if not df_don.empty else 0
            ma_don_moi = f"{prefix}-{count_today + 1:04d}"

            co_may_dang_di = any(map_may_info[lbl][1] == "Đang đi đơn" for lbl in selected_may_labels)
            trang_thai_don = "Đặt trước / Chờ máy về" if co_may_dang_di else "Đang đi đơn"

            # 1. Ghi đơn hàng
            sh.worksheet("don_hang").append_row([ma_don_moi, doi_tac.strip(), ngay_xuat_str, ngay_ve_str, trang_thai_don])

            # 2. Ghi chi tiết máy & cập nhật trạng thái máy nếu sẵn sàng
            ws_ct = sh.worksheet("chi_tiet_don")
            ws_may = sh.worksheet("may_moc")
            next_ct_id = len(df_ct) + 1
            for lbl in selected_may_labels:
                ma_m, t_thai, _, _ = map_may_info[lbl]
                ws_ct.append_row([next_ct_id, ma_don_moi, ma_m, 0, ""])
                next_ct_id += 1
                if t_thai == "Sẵn sàng":
                    cell = ws_may.find(ma_m, in_column=1)
                    if cell:
                        ws_may.update_cell(cell.row, 4, "Đang đi đơn")

            # 3. Ghi nhân sự đi kèm
            ws_ns = sh.worksheet("nhan_su_don_hang")
            next_ns_id = len(df_ns) + 1
            for ns in st.session_state.ns_tam:
                ws_ns.append_row([
                    next_ns_id, ma_don_moi, ns["Vai Trò"], ns["Họ Tên"],
                    ns["ĐVT"], ns["SL"], ns["Đơn Giá"], ns["Thành Tiền"]
                ])
                next_ns_id += 1

            st.session_state.ns_tam = []
            st.success(f"🎉 Đã tạo đơn {ma_don_moi} ({trang_thai_don}) thành công! Sếp và các bộ phận đã có thể thấy đơn này.")
            st.rerun()

# ================= TAB 3: THEO DÕI, THU HỒI, IN BÀN GIAO & ĐỐI CHIẾU =================
with tab3:
    st.subheader("Quản Lý Đơn Hàng, Thu Hồi Máy & Xuất Báo Cáo")

    if df_don.empty:
        st.info("Chưa có đơn hàng nào trên hệ thống.")
    else:
        st.dataframe(df_don, use_container_width=True, hide_index=True)

        st.markdown("### Thao tác trên từng đơn hàng")
        chon_ma_don = st.selectbox("Chọn Mã Đơn để thao tác:", df_don["ma_don"].tolist())
        don_info = df_don[df_don["ma_don"] == chon_ma_don].iloc[0]

        col_act1, col_act2, col_act3 = st.columns(3)

        # 1. Kích hoạt đơn đặt trước
        with col_act1:
            if don_info["trang_thai"] == "Đặt trước / Chờ máy về":
                if st.button("🚀 Kích Hoạt Xuất Kho Đơn Đặt Trước"):
                    ct_don = df_ct[df_ct["ma_don"] == chon_ma_don]
                    mays = ct_don["ma_may"].tolist()
                    chua_ve = df_may[(df_may["ma_may"].isin(mays)) & (df_may["trang_thai"] != "Sẵn sàng")]["ma_may"].tolist()
                    if chua_ve:
                        st.error(f"Các máy sau chưa trả về kho: {', '.join(chua_ve)}")
                    else:
                        ws_may = sh.worksheet("may_moc")
                        for m in mays:
                            cell = ws_may.find(m, in_column=1)
                            if cell:
                                ws_may.update_cell(cell.row, 4, "Đang đi đơn")
                        ws_don = sh.worksheet("don_hang")
                        cell_d = ws_don.find(chon_ma_don, in_column=1)
                        if cell_d:
                            ws_don.update_cell(cell_d.row, 5, "Đang đi đơn")
                        st.success("Đã kích hoạt xuất kho chính thức!")
                        st.rerun()

        # 2. Thu hồi máy (lẻ hoặc toàn bộ)
        with col_act2:
            if don_info["trang_thai"] == "Đang đi đơn":
                ct_chua_tra = df_ct[(df_ct["ma_don"] == chon_ma_don) & (df_ct["da_tra"].astype(str) == "0")] if not df_ct.empty else pd.DataFrame()
                if ct_chua_tra.empty:
                    if st.button("🏁 Chốt Hoàn Thành Đơn (Chỉ Nhân Sự)"):
                        ws_don = sh.worksheet("don_hang")
                        cell_d = ws_don.find(chon_ma_don, in_column=1)
                        ws_don.update_cell(cell_d.row, 5, "Hoàn thành")
                        st.rerun()
                else:
                    may_chon_tra = st.multiselect("Chọn máy thu hồi về kho hôm nay:", ct_chua_tra["ma_may"].tolist(), default=ct_chua_tra["ma_may"].tolist())
                    if st.button("📥 Xác Nhận Thu Hồi Các Máy Đã Chọn"):
                        ngay_hn = datetime.now().strftime("%d/%m/%Y")
                        ws_ct = sh.worksheet("chi_tiet_don")
                        ws_may = sh.worksheet("may_moc")
                        all_ct = ws_ct.get_all_records()
                        for idx, row in enumerate(all_ct, start=2):
                            if str(row["ma_don"]) == chon_ma_don and str(row["ma_may"]) in may_chon_tra and str(row["da_tra"]) == "0":
                                ws_ct.update_cell(idx, 4, 1)
                                ws_ct.update_cell(idx, 5, ngay_hn)
                                cell_m = ws_may.find(str(row["ma_may"]), in_column=1)
                                if cell_m:
                                    ws_may.update_cell(cell_m.row, 4, "Sẵn sàng")
                        
                        # Kiểm tra xem đã trả hết chưa
                        df_ct_moi = doc_du_lieu("chi_tiet_don")
                        con_lai = len(df_ct_moi[(df_ct_moi["ma_don"] == chon_ma_don) & (df_ct_moi["da_tra"].astype(str) == "0")])
                        if con_lai == 0:
                            ws_don = sh.worksheet("don_hang")
                            cell_d = ws_don.find(chon_ma_don, in_column=1)
                            ws_don.update_cell(cell_d.row, 4, ngay_hn)
                            ws_don.update_cell(cell_d.row, 5, "Hoàn thành")
                        st.success("Đã cập nhật thu hồi thiết bị lên hệ thống!")
                        st.rerun()

        # 3. Xuất Biên Bản Bàn Giao Word
        with col_act3:
            doc = Document()
            p_h = doc.add_paragraph()
            p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_h.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n-------------------\n\n").bold = True
            r_t = p_h.add_run("BIÊN BẢN BÀN GIAO THIẾT BỊ & NHÂN SỰ Y TẾ\n")
            r_t.bold = True
            r_t.font.size = Pt(14)
            p_h.add_run(f"Số đơn: {chon_ma_don} | Ngày xuất: {don_info['ngay_xuat']}").italic = True

            doc.add_paragraph(f"Đơn vị nhận bàn giao: {don_info['doi_tac']} | Dự kiến kết thúc: {don_info['ngay_ve']}")

            # Bảng máy
            mays_don = df_ct[df_ct["ma_don"] == chon_ma_don]["ma_may"].tolist() if not df_ct.empty else []
            if mays_don:
                doc.add_paragraph("I. DANH MỤC THIẾT BỊ BÀN GIAO:").runs[0].bold = True
                t1 = doc.add_table(rows=1, cols=4)
                t1.style = "Table Grid"
                for i, h in enumerate(["STT", "Mã Máy / Serial", "Tên Thiết Bị", "Phụ Kiện / Tình Trạng"]):
                    t1.rows[0].cells[i].text = h
                for idx, mm in enumerate(mays_don, 1):
                    m_row = df_may[df_may["ma_may"] == mm].iloc[0]
                    rc = t1.add_row().cells
                    rc[0].text = str(idx)
                    rc[1].text = mm
                    rc[2].text = m_row["ten_may"]
                    rc[3].text = str(m_row["ghi_chu"]) or "Đầy đủ phụ kiện"

            # Bảng nhân sự
            ns_don = df_ns[df_ns["ma_don"] == chon_ma_don] if not df_ns.empty else pd.DataFrame()
            if not ns_don.empty:
                doc.add_paragraph("\nII. NHÂN SỰ CHUYÊN MÔN ĐI KÈM:").runs[0].bold = True
                t2 = doc.add_table(rows=1, cols=5)
                t2.style = "Table Grid"
                for i, h in enumerate(["STT", "Vai Trò", "Họ Tên", "ĐVT", "Số Lượng"]):
                    t2.rows[0].cells[i].text = h
                for idx, (_, nr) in enumerate(ns_don.iterrows(), 1):
                    rc = t2.add_row().cells
                    rc[0].text = str(idx)
                    rc[1].text = str(nr["vai_tro"])
                    rc[2].text = str(nr["ho_ten"])
                    rc[3].text = str(nr["don_vi_tinh"])
                    rc[4].text = str(nr["so_luong"])

            sig = doc.add_table(rows=1, cols=2)
            sig.rows[0].cells[0].paragraphs[0].add_run("\nĐẠI DIỆN BÊN GIAO\n(Ký, ghi rõ họ tên)").bold = True
            sig.rows[0].cells[1].paragraphs[0].add_run("\nĐẠI DIỆN BÊN NHẬN\n(Ký, ghi rõ họ tên)").bold = True

            buf_doc = io.BytesIO()
            doc.save(buf_doc)
            st.download_button(
                label="📄 Tải Biên Bản Bàn Giao (Word)",
                data=buf_doc.getvalue(),
                file_name=f"BienBanBanGiao_{chon_ma_don}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    st.markdown("---")
    st.subheader("📊 Xuất Bảng Đối Chiếu Cuối Tháng (Excel)")
    thang_dc = st.text_input("Nhập Tháng/Năm cần đối chiếu (MM/YYYY):", datetime.now().strftime("%m/%Y"))
    
    if st.button("Tạo File Excel Đối Chiếu Tổng Hợp"):
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "1_Doi_Chieu_May_Moc"
        ws1.append(["Mã Đơn", "Đối Tác", "Ngày Xuất", "Ngày Thu Hồi", "Số Ngày", "Mã Máy", "Tên Thiết Bị", "Trạng Thái"])

        if not df_don.empty and not df_ct.empty:
            don_thang = df_don[df_don["ngay_xuat"].str.endswith(f"/{thang_dc}")]
            for _, d in don_thang.iterrows():
                cts = df_ct[df_ct["ma_don"] == d["ma_don"]]
                for _, c in cts.iterrows():
                    n_ve = c["ngay_tra_thuc_te"] if str(c["da_tra"]) == "1" and c["ngay_tra_thuc_te"] else d["ngay_ve"]
                    try:
                        so_ngay = max(1, (datetime.strptime(n_ve, "%d/%m/%Y") - datetime.strptime(d["ngay_xuat"], "%d/%m/%Y")).days + 1)
                    except:
                        so_ngay = 1
                    ten_m = df_may[df_may["ma_may"] == c["ma_may"]]["ten_may"].values[0] if not df_may[df_may["ma_may"] == c["ma_may"]].empty else ""
                    tt = "Đã thu hồi" if str(c["da_tra"]) == "1" else d["trang_thai"]
                    ws1.append([d["ma_don"], d["doi_tac"], d["ngay_xuat"], n_ve, so_ngay, c["ma_may"], ten_m, tt])

        ws2 = wb.create_sheet("2_Doi_Chieu_BS_ThuKy")
        ws2.append(["Mã Đơn", "Đối Tác", "Ngày Xuất", "Vai Trò", "Họ Tên", "ĐVT", "Số Lượng", "Đơn Giá", "Thành Tiền"])
        if not df_don.empty and not df_ns.empty:
            don_thang = df_don[df_don["ngay_xuat"].str.endswith(f"/{thang_dc}")]
            for _, d in don_thang.iterrows():
                nss = df_ns[df_ns["ma_don"] == d["ma_don"]]
                for _, n in nss.iterrows():
                    ws2.append([d["ma_don"], d["doi_tac"], d["ngay_xuat"], n["vai_tro"], n["ho_ten"], n["don_vi_tinh"], n["so_luong"], n["don_gia"], n["thanh_tien"]])

        buf_xls = io.BytesIO()
        wb.save(buf_xls)
        st.download_button(
            label=f"📥 Tải File Excel Đối Chiếu Tháng {thang_dc}",
            data=buf_xls.getvalue(),
            file_name=f"DoiChieu_Thang_{thang_dc.replace('/', '-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
