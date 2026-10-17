import datetime
import io
import json
import os
import uuid
import streamlit as st
import pandas as pd
import core_engine
import auth_manager
import storage_manager


def safe_columns(spec, vertical_alignment=None, **kwargs):
    """Bọc an toàn st.columns để tương thích mọi phiên bản Streamlit (kể cả cũ hơn 1.38.0)."""
    try:
        if vertical_alignment:
            return st.columns(spec, vertical_alignment=vertical_alignment, **kwargs)
        return st.columns(spec, **kwargs)
    except TypeError:
        return st.columns(spec, **kwargs)


# ==================== CẤU HÌNH TRANG STREAMLIT ====================
st.set_page_config(
    page_title="Đối Soát BOM Sản Xuất & Tồn Kho (FIFO)",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS GIAO DIỆN HIỆN ĐẠI, TO RÕ & SỐNG ĐỘNG ====================
st.markdown("""
<style>
    /* Google Fonts & Base styling */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Giao diện tràn viền rộng rãi, thoáng đãng */
    .block-container {
        max-width: 98% !important;
        padding-top: 1.0rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    
    /* Ẩn triệt để dấu tăng giảm +/- spin buttons trên các ô số lượng */
    input[type=number]::-webkit-inner-spin-button, 
    input[type=number]::-webkit-outer-spin-button { 
        -webkit-appearance: none !important; 
        margin: 0 !important; 
    }
    input[type=number] {
        -moz-appearance: textfield !important;
    }
    button[data-testid="stNumberInputStepDown"],
    button[data-testid="stNumberInputStepUp"] {
        display: none !important;
    }

    /* Hero Banner hiện đại, cao cấp */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 18px;
        padding: 22px 30px;
        color: #FFFFFF;
        box-shadow: 0 12px 30px -5px rgba(30, 58, 138, 0.3), 0 8px 12px -6px rgba(30, 58, 138, 0.15);
        margin-bottom: 20px;
        position: relative;
        overflow: hidden;
    }
    .hero-banner::after {
        content: '';
        position: absolute;
        top: -40px;
        right: -40px;
        width: 220px;
        height: 220px;
        background: radial-gradient(circle, rgba(255,255,255,0.14) 0%, rgba(255,255,255,0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }
    .hero-title {
        font-size: 26px;
        font-weight: 850;
        letter-spacing: -0.5px;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-desc {
        font-size: 14px;
        color: #E2E8F0;
        font-weight: 450;
        line-height: 1.5;
    }
    .hero-tag {
        display: inline-block;
        background: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        padding: 4px 14px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 700;
        margin-top: 10px;
        border: 1px solid rgba(255, 255, 255, 0.25);
    }

    /* KPI Metric Cards TO, SỐNG ĐỘNG */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 16px;
        margin-bottom: 24px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1.5px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        position: relative;
        overflow: hidden;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px rgba(0, 0, 0, 0.08);
    }
    .kpi-card-border-blue { border-top: 6px solid #2563EB; }
    .kpi-card-border-emerald { 
        border-top: 6px solid #059669; 
        background: linear-gradient(180deg, #FFFFFF 0%, #F0FDF4 100%);
        border-color: #86EFAC;
    }
    .kpi-card-border-green { border-top: 6px solid #10B981; }
    .kpi-card-border-purple { border-top: 6px solid #8B5CF6; }
    .kpi-card-border-amber { border-top: 6px solid #F59E0B; }

    .kpi-label {
        font-size: 12px;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 900;
        color: #0F172A;
        line-height: 1.1;
        margin-top: 6px;
    }
    .kpi-unit {
        font-size: 14px;
        font-weight: 600;
        color: #64748B;
        margin-left: 4px;
    }

    /* Hộp nhập liệu TO & NỔI BẬT */
    .floating-box {
        background: #FFFFFF;
        border: 2px solid #2563EB;
        border-radius: 16px;
        padding: 22px 26px;
        box-shadow: 0 12px 30px -5px rgba(37, 99, 235, 0.14);
        margin-bottom: 20px;
    }
    .box-header {
        font-size: 19px;
        font-weight: 800;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 14px;
    }

    /* Nút chính to bản */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: white !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 10px !important;
        padding: 8px 20px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.45) !important;
    }

    /* Thẻ tóm tắt đơn hàng kết quả */
    .order-card {
        background: #FFFFFF;
        border: 1.5px solid #E2E8F0;
        border-radius: 14px;
        padding: 18px 22px;
        margin-bottom: 14px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.03);
    }
    .order-card-ok {
        border-left: 6px solid #10B981;
    }
    .order-card-missing {
        border-left: 6px solid #EF4444;
    }

    /* Bảng HTML đối soát chuyên nghiệp */
    .styled-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 13px;
        border-radius: 10px;
        overflow: hidden;
        margin-top: 10px;
        margin-bottom: 14px;
        border: 1px solid #CBD5E1;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
    .styled-table thead tr {
        background: linear-gradient(135deg, #1E293B 0%, #1E3A8A 100%);
        color: #FFFFFF;
        text-align: left;
        font-weight: 700;
        font-size: 12.5px;
    }
    .styled-table th {
        padding: 10px 12px;
        border-right: 1px solid rgba(255, 255, 255, 0.1);
        vertical-align: middle;
    }
    .styled-table td {
        padding: 8px 12px;
        border-bottom: 1px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
        vertical-align: middle;
    }
    .styled-table tbody tr:last-child td {
        border-bottom: none;
    }
    .styled-table tbody tr.row-ok {
        background-color: #F0FDF4;
        color: #166534;
    }
    .styled-table tbody tr.row-missing {
        background-color: #FEF2F2;
        color: #991B1B;
        font-weight: 500;
    }
    .styled-table tbody tr:hover {
        filter: brightness(0.96);
        transition: filter 0.15s ease;
    }

    /* Badges trạng thái TO & RÕ */
    .badge-ok {
        background-color: #DCFCE7;
        color: #15803D;
        border: 1px solid #86EFAC;
        padding: 4px 10px;
        border-radius: 9999px;
        font-weight: 800;
        font-size: 11.5px;
        display: inline-block;
    }
    .badge-missing {
        background-color: #FEE2E2;
        color: #B91C1C;
        border: 1px solid #FCA5A5;
        padding: 4px 10px;
        border-radius: 9999px;
        font-weight: 800;
        font-size: 11.5px;
        display: inline-block;
    }
    .badge-order {
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 800;
        font-size: 13.5px;
        display: inline-block;
    }
    .badge-role {
        padding: 3px 10px;
        border-radius: 9999px;
        font-weight: 800;
        font-size: 12px;
        display: inline-block;
    }

    /* Login Card UI */
    .login-container {
        max-width: 900px;
        margin: 40px auto;
        background: #FFFFFF;
        border-radius: 20px;
        padding: 36px 40px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.08);
        border: 1px solid #E2E8F0;
    }
    .quick-role-card {
        border-radius: 12px;
        padding: 16px;
        border: 1.5px solid #E2E8F0;
        background: #F8FAFC;
        text-align: center;
        transition: all 0.2s ease;
    }
    .quick-role-card:hover {
        border-color: #2563EB;
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(37,99,235,0.12);
    }
</style>
""", unsafe_allow_html=True)

# Khởi tạo Auth & Storage
auth_manager.init_auth_system()

# Quản lý Session State
if 'current_user' not in st.session_state:
    # Mặc định tự động đăng nhập tài khoản planner để sẵn sàng dùng ngay hoặc chọn đổi tài khoản
    st.session_state.current_user = None

if 'workspace_mode' not in st.session_state:
    st.session_state.workspace_mode = storage_manager.WORKSPACE_PRIVATE

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PERSIST_DIR = os.path.join(APP_DIR, ".persistent_data")
os.makedirs(DATA_PERSIST_DIR, exist_ok=True)
STOCK_PERSIST_FILE = os.path.join(DATA_PERSIST_DIR, "custom_stock.xlsx")
BOM_PERSIST_FILE = os.path.join(DATA_PERSIST_DIR, "custom_bom.xlsx")
PLAN_PERSIST_FILE = os.path.join(DATA_PERSIST_DIR, "custom_plan.xlsx")
META_PERSIST_FILE = os.path.join(DATA_PERSIST_DIR, "metadata.json")

def load_meta():
    return storage_manager.load_shared_meta()

def save_meta(update_dict):
    storage_manager.update_shared_meta(update_dict)

meta = load_meta()


# ==================== MÀN HÌNH ĐĂNG NHẬP (KHI CHƯA XÁC THỰC) ====================
if st.session_state.current_user is None:
    st.markdown("""
    <div style="text-align: center; margin-top: 20px; margin-bottom: 25px;">
        <h1 style="color: #1E3A8A; font-size: 32px; font-weight: 900; margin-bottom: 8px;">
            📦 HỆ THỐNG ĐỐI SOÁT BOM SẢN XUẤT & TỒN KHO (FIFO)
        </h1>
        <p style="color: #64748B; font-size: 16px; max-width: 750px; margin: 0 auto;">
            Hệ thống quản lý đối soát vật tư đa người dùng. Vui lòng đăng nhập hoặc chọn <b>Đăng nhập nhanh 1 chạm</b> theo phân quyền công việc của bạn.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_l1, col_l2, col_l3 = safe_columns([1, 8, 1])
    with col_l2:
        st.markdown("""
        <div style="background:#EFF6FF; border: 1.5px solid #BFDBFE; border-radius: 14px; padding: 16px 22px; margin-bottom: 20px;">
            <div style="font-weight: 800; color: #1D4ED8; font-size: 15px; margin-bottom: 4px;">
                ⚡ ĐĂNG NHẬP NHANH 1 CHẠM (DÀNH CHO TRẢI NGHIỆM & KIỂM THỬ)
            </div>
            <div style="font-size: 13px; color: #475569;">
                Bấm vào một trong các vai trò dưới đây để vào ngay hệ thống với đúng phân quyền tương ứng:
            </div>
        </div>
        """, unsafe_allow_html=True)

        c_r1, c_r2, c_r3, c_r4 = safe_columns(4)
        with c_r1:
            st.markdown("""
            <div class="quick-role-card" style="border-top: 4px solid #DC2626;">
                <div style="font-size: 18px;">👑</div>
                <div style="font-weight: 850; color: #DC2626; font-size: 14px; margin-top: 4px;">Quản Trị Viên</div>
                <div style="font-size: 11.5px; color: #64748B; margin-top: 2px;">Tài khoản: <b>admin</b></div>
                <div style="font-size: 11px; color: #475569; margin-top: 6px;">Toàn quyền hệ thống & quản lý tài khoản</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🚀 Vào với quyền Admin", key="btn_quick_admin", use_container_width=True, type="primary"):
                ok, u_info, _ = auth_manager.authenticate("admin", "admin123")
                if ok:
                    st.session_state.current_user = u_info
                    st.rerun()

        with c_r2:
            st.markdown("""
            <div class="quick-role-card" style="border-top: 4px solid #2563EB;">
                <div style="font-size: 18px;">📋</div>
                <div style="font-weight: 850; color: #2563EB; font-size: 14px; margin-top: 4px;">Kế Hoạch & ĐĐSX</div>
                <div style="font-size: 11.5px; color: #64748B; margin-top: 2px;">Tài khoản: <b>planner</b></div>
                <div style="font-size: 11px; color: #475569; margin-top: 6px;">Quản lý hàng chờ FIFO & xuất Excel</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🚀 Vào với quyền Planner", key="btn_quick_plan", use_container_width=True, type="primary"):
                ok, u_info, _ = auth_manager.authenticate("planner", "plan123")
                if ok:
                    st.session_state.current_user = u_info
                    st.rerun()

        with c_r3:
            st.markdown("""
            <div class="quick-role-card" style="border-top: 4px solid #059669;">
                <div style="font-size: 18px;">🏬</div>
                <div style="font-weight: 850; color: #059669; font-size: 14px; margin-top: 4px;">Thủ Kho & Vật Tư</div>
                <div style="font-size: 11.5px; color: #64748B; margin-top: 2px;">Tài khoản: <b>warehouse</b></div>
                <div style="font-size: 11px; color: #475569; margin-top: 6px;">Cập nhật Tồn kho & kiểm tra mã thiếu</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🚀 Vào với quyền Kho", key="btn_quick_ware", use_container_width=True, type="primary"):
                ok, u_info, _ = auth_manager.authenticate("warehouse", "kho123")
                if ok:
                    st.session_state.current_user = u_info
                    st.rerun()

        with c_r4:
            st.markdown("""
            <div class="quick-role-card" style="border-top: 4px solid #64748B;">
                <div style="font-size: 18px;">👁️</div>
                <div style="font-weight: 850; color: #475569; font-size: 14px; margin-top: 4px;">Người Xem / Công Nhân</div>
                <div style="font-size: 11.5px; color: #64748B; margin-top: 2px;">Tài khoản: <b>viewer</b></div>
                <div style="font-size: 11px; color: #475569; margin-top: 6px;">Chỉ xem báo cáo & tải file Excel</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🚀 Vào với quyền Viewer", key="btn_quick_view", use_container_width=True):
                ok, u_info, _ = auth_manager.authenticate("viewer", "view123")
                if ok:
                    st.session_state.current_user = u_info
                    st.rerun()

        st.markdown("<hr style='margin: 24px 0; border: none; border-top: 1px dashed #CBD5E1;'>", unsafe_allow_html=True)
        
        # Form đăng nhập tài khoản tùy chỉnh
        with st.form("custom_login_form"):
            st.markdown("#### 🔐 Hoặc Đăng Nhập Với Tài Khoản Riêng:")
            col_inp_u, col_inp_p = safe_columns(2)
            with col_inp_u:
                username_input = st.text_input("Tên đăng nhập:", placeholder="Ví dụ: admin, planner...")
            with col_inp_p:
                password_input = st.text_input("Mật khẩu:", type="password", placeholder="Nhập mật khẩu...")
            btn_submit_login = st.form_submit_button("Đăng Nhập", type="primary", use_container_width=True)
            if btn_submit_login:
                ok, user_res, err_msg = auth_manager.authenticate(username_input, password_input)
                if ok:
                    st.session_state.current_user = user_res
                    st.success(f"Xin chào {user_res['full_name']}!")
                    st.rerun()
                else:
                    st.error(f"⚠️ {err_msg}")
    st.stop()


# ==================== ĐÃ ĐĂNG NHẬP: KHỞI TẠO DỮ LIỆU & QUYỀN HẠN ====================
current_user = st.session_state.current_user
username = current_user['username']
user_role = current_user['role']
user_fullname = current_user['full_name']
user_permissions = current_user.get('permissions', {})
role_meta = auth_manager.ROLES_METADATA.get(user_role, auth_manager.ROLES_METADATA[auth_manager.ROLE_VIEWER])

# Khởi tạo hàng chờ theo Không gian làm việc
if 'orders_queue' not in st.session_state:
    st.session_state.orders_queue = storage_manager.load_queue(username, st.session_state.workspace_mode)

if 'custom_bom_bytes' not in st.session_state:
    if os.path.exists(BOM_PERSIST_FILE):
        try:
            with open(BOM_PERSIST_FILE, 'rb') as f:
                st.session_state.custom_bom_bytes = f.read()
            st.session_state.custom_bom_name = meta.get('bom_name', os.path.basename(BOM_PERSIST_FILE))
        except Exception:
            st.session_state.custom_bom_bytes = None
            st.session_state.custom_bom_name = None
    else:
        st.session_state.custom_bom_bytes = None
        st.session_state.custom_bom_name = None

if 'custom_stock_bytes' not in st.session_state:
    if os.path.exists(STOCK_PERSIST_FILE):
        try:
            with open(STOCK_PERSIST_FILE, 'rb') as f:
                st.session_state.custom_stock_bytes = f.read()
            st.session_state.custom_stock_name = meta.get('stock_name', os.path.basename(STOCK_PERSIST_FILE))
        except Exception:
            st.session_state.custom_stock_bytes = None
            st.session_state.custom_stock_name = None
    else:
        st.session_state.custom_stock_bytes = None
        st.session_state.custom_stock_name = None

if 'plan_cleared' not in st.session_state:
    st.session_state.plan_cleared = meta.get('plan_cleared', False)

if 'custom_plan_bytes' not in st.session_state:
    if not st.session_state.plan_cleared and os.path.exists(PLAN_PERSIST_FILE):
        try:
            with open(PLAN_PERSIST_FILE, 'rb') as f:
                st.session_state.custom_plan_bytes = f.read()
            st.session_state.custom_plan_name = meta.get('plan_name', os.path.basename(PLAN_PERSIST_FILE))
        except Exception:
            st.session_state.custom_plan_bytes = None
            st.session_state.custom_plan_name = None
    else:
        st.session_state.custom_plan_bytes = None
        st.session_state.custom_plan_name = None

if 'quick_selected_order' not in st.session_state:
    st.session_state.quick_selected_order = "8393"
if 'quick_selected_qty' not in st.session_state:
    st.session_state.quick_selected_qty = "100"


def find_data_file(keywords):
    """Tìm file dữ liệu tự động theo từ khóa trong thư mục script, thư mục làm việc và thư mục cha."""
    search_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.getcwd(),
        os.path.join(os.getcwd(), ".."),
    ]
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        try:
            for fname in os.listdir(d):
                if not fname.endswith(('.xlsx', '.xls')):
                    continue
                fname_lower = fname.lower()
                if all(kw.lower() in fname_lower for kw in keywords):
                    return os.path.join(d, fname)
        except Exception:
            pass
    return None


# ==================== HÀM LOAD DỮ LIỆU CÓ CACHE ====================
@st.cache_data
def get_loaded_data(bom_bytes, bom_name, stock_bytes, stock_name):
    if bom_bytes:
        bom_source = io.BytesIO(bom_bytes)
    else:
        bom_source = find_data_file(['production', 'bom']) or find_data_file(['bom'])

    if stock_bytes:
        stock_source = io.BytesIO(stock_bytes)
    else:
        stock_source = (
            find_data_file(['stock', '61']) 
            or find_data_file(['stock', 'balance']) 
            or find_data_file(['stock'])
        )

    empty_readiness = {'ready_orders': [], 'not_ready_orders': [], 'total_orders': 0, 'ready_count': 0, 'not_ready_count': 0, 'ready_percent': 0.0}

    if bom_source is None or stock_source is None:
        missing = []
        if bom_source is None:
            missing.append("Production BOM (.xlsx)")
        if stock_source is None:
            missing.append("Tồn kho (Stock 61 / Stock Balance) (.xlsx)")
        return None, None, {}, empty_readiness, f"Chưa tìm thấy file mẫu: {', '.join(missing)}"

    try:
        df_bom = core_engine.load_bom_data(bom_source)
        df_stock = core_engine.load_stock_data(stock_source)
        orders_info = core_engine.get_available_orders_info(df_bom)
        readiness_data = core_engine.analyze_orders_readiness(df_bom, df_stock)
        return df_bom, df_stock, orders_info, readiness_data, None
    except Exception as e:
        return None, None, {}, empty_readiness, f"Lỗi đọc dữ liệu Excel: {str(e)}"

@st.cache_data
def get_loaded_plan_data(plan_bytes, plan_name, is_cleared=False):
    if is_cleared:
        return None, None
    if plan_bytes:
        plan_source = io.BytesIO(plan_bytes)
    else:
        plan_source = (
            find_data_file(['kế hoạch'])
            or find_data_file(['ke hoach'])
            or find_data_file(['plan'])
        )
    if plan_source is None:
        return None, None
    try:
        df_plan = core_engine.load_plan_data(plan_source)
        source_name = plan_name if plan_bytes else os.path.basename(plan_source)
        return df_plan, source_name
    except Exception:
        return None, None

@st.cache_data
def get_cached_readiness_excel(cache_key, _readiness_data):
    return core_engine.export_readiness_report_to_excel(_readiness_data)

@st.cache_data
def get_cached_single_order_excel(cache_key, _order_result):
    return core_engine.export_order_to_excel(_order_result)

@st.cache_data
def get_cached_all_excel(cache_key, _calculation_results):
    return core_engine.export_all_orders_to_excel(_calculation_results)


# ==================== SIDEBAR QUẢN LÝ TÀI KHOẢN & DỮ LIỆU NGUỒN ====================
with st.sidebar:
    # 1. Thẻ thông tin người dùng đang đăng nhập
    st.markdown(f"""
    <div style="background: {role_meta['bg_color']}; border: 1.5px solid {role_meta['badge_color']}; border-radius: 12px; padding: 12px 14px; margin-bottom: 12px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <span style="font-weight: 850; color: #0F172A; font-size: 14px;">👤 {user_fullname}</span>
            <span class="badge-role" style="background:{role_meta['badge_color']}; color:white;">{role_meta['name'].split('(')[0].strip()}</span>
        </div>
        <div style="font-size: 12px; color: #475569; margin-top: 4px;">
            Bộ phận: <b>{current_user.get('department', 'Sản xuất')}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col_sb_act1, col_sb_act2 = safe_columns(2)
    with col_sb_act1:
        with st.popover("🔑 Đổi mật khẩu", use_container_width=True):
            st.markdown("##### Đổi mật khẩu cá nhân")
            curr_pw = st.text_input("Mật khẩu hiện tại:", type="password", key="chg_curr_pw")
            new_pw = st.text_input("Mật khẩu mới:", type="password", key="chg_new_pw")
            new_pw_cf = st.text_input("Xác nhận MK mới:", type="password", key="chg_new_cf")
            if st.button("Lưu mật khẩu mới", use_container_width=True, type="primary"):
                if new_pw != new_pw_cf:
                    st.error("Xác nhận mật khẩu mới không khớp.")
                else:
                    chg_ok, chg_msg = auth_manager.change_own_password(username, curr_pw, new_pw)
                    if chg_ok:
                        st.success(chg_msg)
                    else:
                        st.error(chg_msg)
    with col_sb_act2:
        if st.button("🚪 Đăng xuất", use_container_width=True):
            auth_manager.log_activity(username, "LOGOUT", f"Người dùng {user_fullname} đã đăng xuất")
            st.session_state.current_user = None
            st.rerun()

    st.divider()

    # 2. Điều khiển Không gian làm việc (Private vs Shared Workspace)
    st.markdown("#### 🌐 Không Gian Làm Việc (Workspace)")
    current_mode = st.session_state.workspace_mode
    mode_options = [
        ("private", "👤 Không gian Cá nhân (Riêng tư)"),
        ("shared", "👥 Không gian Đội nhóm (Chung)")
    ]
    mode_choice = st.radio(
        "Chế độ hiển thị hàng chờ:",
        options=[m[0] for m in mode_options],
        format_func=lambda x: next(m[1] for m in mode_options if m[0] == x),
        index=0 if current_mode == "private" else 1,
        help="Không gian cá nhân giúp bạn thoải mái thử nghiệm đối soát độc lập, không làm xáo trộn màn hình người khác!"
    )
    if mode_choice != st.session_state.workspace_mode:
        st.session_state.workspace_mode = mode_choice
        st.session_state.orders_queue = storage_manager.load_queue(username, mode_choice)
        st.rerun()

    # Nút đồng bộ hàng chờ
    if user_permissions.get('can_manage_queue', False):
        col_sync1, col_sync2 = safe_columns(2)
        with col_sync1:
            if st.button("📤 Xuất bản Chung", help="Đưa hàng chờ cá nhân của bạn thành hàng chờ chung cho toàn nhà máy", use_container_width=True):
                s_ok, s_msg = storage_manager.copy_queue(storage_manager.WORKSPACE_PRIVATE, storage_manager.WORKSPACE_SHARED, username)
                if s_ok:
                    auth_manager.log_activity(username, "SYNC_QUEUE_PUBLISH", "Xuất bản hàng chờ cá nhân lên hàng chờ chung")
                    st.toast(s_msg, icon="✅")
                    if st.session_state.workspace_mode == storage_manager.WORKSPACE_SHARED:
                        st.session_state.orders_queue = storage_manager.load_queue(username, storage_manager.WORKSPACE_SHARED)
                        st.rerun()
        with col_sync2:
            if st.button("📥 Sao chép về", help="Sao chép hàng chờ chung về không gian cá nhân của bạn để chỉnh sửa độc lập", use_container_width=True):
                s_ok, s_msg = storage_manager.copy_queue(storage_manager.WORKSPACE_SHARED, storage_manager.WORKSPACE_PRIVATE, username)
                if s_ok:
                    auth_manager.log_activity(username, "SYNC_QUEUE_PULL", "Sao chép hàng chờ chung về không gian cá nhân")
                    st.toast(s_msg, icon="✅")
                    if st.session_state.workspace_mode == storage_manager.WORKSPACE_PRIVATE:
                        st.session_state.orders_queue = storage_manager.load_queue(username, storage_manager.WORKSPACE_PRIVATE)
                        st.rerun()

    st.divider()

    # 3. Quản lý Dữ liệu Nguồn có Phân Quyền
    st.markdown("### ⚙️ Dữ Liệu Nguồn (Upload)")
    
    # Upload Tồn kho (Chỉ Admin & Warehouse)
    st.markdown("#### 1. File Tồn kho (Stock 61 / Stock Balance)")
    if user_permissions.get('can_upload_stock', False):
        uploaded_stock = st.file_uploader(
            "Tải lên file Tồn kho mới",
            type=["xlsx", "xls"],
            key="uploader_stock",
            help="Hỗ trợ file Stock 61 (Item code, Wh, Qty) hoặc Stock Balance (Stock Code, Qty...)"
        )
        if uploaded_stock is not None:
            stock_bytes_val = uploaded_stock.getvalue()
            if (st.session_state.custom_stock_bytes != stock_bytes_val) or (st.session_state.custom_stock_name != uploaded_stock.name):
                st.session_state.custom_stock_bytes = stock_bytes_val
                st.session_state.custom_stock_name = uploaded_stock.name
                try:
                    with open(STOCK_PERSIST_FILE, 'wb') as f:
                        f.write(stock_bytes_val)
                    save_meta({'stock_name': uploaded_stock.name})
                    auth_manager.log_activity(username, "UPLOAD_STOCK", f"Tải lên file tồn kho mới: {uploaded_stock.name}")
                except Exception:
                    pass
                st.cache_data.clear()
                st.rerun()
    else:
        st.info("🔒 Quyền **Thủ kho** hoặc **Admin** mới được nạp file Tồn kho mới.")

    st.divider()
    
    # Upload Production BOM (Chỉ Admin & Planner)
    st.markdown("#### 2. File Định mức (Production BOM)")
    if user_permissions.get('can_upload_bom', False):
        uploaded_bom = st.file_uploader(
            "Tải lên file Production BOM mới",
            type=["xlsx", "xls"],
            key="uploader_bom",
            help="Cột quan trọng: Product Code, Component Code, Qty"
        )
        if uploaded_bom is not None:
            bom_bytes_val = uploaded_bom.getvalue()
            if (st.session_state.custom_bom_bytes != bom_bytes_val) or (st.session_state.custom_bom_name != uploaded_bom.name):
                st.session_state.custom_bom_bytes = bom_bytes_val
                st.session_state.custom_bom_name = uploaded_bom.name
                try:
                    with open(BOM_PERSIST_FILE, 'wb') as f:
                        f.write(bom_bytes_val)
                    save_meta({'bom_name': uploaded_bom.name})
                    auth_manager.log_activity(username, "UPLOAD_BOM", f"Tải lên file BOM mới: {uploaded_bom.name}")
                except Exception:
                    pass
                st.cache_data.clear()
                st.rerun()
    else:
        st.info("🔒 Quyền **Kế hoạch** hoặc **Admin** mới được nạp file BOM mới.")

    st.divider()

    # Upload Kế hoạch sản xuất (Chỉ Admin & Planner)
    st.markdown("#### 3. File Kế hoạch sản xuất")
    if user_permissions.get('can_upload_plan', False):
        uploaded_plan = st.file_uploader(
            "Tải lên file Kế hoạch mới",
            type=["xlsx", "xls"],
            key="uploader_plan",
            help="Cột quan trọng: MODEL và Qty (và Due date nếu có)"
        )
        if uploaded_plan is not None:
            plan_bytes_val = uploaded_plan.getvalue()
            if (st.session_state.custom_plan_bytes != plan_bytes_val) or (st.session_state.custom_plan_name != uploaded_plan.name):
                st.session_state.custom_plan_bytes = plan_bytes_val
                st.session_state.custom_plan_name = uploaded_plan.name
                st.session_state.plan_cleared = False
                try:
                    with open(PLAN_PERSIST_FILE, 'wb') as f:
                        f.write(plan_bytes_val)
                    save_meta({'plan_name': uploaded_plan.name, 'plan_cleared': False})
                    auth_manager.log_activity(username, "UPLOAD_PLAN", f"Tải lên file Kế hoạch mới: {uploaded_plan.name}")
                except Exception:
                    pass
                st.cache_data.clear()
                st.rerun()

        # Nút xóa dữ liệu Kế hoạch trên sidebar
        has_plan_active = not st.session_state.get('plan_cleared', False)
        if has_plan_active:
            if st.button("🗑️ Xóa dữ liệu Kế hoạch", key="btn_clear_plan_sidebar", use_container_width=True, help="Xóa dữ liệu file kế hoạch sản xuất để hệ thống tự động tính tồn kho đối soát các đơn đồng bộ"):
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = True
                if os.path.exists(PLAN_PERSIST_FILE):
                    try:
                        os.remove(PLAN_PERSIST_FILE)
                    except Exception:
                        pass
                save_meta({'plan_name': None, 'plan_cleared': True})
                st.session_state.orders_queue = [
                    item for item in st.session_state.orders_queue
                    if item.get('source') != 'Kế hoạch' and not str(item.get('id', '')).startswith('plan_')
                ]
                storage_manager.save_queue(st.session_state.orders_queue, username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "CLEAR_PLAN", f"Người dùng {user_fullname} xóa dữ liệu file Kế hoạch sản xuất")
                st.cache_data.clear()
                st.toast("✅ Đã xóa dữ liệu file Kế hoạch!", icon="🗑️")
                st.rerun()
        else:
            st.caption("ℹ️ Đã xóa dữ liệu Kế hoạch. Hệ thống đang tự động tính tồn kho đối soát những đơn đồng bộ chạy được.")
            if st.button("🔄 Nạp lại Kế hoạch mặc định", key="btn_restore_default_plan", use_container_width=True):
                st.session_state.plan_cleared = False
                save_meta({'plan_cleared': False})
                auth_manager.log_activity(username, "RESTORE_PLAN", f"Người dùng {user_fullname} khôi phục nạp file Kế hoạch mặc định")
                st.cache_data.clear()
                st.toast("✅ Đã nạp lại file Kế hoạch mặc định!", icon="✅")
                st.rerun()
    else:
        st.info("🔒 Quyền **Kế hoạch** hoặc **Admin** mới được nạp file Kế hoạch mới.")


# Tải dữ liệu thực tế
active_stock_bytes = st.session_state.custom_stock_bytes
active_stock_name = st.session_state.custom_stock_name
active_bom_bytes = st.session_state.custom_bom_bytes
active_bom_name = st.session_state.custom_bom_name
active_plan_bytes = st.session_state.custom_plan_bytes
active_plan_name = st.session_state.custom_plan_name

df_bom, df_stock, orders_info, readiness_data, load_error = get_loaded_data(
    active_bom_bytes,
    active_bom_name,
    active_stock_bytes,
    active_stock_name
)

df_plan, plan_file_display_name = get_loaded_plan_data(
    active_plan_bytes,
    active_plan_name,
    is_cleared=st.session_state.get('plan_cleared', False)
)

# Hiển thị tóm tắt dữ liệu ở Sidebar
with st.sidebar:
    st.markdown("---")
    st.markdown("#### 📊 Dữ liệu hiện tại:")
    default_stock_path = find_data_file(['stock', '61']) or find_data_file(['stock', 'balance']) or find_data_file(['stock'])
    default_stock_label = f"{os.path.basename(default_stock_path)} (Mặc định)" if default_stock_path else "Stock 61.xlsx (Mặc định)"
    current_stock_name = active_stock_name or default_stock_label
    if df_stock is not None:
        st.caption(f"📁 **{current_stock_name}** ({len(df_stock):,} dòng, {len(df_stock['Stock_Code'].unique()):,} mã)")
    else:
        st.caption(f"📁 **{current_stock_name}** (Chưa nạp)")

    current_bom_name = active_bom_name or "Production BOM List 10.2.xlsx (Mặc định)"
    if df_bom is not None:
        st.caption(f"📁 **{current_bom_name}** ({len(df_bom):,} dòng, {len(orders_info):,} đơn hàng)")
    else:
        st.caption(f"📁 **{current_bom_name}** (Chưa nạp)")

    current_plan_name = active_plan_name or plan_file_display_name or "Kế hoạch.xlsx (Mặc định)"
    if df_plan is not None:
        plan_orders_temp = core_engine.extract_orders_from_plan(df_plan)
        st.caption(f"📁 **{current_plan_name}** ({len(df_plan):,} dòng, {len(plan_orders_temp):,} đơn)")
    else:
        st.caption(f"📁 **File Kế hoạch** (Đã xóa dữ liệu / Chưa nạp)")

    # Nút khôi phục dữ liệu gốc (chỉ Admin)
    if user_permissions.get('can_reset_data', False):
        if st.session_state.custom_bom_bytes or st.session_state.custom_stock_bytes or st.session_state.custom_plan_bytes or st.session_state.get('plan_cleared', False) or os.path.exists(STOCK_PERSIST_FILE) or os.path.exists(BOM_PERSIST_FILE) or os.path.exists(PLAN_PERSIST_FILE):
            st.divider()
            if st.button("🔄 Khôi phục dữ liệu mẫu gốc", use_container_width=True):
                st.session_state.custom_bom_bytes = None
                st.session_state.custom_bom_name = None
                st.session_state.custom_stock_bytes = None
                st.session_state.custom_stock_name = None
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = False
                st.session_state.orders_queue = []
                for p_file in [STOCK_PERSIST_FILE, BOM_PERSIST_FILE, PLAN_PERSIST_FILE, META_PERSIST_FILE]:
                    if os.path.exists(p_file):
                        try: os.remove(p_file)
                        except Exception: pass
                storage_manager.save_queue([], username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "RESET_DATA", "Khôi phục dữ liệu mẫu mặc định của toàn hệ thống")
                st.cache_data.clear()
                st.success("Đã khôi phục dữ liệu mẫu mặc định!")
                st.rerun()


# ==================== MAIN BANNER ====================
workspace_title = "Không Gian Cá Nhân (Riêng Tư)" if st.session_state.workspace_mode == storage_manager.WORKSPACE_PRIVATE else "Không Gian Đội Nhóm (Chung Nhà Máy)"
st.markdown(f"""
<div class="hero-banner">
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
        <div class="hero-title">📦 HỆ THỐNG ĐỐI SOÁT BOM SẢN XUẤT & TỒN KHO</div>
        <div style="background: rgba(255,255,255,0.2); padding: 5px 14px; border-radius: 9999px; font-size: 13px; font-weight: 800; border: 1px solid rgba(255,255,255,0.3);">
            👤 {user_fullname} | {role_meta['name'].split('(')[0]} | 🌐 {workspace_title}
        </div>
    </div>
    <div class="hero-desc">
        Tự động tính định mức nguyên đơn <b>(Trái + Phải)</b> | Phân bổ trừ tồn kho tự động <b>FIFO (Trái trước - Phải sau)</b> | Quản lý & phân quyền đa người dùng mượt mà, chính xác.
    </div>
</div>
""", unsafe_allow_html=True)

# KIỂM TRA LỖI NẠP FILE
if df_bom is None or df_stock is None:
    st.warning(f"⚠️ **Thông báo hệ thống**: {load_error or 'Chưa thể nạp dữ liệu mặc định.'}")
    st.stop()


# ==================== TÍNH TOÁN ĐỐI SOÁT FIFO HÀNG CHỜ ====================
calculation_results = []
actual_ok_orders = 0
actual_missing_orders = 0
queue_len = len(st.session_state.orders_queue)
if queue_len > 0:
    calculation_results = core_engine.calculate_inventory_allocation(
        st.session_state.orders_queue,
        df_bom,
        df_stock
    )
    actual_ok_orders = sum(1 for r in calculation_results if r.get('is_fully_ok', False))
    actual_missing_orders = len(calculation_results) - actual_ok_orders

# 5 THẺ KPI TO BẢN
col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
with col_m1:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-blue">
        <div class="kpi-label">📁 Tổng Đơn Trong BOM</div>
        <div class="kpi-value">{len(orders_info):,}<span class="kpi-unit">đơn</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    if queue_len > 0:
        st.markdown(f"""
        <div class="kpi-card kpi-card-border-emerald">
            <div class="kpi-label">🎯 Đơn Đủ Mã Đồng Bộ</div>
            <div class="kpi-value" style="color:#059669;">{actual_ok_orders}<span class="kpi-unit" style="color:#059669;">/ {queue_len}</span></div>
            <div style="font-size:11.5px; font-weight:700; color:#059669; margin-top:2px;">✨ {actual_ok_orders} đơn thực tế đối soát OK</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="kpi-card kpi-card-border-emerald">
            <div class="kpi-label">🎯 Đơn Đủ Mã Đồng Bộ</div>
            <div class="kpi-value" style="color:#059669;">{readiness_data['ready_count']}<span class="kpi-unit" style="color:#059669;">/ {readiness_data['total_orders']}</span></div>
            <div style="font-size:11.5px; font-weight:700; color:#059669; margin-top:2px;">✨ {readiness_data['ready_percent']}% đơn sẵn sàng chạy</div>
        </div>
        """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-green">
        <div class="kpi-label">🧩 Linh Kiện Trong Kho</div>
        <div class="kpi-value">{len(df_stock['Stock_Code'].unique()):,}<span class="kpi-unit">mã</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m4:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-purple">
        <div class="kpi-label">📦 Tổng Lượng Tồn Kho</div>
        <div class="kpi-value">{int(df_stock['Qty'].sum()):,}<span class="kpi-unit">cái</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m5:
    sub_kpi = f"""<div style="font-size:11.5px; font-weight:700; color:{'#059669' if actual_ok_orders > 0 else '#DC2626'}; margin-top:2px;">✨ {actual_ok_orders} / {queue_len} đơn đối soát OK</div>""" if queue_len > 0 else ""
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-amber">
        <div class="kpi-label">⏳ Đơn Trong Hàng Chờ</div>
        <div class="kpi-value">{queue_len}<span class="kpi-unit">đơn</span></div>
        {sub_kpi}
    </div>
    """, unsafe_allow_html=True)


# ==================== ĐIỀU HƯỚNG TABS CHÍNH ====================
tabs_list = ["📊 Đối Soát & Hàng Chờ FIFO", "🚀 Đơn Sẵn Sàng Chạy", "🔍 Kiểm Tra Tính Toàn Vẹn & Chính Xác"]
if user_permissions.get('can_view_audit', False):
    tabs_list.append("📜 Nhật Ký Hoạt Động (Audit Log)")
if user_permissions.get('can_manage_users', False):
    tabs_list.append("🛡️ Quản Lý Người Dùng & Phân Quyền")

tabs = st.tabs(tabs_list)

# TAB 1: ĐỐI SOÁT & HÀNG CHỜ FIFO
with tabs[0]:
    # 1. BẢNG NHẬP ĐƠN HÀNG
    st.markdown("""
    <div class="floating-box">
        <div class="box-header">
            <span>🎯 NHẬP ĐƠN HÀNG & SỐ LƯỢNG CẦN CHẠY</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    sorted_order_keys = sorted(orders_info.keys())
    ready_orders_dict = {ro['order_no']: ro for ro in readiness_data['ready_orders']}
    
    # Gợi ý chọn nhanh 1 chạm
    st.markdown("**⚡ Gợi ý chọn nhanh các đơn đủ mã đồng bộ (1 Chạm để điền mã & SL tối đa):**")
    quick_sample_orders = ['5151', '5171', '8333', '4151', '1009', '5804']
    chip_cols = safe_columns(len(quick_sample_orders))
    for c_i, ord_code in enumerate(quick_sample_orders):
        if ord_code in orders_info:
            ro_m = ready_orders_dict.get(ord_code)
            chip_label = f"🟢 Đơn {ord_code} ({ro_m['max_runnable_qty']:,} pcs)" if ro_m else f"⚠️ Đơn {ord_code}"
            if chip_cols[c_i].button(chip_label, key=f"quick_{ord_code}", use_container_width=True):
                st.session_state.quick_selected_order = ord_code
                if ro_m:
                    st.session_state.quick_selected_qty = str(int(ro_m['max_runnable_qty']))
                st.rerun()

    def format_order_item(k):
        ro = ready_orders_dict.get(k)
        if ro:
            return f"🟢 Đơn {k} - {orders_info[k]['description']} (✅ ĐỦ 100% MÃ - Tối đa: {ro['max_runnable_qty']:,} pcs)"
        else:
            return f"⚪ Đơn {k} - {orders_info[k]['description']} (⚠️ Thiếu linh kiện)"

    col_input1, col_input2, col_input3 = safe_columns([5, 3, 2], vertical_alignment="bottom")
    with col_input1:
        curr_target = st.session_state.get('quick_selected_order', '5151')
        default_idx = sorted_order_keys.index(curr_target) if curr_target in sorted_order_keys else 0
        selected_code = st.selectbox(
            "1. Chọn hoặc gõ tìm mã đơn hàng (Ví dụ: 5151, 8333...):",
            options=sorted_order_keys,
            format_func=format_order_item,
            index=default_idx if sorted_order_keys else None
        )
    with col_input2:
        qty_input_str = st.text_input(
            "2. Nhập số lượng cần chạy (bộ):",
            value=st.session_state.get('quick_selected_qty', "100"),
            placeholder="Nhập số lượng..."
        )
    with col_input3:
        can_queue = user_permissions.get('can_manage_queue', False)
        btn_add = st.button("➕ BẤM OK ĐỂ TÍNH", type="primary", use_container_width=True, disabled=not can_queue)

    if not can_queue:
        st.caption("🔒 *Tài khoản của bạn chỉ có quyền Xem (Viewer) hoặc Thủ kho, không được chỉnh sửa hàng chờ.*")

    if selected_code:
        curr_info = orders_info.get(selected_code, {})
        ready_match = ready_orders_dict.get(selected_code)
        if ready_match:
            ready_badge_str = f"<span class='badge-ok'>✅ ĐỦ 100% MÃ ĐỒNG BỘ (Tối đa chạy được {ready_match['max_runnable_qty']:,} pcs)</span>"
            bottleneck_str = f" | <span style='color:#D97706; font-weight:700;'>⚠️ Điểm nghẽn: <b>{ready_match['bottleneck_desc']}</b> (Còn: {int(ready_match['bottleneck_stock']):,} cái)</span>"
        else:
            ready_badge_str = "<span class='badge-missing'>⚠️ Đang thiếu linh kiện trong kho</span>"
            bottleneck_str = ""
        st.markdown(
            f"<div style='background:#F1F5F9; border-radius:10px; padding:10px 14px; font-size:13.5px; color:#334155; margin-top:8px; border-left:4px solid #2563EB;'>"
            f"ℹ️ <b>Đơn hàng {selected_code}</b>: {curr_info.get('description', '')} | <b>Nguyên đơn</b>: Trái + Phải | "
            f"Độ sẵn sàng: {ready_badge_str}{bottleneck_str}"
            f"</div>",
            unsafe_allow_html=True
        )

    # Thêm vào hàng chờ
    if btn_add and can_queue:
        try:
            val = float(str(qty_input_str).replace(',', '').strip())
            if val > 0 and selected_code:
                st.session_state.orders_queue.append({
                    'id': f"{selected_code}_{len(st.session_state.orders_queue)+1}_{uuid.uuid4().hex[:8]}",
                    'order_no': selected_code,
                    'description': orders_info[selected_code]['description'],
                    'qty': val,
                    'created_at': datetime.datetime.now().strftime("%H:%M:%S")
                })
                storage_manager.save_queue(st.session_state.orders_queue, username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "ADD_QUEUE", f"Thêm đơn {selected_code} ({val:,g} bộ) vào hàng chờ ({st.session_state.workspace_mode})")
                st.toast(f"✅ Đã thêm đơn {selected_code} ({val:,g} bộ) vào hàng chờ!", icon="✅")
                st.rerun()
            else:
                st.error("⚠️ Số lượng cần chạy phải lớn hơn 0.")
        except ValueError:
            st.error("⚠️ Vui lòng nhập số lượng hợp lệ.")

    # 2. KHUNG HÀNG CHỜ FIFO
    st.markdown("---")
    col_q_title, col_q_clear = safe_columns([8, 2], vertical_alignment="center")
    with col_q_title:
        st.markdown(f"### ⏳ Hàng Chờ Tính Toán FIFO ({len(st.session_state.orders_queue)} đơn) - *{workspace_title}*")
    with col_q_clear:
        if can_queue and len(st.session_state.orders_queue) > 0:
            if st.button("🗑️ Xóa hàng chờ", type="secondary", use_container_width=True):
                st.session_state.orders_queue = []
                storage_manager.save_queue([], username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "CLEAR_QUEUE", f"Xóa toàn bộ hàng chờ ({st.session_state.workspace_mode})")
                st.rerun()

    # Nạp từ file kế hoạch nếu có
    if df_plan is not None and len(df_plan) > 0 and can_queue:
        plan_orders_extracted = core_engine.extract_orders_from_plan(df_plan)
        col_p_btn, col_p_del, col_p_info = safe_columns([5, 2.5, 4.5], vertical_alignment="center")
        with col_p_btn:
            if st.button(f"🚀 NẠP TOÀN BỘ {len(plan_orders_extracted)} ĐƠN KẾ HOẠCH VÀO HÀNG CHỜ FIFO", type="primary", use_container_width=True):
                st.session_state.orders_queue = plan_orders_extracted
                storage_manager.save_queue(st.session_state.orders_queue, username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "LOAD_PLAN_QUEUE", f"Nạp {len(plan_orders_extracted)} đơn từ Kế hoạch vào hàng chờ ({st.session_state.workspace_mode})")
                st.toast(f"✅ Đã nạp thành công {len(plan_orders_extracted)} đơn hàng từ file Kế hoạch!", icon="✅")
                st.rerun()
        with col_p_del:
            if st.button("🗑️ Xóa dữ liệu Kế hoạch", key="btn_clear_plan_tab1", type="secondary", use_container_width=True, help="Xóa dữ liệu kế hoạch sản xuất để hệ thống tự động tính tồn kho đối soát các đơn đồng bộ"):
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = True
                if os.path.exists(PLAN_PERSIST_FILE):
                    try:
                        os.remove(PLAN_PERSIST_FILE)
                    except Exception:
                        pass
                save_meta({'plan_name': None, 'plan_cleared': True})
                st.session_state.orders_queue = [
                    item for item in st.session_state.orders_queue
                    if item.get('source') != 'Kế hoạch' and not str(item.get('id', '')).startswith('plan_')
                ]
                storage_manager.save_queue(st.session_state.orders_queue, username, st.session_state.workspace_mode)
                auth_manager.log_activity(username, "CLEAR_PLAN", f"Người dùng {user_fullname} xóa dữ liệu file Kế hoạch sản xuất")
                st.cache_data.clear()
                st.toast("✅ Đã xóa dữ liệu file Kế hoạch!", icon="🗑️")
                st.rerun()
        with col_p_info:
            st.caption(f"File Kế hoạch gồm {len(df_plan)} dòng Model, ghép thành {len(plan_orders_extracted)} đơn sản xuất.")
    elif st.session_state.get('plan_cleared', False) and can_queue:
        col_p_nt1, col_p_nt2 = safe_columns([8, 2], vertical_alignment="center")
        with col_p_nt1:
            st.caption("ℹ️ Dữ liệu File Kế hoạch đã được xóa (chưa có đơn hàng chỉ định chạy). Hệ thống tự động tính tồn kho để đối soát những đơn đồng bộ chạy được.")
        with col_p_nt2:
            if st.button("🔄 Nạp lại Kế hoạch", key="btn_restore_plan_tab1", use_container_width=True):
                st.session_state.plan_cleared = False
                save_meta({'plan_cleared': False})
                auth_manager.log_activity(username, "RESTORE_PLAN", f"Người dùng {user_fullname} khôi phục nạp file Kế hoạch mặc định")
                st.cache_data.clear()
                st.toast("✅ Đã nạp lại file Kế hoạch mặc định!", icon="✅")
                st.rerun()

    if len(st.session_state.orders_queue) == 0:
        st.info("💡 Hiện chưa có đơn hàng nào trong hàng chờ. Vui lòng chọn mã đơn ở trên rồi bấm **BẤM OK ĐỂ TÍNH**.")
    else:
        # Bảng hiển thị hàng chờ
        cols_header = safe_columns([1, 2, 3, 2, 3, 1, 1])
        cols_header[0].markdown("**Ưu tiên**")
        cols_header[1].markdown("**Mã đơn**")
        cols_header[2].markdown("**Tên sản phẩm**")
        cols_header[3].markdown("**Số lượng đặt**")
        cols_header[4].markdown("**Đồng bộ tồn kho FIFO**")
        cols_header[5].markdown("**Thời gian**")
        cols_header[6].markdown("**Xóa**")

        to_delete_idx = None
        for idx, item in enumerate(st.session_state.orders_queue):
            c = safe_columns([1, 2, 3, 2, 3, 1, 1], vertical_alignment="center")
            c[0].markdown(f"**#{idx + 1}**")
            c[1].markdown(f"<span class='badge-order'>{item['order_no']}</span>", unsafe_allow_html=True)
            c[2].write(item.get('description') or orders_info.get(item['order_no'], {}).get('description', ''))
            
            raw_qty = item.get('qty', 0)
            try: qty_num = float(raw_qty)
            except Exception: qty_num = 0.0
            qty_l = float(item.get('qty_left', qty_num))
            qty_r = float(item.get('qty_right', qty_num))
            if qty_l == qty_r:
                c[3].markdown(f"**{qty_num:,g}** bộ")
            else:
                c[3].markdown(f"T: **{qty_l:,g}** | P: **{qty_r:,g}**")

            if idx < len(calculation_results):
                res_item = calculation_results[idx]
                if res_item.get('is_fully_ok', False):
                    c[4].markdown("<span class='badge-ok'>✅ ĐỦ 100% VẬT TƯ</span>", unsafe_allow_html=True)
                else:
                    m_count = res_item.get('missing_count', 0)
                    c[4].markdown(f"<span class='badge-missing'>⚠️ THIẾU {m_count} MÃ</span>", unsafe_allow_html=True)
            else:
                c[4].markdown("<span>⏳ Đang tính...</span>", unsafe_allow_html=True)

            c[5].write(item.get('due_date') or item.get('created_at', ''))
            if can_queue:
                if c[6].button("❌", key=f"del_{item['id']}"):
                    to_delete_idx = idx
            else:
                c[6].write("-")

        if to_delete_idx is not None:
            del_item = st.session_state.orders_queue.pop(to_delete_idx)
            storage_manager.save_queue(st.session_state.orders_queue, username, st.session_state.workspace_mode)
            auth_manager.log_activity(username, "DELETE_QUEUE", f"Xóa đơn {del_item.get('order_no')} khỏi hàng chờ ({st.session_state.workspace_mode})")
            st.rerun()

    # 3. KẾT QUẢ ĐỐI SOÁT & XUẤT EXCEL
    if len(st.session_state.orders_queue) > 0 and calculation_results:
        st.markdown("---")
        st.markdown(f"### 📊 Báo Cáo Chi Tiết Đối Soát Từng Đơn ({len(calculation_results)} đơn)")
        
        # Nút xuất toàn bộ Excel
        queue_cache_key = tuple(
            (item.get('id', ''), item['order_no'], item['order_qty'], item.get('ok_count', 0), item.get('missing_count', 0))
            for item in calculation_results
        )
        all_excel_bytes = get_cached_all_excel(queue_cache_key, calculation_results)
        st.download_button(
            label=f"📥 Tải Trọn Bộ Excel Đối Soát ({len(calculation_results)} đơn hàng - mỗi đơn 1 sheet)",
            data=all_excel_bytes,
            file_name=f"Doi_Soat_Tong_Hop_{len(calculation_results)}_Don_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        st.markdown("<br>", unsafe_allow_html=True)
        for r in calculation_results:
            order_no = r['order_no']
            order_qty = r['order_qty']
            is_fully_ok = r.get('is_fully_ok', False)
            stt = r['stt']
            status_badge = "<span class='badge-ok'>✅ ĐỦ VẬT TƯ (100% OK)</span>" if is_fully_ok else f"<span class='badge-missing'>⚠️ THIẾU {r['missing_count']} ITEM CODE</span>"
            card_class = "order-card-ok" if is_fully_ok else "order-card-missing"
            
            qty_l = float(r.get('qty_left', order_qty))
            qty_r = float(r.get('qty_right', order_qty))
            qty_str = f"<b>{order_qty:,g} bộ</b> (Trái + Phải)" if qty_l == qty_r else f"Trái: <b>{qty_l:,g}</b> | Phải: <b>{qty_r:,g}</b>"

            st.markdown(f"""
            <div class="order-card {card_class}">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                    <div>
                        <span style="font-size: 17px; font-weight: 850; color: #1E3A8A;">Ưu tiên #{stt} - Đơn hàng {order_no}</span>
                        <span style="font-size: 14.5px; color: #475569; margin-left: 6px;">({r.get('product_title', '')})</span>
                        <div style="margin-top: 3px; font-size: 13.5px; color: #64748B;">
                            Số lượng: {qty_str} | Trạng thái: {status_badge}
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_ck1, col_ck2 = safe_columns([6, 4], vertical_alignment="center")
            with col_ck1:
                k1, k2, k3 = safe_columns(3)
                k1.metric("Tổng mã LK", f"{r['total_items']} mã")
                k2.metric("Mã ĐỦ", f"{r['ok_count']} mã")
                k3.metric("Mã THIẾU", f"{r['missing_count']} mã", delta=f"-{r['missing_count']}" if r['missing_count'] > 0 else "0", delta_color="inverse")
            with col_ck2:
                single_cache_key = f"{r.get('id', '')}_{order_no}_{order_qty}_{r.get('ok_count', 0)}_{r.get('missing_count', 0)}"
                single_excel = get_cached_single_order_excel(single_cache_key, r)
                st.download_button(
                    label=f"📥 Tải Excel Đơn {order_no}",
                    data=single_excel,
                    file_name=f"Doi_Soat_Don_{order_no}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_{r['id']}",
                    use_container_width=True
                )
            st.markdown("<hr style='margin: 12px 0; border: none; border-top: 1px dashed #CBD5E1;'>", unsafe_allow_html=True)


# TAB 2: ĐƠN SẴN SÀNG CHẠY
with tabs[1]:
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #064E3B 0%, #059669 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
        <div style="font-size: 20px; font-weight: 850;">
            🚀 TỔNG HỢP {readiness_data['ready_count']} ĐƠN HÀNG ĐỦ 100% MÃ ĐỒNG BỘ TRONG KHO
        </div>
        <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
            Hệ thống tự động tính toán số lượng tối đa có thể sản xuất ngay và chỉ rõ linh kiện nghẽn số lượng ít nhất.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Nút tải Excel báo cáo
    col_dl_r1, col_dl_r2 = safe_columns([5, 5], vertical_alignment="center")
    with col_dl_r1:
        readiness_key = (readiness_data['ready_count'], readiness_data['not_ready_count'], readiness_data['total_orders'])
        readiness_excel = get_cached_readiness_excel(readiness_key, readiness_data)
        st.download_button(
            label=f"📥 Tải Báo Cáo Excel {readiness_data['ready_count']} Đơn Đủ Mã (Kèm Đơn Thiếu)",
            data=readiness_excel,
            file_name=f"Bao_Cao_Don_Du_Ma_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    with col_dl_r2:
        st.caption("File Excel gồm 2 sheet: Đơn đủ mã sẵn sàng chạy và Đơn chưa đủ mã kèm danh mục linh kiện thiếu.")

    # Tìm kiếm & phân trang bảng đơn đủ mã
    col_s1, col_s2, col_s3 = safe_columns([5, 3, 2], vertical_alignment="bottom")
    with col_s1:
        search_ready = st.text_input("🔍 Tìm kiếm đơn hàng:", placeholder="Gõ mã đơn hoặc tên sản phẩm...")
    with col_s2:
        sort_mode = st.selectbox("Sắp xếp:", ["🔥 SL chạy tối đa (Giảm dần)", "🔢 SL chạy tối đa (Tăng dần)", "🏷️ Mã đơn (A-Z)"])
    with col_s3:
        page_size = st.selectbox("Số dòng / trang:", [15, 25, 50, 100], index=0)

    # Lọc danh sách
    filtered_ready = list(readiness_data['ready_orders'])
    if search_ready.strip():
        kw = search_ready.strip().lower()
        filtered_ready = [ro for ro in filtered_ready if kw in ro['order_no'].lower() or kw in ro['description'].lower()]

    if "Giảm dần" in sort_mode:
        filtered_ready.sort(key=lambda x: x['max_runnable_qty'], reverse=True)
    elif "Tăng dần" in sort_mode:
        filtered_ready.sort(key=lambda x: x['max_runnable_qty'])
    else:
        filtered_ready.sort(key=lambda x: x['order_no'])

    total_pages = max(1, (len(filtered_ready) + page_size - 1) // page_size)
    if 'ready_page_idx' not in st.session_state:
        st.session_state.ready_page_idx = 1
    if st.session_state.ready_page_idx > total_pages:
        st.session_state.ready_page_idx = 1

    start_idx = (st.session_state.ready_page_idx - 1) * page_size
    page_items = filtered_ready[start_idx : start_idx + page_size]

    # Bảng HTML hiển thị mượt mà
    ready_html = """
    <div style="overflow-x: auto;">
    <table class="styled-table">
        <thead>
            <tr>
                <th style="width: 50px; text-align:center;">STT</th>
                <th style="width: 120px; text-align:center;">Mã Đơn</th>
                <th>Tên Sản Phẩm</th>
                <th style="text-align:center; width: 200px; background:#059669; color:#FFFFFF;">SL TỐI ĐA CHẠY ĐƯỢC</th>
                <th style="text-align:center; width: 140px;">Trạng Thái</th>
            </tr>
        </thead>
        <tbody>
    """
    for idx_row, r_item in enumerate(page_items, start_idx + 1):
        hover_info = f"Điểm nghẽn: {r_item['bottleneck_desc']} (Kho còn: {int(r_item['bottleneck_stock']):,} cái)"
        ready_html += f"""
            <tr class="row-ok">
                <td style="text-align:center;"><b>#{idx_row}</b></td>
                <td style="text-align:center;"><span class="badge-order">{r_item['order_no']}</span></td>
                <td><b>{r_item['description']}</b></td>
                <td style="text-align:center; font-size:16px; font-weight:900; color:#059669; background:#ECFDF5;" title="{hover_info}">
                    {r_item['max_runnable_qty']:,} pcs
                </td>
                <td style="text-align:center;"><span class="badge-ok">✅ ĐỦ 100% MÃ</span></td>
            </tr>
        """
    ready_html += "</tbody></table></div>"
    st.markdown(ready_html, unsafe_allow_html=True)

    # Điều khiển phân trang
    col_p1, col_p2, col_p3 = safe_columns([3, 4, 3], vertical_alignment="center")
    with col_p1:
        st.caption(f"Hiển thị {len(page_items)} / {len(filtered_ready)} đơn (Trang {st.session_state.ready_page_idx} / {total_pages})")
    with col_p2:
        c_prev, c_txt, c_next = safe_columns([1, 2, 1], vertical_alignment="center")
        if c_prev.button("◀ Trang trước", disabled=st.session_state.ready_page_idx <= 1, use_container_width=True):
            st.session_state.ready_page_idx -= 1
            st.rerun()
        c_txt.markdown(f"<div style='text-align:center; font-weight:700;'>Trang {st.session_state.ready_page_idx}</div>", unsafe_allow_html=True)
        if c_next.button("Trang sau ▶", disabled=st.session_state.ready_page_idx >= total_pages, use_container_width=True):
            st.session_state.ready_page_idx += 1
            st.rerun()


# TAB 3: KIỂM TRA TÍNH TOÀN VẸN & CHÍNH XÁC
with tabs[2]:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
        <div style="font-size: 20px; font-weight: 850;">
            🔍 KIỂM SOÁT TÍNH CHÍNH XÁC & ĐỐI SOÁT TOÀN VẸN DỮ LIỆU
        </div>
        <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
            Tự động rà soát sai lệch giữa Định mức BOM, File Tồn kho và File Kế hoạch để đảm bảo tính toán luôn chuẩn xác 100%.
        </div>
    </div>
    """, unsafe_allow_html=True)

    integrity_report = core_engine.check_data_integrity(df_bom, df_stock, df_plan)
    if integrity_report:
        c_h1, c_h2, c_h3, c_h4 = safe_columns(4)
        c_h1.metric("Điểm Sức Khỏe Dữ Liệu", f"{integrity_report['overall_health']}%")
        c_h2.metric("Mã LK Khớp Trong Kho", f"{integrity_report['matched_comps']} / {integrity_report['total_bom_comps']} mã")
        c_h3.metric("Mã LK Chưa Có Trong Kho", f"{integrity_report['unmatched_comps_count']} mã", delta=f"-{integrity_report['unmatched_comps_count']}", delta_color="inverse")
        c_h4.metric("Sai Lệch Kế Hoạch Sản Xuất", f"{integrity_report['plan_issues_count']} đơn")

        st.markdown("---")
        st.markdown(f"#### ⚠️ 1. Danh sách {integrity_report['unmatched_comps_count']} mã linh kiện trong BOM chưa có trong kho:")
        st.caption("Các mã linh kiện này xuất hiện trong cấu trúc BOM nhưng hoàn toàn chưa có dòng nào trong file Tồn kho (Stock 61):")
        if len(integrity_report['df_unmatched_comps']) > 0:
            st.dataframe(integrity_report['df_unmatched_comps'], use_container_width=True, height=280)
        else:
            st.success("✅ Tuyệt vời: 100% linh kiện trong BOM đều tồn tại trong danh mục tồn kho.")

        st.markdown("#### 📅 2. Kiểm tra tính đồng bộ của File Kế hoạch với BOM:")
        if df_plan is not None:
            if integrity_report['plan_issues_count'] > 0:
                st.dataframe(integrity_report['df_plan_issues'], use_container_width=True)
            else:
                st.success("✅ 100% các đơn hàng trong File Kế hoạch đều có cấu trúc định mức BOM đầy đủ, sẵn sàng chạy đối soát.")
        else:
            st.info("ℹ️ Chưa nạp File Kế hoạch (hoặc dữ liệu Kế hoạch đã được xóa). Hệ thống tính tồn kho đối soát theo toàn bộ đơn hàng trong BOM.")

        st.markdown("#### 📦 3. Kiểm tra các dòng tồn kho âm hoặc bất thường:")
        if integrity_report['zero_stock_count'] > 0:
            st.dataframe(integrity_report['df_zero_stock'], use_container_width=True)
        else:
            st.success("✅ Dữ liệu số lượng tồn kho hợp lệ, không có mã nào bị âm.")


# TAB 4: NHẬT KÝ HOẠT ĐỘNG (AUDIT LOG - ADMIN & PLANNER)
if user_permissions.get('can_view_audit', False):
    tab_audit_idx = 3
    with tabs[tab_audit_idx]:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #0F172A 0%, #334155 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
            <div style="font-size: 20px; font-weight: 850;">
                📜 SỔ NHẬT KÝ HOẠT ĐỘNG HỆ THỐNG (AUDIT LOG)
            </div>
            <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
                Theo dõi minh bạch mọi thao tác khi chia sẻ cho nhiều người: Ai đăng nhập, ai nạp file, ai thêm/xóa đơn trong hàng chờ.
            </div>
        </div>
        """, unsafe_allow_html=True)

        logs = auth_manager.get_audit_logs(limit=200)
        col_l_filter, col_l_clear = safe_columns([8, 2], vertical_alignment="center")
        with col_l_filter:
            st.caption(f"Tổng số bản ghi nhật ký: **{len(logs)}** thao tác gần nhất")
        with col_l_clear:
            if user_permissions.get('can_manage_users', False):
                if st.button("🗑️ Dọn sạch nhật ký", type="secondary", use_container_width=True):
                    auth_manager.clear_audit_logs(username)
                    st.rerun()

        if logs:
            df_logs = pd.DataFrame(logs).rename(columns={
                'timestamp': 'Thời gian',
                'username': 'Tài khoản',
                'action': 'Hành động',
                'details': 'Chi tiết thao tác'
            })
            st.dataframe(df_logs, use_container_width=True, height=400)
        else:
            st.info("Chưa có nhật ký hoạt động nào được ghi nhận.")


# TAB 5: QUẢN LÝ NGƯỜI DÙNG & PHÂN QUYỀN (CHỈ ADMIN)
if user_permissions.get('can_manage_users', False):
    tab_user_idx = len(tabs_list) - 1
    with tabs[tab_user_idx]:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #7F1D1D 0%, #DC2626 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
            <div style="font-size: 20px; font-weight: 850;">
                🛡️ QUẢN TRỊ TÀI KHOẢN & PHÂN QUYỀN TRUY CẬP (RBAC)
            </div>
            <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
                Quản lý danh sách người dùng, cấp quyền truy cập, đặt lại mật khẩu và khóa tài khoản khi chia sẻ app cho nhà máy.
            </div>
        </div>
        """, unsafe_allow_html=True)

        all_users = auth_manager.get_all_users()
        
        # Danh sách người dùng hiện có
        st.markdown("#### 👥 Danh sách tài khoản hệ thống:")
        user_rows = []
        for uname_k, uval in all_users.items():
            role_info = auth_manager.ROLES_METADATA.get(uval['role'], {})
            user_rows.append({
                'Tên đăng nhập': uname_k,
                'Họ và tên': uval.get('full_name', ''),
                'Vai trò': role_info.get('name', uval['role']),
                'Bộ phận': uval.get('department', ''),
                'Trạng thái': '🟢 Đang hoạt động' if uval.get('is_active', True) else '🔴 Đã khóa',
                'Đăng nhập cuối': uval.get('last_login') or 'Chưa đăng nhập'
            })
        st.dataframe(pd.DataFrame(user_rows), use_container_width=True)

        st.markdown("---")
        col_u1, col_u2 = safe_columns(2)
        
        # Form thêm tài khoản mới
        with col_u1:
            st.markdown("##### ➕ Thêm tài khoản người dùng mới:")
            with st.form("form_create_user"):
                new_uname = st.text_input("Tên đăng nhập (viết liền không dấu):", placeholder="ví dụ: tranvanb")
                new_pw = st.text_input("Mật khẩu ban đầu:", type="password", placeholder="Tối thiểu 4 ký tự")
                new_fname = st.text_input("Họ và tên:", placeholder="Ví dụ: Trần Văn B")
                new_dept = st.text_input("Bộ phận / Phòng ban:", placeholder="Ví dụ: Xưởng Lắp Ráp 1")
                new_role = st.selectbox(
                    "Vai trò & Quyền hạn:",
                    options=list(auth_manager.ROLES_METADATA.keys()),
                    format_func=lambda r: auth_manager.ROLES_METADATA[r]['name']
                )
                btn_create = st.form_submit_button("Tạo Tài Khoản", type="primary", use_container_width=True)
                if btn_create:
                    c_ok, c_msg = auth_manager.create_user(username, new_uname, new_pw, new_fname, new_role, new_dept)
                    if c_ok:
                        st.success(c_msg)
                        st.rerun()
                    else:
                        st.error(c_msg)

        # Form quản lý tài khoản hiện có
        with col_u2:
            st.markdown("##### ⚙️ Điều chỉnh tài khoản & Đặt lại mật khẩu:")
            target_user_opt = [u for u in all_users.keys() if u != 'admin']
            if target_user_opt:
                selected_target = st.selectbox("Chọn tài khoản cần thao tác:", options=target_user_opt)
                target_user_data = all_users[selected_target]
                
                # Đổi vai trò
                col_r_sel, col_r_btn = safe_columns([7, 3], vertical_alignment="bottom")
                with col_r_sel:
                    current_r_idx = list(auth_manager.ROLES_METADATA.keys()).index(target_user_data['role'])
                    new_assigned_role = st.selectbox(
                        "Đổi vai trò:",
                        options=list(auth_manager.ROLES_METADATA.keys()),
                        format_func=lambda r: auth_manager.ROLES_METADATA[r]['name'],
                        index=current_r_idx,
                        key="sel_new_assigned_role"
                    )
                with col_r_btn:
                    if st.button("Cập nhật quyền", key="btn_update_r", use_container_width=True):
                        u_ok, u_msg = auth_manager.update_user_role(username, selected_target, new_assigned_role)
                        if u_ok:
                            st.success(u_msg)
                            st.rerun()
                        else:
                            st.error(u_msg)

                st.markdown("<br>", unsafe_allow_html=True)
                # Đặt lại mật khẩu
                col_pw_in, col_pw_btn = safe_columns([7, 3], vertical_alignment="bottom")
                with col_pw_in:
                    reset_pw_val = st.text_input(f"Đặt lại mật khẩu cho {selected_target}:", type="password", key="inp_reset_pw")
                with col_pw_btn:
                    if st.button("Đặt lại MK", key="btn_reset_pw", use_container_width=True):
                        p_ok, p_msg = auth_manager.reset_user_password(username, selected_target, reset_pw_val)
                        if p_ok:
                            st.success(p_msg)
                        else:
                            st.error(p_msg)

                # Khóa / Mở khóa
                is_target_active = target_user_data.get('is_active', True)
                toggle_btn_label = "🔒 Khóa tài khoản này" if is_target_active else "🔓 Mở khóa tài khoản này"
                if st.button(toggle_btn_label, key="btn_toggle_active", use_container_width=True):
                    t_ok, t_msg = auth_manager.toggle_user_active(username, selected_target)
                    if t_ok:
                        st.success(t_msg)
                        st.rerun()
                    else:
                        st.error(t_msg)
            else:
                st.info("Chưa có tài khoản phụ nào ngoài Admin.")


if __name__ == "__main__":
    from streamlit.runtime import exists
    if not exists():
        import sys
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", __file__]
        sys.exit(stcli.main())
