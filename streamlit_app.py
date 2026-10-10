import datetime
import io
import os
import uuid
import streamlit as st
import pandas as pd
import core_engine


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
        padding-top: 1.2rem !important;
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
        padding: 26px 34px;
        color: #FFFFFF;
        box-shadow: 0 12px 30px -5px rgba(30, 58, 138, 0.3), 0 8px 12px -6px rgba(30, 58, 138, 0.15);
        margin-bottom: 22px;
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
        font-size: 28px;
        font-weight: 850;
        letter-spacing: -0.5px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-desc {
        font-size: 15px;
        color: #E2E8F0;
        font-weight: 450;
        line-height: 1.6;
    }
    .hero-tag {
        display: inline-block;
        background: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        padding: 5px 14px;
        border-radius: 9999px;
        font-size: 12.5px;
        font-weight: 700;
        margin-top: 12px;
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
        padding: 20px 22px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        position: relative;
        overflow: hidden;
        min-height: 125px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
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
        font-size: 12.5px;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .kpi-value {
        font-size: 32px;
        font-weight: 900;
        color: #0F172A;
        line-height: 1.1;
        margin-top: 8px;
    }
    .kpi-unit {
        font-size: 15px;
        font-weight: 600;
        color: #64748B;
        margin-left: 4px;
    }

    /* Hộp nhập liệu TO & NỔI BẬT */
    .floating-box {
        background: #FFFFFF;
        border: 2.5px solid #2563EB;
        border-radius: 18px;
        padding: 26px 30px;
        box-shadow: 0 14px 35px -5px rgba(37, 99, 235, 0.16), 0 10px 15px -6px rgba(37, 99, 235, 0.08);
        margin-bottom: 24px;
    }
    .box-header {
        font-size: 20px;
        font-weight: 800;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 16px;
    }

    /* Nút chính to bản, rực rỡ */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: white !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.5) !important;
    }

    /* Thẻ tóm tắt đơn hàng kết quả (Order Result Card) */
    .order-card {
        background: #FFFFFF;
        border: 1.5px solid #E2E8F0;
        border-radius: 14px;
        padding: 20px 24px;
        margin-bottom: 16px;
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
        font-size: 13.5px;
        border-radius: 10px;
        overflow: hidden;
        margin-top: 12px;
        margin-bottom: 16px;
        border: 1px solid #CBD5E1;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
    .styled-table thead tr {
        background: linear-gradient(135deg, #1E293B 0%, #1E3A8A 100%);
        color: #FFFFFF;
        text-align: left;
        font-weight: 700;
        font-size: 13px;
    }
    .styled-table th {
        padding: 11px 12px;
        border-right: 1px solid rgba(255, 255, 255, 0.1);
        vertical-align: middle;
    }
    .styled-table td {
        padding: 9px 12px;
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
        padding: 5px 12px;
        border-radius: 9999px;
        font-weight: 800;
        font-size: 12px;
        display: inline-block;
        letter-spacing: 0.3px;
    }
    .badge-missing {
        background-color: #FEE2E2;
        color: #B91C1C;
        border: 1px solid #FCA5A5;
        padding: 5px 12px;
        border-radius: 9999px;
        font-weight: 800;
        font-size: 12px;
        display: inline-block;
        letter-spacing: 0.3px;
    }
    .badge-order {
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        padding: 5px 12px;
        border-radius: 8px;
        font-weight: 800;
        font-size: 14px;
        display: inline-block;
    }

    /* Thẻ Spotlight Báo Cáo Nhanh Đơn Sẵn Sàng Chạy */
    .spotlight-card {
        background: #FFFFFF;
        border: 1.5px solid #CBD5E1;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
        border-top: 5px solid #059669;
        margin-bottom: 8px;
        min-height: 215px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .spotlight-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(5, 150, 105, 0.15);
        border-color: #10B981;
    }
    .spotlight-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
    }
    .spotlight-order {
        font-size: 19px;
        font-weight: 850;
        color: #1E3A8A;
    }
    .spotlight-desc {
        font-size: 13px;
        color: #475569;
        font-weight: 600;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        margin-bottom: 8px;
    }
    .spotlight-qty-box {
        background: linear-gradient(180deg, #F0FDF4 0%, #DCFCE7 100%);
        border: 1px solid #86EFAC;
        border-radius: 10px;
        padding: 8px 12px;
        text-align: center;
        margin-bottom: 8px;
    }
    .spotlight-qty-label {
        font-size: 11px;
        font-weight: 700;
        color: #166534;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .spotlight-qty-val {
        font-size: 26px;
        font-weight: 900;
        color: #059669;
        line-height: 1.1;
    }
    .spotlight-qty-unit {
        font-size: 13px;
        font-weight: 700;
        color: #15803D;
        margin-left: 3px;
    }
    .spotlight-bottleneck {
        font-size: 11.5px;
        color: #64748B;
        background: #F8FAFC;
        padding: 5px 8px;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        margin-bottom: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
</style>
""", unsafe_allow_html=True)

# ==================== SESSION STATE ====================
if 'orders_queue' not in st.session_state:
    st.session_state.orders_queue = []
if 'custom_bom_bytes' not in st.session_state:
    st.session_state.custom_bom_bytes = None
if 'custom_bom_name' not in st.session_state:
    st.session_state.custom_bom_name = None
if 'custom_stock_bytes' not in st.session_state:
    st.session_state.custom_stock_bytes = None
if 'custom_stock_name' not in st.session_state:
    st.session_state.custom_stock_name = None
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
    # Xác định file BOM
    if bom_bytes:
        bom_source = io.BytesIO(bom_bytes)
    else:
        bom_source = find_data_file(['production', 'bom']) or find_data_file(['bom'])

    # Xác định file Tồn kho (Stock Balance)
    if stock_bytes:
        stock_source = io.BytesIO(stock_bytes)
    else:
        stock_source = find_data_file(['stock', 'balance']) or find_data_file(['stock'])

    empty_readiness = {'ready_orders': [], 'not_ready_orders': [], 'total_orders': 0, 'ready_count': 0, 'not_ready_count': 0, 'ready_percent': 0.0}

    if bom_source is None or stock_source is None:
        missing = []
        if bom_source is None:
            missing.append("Production BOM (.xlsx)")
        if stock_source is None:
            missing.append("Stock Balance (.xlsx)")
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
def get_cached_readiness_excel(_readiness_data):
    return core_engine.export_readiness_report_to_excel(_readiness_data)

@st.cache_data
def get_cached_single_order_excel(_order_result):
    return core_engine.export_order_to_excel(_order_result)

@st.cache_data
def get_cached_all_excel(_calculation_results):
    return core_engine.export_all_orders_to_excel(_calculation_results)

# Tải dữ liệu
df_bom, df_stock, orders_info, readiness_data, load_error = get_loaded_data(
    st.session_state.custom_bom_bytes,
    st.session_state.custom_bom_name,
    st.session_state.custom_stock_bytes,
    st.session_state.custom_stock_name
)

# ==================== SIDEBAR QUẢN LÝ DỮ LIỆU (LUÔN HIỂN THỊ) ====================
with st.sidebar:
    st.markdown("### ⚙️ Dữ Liệu Nguồn")
    st.caption("Quản lý file Tồn kho và Định mức BOM sản xuất.")
    
    # Nút upload Stock Balance
    st.markdown("#### 1. File Tồn kho (Stock Balance)")
    uploaded_stock = st.file_uploader(
        "Tải lên file Stock Balance mới",
        type=["xlsx", "xls"],
        key="uploader_stock",
        help="Cột quan trọng: Stock Code, Qty"
    )
    if uploaded_stock is not None:
        if st.session_state.custom_stock_name != uploaded_stock.name:
            st.session_state.custom_stock_bytes = uploaded_stock.getvalue()
            st.session_state.custom_stock_name = uploaded_stock.name
            st.cache_data.clear()
            st.success(f"Đã nạp file tồn kho: {uploaded_stock.name}")
            st.rerun()
            
    current_stock_name = st.session_state.custom_stock_name or "Stock Balance With Batch (3).xlsx (Mặc định)"
    if df_stock is not None:
        st.caption(f"📁 **{current_stock_name}** ({len(df_stock):,} dòng, {len(df_stock['Stock_Code'].unique()):,} mã)")
    else:
        st.caption(f"📁 **{current_stock_name}** (Chưa nạp được dữ liệu)")
        
    st.divider()
    
    # Nút upload Production BOM
    st.markdown("#### 2. File Định mức (Production BOM)")
    uploaded_bom = st.file_uploader(
        "Tải lên file Production BOM mới",
        type=["xlsx", "xls"],
        key="uploader_bom",
        help="Cột quan trọng: Product Code, Component Code, Qty"
    )
    if uploaded_bom is not None:
        if st.session_state.custom_bom_name != uploaded_bom.name:
            st.session_state.custom_bom_bytes = uploaded_bom.getvalue()
            st.session_state.custom_bom_name = uploaded_bom.name
            st.cache_data.clear()
            st.success(f"Đã nạp file BOM: {uploaded_bom.name}")
            st.rerun()
            
    current_bom_name = st.session_state.custom_bom_name or "Production BOM List 10.2.xlsx (Mặc định)"
    if df_bom is not None:
        st.caption(f"📁 **{current_bom_name}** ({len(df_bom):,} dòng, {len(orders_info):,} đơn hàng)")
    else:
        st.caption(f"📁 **{current_bom_name}** (Chưa nạp được dữ liệu)")
        
    st.divider()
    
    # Nút reset về file gốc
    if st.session_state.custom_bom_bytes or st.session_state.custom_stock_bytes:
        if st.button("🔄 Khôi phục 2 file mẫu mặc định", use_container_width=True):
            st.session_state.custom_bom_bytes = None
            st.session_state.custom_bom_name = None
            st.session_state.custom_stock_bytes = None
            st.session_state.custom_stock_name = None
            st.cache_data.clear()
            st.success("Đã khôi phục dữ liệu mẫu!")
            st.rerun()
            
    st.markdown("---")
    st.markdown("#### 💡 Cơ chế tính toán:")
    st.markdown("""
    - **1 Đơn hàng = 1 Bộ nguyên đơn**: Tự động lấy đủ cả 2 vế **Trái và Phải**.
    - **Định mức nhân số lượng**: Hệ thống nhân riêng `ĐM Trái × SL` và `ĐM Phải × SL` rồi cộng lại thành tổng BOM cần chạy.
    - **Phân bổ FIFO**: Đơn nhập trước được ưu tiên trừ tồn kho trước.
    """)

# ==================== MAIN BANNER (LUÔN HIỂN THỊ) ====================
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">📦 HỆ THỐNG ĐỐI SOÁT BOM SẢN XUẤT & TỒN KHO</div>
    <div class="hero-desc">
        Tự động tính toán định mức nguyên đơn <b>(Trái + Phải)</b> theo số lượng cần chạy | 
        Phân bổ trừ tồn kho tự động theo thứ tự ưu tiên <b>FIFO</b> | 
        Báo cáo Excel chuẩn hóa tiêu đề: <b>ITEM CODE</b> & <b>DESCRIPTION (tên mô tả)</b>.
    </div>
    <div class="hero-tag">✨ Sẵn sàng lên kế hoạch sản xuất</div>
</div>
""", unsafe_allow_html=True)

# KIỂM TRA TRẠNG THÁI DỮ LIỆU
if df_bom is None or df_stock is None:
    st.warning(f"⚠️ **Thông báo hệ thống**: {load_error or 'Chưa thể nạp dữ liệu mặc định.'}")
    st.markdown("""
    <div style="background: #F8FAFC; border: 2px dashed #94A3B8; border-radius: 14px; padding: 24px; text-align: center; margin-top: 16px;">
        <h4 style="color: #1E293B; margin-bottom: 8px;">📁 Vui lòng tải lên file dữ liệu ở thanh bên trái (Sidebar)</h4>
        <p style="color: #64748B; font-size: 14px; max-width: 600px; margin: 0 auto;">
            Hệ thống cần 2 file: <b>Stock Balance</b> (tồn kho hiện có) và <b>Production BOM</b> (định mức sản xuất).
            Vui lòng mở menu thanh bên trái và tải file Excel lên để hệ thống tự động đối soát ngay lập tức!
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ==================== CÁC THẺ KPI TO BẢN & SỐNG ĐỘNG (5 THẺ) ====================
col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
with col_m1:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-blue">
        <div class="kpi-label">📁 Tổng Đơn Trong BOM</div>
        <div class="kpi-value">{len(orders_info):,}<span class="kpi-unit">đơn</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-emerald">
        <div class="kpi-label">🎯 Đơn Đủ Mã Đồng Bộ</div>
        <div class="kpi-value" style="color:#059669;">{readiness_data['ready_count']}<span class="kpi-unit" style="color:#059669;">/ {readiness_data['total_orders']}</span></div>
        <div style="font-size:12px; font-weight:700; color:#059669; margin-top:2px;">✨ {readiness_data['ready_percent']}% đơn sẵn sàng chạy</div>
    </div>
    """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-green">
        <div class="kpi-label">🧩 Mã Linh Kiện Trong Kho</div>
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
    queue_len = len(st.session_state.orders_queue)
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-amber">
        <div class="kpi-label">⏳ Đơn Trong Hàng Chờ</div>
        <div class="kpi-value">{queue_len}<span class="kpi-unit">đơn</span></div>
    </div>
    """, unsafe_allow_html=True)



# ==================== BẢNG NỔI NHẬP ĐƠN HÀNG TO & NỔI BẬT ====================
st.markdown("""
<div class="floating-box">
    <div class="box-header">
        <span>🎯 NHẬP ĐƠN HÀNG & SỐ LƯỢNG CẦN CHẠY</span>
    </div>
</div>
""", unsafe_allow_html=True)

with st.container():
    sorted_order_keys = sorted(orders_info.keys())
    
    # Gợi ý chọn nhanh các đơn đủ mã đồng bộ sẵn sàng chạy
    st.markdown("**⚡ Gợi ý chọn nhanh các đơn đủ mã đồng bộ (1 Chạm để điền mã & SL tối đa):**")
    quick_sample_orders = ['5151', '5171', '8333', '4151', '1009', '5804']
    chip_cols = safe_columns(len(quick_sample_orders))
    for c_i, ord_code in enumerate(quick_sample_orders):
        if ord_code in orders_info:
            ro_m = next((ro for ro in readiness_data['ready_orders'] if ro['order_no'] == ord_code), None)
            chip_label = f"🟢 Đơn {ord_code} ({ro_m['max_runnable_qty']:,} pcs)" if ro_m else f"⚠️ Đơn {ord_code}"
            if chip_cols[c_i].button(chip_label, key=f"quick_{ord_code}", use_container_width=True):
                st.session_state.quick_selected_order = ord_code
                if ro_m:
                    st.session_state.quick_selected_qty = str(int(ro_m['max_runnable_qty']))
                st.rerun()

    # Chuẩn bị danh sách đơn hàng hiển thị gọn gàng và trực quan
    def format_order_item(k):
        ro = next((x for x in readiness_data['ready_orders'] if x['order_no'] == k), None)
        if ro:
            return f"🟢 Đơn {k} - {orders_info[k]['description']} (✅ ĐỦ 100% MÃ - Tối đa: {ro['max_runnable_qty']:,} pcs)"
        else:
            return f"⚪ Đơn {k} - {orders_info[k]['description']} (⚠️ Thiếu linh kiện)"
            
    col_input1, col_input2, col_input3 = safe_columns([5, 3, 2], vertical_alignment="bottom")
    
    with col_input1:
        # Tìm index mặc định theo order đang chọn
        curr_target = st.session_state.get('quick_selected_order', '5151')
        default_idx = sorted_order_keys.index(curr_target) if curr_target in sorted_order_keys else 0
        
        selected_code = st.selectbox(
            "1. Chọn hoặc gõ tìm mã đơn hàng (Ví dụ: 5151, 8333...):",
            options=sorted_order_keys,
            format_func=format_order_item,
            index=default_idx if sorted_order_keys else None,
            help="Hệ thống gắn nhãn 🟢 [ĐỦ 100% MÃ] cho các đơn sẵn sàng chạy ngay kèm số lượng tối đa."
        )
        
    with col_input2:
        # Ô nhập số lượng bằng tay (KHÔNG có nút tăng giảm +/-)
        qty_input_str = st.text_input(
            "2. Nhập số lượng cần chạy (bộ):",
            value=st.session_state.get('quick_selected_qty', "100"),
            placeholder="Nhập số lượng (Ví dụ: 100)...",
            help="Nhập số lượng nguyên đơn bằng tay. Hệ thống tự động nhân với định mức cả Trái và Phải."
        )
        
    with col_input3:
        btn_add = st.button("➕ BẤM OK ĐỂ TÍNH", type="primary", use_container_width=True, help="Đưa đơn hàng và số lượng đã nhập vào bảng chờ đối soát FIFO")

    # Hiển thị thông tin tóm tắt đơn đã chọn
    if selected_code:
        curr_order_id = selected_code
        curr_info = orders_info.get(curr_order_id, {})
        # Kiểm tra xem đơn này có nằm trong danh sách đủ mã không
        ready_match = next((ro for ro in readiness_data['ready_orders'] if ro['order_no'] == curr_order_id), None)
        if ready_match:
            ready_badge_str = f"<span class='badge-ok'>✅ ĐỦ 100% MÃ ĐỒNG BỘ (Tối đa chạy được {ready_match['max_runnable_qty']:,} pcs)</span>"
            bottleneck_str = f" | <span style='color:#D97706; font-weight:700;'>⚠️ Đang vướng mã có SL ít nhất: <b>{ready_match['bottleneck_code']}</b> ({ready_match['bottleneck_desc']} - tồn {int(ready_match['bottleneck_stock']):,} cái)</span>"
        else:
            ready_badge_str = "<span class='badge-missing'>⚠️ Đang thiếu linh kiện trong kho</span>"
            bottleneck_str = ""
        
        st.markdown(
            f"<div style='background:#F1F5F9; border-radius:10px; padding:12px 16px; font-size:14px; color:#334155; margin-top:10px; border-left:4px solid #2563EB;'>"
            f"ℹ️ <b>Đơn hàng {curr_order_id}</b>: {curr_info.get('description', '')} | "
            f"<b>Nguyên đơn</b>: 1 vế Trái + 1 vế Phải | "
            f"Định mức: <b>{curr_info.get('unique_components', 0)} ITEM CODE</b> | "
            f"Độ sẵn sàng: {ready_badge_str}"
            f"{bottleneck_str}"
            f"</div>",
            unsafe_allow_html=True
        )

    # Xử lý khi bấm nút OK
    if btn_add:
        clean_qty_val = None
        try:
            val = float(str(qty_input_str).replace(',', '').strip())
            if val > 0:
                clean_qty_val = val
            else:
                st.error("⚠️ Số lượng cần chạy phải lớn hơn 0.")
        except ValueError:
            st.error("⚠️ Vui lòng nhập số lượng hợp lệ (Ví dụ: 100).")
            
        if selected_code and clean_qty_val is not None:
            order_id = selected_code
            order_desc = orders_info[order_id]['description']
            timestamp_str = datetime.datetime.now().strftime("%H:%M:%S")
            
            # Thêm vào hàng chờ FIFO
            st.session_state.orders_queue.append({
                'id': f"{order_id}_{len(st.session_state.orders_queue)+1}_{uuid.uuid4().hex[:8]}",
                'order_no': order_id,
                'description': order_desc,
                'qty': clean_qty_val,
                'created_at': timestamp_str
            })
            st.success(f"✅ Đã thêm đơn hàng **{order_id}** (Số lượng: **{clean_qty_val:,g}** bộ) vào hàng chờ đối soát!")
            st.rerun()

# ==================== KHUNG HÀNG CHỜ TÍNH TOÁN (FIFO) ====================
st.markdown("### ⏳ Khung ghi nhận hàng chờ tính toán (FIFO)")

if len(st.session_state.orders_queue) == 0:
    st.info("💡 Hiện chưa có đơn hàng nào trong hàng chờ. Vui lòng nhập mã đơn hàng (Ví dụ: **8393**, **1009**) và số lượng ở trên rồi bấm **BẤM OK ĐỂ TÍNH**.")
else:
    col_q_left, col_q_right = safe_columns([8, 2], vertical_alignment="center")
    with col_q_left:
        st.caption("Nguyên tắc FIFO: Đơn hàng ở hàng trên được ưu tiên trừ tồn kho trước; đơn sau lấy phần tồn kho còn lại.")
    with col_q_right:
        if st.button("🗑️ Xóa toàn bộ hàng chờ", type="secondary", use_container_width=True):
            st.session_state.orders_queue = []
            st.rerun()
            
    # Hiển thị hàng chờ dạng bảng trực quan kèm cột đồng bộ tồn kho
    cols_header = safe_columns([1, 2, 3, 2, 3, 1, 1])
    cols_header[0].markdown("**Ưu tiên**")
    cols_header[1].markdown("**Mã đơn**")
    cols_header[2].markdown("**Tên sản phẩm**")
    cols_header[3].markdown("**Số lượng đặt**")
    cols_header[4].markdown("**Đồng bộ tồn kho**")
    cols_header[5].markdown("**Thời gian**")
    cols_header[6].markdown("**Xóa**")
    
    to_delete_idx = None
    for idx, item in enumerate(st.session_state.orders_queue):
        c = safe_columns([1, 2, 3, 2, 3, 1, 1], vertical_alignment="center")
        c[0].markdown(f"**#{idx + 1}**")
        c[1].markdown(f"<span class='badge-order'>{item['order_no']}</span>", unsafe_allow_html=True)
        c[2].write(item['description'])
        
        raw_qty = item.get('qty', 0)
        try:
            qty_num = float(raw_qty)
        except (ValueError, TypeError):
            qty_num = 0.0
        c[3].markdown(f"**{qty_num:,g}** bộ")
        
        ord_no_str = str(item.get('order_no', '')).strip()
        ro_match = next((x for x in readiness_data['ready_orders'] if str(x.get('order_no', '')).strip() == ord_no_str), None)
        if ro_match:
            max_pcs = float(ro_match.get('max_runnable_qty', 0))
            if qty_num <= max_pcs:
                c[4].markdown(f"<span class='badge-ok' style='font-size:12px; font-weight:800;'>✅ ĐỦ 100% MÃ (Tối đa {int(max_pcs):,} pcs)</span>", unsafe_allow_html=True)
            else:
                c[4].markdown(f"<span class='badge-missing' style='font-size:12px; font-weight:800;'>⚠️ VƯỢT TỒN KHO (Tối đa {int(max_pcs):,} pcs)</span>", unsafe_allow_html=True)
        else:
            nr_match = next((x for x in readiness_data['not_ready_orders'] if str(x.get('order_no', '')).strip() == ord_no_str), None)
            m_count = nr_match.get('missing_count', 'nhiều') if nr_match else 'nhiều'
            c[4].markdown(f"<span class='badge-missing' style='font-size:12px; font-weight:800;'>⚠️ THIẾU {m_count} MÃ TRONG KHO</span>", unsafe_allow_html=True)
            
        c[5].write(item.get('created_at', ''))
        if c[6].button("❌", key=f"del_{item['id']}"):
            to_delete_idx = idx
            
    if to_delete_idx is not None:
        st.session_state.orders_queue.pop(to_delete_idx)
        st.rerun()

st.divider()

# ==================== BÁO CÁO TRỰC QUAN NHANH TOÀN BỘ 39 ĐƠN ĐỒNG BỘ ĐỦ MÃ SẴN SÀNG CHẠY ====================
st.markdown(f"""
<div style="background: linear-gradient(135deg, #064E3B 0%, #059669 100%); border-radius: 16px; padding: 22px 28px; color: #FFFFFF; margin-top: 10px; margin-bottom: 20px; box-shadow: 0 10px 25px -5px rgba(5, 150, 105, 0.3);">
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <div>
            <div style="font-size: 22px; font-weight: 850; display:flex; align-items:center; gap:10px;">
                🚀 BÁO CÁO NHANH: TOÀN BỘ {readiness_data['ready_count']} ĐƠN HÀNG ĐỦ MÃ ĐỒNG BỘ SẴN SÀNG CHẠY
            </div>
            <div style="font-size: 14.5px; opacity: 0.95; margin-top: 6px;">
                Tổng hợp đầy đủ <b>{readiness_data['ready_count']}</b> đơn hàng đủ 100% Item Code trong kho | Tự động tính sẵn số lượng chạy tối đa theo tồn kho thực tế và thông báo mã hàng vướng giới hạn số lượng ít nhất.
            </div>
        </div>
        <div style="background: rgba(255,255,255,0.22); padding: 8px 18px; border-radius: 9999px; font-weight: 800; font-size: 14px; border: 1px solid rgba(255,255,255,0.35);">
            🎯 {readiness_data['ready_count']} / {readiness_data['total_orders']} Đơn Hàng Sẵn Sàng ({readiness_data['ready_percent']}%)
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KHUNG ĐIỀU KHIỂN: LỰA CHỌN ĐƠN ĐỂ ƯU TIÊN CHẠY TRƯỚC (FIFO)
st.markdown("""
<div style="background: #FFFFFF; border: 2px solid #059669; border-radius: 14px; padding: 18px 22px; margin-bottom: 18px; box-shadow: 0 4px 14px rgba(5, 150, 105, 0.08);">
    <div style="font-size: 16px; font-weight: 850; color: #064E3B; margin-bottom: 4px; display: flex; align-items: center; gap: 8px;">
        <span>🎯 LỰA CHỌN ĐƠN ĐỂ ƯU TIÊN CHẠY TRƯỚC VÀO HÀNG CHỜ (FIFO)</span>
        <span style="font-size: 12px; font-weight: 700; color: #059669; background: #ECFDF5; padding: 3px 10px; border-radius: 9999px; border: 1px solid #A7F3D0;">Tùy chọn đa năng</span>
    </div>
    <div style="font-size: 13.5px; color: #475569;">
        Chọn một hoặc nhiều đơn từ danh sách 39 đơn đủ mã. Thứ tự bạn chọn trong ô chính là <b>thứ tự ưu tiên trừ tồn kho FIFO</b> trong hàng chờ sản xuất.
    </div>
</div>
""", unsafe_allow_html=True)

col_prio_select, col_prio_btn = safe_columns([7, 3], vertical_alignment="bottom")

with col_prio_select:
    ready_order_choices = [ro['order_no'] for ro in readiness_data['ready_orders']]
    selected_priority_orders = st.multiselect(
        "1. Chọn các đơn hàng muốn ưu tiên chạy (thứ tự chọn = thứ tự ưu tiên trừ tồn kho FIFO):",
        options=ready_order_choices,
        format_func=lambda code: f"Đơn {code} - {orders_info.get(code, {}).get('description', '')} (Tối đa: {next((x['max_runnable_qty'] for x in readiness_data['ready_orders'] if x['order_no'] == code), 0):,} pcs)",
        placeholder="Gõ hoặc chọn các đơn cần chạy (Ví dụ: 5151, 8333, 4151, 1009, 5171...)...",
        key="multiselect_prio_orders"
    )

with col_prio_btn:
    btn_add_prio = st.button(
        f"🚀 ĐƯA {len(selected_priority_orders)} ĐƠN ĐÃ CHỌN VÀO HÀNG CHỜ" if selected_priority_orders else "🚀 ĐƯA CÁC ĐƠN VÀO HÀNG CHỜ",
        type="primary",
        use_container_width=True,
        disabled=len(selected_priority_orders) == 0,
        help="Thêm toàn bộ các đơn đã chọn vào hàng chờ FIFO phía trên theo đúng thứ tự ưu tiên đã chọn."
    )

if btn_add_prio and selected_priority_orders:
    for code in selected_priority_orders:
        ro_match = next((x for x in readiness_data['ready_orders'] if x['order_no'] == code), None)
        if ro_match:
            st.session_state.orders_queue.append({
                'id': f"{code}_{len(st.session_state.orders_queue)+1}_{uuid.uuid4().hex[:8]}",
                'order_no': code,
                'description': ro_match['description'],
                'qty': float(ro_match['max_runnable_qty']),
                'created_at': datetime.datetime.now().strftime("%H:%M:%S")
            })
    st.success(f"✅ Đã thêm {len(selected_priority_orders)} đơn hàng vào hàng chờ ưu tiên sản xuất!")
    st.rerun()

# Phím tắt thêm nhanh 1 chạm cho các mã đơn phổ biến
st.markdown("**⚡ Phím tắt thêm nhanh 1 chạm (Tự động nạp SL tối đa vào hàng chờ):**")
sample_quick_codes = ['5151', '5171', '8333', '4151', '1009', '5804']
chip_quick_cols = safe_columns(len(sample_quick_codes))
for idx_c, qc in enumerate(sample_quick_codes):
    ro_c = next((x for x in readiness_data['ready_orders'] if x['order_no'] == qc), None)
    if ro_c:
        with chip_quick_cols[idx_c]:
            if st.button(f"➕ {qc} ({ro_c['max_runnable_qty']:,} pcs)", key=f"fast_chip_{qc}", use_container_width=True, help=f"Thêm ngay đơn {qc} ({ro_c['max_runnable_qty']:,} pcs) vào hàng chờ"):
                st.session_state.orders_queue.append({
                    'id': f"{qc}_{len(st.session_state.orders_queue)+1}_{uuid.uuid4().hex[:8]}",
                    'order_no': qc,
                    'description': ro_c['description'],
                    'qty': float(ro_c['max_runnable_qty']),
                    'created_at': datetime.datetime.now().strftime("%H:%M:%S")
                })
                st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# THANH TRA CỨU, SẮP XẾP VÀ XUẤT EXCEL
col_search, col_sort, col_dl_ready = safe_columns([5, 3, 2], vertical_alignment="bottom")

with col_search:
    search_ready = st.text_input(
        "🔍 Tìm kiếm nhanh trong 39 đơn (gõ 5151, 8333, tên SP...):",
        value="",
        placeholder="Gõ mã đơn hoặc tên sản phẩm...",
        key="search_ready_input"
    )

with col_sort:
    sort_option = st.selectbox(
        "📊 Sắp xếp danh sách 39 đơn:",
        options=[
            "🔥 SL tối đa chạy được (Cao ➜ Thấp)",
            "🔢 SL tối đa chạy được (Thấp ➜ Cao)",
            "🏷️ Mã đơn hàng (A ➜ Z)",
            "📝 Tên sản phẩm (A ➜ Z)"
        ],
        index=0,
        key="sort_ready_orders_key"
    )

with col_dl_ready:
    readiness_excel = get_cached_readiness_excel(readiness_data)
    st.download_button(
        label="📥 Tải Excel 39 Đơn Sẵn Sàng",
        data=readiness_excel,
        file_name=f"Bao_Cao_Don_Du_Ma_San_Sang_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

# Lọc và sắp xếp danh sách 39 đơn
filtered_ready = [
    ro for ro in readiness_data['ready_orders']
    if search_ready.strip().lower() in ro['order_no'].lower() or search_ready.strip().lower() in ro['description'].lower()
] if search_ready.strip() else list(readiness_data['ready_orders'])

if "Cao ➜ Thấp" in sort_option:
    filtered_ready.sort(key=lambda x: x['max_runnable_qty'], reverse=True)
elif "Thấp ➜ Cao" in sort_option:
    filtered_ready.sort(key=lambda x: x['max_runnable_qty'], reverse=False)
elif "Mã đơn hàng" in sort_option:
    filtered_ready.sort(key=lambda x: x['order_no'])
elif "Tên sản phẩm" in sort_option:
    filtered_ready.sort(key=lambda x: x['description'])

tab_r1, tab_r2 = st.tabs([
    f"✅ Bảng Toàn Bộ Đơn Hàng Đủ Mã Đồng Bộ ({len(filtered_ready)} / {readiness_data['ready_count']} đơn - Sẵn Sàng Chạy)",
    f"⚠️ Danh Sách Đơn Hàng Chưa Đủ Mã ({readiness_data['not_ready_count']} đơn - Cần Nhập Thêm Linh Kiện)"
])

with tab_r1:
    st.caption(f"Hiển thị đầy đủ {len(filtered_ready)} đơn hàng đồng bộ đủ 100% linh kiện trong kho (Ẩn hoàn toàn đoạn cost, thể hiện trực quan thông tin quan trọng & thông báo mã hàng vướng số lượng ít nhất):")
    
    # BẢNG GIAO DIỆN CHÍNH: ẨN TOÀN BỘ ĐOẠN COST, CHỈ HIỂN THỊ THÔNG TIN QUAN TRỌNG VÀ THÔNG BÁO MÃ VƯỚNG SL ÍT NHẤT
    ready_table_html = """
    <div style="overflow-x: auto;">
    <table class="styled-table">
        <thead>
            <tr>
                <th style="width: 50px; text-align:center;">STT</th>
                <th style="width: 120px; text-align:center;">Mã Đơn Hàng</th>
                <th>Tên Sản Phẩm (Mô tả nguyên đơn)</th>
                <th style="text-align:center; width: 190px; background:#059669; color:#FFFFFF; font-size:14.5px;">SL TỐI ĐA CHẠY ĐƯỢC</th>
                <th>Thông Báo Mã Hàng Đang Vướng (Số Lượng Ít Nhất)</th>
                <th style="text-align:center; width: 140px;">Trạng Thái</th>
            </tr>
        </thead>
        <tbody>
    """
    for idx_r, r_item in enumerate(filtered_ready, 1):
        bottleneck_info = f"<span style='color:#D97706; font-weight:800;'>⚠️ Vướng mã: {r_item['bottleneck_code']}</span> - <span style='color:#1E293B; font-weight:600;'>{r_item['bottleneck_desc']}</span> <span style='color:#64748B; font-size:12px;'>(Tồn kho còn: {int(r_item['bottleneck_stock']):,} cái)</span>"
        ready_table_html += f"""
            <tr class="row-ok">
                <td style="text-align:center;"><b>#{idx_r}</b></td>
                <td style="text-align:center;"><span class="badge-order">{r_item['order_no']}</span></td>
                <td><b>{r_item['description']}</b></td>
                <td style="text-align:center; font-size:17px; font-weight:900; color:#059669; background:#ECFDF5;">
                    <b>{r_item['max_runnable_qty']:,} pcs</b>
                </td>
                <td>{bottleneck_info}</td>
                <td style="text-align:center;"><span class="badge-ok">✅ ĐỦ 100% MÃ</span></td>
            </tr>
        """
    ready_table_html += """
        </tbody>
    </table>
    </div>
    """
    st.markdown(ready_table_html, unsafe_allow_html=True)
    
    # Thanh nạp nhanh 1 đơn trực tiếp từ bảng
    if filtered_ready:
        col_sel_opt, col_act1, col_act2 = safe_columns([6, 2, 2], vertical_alignment="bottom")
        with col_sel_opt:
            ready_opts = [f"{ro['order_no']} - {ro['description']} (Tối đa: {ro['max_runnable_qty']:,} pcs)" for ro in filtered_ready]
            sel_opt = st.selectbox("⚡ Hoặc chọn 1 đơn bất kỳ từ bảng trên để nạp hoặc chạy ngay:", options=ready_opts, key="sel_ready_fast_table")
        with col_act1:
            if st.button("⚡ Nạp lên ô nhập", key="btn_act_load", use_container_width=True):
                chosen_code = sel_opt.split(" - ")[0].strip()
                chosen_item = next((x for x in filtered_ready if x['order_no'] == chosen_code), None)
                if chosen_item:
                    st.session_state.quick_selected_order = chosen_code
                    st.session_state.quick_selected_qty = str(int(chosen_item['max_runnable_qty']))
                    st.rerun()
        with col_act2:
            if st.button("➕ Thêm vào hàng chờ", key="btn_act_add", use_container_width=True):
                chosen_code = sel_opt.split(" - ")[0].strip()
                chosen_item = next((x for x in filtered_ready if x['order_no'] == chosen_code), None)
                if chosen_item:
                    st.session_state.orders_queue.append({
                        'id': f"{chosen_code}_{len(st.session_state.orders_queue)+1}_{uuid.uuid4().hex[:8]}",
                        'order_no': chosen_code,
                        'description': chosen_item['description'],
                        'qty': float(chosen_item['max_runnable_qty']),
                        'created_at': datetime.datetime.now().strftime("%H:%M:%S")
                    })
                    st.rerun()

with tab_r2:
    st.caption("Danh sách các đơn hàng chưa thể sản xuất do thiếu linh kiện trong kho (Đã ẩn đoạn cost và thông tin kỹ thuật rườm rà):")
    not_ready_table_html = """
    <div style="overflow-x: auto;">
    <table class="styled-table">
        <thead>
            <tr>
                <th style="width: 50px; text-align:center;">STT</th>
                <th style="width: 120px; text-align:center;">Mã Đơn Hàng</th>
                <th>Tên Sản Phẩm</th>
                <th style="text-align:center; width: 140px;">Tổng Item Code</th>
                <th style="text-align:center; width: 140px;">Có Trong Kho</th>
                <th style="text-align:center; width: 150px; background:#DC2626; color:#FFFFFF;">Số Mã THIẾU</th>
                <th style="text-align:center; width: 140px;">Trạng Thái</th>
            </tr>
        </thead>
        <tbody>
    """
    for idx_nr, nr_item in enumerate(readiness_data['not_ready_orders'], 1):
        not_ready_table_html += f"""
            <tr class="row-missing">
                <td style="text-align:center;"><b>#{idx_nr}</b></td>
                <td style="text-align:center;"><span class="badge-order" style="background:#FEE2E2; color:#991B1B; border-color:#FCA5A5;">{nr_item['order_no']}</span></td>
                <td>{nr_item['description']}</td>
                <td style="text-align:center;">{nr_item['total_components']} mã</td>
                <td style="text-align:center;">{nr_item['components_in_stock']} mã</td>
                <td style="text-align:center; font-weight:800; color:#DC2626;">Thiếu {nr_item['missing_count']} mã</td>
                <td style="text-align:center;"><span class="badge-missing">⚠️ Chưa đủ linh kiện</span></td>
            </tr>
        """
    not_ready_table_html += """
        </tbody>
    </table>
    </div>
    """
    st.markdown(not_ready_table_html, unsafe_allow_html=True)

st.divider()


# ==================== TÍNH TOÁN SO SÁNH VÀ KẾT QUẢ ====================
if len(st.session_state.orders_queue) > 0:
    # Header kết quả kèm nút điều khiển ẩn/hiện bảng chi tiết
    col_res_title, col_res_toggle = safe_columns([7, 3], vertical_alignment="center")
    with col_res_title:
        st.markdown("### 📊 Kết Quả Đối Soát Đơn Hàng (FIFO)")
    with col_res_toggle:
        # Nút bật tắt hiển thị bảng chi tiết để không bị rối mắt
        expand_all = st.toggle("👁️ Mở rộng tất cả bảng chi tiết BOM", value=False, help="Bật để mở bung toàn bộ bảng BOM của tất cả đơn hàng, tắt để ẩn gọn gàng.")
    
    with st.spinner("Đang tính toán phân bổ tồn kho FIFO cho cả Trái & Phải..."):
        calculation_results = core_engine.calculate_inventory_allocation(
            st.session_state.orders_queue,
            df_bom,
            df_stock
        )
        
    # Nút tải file tổng hợp tất cả các đơn
    col_dl_all, _ = safe_columns([5, 5])
    with col_dl_all:
        all_excel_bytes = get_cached_all_excel(calculation_results)
        st.download_button(
            label="📥 Tải trọn bộ Excel tất cả đơn hàng (Mỗi đơn 1 sheet)",
            data=all_excel_bytes,
            file_name=f"Doi_Soat_Tong_Hop_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Hiển thị từng đơn hàng trong danh sách chờ dưới dạng thẻ tóm tắt lớn
    for r in calculation_results:
        order_no = r['order_no']
        order_qty = r['order_qty']
        is_fully_ok = r.get('is_fully_ok', False)
        stt = r['stt']
        
        status_badge = "<span class='badge-ok'>✅ ĐỦ VẬT TƯ (100% OK)</span>" if is_fully_ok else f"<span class='badge-missing'>⚠️ THIẾU {r['missing_count']} ITEM CODE</span>"
        card_class = "order-card-ok" if is_fully_ok else "order-card-missing"
        
        # Thẻ thông tin lớn, rõ ràng cho từng đơn hàng
        st.markdown(f"""
        <div class="order-card {card_class}">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <span style="font-size: 18px; font-weight: 800; color: #1E3A8A;">Ưu tiên #{stt} - Đơn hàng {order_no}</span>
                    <span style="font-size: 15px; color: #475569; margin-left: 8px;">({r.get('product_title', '')})</span>
                    <div style="margin-top: 4px; font-size: 14px; color: #64748B;">
                        Số lượng đặt: <b>{order_qty:,g} bộ</b> (Trái + Phải) | Trạng thái: {status_badge}
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Hàng nút tải Excel và KPI nhanh của đơn
        col_card_kpi, col_card_dl = safe_columns([7, 3], vertical_alignment="center")
        with col_card_kpi:
            k1, k2, k3 = safe_columns(3)
            with k1:
                st.metric("Tổng Item Code", f"{r['total_items']} mã")
            with k2:
                st.metric("Item Code ĐỦ", f"{r['ok_count']} mã", delta=None)
            with k3:
                st.metric("Item Code THIẾU", f"{r['missing_count']} mã", delta=f"-{r['missing_count']}" if r['missing_count'] > 0 else "0", delta_color="inverse")
        with col_card_dl:
            single_excel = get_cached_single_order_excel(r)
            st.download_button(
                label=f"📥 Tải Excel đơn {order_no}",
                data=single_excel,
                file_name=f"Doi_Soat_Don_{order_no}_SL{int(order_qty)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_single_{r['id']}",
                use_container_width=True
            )
            
        # ==================== BẢNG CHI TIẾT ĐỐI SOÁT (MẶC ĐỊNH ẨN CHO ĐỠ RỐI MẮT) ====================
        with st.expander(f"📋 Bảng Chi Tiết BOM Cần Chạy & Tồn Kho - Đơn {order_no} (Nhấn để Mở / Đóng)", expanded=expand_all):
            tab_summary, tab_detail = st.tabs([
                "📋 Bảng Tổng Hợp BOM Cần Chạy & Tồn Kho (So Sánh)",
                "🔍 Chi Tiết Định Mức Từng Vế Trái & Phải"
            ])
            
            # Tab 1: Tổng hợp BOM Cần Chạy & Tồn Kho (Đã đổi sang ITEM CODE & DESCRIPTION)
            with tab_summary:
                df_s = r['summary_table']
                if len(df_s) == 0:
                    st.warning("Không có dữ liệu linh kiện.")
                else:
                    st.caption("Bảng tinh gọn: Loại bỏ thông tin Batch/Bin, tập trung vào so sánh Số lượng BOM cần chạy (Trái + Phải) với Tồn kho hiện có.")
                    
                    table_html = """
                    <div style="overflow-x: auto;">
                    <table class="styled-table">
                        <thead>
                            <tr>
                                <th>ITEM CODE</th>
                                <th>DESCRIPTION (tên mô tả)</th>
                                <th style="text-align:right;">ĐM Trái</th>
                                <th style="text-align:right;">ĐM Phải</th>
                                <th style="text-align:right;">Tổng ĐM/bộ</th>
                                <th style="text-align:right;">SL Cần Trái</th>
                                <th style="text-align:right;">SL Cần Phải</th>
                                <th style="text-align:right; background:#1E3A8A; color:#FDE047;">Tổng BOM cần chạy</th>
                                <th style="text-align:right; background:#0F172A; color:#67E8F9;">Tồn kho hiện có</th>
                                <th style="text-align:right;">SL Cấp</th>
                                <th style="text-align:right;">SL Thiếu</th>
                                <th style="text-align:center;">Trạng thái</th>
                                <th>Ghi chú</th>
                            </tr>
                        </thead>
                        <tbody>
                    """
                    for _, row_item in df_s.iterrows():
                        is_ok_row = (row_item['Trạng thái'] == 'OK')
                        row_class = "row-ok" if is_ok_row else "row-missing"
                        badge = "<span class='badge-ok'>OK</span>" if is_ok_row else "<span class='badge-missing'>THIẾU</span>"
                        
                        table_html += f"""
                            <tr class="{row_class}">
                                <td><b>{row_item['ITEM CODE']}</b></td>
                                <td>{row_item['DESCRIPTION (tên mô tả)']}</td>
                                <td style="text-align:right;">{row_item['Định mức Trái']:,g}</td>
                                <td style="text-align:right;">{row_item['Định mức Phải']:,g}</td>
                                <td style="text-align:right;"><b>{row_item['Tổng định mức (1 bộ)']:,g}</b></td>
                                <td style="text-align:right;">{row_item['SL Cần Trái']:,g}</td>
                                <td style="text-align:right;">{row_item['SL Cần Phải']:,g}</td>
                                <td style="text-align:right; font-weight:700; color:#1E3A8A;">{row_item['SL BOM cần chạy']:,g}</td>
                                <td style="text-align:right; font-weight:700;">{row_item['Số lượng tồn kho']:,g}</td>
                                <td style="text-align:right;">{row_item['Số lượng cấp']:,g}</td>
                                <td style="text-align:right; color:{'#DC2626' if not is_ok_row else 'inherit'}; font-weight:700;">{row_item['Số lượng thiếu']:,g}</td>
                                <td style="text-align:center;">{badge}</td>
                                <td><b>{row_item['Ghi chú']}</b></td>
                            </tr>
                        """
                    table_html += """
                        </tbody>
                    </table>
                    </div>
                    """
                    st.markdown(table_html, unsafe_allow_html=True)
                    
            # Tab 2: Chi tiết theo từng vế Trái & Phải
            with tab_detail:
                df_d = r['detail_table']
                if len(df_d) == 0:
                    st.warning("Không có dữ liệu chi tiết BOM.")
                else:
                    detail_html = """
                    <div style="overflow-x: auto;">
                    <table class="styled-table">
                        <thead>
                            <tr>
                                <th>Vế</th>
                                <th>Mã SP (Product Code)</th>
                                <th>ITEM CODE</th>
                                <th>DESCRIPTION (tên mô tả)</th>
                                <th style="text-align:right;">Định mức BOM</th>
                                <th style="text-align:right;">Số lượng đơn</th>
                                <th style="text-align:right;">Số lượng cần</th>
                                <th style="text-align:right;">Tồn kho khả dụng</th>
                                <th style="text-align:center;">Trạng thái</th>
                                <th>Ghi chú</th>
                            </tr>
                        </thead>
                        <tbody>
                    """
                    for _, row_item in df_d.iterrows():
                        is_ok_row = (row_item['Trạng thái'] == 'OK')
                        row_class = "row-ok" if is_ok_row else "row-missing"
                        badge = "<span class='badge-ok'>OK</span>" if is_ok_row else "<span class='badge-missing'>THIẾU</span>"
                        side_badge = f"<span style='background:#E0E7FF; color:#3730A3; padding:2px 8px; border-radius:6px; font-weight:700;'>{row_item['Vế']}</span>"
                        
                        detail_html += f"""
                            <tr class="{row_class}">
                                <td>{side_badge}</td>
                                <td>{row_item['Mã sản phẩm (Product Code)']}</td>
                                <td><b>{row_item['ITEM CODE']}</b></td>
                                <td>{row_item['DESCRIPTION (tên mô tả)']}</td>
                                <td style="text-align:right;">{row_item['Định mức BOM']:,g}</td>
                                <td style="text-align:right;">{row_item['Số lượng đơn']:,g}</td>
                                <td style="text-align:right;"><b>{row_item['Số lượng cần']:,g}</b></td>
                                <td style="text-align:right;">{row_item['Tồn kho khả dụng']:,g}</td>
                                <td style="text-align:center;">{badge}</td>
                                <td><b>{row_item['Ghi chú']}</b></td>
                            </tr>
                        """
                    detail_html += """
                        </tbody>
                    </table>
                    </div>
                    """
                    st.markdown(detail_html, unsafe_allow_html=True)
                    
        st.markdown("<hr style='margin: 16px 0; border: none; border-top: 1px dashed #CBD5E1;'>", unsafe_allow_html=True)


if __name__ == "__main__":
    from streamlit.runtime import exists
    if not exists():
        import sys
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", __file__]
        sys.exit(stcli.main())

