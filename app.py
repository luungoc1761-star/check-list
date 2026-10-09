import datetime
import io
import streamlit as st
import pandas as pd
import core_engine

# ==================== CẤU HÌNH TRANG STREAMLIT ====================
st.set_page_config(
    page_title="Đối Soát BOM Sản Xuất & Tồn Kho (FIFO)",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS GIAO DIỆN HIỆN ĐẠI & SỐNG ĐỘNG ====================
st.markdown("""
<style>
    /* Google Fonts & Base styling */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Ẩn dấu tăng giảm +/- spin buttons trên các ô số lượng */
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
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 55%, #2563EB 100%);
        border-radius: 16px;
        padding: 26px 32px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25), 0 8px 10px -6px rgba(30, 58, 138, 0.15);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
    }
    .hero-banner::after {
        content: '';
        position: absolute;
        top: -40px;
        right: -40px;
        width: 180px;
        height: 180px;
        background: radial-gradient(circle, rgba(255,255,255,0.12) 0%, rgba(255,255,255,0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }
    .hero-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-desc {
        font-size: 14px;
        color: #E2E8F0;
        font-weight: 400;
        line-height: 1.5;
    }
    .hero-tag {
        display: inline-block;
        background: rgba(255, 255, 255, 0.18);
        backdrop-filter: blur(8px);
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
        margin-top: 10px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }

    /* KPI Metric Cards cao cấp */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 14px;
        margin-bottom: 24px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        position: relative;
        overflow: hidden;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 14px rgba(0, 0, 0, 0.07);
    }
    .kpi-card-border-blue { border-top: 4px solid #2563EB; }
    .kpi-card-border-green { border-top: 4px solid #10B981; }
    .kpi-card-border-purple { border-top: 4px solid #8B5CF6; }
    .kpi-card-border-amber { border-top: 4px solid #F59E0B; }

    .kpi-label {
        font-size: 12px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 800;
        color: #0F172A;
    }

    /* Bảng nổi nhập đơn hàng */
    .floating-box {
        background: #FFFFFF;
        border: 1.5px solid #2563EB;
        border-radius: 14px;
        padding: 22px;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.12), 0 8px 10px -6px rgba(37, 99, 235, 0.08);
        margin-bottom: 24px;
    }
    .box-header {
        font-size: 17px;
        font-weight: 700;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 12px;
    }

    /* Chip gợi ý đơn hàng */
    .chip-container {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
        margin-bottom: 14px;
    }
    .chip-title {
        font-size: 12px;
        font-weight: 600;
        color: #475569;
    }

    /* Bảng HTML đối soát chuyên nghiệp */
    .styled-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 13px;
        border-radius: 10px;
        overflow: hidden;
        margin-top: 12px;
        margin-bottom: 16px;
        border: 1px solid #CBD5E1;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
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

    /* Badges trạng thái */
    .badge-ok {
        background-color: #DCFCE7;
        color: #15803D;
        border: 1px solid #86EFAC;
        padding: 3px 9px;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 11px;
        display: inline-block;
        letter-spacing: 0.3px;
    }
    .badge-missing {
        background-color: #FEE2E2;
        color: #B91C1C;
        border: 1px solid #FCA5A5;
        padding: 3px 9px;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 11px;
        display: inline-block;
        letter-spacing: 0.3px;
    }
    .badge-order {
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        padding: 4px 10px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
    
    /* Khung danh sách chờ */
    .queue-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 24px;
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

# ==================== HÀM LOAD DỮ LIỆU CÓ CACHE ====================
@st.cache_data
def get_loaded_data(bom_bytes, bom_name, stock_bytes, stock_name):
    bom_source = io.BytesIO(bom_bytes) if bom_bytes else 'Production BOM List 10.2.xlsx'
    stock_source = io.BytesIO(stock_bytes) if stock_bytes else 'Stock Balance With Batch (3).xlsx'
    
    df_bom = core_engine.load_bom_data(bom_source)
    df_stock = core_engine.load_stock_data(stock_source)
    orders_info = core_engine.get_available_orders_info(df_bom)
    
    return df_bom, df_stock, orders_info

# Tải dữ liệu
df_bom, df_stock, orders_info = get_loaded_data(
    st.session_state.custom_bom_bytes,
    st.session_state.custom_bom_name,
    st.session_state.custom_stock_bytes,
    st.session_state.custom_stock_name
)

# ==================== SIDEBAR QUẢN LÝ DỮ LIỆU ====================
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
    st.caption(f"📁 **{current_stock_name}** ({len(df_stock):,} dòng, {len(df_stock['Stock_Code'].unique()):,} mã)")
    
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
    st.caption(f"📁 **{current_bom_name}** ({len(df_bom):,} dòng, {len(orders_info):,} đơn hàng)")
    
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
    - **1 Đơn hàng = 1 Bộ nguyên đơn**: Tự động lấy cả 2 vế **Trái và Phải**.
    - **Định mức nhân số lượng**: Hệ thống nhân riêng `ĐM Trái × SL` và `ĐM Phải × SL` rồi cộng lại thành tổng BOM cần chạy.
    - **Phân bổ FIFO**: Đơn nhập trước được ưu tiên trừ tồn kho trước.
    """)

# ==================== MAIN BANNER ====================
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">📦 HỆ THỐNG ĐỐI SOÁT BOM SẢN XUẤT & TỒN KHO</div>
    <div class="hero-desc">
        Tự động tính toán định mức nguyên đơn <b>(Trái + Phải)</b> theo số lượng cần chạy | 
        Phân bổ trừ tồn kho tự động theo thứ tự ưu tiên <b>FIFO</b> | 
        Bảng so sánh trực quan, tinh gọn, không rối mắt.
    </div>
    <div class="hero-tag">✨ Sẵn sàng vận hành sản xuất</div>
</div>
""", unsafe_allow_html=True)

# Thống kê nhanh dữ liệu (KPI Cards)
col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-blue">
        <div class="kpi-label">Tổng đơn hàng trong BOM</div>
        <div class="kpi-value">{len(orders_info):,} <span style="font-size:14px; font-weight:500; color:#64748B;">đơn</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-green">
        <div class="kpi-label">Mã linh kiện trong kho</div>
        <div class="kpi-value">{len(df_stock['Stock_Code'].unique()):,} <span style="font-size:14px; font-weight:500; color:#64748B;">mã</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-purple">
        <div class="kpi-label">Tổng lượng tồn kho</div>
        <div class="kpi-value">{int(df_stock['Qty'].sum()):,} <span style="font-size:14px; font-weight:500; color:#64748B;">cái</span></div>
    </div>
    """, unsafe_allow_html=True)
with col_m4:
    queue_len = len(st.session_state.orders_queue)
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-amber">
        <div class="kpi-label">Hàng chờ tính toán (FIFO)</div>
        <div class="kpi-value">{queue_len} <span style="font-size:14px; font-weight:500; color:#64748B;">đơn</span></div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==================== BẢNG NỔI NHẬP ĐƠN HÀNG ====================
st.markdown("""
<div class="floating-box">
    <div class="box-header">
        <span>🎯 NHẬP ĐƠN HÀNG & SỐ LƯỢNG CẦN CHẠY</span>
    </div>
</div>
""", unsafe_allow_html=True)

with st.container():
    sorted_order_keys = sorted(orders_info.keys())
    
    # Gợi ý đơn hàng nhanh (Chips)
    st.markdown("**Gợi ý đơn hàng mẫu:**")
    quick_sample_orders = ['8393', '1220', '1573', '1009', '3576']
    chip_cols = st.columns(len(quick_sample_orders) + 3)
    for c_i, ord_code in enumerate(quick_sample_orders):
        if ord_code in orders_info:
            if chip_cols[c_i].button(f"👉 Đơn {ord_code}", key=f"quick_{ord_code}", use_container_width=True):
                st.session_state.quick_selected_order = ord_code
                st.rerun()

    # Chuẩn bị danh sách đơn hàng hiển thị gọn gàng (KHÔNG để chữ Trái hay Phải)
    order_options_clean = [f"{k} - {orders_info[k]['description']}" for k in sorted_order_keys]
    
    # Tìm index mặc định theo order đang chọn
    curr_target = st.session_state.get('quick_selected_order', '8393')
    default_idx = 0
    for idx_item, opt_str in enumerate(order_options_clean):
        if opt_str.startswith(f"{curr_target} - ") or opt_str == curr_target:
            default_idx = idx_item
            break
            
    col_input1, col_input2, col_input3 = st.columns([5, 3, 2], vertical_alignment="bottom")
    
    with col_input1:
        # Ô chọn / tìm kiếm mã đơn hàng (Định dạng gọn: 8393 - Tên SP, KHÔNG có Trái/Phải)
        selected_option = st.selectbox(
            "1. Chọn hoặc gõ tìm mã đơn hàng (Ví dụ: 8393):",
            options=order_options_clean,
            index=default_idx if order_options_clean else None,
            help="Chỉ cần nhập hoặc chọn mã đơn hàng (Ví dụ: 8393, 1220, 1009...). Không cần chọn Trái hay Phải."
        )
        
    with col_input2:
        # Ô nhập số lượng bằng tay (KHÔNG có nút tăng giảm +/-)
        qty_input_str = st.text_input(
            "2. Nhập số lượng cần chạy (bộ):",
            value="100",
            placeholder="Nhập số lượng (Ví dụ: 100)...",
            help="Nhập số lượng nguyên đơn bằng tay. Hệ thống tự động nhân với định mức cả Trái và Phải."
        )
        
    with col_input3:
        btn_add = st.button("➕ BẤM OK ĐỂ TÍNH", type="primary", use_container_width=True)

    # Hiển thị thông tin tóm tắt đơn đã chọn
    if selected_option:
        curr_order_id = selected_option.split(" - ")[0].strip()
        curr_info = orders_info.get(curr_order_id, {})
        st.markdown(
            f"<div style='background:#F1F5F9; border-radius:8px; padding:10px 14px; font-size:13px; color:#334155; margin-top:8px;'>"
            f"ℹ️ <b>Đơn hàng {curr_order_id}</b>: {curr_info.get('description', '')} | "
            f"<b>Nguyên đơn</b>: Tự động gom đủ 1 vế Trái + 1 vế Phải | "
            f"Định mức: <b>{curr_info.get('unique_components', 0)} mã linh kiện</b>"
            f"</div>",
            unsafe_allow_html=True
        )

    # Xử lý khi bấm nút OK
    if btn_add:
        # Kiểm tra tính hợp lệ của số lượng nhập tay
        clean_qty_val = None
        try:
            val = float(str(qty_input_str).replace(',', '').strip())
            if val > 0:
                clean_qty_val = val
            else:
                st.error("⚠️ Số lượng cần chạy phải lớn hơn 0.")
        except ValueError:
            st.error("⚠️ Vui lòng nhập số lượng hợp lệ (Ví dụ: 100).")
            
        if selected_option and clean_qty_val is not None:
            order_id = selected_option.split(" - ")[0].strip()
            order_desc = orders_info[order_id]['description']
            timestamp_str = datetime.datetime.now().strftime("%H:%M:%S")
            
            # Thêm vào hàng chờ FIFO
            st.session_state.orders_queue.append({
                'id': f"{order_id}_{len(st.session_state.orders_queue)+1}_{datetime.datetime.now().strftime('%f')}",
                'order_no': order_id,
                'description': order_desc,
                'qty': clean_qty_val,
                'created_at': timestamp_str
            })
            st.success(f"✅ Đã thêm đơn hàng **{order_id}** (Số lượng: **{clean_qty_val:,g}** bộ) vào hàng chờ tính toán!")
            st.rerun()

# ==================== KHUNG HÀNG CHỜ TÍNH TOÁN (FIFO) ====================
st.markdown("### ⏳ Khung ghi nhận hàng chờ tính toán (FIFO)")

if len(st.session_state.orders_queue) == 0:
    st.info("💡 Hiện chưa có đơn hàng nào trong hàng chờ. Vui lòng nhập mã đơn hàng (Ví dụ: **8393**) và số lượng ở trên rồi bấm **BẤM OK ĐỂ TÍNH**.")
else:
    col_q_left, col_q_right = st.columns([8, 2], vertical_alignment="center")
    with col_q_left:
        st.caption("Nguyên tắc FIFO: Đơn hàng ở hàng trên được ưu tiên trừ tồn kho trước; đơn sau lấy phần tồn kho còn lại.")
    with col_q_right:
        if st.button("🗑️ Xóa toàn bộ hàng chờ", type="secondary", use_container_width=True):
            st.session_state.orders_queue = []
            st.rerun()
            
    # Hiển thị hàng chờ dạng bảng trực quan kèm nút xóa
    cols_header = st.columns([1, 2, 4, 2, 2, 1])
    cols_header[0].markdown("**Ưu tiên**")
    cols_header[1].markdown("**Mã đơn**")
    cols_header[2].markdown("**Tên sản phẩm**")
    cols_header[3].markdown("**Số lượng đặt**")
    cols_header[4].markdown("**Thời gian**")
    cols_header[5].markdown("**Xóa**")
    
    to_delete_idx = None
    for idx, item in enumerate(st.session_state.orders_queue):
        c = st.columns([1, 2, 4, 2, 2, 1], vertical_alignment="center")
        c[0].markdown(f"**#{idx + 1}**")
        c[1].markdown(f"<span class='badge-order'>{item['order_no']}</span>", unsafe_allow_html=True)
        c[2].write(item['description'])
        c[3].markdown(f"**{item['qty']:,g}** bộ")
        c[4].write(item['created_at'])
        if c[5].button("❌", key=f"del_{item['id']}"):
            to_delete_idx = idx
            
    if to_delete_idx is not None:
        st.session_state.orders_queue.pop(to_delete_idx)
        st.rerun()

st.divider()

# ==================== TÍNH TOÁN SO SÁNH VÀ KẾT QUẢ ====================
if len(st.session_state.orders_queue) > 0:
    st.markdown("### 📊 Kết Quả So Sánh Tồn Kho & BOM Cần Chạy")
    
    with st.spinner("Đang tính toán phân bổ tồn kho FIFO cho cả Trái & Phải..."):
        calculation_results = core_engine.calculate_inventory_allocation(
            st.session_state.orders_queue,
            df_bom,
            df_stock
        )
        
    # Nút tải file tổng hợp tất cả các đơn
    col_dl_all, _ = st.columns([4, 6])
    with col_dl_all:
        all_excel_bytes = core_engine.export_all_orders_to_excel(calculation_results)
        st.download_button(
            label="📥 Tải trọn bộ Excel tất cả đơn hàng (Mỗi đơn 1 sheet)",
            data=all_excel_bytes,
            file_name=f"Doi_Soat_Tong_Hop_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Hiển thị từng đơn hàng trong danh sách chờ
    for r in calculation_results:
        order_no = r['order_no']
        order_qty = r['order_qty']
        is_fully_ok = r.get('is_fully_ok', False)
        stt = r['stt']
        
        status_badge = "<span class='badge-ok'>✅ ĐỦ VẬT TƯ (100% OK)</span>" if is_fully_ok else f"<span class='badge-missing'>⚠️ THIẾU {r['missing_count']} MÃ LINH KIỆN</span>"
        
        with st.expander(
            f"Ưu tiên #{stt} - Đơn hàng {order_no} | SL: {order_qty:,g} bộ | {r.get('product_title', '')} - {'✅ ĐỦ HÀNG' if is_fully_ok else '⚠️ THIẾU HÀNG'}",
            expanded=True
        ):
            # Header đơn hàng
            col_h1, col_h2, col_h3, col_h4 = st.columns([3, 2, 2, 3], vertical_alignment="center")
            with col_h1:
                st.markdown(f"**Đơn hàng**: `{order_no}` ({r.get('product_title', '')})")
            with col_h2:
                st.markdown(f"**Số lượng đặt**: `{order_qty:,g}` bộ (Trái + Phải)")
            with col_h3:
                st.markdown(f"**Trạng thái**: {status_badge}", unsafe_allow_html=True)
            with col_h4:
                # Nút tải file Excel riêng cho đơn này
                single_excel = core_engine.export_order_to_excel(r)
                st.download_button(
                    label=f"📥 Tải Excel đơn {order_no}",
                    data=single_excel,
                    file_name=f"Doi_Soat_Don_{order_no}_SL{int(order_qty)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_{r['id']}",
                    use_container_width=True
                )
                
            # Thống kê nhanh đơn
            c_k1, c_k2, c_k3 = st.columns(3)
            with c_k1:
                st.metric("Tổng mã linh kiện", f"{r['total_items']} mã")
            with c_k2:
                st.metric("Mã linh kiện ĐỦ", f"{r['ok_count']} mã", delta=None)
            with c_k3:
                st.metric("Mã linh kiện THIẾU", f"{r['missing_count']} mã", delta=f"-{r['missing_count']}" if r['missing_count'] > 0 else "0", delta_color="inverse")
                
            tab_summary, tab_detail = st.tabs([
                "📋 Bảng Tổng Hợp BOM Cần Chạy & Tồn Kho (So Sánh)",
                "🔍 Chi Tiết Định Mức Từng Vế Trái & Phải"
            ])
            
            # Tab 1: Tổng hợp BOM Cần Chạy & Tồn Kho (Đã loại bỏ BATCH và BIN)
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
                                <th>Mã linh kiện</th>
                                <th>Tên linh kiện</th>
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
                                <td><b>{row_item['Mã linh kiện']}</b></td>
                                <td>{row_item['Tên linh kiện']}</td>
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
                                <th>Mã linh kiện (Component Code)</th>
                                <th>Tên linh kiện</th>
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
                                <td><b>{row_item['Mã linh kiện (Component Code)']}</b></td>
                                <td>{row_item['Tên linh kiện']}</td>
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
