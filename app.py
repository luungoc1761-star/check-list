import datetime
import streamlit as st
import pandas as pd
import core_engine

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Đối soát BOM Sản Xuất & Tồn Kho (FIFO)",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện hiện đại, chuyên nghiệp
st.markdown("""
<style>
    /* Tổng thể font & spacing */
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 4px;
    }
    .sub-title {
        font-size: 14px;
        color: #4B5563;
        margin-bottom: 20px;
    }
    
    /* Hộp nổi nhập liệu */
    .floating-box {
        background: #FFFFFF;
        border: 2px solid #3B82F6;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.15), 0 8px 10px -6px rgba(59, 130, 246, 0.1);
        margin-bottom: 24px;
    }
    .box-header {
        font-size: 17px;
        font-weight: 700;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 14px;
    }
    
    /* Khung danh sách chờ */
    .queue-container {
        background: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 24px;
    }
    
    /* Bảng HTML tô màu trực quan */
    .styled-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        font-family: inherit;
        border-radius: 8px;
        overflow: hidden;
        margin-top: 10px;
        margin-bottom: 14px;
    }
    .styled-table thead tr {
        background-color: #1E3A8A;
        color: #ffffff;
        text-align: left;
        font-weight: 600;
    }
    .styled-table th, .styled-table td {
        padding: 8px 12px;
        border: 1px solid #E5E7EB;
    }
    .styled-table tbody tr.row-ok {
        background-color: #EBFEEB;
        color: #14532D;
    }
    .styled-table tbody tr.row-missing {
        background-color: #FEF2F2;
        color: #991B1B;
        font-weight: 500;
    }
    .styled-table tbody tr:hover {
        filter: brightness(0.97);
    }
    
    .badge-ok {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 3px 8px;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 11px;
        display: inline-block;
    }
    .badge-missing {
        background-color: #FDE8E8;
        color: #9B1C1C;
        padding: 3px 8px;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 11px;
        display: inline-block;
    }
    .badge-order {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 4px 10px;
        border-radius: 8px;
        font-weight: 600;
    }
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 12px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# Khởi tạo session state
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

# Hàm load dữ liệu có cache
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

# ==================== SIDEBAR ====================
with st.sidebar:
    st.markdown("### ⚙️ Quản lý Dữ liệu")
    st.markdown("Hệ thống tự động nhận diện cấu trúc tiêu đề file mẫu.")
    
    # Nút upload Stock Balance
    st.markdown("#### 1. File Tồn kho (Stock Balance)")
    uploaded_stock = st.file_uploader(
        "Nhập file Stock Balance mới",
        type=["xlsx", "xls"],
        key="uploader_stock",
        help="Cột quan trọng: Stock Code, BATCH, BIN, Qty"
    )
    if uploaded_stock is not None:
        if st.session_state.custom_stock_name != uploaded_stock.name:
            st.session_state.custom_stock_bytes = uploaded_stock.getvalue()
            st.session_state.custom_stock_name = uploaded_stock.name
            st.cache_data.clear()
            st.success(f"Đã nạp file tồn kho: {uploaded_stock.name}")
            st.rerun()
            
    current_stock_name = st.session_state.custom_stock_name or "Stock Balance With Batch (3).xlsx (Mẫu)"
    st.caption(f"📁 Đang dùng: **{current_stock_name}** ({len(df_stock):,} dòng, {len(df_stock['Stock_Code'].unique()):,} mã)")
    
    st.divider()
    
    # Nút upload Production BOM
    st.markdown("#### 2. File Định mức (Production BOM)")
    uploaded_bom = st.file_uploader(
        "Nhập file Production BOM mới",
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
            
    current_bom_name = st.session_state.custom_bom_name or "Production BOM List 10.2.xlsx (Mẫu)"
    st.caption(f"📁 Đang dùng: **{current_bom_name}** ({len(df_bom):,} dòng, {len(orders_info):,} đơn hàng)")
    
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
    st.markdown("#### 📌 Quy tắc nhận diện đơn:")
    st.markdown("""
    - **Số đầu = 1**: 4 số đầu là mã đơn hàng, 2 số tiếp (`19` = Trái, `20` = Phải).
    - **Số đầu = 8**: Bỏ số 8, 4 số tiếp là mã đơn hàng, số thứ 6 (`1` = Trái, `2` = Phải).
    - **FIFO**: Đơn nhập trước được ưu tiên trừ kho trước.
    """)

# ==================== MAIN CONTENT ====================
st.markdown('<div class="main-title">📦 HỆ THỐNG ĐỐI SOÁT BOM SẢN XUẤT & TỒN KHO KHO HÀNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Tính toán trừ kho tự động theo thứ tự nhập (FIFO) cho từng đơn hàng có đủ hai vế Trái và Phải</div>', unsafe_allow_html=True)

# Thống kê nhanh dữ liệu
col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.metric("Tổng đơn hàng trong BOM", f"{len(orders_info):,} đơn")
with col_m2:
    st.metric("Mã linh kiện trong kho", f"{len(df_stock['Stock_Code'].unique()):,} mã")
with col_m3:
    st.metric("Tổng lượng tồn kho", f"{int(df_stock['Qty'].sum()):,} cái")
with col_m4:
    st.metric("Đơn hàng trong hàng chờ", f"{len(st.session_state.orders_queue)} đơn")

st.markdown("<br>", unsafe_allow_html=True)

# ==================== BẢNG NỔI ĐIỀN ĐƠN HÀNG ====================
st.markdown("""
<div class="floating-box">
    <div class="box-header">
        <span>📝 BẢNG NỔI NHẬP ĐƠN HÀNG & SỐ LƯỢNG CẦN CHẠY</span>
    </div>
</div>
""", unsafe_allow_html=True)

with st.container():
    # Chuẩn bị danh sách gợi ý đơn hàng
    sorted_order_keys = sorted(orders_info.keys())
    
    col_input1, col_input2, col_input3 = st.columns([4, 3, 2], vertical_alignment="bottom")
    
    with col_input1:
        # Hỗ trợ cả gõ tay và chọn từ dropdown
        order_options = [f"{k} - {orders_info[k]['description']}" for k in sorted_order_keys]
        selected_option = st.selectbox(
            "1. Chọn hoặc tìm kiếm Đơn hàng:",
            options=order_options,
            index=0 if order_options else None,
            help="Tìm theo mã đơn (ví dụ: 1573, 1220, 1009, 3576,...) hoặc tên sản phẩm"
        )
        
    with col_input2:
        input_qty = st.number_input(
            "2. Nhập Số lượng (bộ):",
            min_value=1,
            max_value=1000000,
            value=100,
            step=10,
            help="Số lượng sản phẩm sẽ được tự động nhân với định mức BOM của cả Trái và Phải"
        )
        
    with col_input3:
        btn_add = st.button("➕ BẤM OK", type="primary", use_container_width=True)

    # Hiển thị thông tin tóm tắt đơn đã chọn
    if selected_option:
        curr_order_id = selected_option.split(" - ")[0].strip()
        curr_info = orders_info.get(curr_order_id, {})
        sides_str = ", ".join(curr_info.get('sides', []))
        st.caption(f"ℹ️ **Đơn {curr_order_id}**: {curr_info.get('description', '')} | Vế: **{sides_str}** | Định mức: **{curr_info.get('unique_components', 0)} mã linh kiện** ({curr_info.get('total_bom_rows', 0)} dòng BOM)")

    # Xử lý khi bấm OK
    if btn_add:
        if selected_option and input_qty > 0:
            order_id = selected_option.split(" - ")[0].strip()
            order_desc = orders_info[order_id]['description']
            timestamp_str = datetime.datetime.now().strftime("%H:%M:%S")
            
            # Thêm vào hàng chờ
            st.session_state.orders_queue.append({
                'id': f"{order_id}_{len(st.session_state.orders_queue)+1}_{datetime.datetime.now().strftime('%f')}",
                'order_no': order_id,
                'description': order_desc,
                'qty': input_qty,
                'created_at': timestamp_str
            })
            st.success(f"✅ Đã thêm đơn hàng **{order_id}** (Số lượng: **{input_qty:,}**) vào bảng chờ tính toán!")
            st.rerun()

# ==================== KHUNG GHI LẠI CHỜ ====================
st.markdown("### ⏳ Khung ghi lại chờ tính toán so sánh (FIFO)")

if len(st.session_state.orders_queue) == 0:
    st.info("Hiện chưa có đơn hàng nào trong hàng chờ. Vui lòng điền Đơn hàng và Số lượng ở bảng nổi phía trên rồi bấm **OK** để thêm vào.")
else:
    # Hiển thị danh sách các đơn trong hàng chờ
    col_q_left, col_q_right = st.columns([8, 2], vertical_alignment="center")
    with col_q_left:
        st.caption("Ưu tiên tồn kho: Đơn hàng ở hàng trên được trừ tồn kho trước; các đơn sau sẽ lấy phần tồn kho còn lại.")
    with col_q_right:
        if st.button("🗑️ Xóa toàn bộ hàng chờ", type="secondary", use_container_width=True):
            st.session_state.orders_queue = []
            st.rerun()
            
    # Hiển thị bảng hàng chờ với nút xóa từng dòng
    queue_data = []
    for idx, item in enumerate(st.session_state.orders_queue, 1):
        queue_data.append({
            'Ưu tiên (FIFO)': f"#{idx}",
            'Mã đơn hàng': item['order_no'],
            'Tên sản phẩm': item['description'],
            'Số lượng đặt': f"{item['qty']:,}",
            'Thời gian thêm': item['created_at']
        })
        
    df_queue_display = pd.DataFrame(queue_data)
    
    # Hiển thị hàng chờ dạng bảng trực quan kèm nút xóa
    cols_header = st.columns([1, 2, 4, 2, 2, 1])
    cols_header[0].markdown("**Ưu tiên**")
    cols_header[1].markdown("**Mã đơn**")
    cols_header[2].markdown("**Sản phẩm**")
    cols_header[3].markdown("**Số lượng**")
    cols_header[4].markdown("**Thời gian**")
    cols_header[5].markdown("**Xóa**")
    
    to_delete_idx = None
    for idx, item in enumerate(st.session_state.orders_queue):
        c = st.columns([1, 2, 4, 2, 2, 1], vertical_alignment="center")
        c[0].markdown(f"**#{idx + 1}**")
        c[1].markdown(f"<span class='badge-order'>{item['order_no']}</span>", unsafe_allow_html=True)
        c[2].write(item['description'])
        c[3].markdown(f"**{item['qty']:,}**")
        c[4].write(item['created_at'])
        if c[5].button("❌", key=f"del_{item['id']}"):
            to_delete_idx = idx
            
    if to_delete_idx is not None:
        st.session_state.orders_queue.pop(to_delete_idx)
        st.rerun()

st.divider()

# ==================== TÍNH TOÁN SO SÁNH VÀ KẾT QUẢ ====================
if len(st.session_state.orders_queue) > 0:
    st.markdown("### 📊 Kết Quả So Sánh Tồn Kho & BOM")
    
    with st.spinner("Đang tính toán phân bổ tồn kho FIFO..."):
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
        
        status_badge = "<span class='badge-ok'>ĐỦ VẬT TƯ (100% OK)</span>" if is_fully_ok else f"<span class='badge-missing'>THIẾU {r['missing_count']} MÃ LINH KIỆN</span>"
        
        with st.expander(
            f"Ưu tiên #{stt} - Đơn hàng {order_no} | SL: {order_qty:,} | {r.get('product_title', '')} - {'✅ ĐỦ HÀNG' if is_fully_ok else '⚠️ THIẾU HÀNG'}",
            expanded=True
        ):
            # Header đơn hàng
            col_h1, col_h2, col_h3, col_h4 = st.columns([3, 2, 2, 3], vertical_alignment="center")
            with col_h1:
                st.markdown(f"**Đơn hàng**: `{order_no}` ({r.get('product_title', '')})")
            with col_h2:
                st.markdown(f"**Số lượng đặt**: `{order_qty:,}` bộ (Trái + Phải)")
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
                "📋 Bảng Tổng Hợp Linh Kiện (Xuất Kho)",
                "🔍 Chi Tiết Phân Bổ BOM Trái & Phải"
            ])
            
            # Tab 1: Tổng hợp linh kiện
            with tab_summary:
                df_s = r['summary_table']
                if len(df_s) == 0:
                    st.warning("Không có dữ liệu linh kiện.")
                else:
                    # Tạo bảng HTML với styling dòng xanh / đỏ trực quan
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
                                <th style="text-align:right;">SL Cần</th>
                                <th style="text-align:right;">Tồn kho trước trừ</th>
                                <th style="text-align:right;">SL Cấp</th>
                                <th style="text-align:right;">SL Thiếu</th>
                                <th style="text-align:center;">Trạng thái</th>
                                <th>Ghi chú</th>
                                <th>Vị trí BIN</th>
                                <th>Lô BATCH</th>
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
                                <td style="text-align:right;"><b>{row_item['Số lượng cần']:,g}</b></td>
                                <td style="text-align:right;">{row_item['Tồn kho trước trừ']:,g}</td>
                                <td style="text-align:right;">{row_item['Số lượng cấp']:,g}</td>
                                <td style="text-align:right; color:{'#991B1B' if not is_ok_row else 'inherit'};"><b>{row_item['Số lượng thiếu']:,g}</b></td>
                                <td style="text-align:center;">{badge}</td>
                                <td><b>{row_item['Ghi chú']}</b></td>
                                <td>{row_item['Vị trí BIN']}</td>
                                <td>{row_item['Lô BATCH']}</td>
                            </tr>
                        """
                    table_html += """
                        </tbody>
                    </table>
                    </div>
                    """
                    st.markdown(table_html, unsafe_allow_html=True)
                    
            # Tab 2: Chi tiết từng vế Trái/Phải theo BOM gốc
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
                                <th>Mã linh kiện</th>
                                <th>Tên linh kiện</th>
                                <th style="text-align:right;">Định mức BOM</th>
                                <th style="text-align:right;">SL Đơn</th>
                                <th style="text-align:right;">SL Cần</th>
                                <th style="text-align:right;">Tồn kho khả dụng</th>
                                <th style="text-align:center;">Trạng thái</th>
                                <th>Ghi chú</th>
                                <th>Vị trí BIN</th>
                                <th>Lô BATCH</th>
                            </tr>
                        </thead>
                        <tbody>
                    """
                    for _, row_item in df_d.iterrows():
                        is_ok_row = (row_item['Trạng thái'] == 'OK')
                        row_class = "row-ok" if is_ok_row else "row-missing"
                        badge = "<span class='badge-ok'>OK</span>" if is_ok_row else "<span class='badge-missing'>THIẾU</span>"
                        side_badge = f"<span style='background:#E0E7FF; color:#3730A3; padding:2px 6px; border-radius:4px; font-weight:600;'>{row_item['Vế']}</span>"
                        
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
                                <td>{row_item['Vị trí BIN']}</td>
                                <td>{row_item['Lô BATCH']}</td>
                            </tr>
                        """
                    detail_html += """
                        </tbody>
                    </table>
                    </div>
                    """
                    st.markdown(detail_html, unsafe_allow_html=True)
