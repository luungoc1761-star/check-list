import os
import io
import json
import uuid
import inspect
import datetime
import importlib
import streamlit as st
import pandas as pd
import core_engine
import storage_manager

# Tự động nạp lại module nếu môi trường Streamlit giữ bản cache cũ của core_engine / storage_manager
_APP_ENGINE_VERSION = "2026.11.07.v3"
if getattr(core_engine, "__version__", None) != _APP_ENGINE_VERSION or not hasattr(core_engine, "load_repo_stock_data"):
    core_engine = importlib.reload(core_engine)
if not hasattr(storage_manager, "_normalize_queue_items"):
    storage_manager = importlib.reload(storage_manager)


def _get_stretch_kwargs(widget_fn):
    """Tương thích đa phiên bản Streamlit cho tham số giãn rộng chiều ngang."""
    try:
        params = inspect.signature(widget_fn).parameters
        if "width" in params:
            return {"width": "stretch"}
        if "use_container_width" in params:
            return {"use_container_width": True}
    except Exception:
        pass
    return {"use_container_width": True}


_STRETCH_KW = _get_stretch_kwargs(st.button)
_DF_STRETCH_KW = _get_stretch_kwargs(st.dataframe)
_COLUMNS_SUPPORTS_VALIGN = "vertical_alignment" in inspect.signature(st.columns).parameters
_SELECTBOX_SUPPORTS_NEW = "accept_new_options" in inspect.signature(st.selectbox).parameters


def safe_columns(spec, vertical_alignment=None):
    if vertical_alignment and _COLUMNS_SUPPORTS_VALIGN:
        return st.columns(spec, vertical_alignment=vertical_alignment)
    return st.columns(spec)


def safe_dataframe(df, height=None, hide_index=True):
    kwargs = dict(_DF_STRETCH_KW)
    if height is not None:
        kwargs["height"] = height
    try:
        return st.dataframe(df, hide_index=hide_index, **kwargs)
    except TypeError:
        return st.dataframe(df, **kwargs)


st.set_page_config(
    page_title="Hệ Thống Quản Lý Đối Soát Tồn Kho & BOM Sản Xuất",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CSS GIAO DIỆN HIỆN ĐẠI - RÕ NÉT - TỐC ĐỘ CAO ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .stApp {
        background: linear-gradient(180deg, #F0F4F8 0%, #F8FAFC 40%, #EFF6FF 100%);
        color: #0F172A;
    }

    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 97% !important;
    }

    /* Ẩn các thành phần thừa của Streamlit để giao diện như Web App độc lập */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%) !important;
        border-right: 1px solid #CBD5E1;
    }
    [data-testid="stSidebar"] * {
        color: #1E293B;
    }

    /* BANNER CHÍNH */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 18px;
        padding: 22px 28px;
        color: #FFFFFF !important;
        box-shadow: 0 12px 28px -6px rgba(30, 58, 138, 0.32);
        margin-bottom: 18px;
        border: 1px solid rgba(255, 255, 255, 0.14);
    }
    .hero-title {
        font-size: 25px;
        font-weight: 800;
        letter-spacing: -0.4px;
        color: #FFFFFF !important;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #E2E8F0 !important;
        font-weight: 500;
        line-height: 1.55;
    }

    /* KPI CARDS */
    .kpi-card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 16px 18px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
        margin-bottom: 8px;
    }
    .kpi-card-border-blue { border-left: 5px solid #2563EB; }
    .kpi-card-border-green { border-left: 5px solid #10B981; }
    .kpi-card-border-emerald { border-left: 5px solid #059669; background: linear-gradient(135deg, #FFFFFF 0%, #ECFDF5 100%); }
    .kpi-card-border-purple { border-left: 5px solid #8B5CF6; }
    .kpi-card-border-amber { border-left: 5px solid #F59E0B; }

    .kpi-label {
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748B;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.2;
    }
    .kpi-unit {
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        margin-left: 4px;
    }

    /* KHUNG NỔI (FLOATING BOX) */
    .floating-box {
        background: #FFFFFF;
        border-radius: 16px;
        padding: 18px 22px;
        border: 1.5px solid #CBD5E1;
        box-shadow: 0 8px 20px -4px rgba(15, 23, 42, 0.05);
        margin-bottom: 14px;
    }
    .box-header {
        font-size: 16.5px;
        font-weight: 800;
        color: #0F172A;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    /* THẺ ĐƠN HÀNG ĐỐI SOÁT */
    .order-card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 12px;
        border: 1.5px solid #E2E8F0;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.04);
    }
    .order-card-ok {
        border-left: 7px solid #10B981;
        background: linear-gradient(90deg, #ECFDF5 0%, #FFFFFF 20%);
    }
    .order-card-missing {
        border-left: 7px solid #EF4444;
        background: linear-gradient(90deg, #FEF2F2 0%, #FFFFFF 20%);
    }
    .order-card-partial {
        border-left: 7px solid #F59E0B;
        background: linear-gradient(90deg, #FFFBEB 0%, #FFFFFF 20%);
    }

    /* BADGES */
    .badge-ok {
        background-color: #D1FAE5;
        color: #065F46;
        padding: 5px 12px;
        border-radius: 999px;
        font-weight: 800;
        font-size: 12.5px;
        border: 1px solid #6EE7B7;
        display: inline-block;
    }
    .badge-missing {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 5px 12px;
        border-radius: 999px;
        font-weight: 800;
        font-size: 12.5px;
        border: 1px solid #FCA5A5;
        display: inline-block;
    }
    .badge-runnable {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 4px 10px;
        border-radius: 999px;
        font-weight: 800;
        font-size: 12px;
        border: 1px solid #FCD34D;
        display: inline-block;
    }
    .badge-order {
        background: linear-gradient(135deg, #DBEAFE 0%, #EFF6FF 100%);
        color: #1E40AF;
        padding: 4px 10px;
        border-radius: 8px;
        font-weight: 800;
        font-size: 13.5px;
        border: 1px solid #93C5FD;
    }
    .badge-wh {
        background: linear-gradient(135deg, #F3E8FF 0%, #FAF5FF 100%);
        color: #6B21A8;
        padding: 4px 10px;
        border-radius: 8px;
        font-weight: 800;
        font-size: 12.5px;
        border: 1px solid #D8B4FE;
        display: inline-block;
    }

    /* TABS HIỆN ĐẠI */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #E2E8F0;
        padding: 6px;
        border-radius: 14px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 46px;
        background-color: transparent;
        border-radius: 10px;
        padding: 8px 18px;
        font-weight: 700;
        font-size: 14.5px;
        color: #475569 !important;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background: #FFFFFF !important;
        color: #1D4ED8 !important;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.08) !important;
    }

    /* BẢNG HTML TÙY CHỈNH */
    .styled-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 13.5px;
        background-color: #FFFFFF;
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #CBD5E1;
        margin-bottom: 12px;
    }
    .styled-table thead tr {
        background: linear-gradient(90deg, #0F172A 0%, #1E3A8A 100%);
        color: #FFFFFF;
        text-align: left;
        font-weight: 700;
    }
    .styled-table th, .styled-table td {
        padding: 10px 14px;
        border-bottom: 1px solid #E2E8F0;
    }
    .styled-table tbody tr.row-ok {
        background-color: #F0FDF4;
        color: #064E3B;
    }
    .styled-table tbody tr.row-ok:hover {
        background-color: #DCFCE7;
    }
</style>
""", unsafe_allow_html=True)


# ==================== HÀM TIỆN ÍCH LƯU TRỮ BỀN VỮNG ====================
PERSIST_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".persisted_data")
os.makedirs(PERSIST_DIR, exist_ok=True)
BOM_PERSIST_FILE = os.path.join(PERSIST_DIR, "saved_bom.xlsx")
STOCK_PERSIST_FILE = os.path.join(PERSIST_DIR, "saved_stock.xlsx")
PLAN_PERSIST_FILE = os.path.join(PERSIST_DIR, "saved_plan.xlsx")
META_PERSIST_FILE = os.path.join(PERSIST_DIR, "saved_meta.json")


def load_saved_meta():
    if os.path.exists(META_PERSIST_FILE):
        try:
            with open(META_PERSIST_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_meta(meta_dict):
    try:
        current = load_saved_meta()
        current.update(meta_dict)
        with open(META_PERSIST_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False)
    except Exception:
        pass


def find_data_file(keywords):
    search_dirs = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        for fname in os.listdir(d):
            if fname.startswith('~$') or fname.startswith('.'):
                continue
            if fname.lower().endswith(('.xlsx', '.xls', '.csv')):
                fl = fname.lower()
                if all(k.lower() in fl for k in keywords):
                    return os.path.join(d, fname)
    return None


def get_repo_stock_signature():
    """Tạo chữ ký (signature) tự động thay đổi ngay khi có file tồn kho mới được nạp vào hoặc cập nhật trong repo."""
    try:
        files = core_engine.find_all_stock_files()
        sig_parts = []
        for f in files:
            try:
                st_info = os.stat(f)
                sig_parts.append(f"{os.path.basename(f)}:{int(st_info.st_mtime)}:{st_info.st_size}")
            except Exception:
                sig_parts.append(os.path.basename(f))
        return "|".join(sig_parts)
    except Exception:
        return "default"


# ==================== KHỞI TẠO SESSION STATE ====================
saved_meta = load_saved_meta()

if 'custom_bom_bytes' not in st.session_state:
    if os.path.exists(BOM_PERSIST_FILE):
        with open(BOM_PERSIST_FILE, "rb") as f:
            st.session_state.custom_bom_bytes = f.read()
        st.session_state.custom_bom_name = saved_meta.get('bom_name', 'Saved_BOM.xlsx')
    else:
        st.session_state.custom_bom_bytes = None
        st.session_state.custom_bom_name = None

if 'custom_stock_bytes' not in st.session_state:
    if os.path.exists(STOCK_PERSIST_FILE):
        with open(STOCK_PERSIST_FILE, "rb") as f:
            st.session_state.custom_stock_bytes = f.read()
        st.session_state.custom_stock_name = saved_meta.get('stock_name', 'Saved_Stock.xlsx')
    else:
        st.session_state.custom_stock_bytes = None
        st.session_state.custom_stock_name = None

if 'plan_cleared' not in st.session_state:
    st.session_state.plan_cleared = bool(saved_meta.get('plan_cleared', False))

if 'custom_plan_bytes' not in st.session_state:
    if not st.session_state.plan_cleared and os.path.exists(PLAN_PERSIST_FILE):
        with open(PLAN_PERSIST_FILE, "rb") as f:
            st.session_state.custom_plan_bytes = f.read()
        st.session_state.custom_plan_name = saved_meta.get('plan_name', 'Saved_Plan.xlsx')
    else:
        st.session_state.custom_plan_bytes = None
        st.session_state.custom_plan_name = None

if 'orders_queue' not in st.session_state:
    st.session_state.orders_queue = storage_manager.load_queue("shared", storage_manager.WORKSPACE_SHARED)

if 'queue_manually_cleared' not in st.session_state:
    st.session_state.queue_manually_cleared = False


def save_orders_queue(queue_items):
    """Lưu hàng chờ FIFO bền vững vào bộ nhớ dùng chung của ứng dụng."""
    st.session_state.orders_queue = queue_items
    storage_manager.save_queue(queue_items, "shared", storage_manager.WORKSPACE_SHARED)


# ==================== CÁC HÀM CACHE TỐI ƯU HIỆU NĂNG ====================
@st.cache_data(show_spinner=False)
def get_loaded_data(bom_bytes, stock_bytes, plan_bytes=None, bom_name=None, stock_name=None, plan_name=None, plan_cleared=False, repo_stock_sig=""):
    if bom_bytes is not None:
        bio_bom = io.BytesIO(bom_bytes)
        if bom_name:
            bio_bom.name = bom_name
        df_bom = core_engine.load_bom_data(bio_bom)
    else:
        bom_file = find_data_file(['bom'])
        if not bom_file:
            raise FileNotFoundError("Không tìm thấy file Production BOM List.xlsx trong thư mục gốc.")
        df_bom = core_engine.load_bom_data(bom_file)

    resolved_stock_label = stock_name
    if stock_bytes is not None:
        bio_stock = io.BytesIO(stock_bytes)
        if stock_name:
            bio_stock.name = stock_name
        df_stock = core_engine.load_stock_data(bio_stock)
    else:
        df_stock, repo_stock_label, _ = core_engine.load_repo_stock_data()
        if df_stock is None:
            raise FileNotFoundError("Không tìm thấy file tồn kho hợp lệ trong thư mục repository.")
        resolved_stock_label = repo_stock_label

    df_plan = None
    if not plan_cleared:
        if plan_bytes is not None:
            bio_plan = io.BytesIO(plan_bytes)
            if plan_name:
                bio_plan.name = plan_name
            df_plan = core_engine.load_plan_data(bio_plan)
        else:
            plan_file = find_data_file(['kế hoạch']) or find_data_file(['ke hoach']) or find_data_file(['plan'])
            if plan_file:
                try:
                    df_plan = core_engine.load_plan_data(plan_file)
                except Exception:
                    df_plan = None

    orders_info = core_engine.get_available_orders_info(df_bom)
    readiness_data = core_engine.analyze_orders_readiness(df_bom, df_stock)
    return df_bom, df_stock, df_plan, orders_info, readiness_data, resolved_stock_label


@st.cache_data(show_spinner=False)
def get_cached_readiness_by_wh(cache_key, _df_bom, _df_stock, wh_code):
    return core_engine.analyze_orders_readiness(_df_bom, _df_stock, wh_spec=wh_code)


@st.cache_data(show_spinner=False)
def get_cached_readiness_excel(cache_key, _readiness_data):
    return core_engine.export_readiness_report_to_excel(_readiness_data)


@st.cache_data(show_spinner=False)
def get_cached_all_excel(queue_key, _calculation_results):
    return core_engine.export_all_orders_to_excel(_calculation_results)


@st.cache_data(show_spinner=False)
def get_cached_single_order_excel(order_key, _order_result):
    return core_engine.export_order_to_excel(_order_result)


@st.cache_data(show_spinner=False)
def get_cached_integrity_and_smart(cache_key, _df_bom, _df_stock, _df_plan, _calculation_results, _readiness_data):
    integrity = core_engine.check_data_integrity(_df_bom, _df_stock, _df_plan)
    smart = core_engine.generate_smart_reconciliation_insights(
        _df_bom, _df_stock, _df_plan, _calculation_results, _readiness_data, integrity_report=integrity
    )
    return integrity, smart


def optimize_fifo_queue_order(orders_queue, df_bom, df_stock):
    """
    Sắp xếp tối ưu hàng chờ FIFO sao cho đạt số lượng đơn ĐỦ 100% VẬT TƯ cao nhất:
    - Ở mỗi bước chọn đơn tiếp theo trong hàng chờ:
      + Ưu tiên đơn có thể đáp ứng ĐỦ 100% VẬT TƯ theo đúng số lượng đặt (order_qty) trên tồn kho còn lại.
      + Nếu có nhiều đơn cùng đủ 100%, giữ nguyên thứ tự ưu tiên ban đầu (hoặc ưu tiên đơn tiêu tốn ít mã chung hơn).
      + Nếu các đơn còn lại đều thiếu, ưu tiên đơn có số bộ chạy tối đa (max_runnable_fifo) cao nhất và tỷ lệ đủ mã cao nhất.
    """
    if not orders_queue or len(orders_queue) <= 1:
        return list(orders_queue)

    remaining = list(orders_queue)
    ordered = []

    while remaining:
        best_idx = 0
        best_key = None
        for i, cand in enumerate(remaining):
            trial_queue = ordered + [cand]
            trial_res = core_engine.calculate_inventory_allocation(trial_queue, df_bom, df_stock)[-1]
            is_ok = 1 if trial_res.get('is_fully_ok', False) else 0
            max_run = int(trial_res.get('max_runnable_fifo', 0))
            has_run = 1 if max_run > 0 else 0
            total_it = max(1, int(trial_res.get('total_items', 1)))
            ok_ratio = float(trial_res.get('ok_count', 0)) / total_it
            # Ưu tiên: (1) Đủ 100% theo SL đặt, (2) Giữ nguyên thứ tự gốc nếu cùng đủ 100%, (3) Có thể chạy >= 1 bộ, (4) SL bộ chạy tối đa, (5) Tỷ lệ mã đủ
            if is_ok:
                cand_key = (1, -i, 1, max_run, ok_ratio)
            else:
                cand_key = (0, has_run, max_run, ok_ratio, -i)
            if best_key is None or cand_key > best_key:
                best_key = cand_key
                best_idx = i
        ordered.append(remaining.pop(best_idx))

    return ordered


# ==================== SIDEBAR: NẠP FILE & TRẠNG THÁI DỮ LIỆU ====================
with st.sidebar:
    st.markdown("### 📂 Quản Lý Nguồn Dữ Liệu")
    st.caption("Hỗ trợ tự động nhận diện file tồn kho trong repo hoặc tải lên file Excel/CSV mới:")

    # 1. NẠP FILE TỒN KHO
    st.markdown("**1️⃣ Cập nhật File Tồn Kho (`Stock 61.xlsx` / `Stock 01.xlsx`):**")
    uploaded_stock = st.file_uploader(
        "Chọn file Tồn kho mới (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="uploader_stock"
    )
    if uploaded_stock is not None:
        file_bytes = uploaded_stock.getvalue()
        if st.session_state.get('custom_stock_name') != uploaded_stock.name or st.session_state.get('custom_stock_bytes') != file_bytes:
            try:
                bio_test = io.BytesIO(file_bytes)
                bio_test.name = uploaded_stock.name
                df_test_stock = core_engine.load_stock_data(bio_test)
                st.session_state.custom_stock_bytes = file_bytes
                st.session_state.custom_stock_name = uploaded_stock.name
                with open(STOCK_PERSIST_FILE, "wb") as f:
                    f.write(file_bytes)
                save_meta({'stock_name': uploaded_stock.name})
                st.cache_data.clear()
                st.toast(f"✅ Đã nạp thành công file Tồn kho: {uploaded_stock.name} ({len(df_test_stock):,} dòng)", icon="✅")
                st.rerun()
            except Exception as e:
                st.error(f"❌ File Tồn kho không hợp lệ: {e}")

    if st.button("🔄 Quét lại file Tồn kho trong Repo", key="btn_rescan_repo_stock", **_STRETCH_KW):
        st.session_state.custom_stock_bytes = None
        st.session_state.custom_stock_name = None
        if os.path.exists(STOCK_PERSIST_FILE):
            try:
                os.remove(STOCK_PERSIST_FILE)
            except Exception:
                pass
        save_meta({'stock_name': None})
        st.cache_data.clear()
        st.toast("✅ Đã quét và đồng bộ lại dữ liệu tồn kho từ thư mục Repo!", icon="🔄")
        st.rerun()

    st.markdown("---")

    # 2. NẠP FILE BOM
    st.markdown("**2️⃣ Cập nhật File Định Mức (`Production BOM List.xlsx`):**")
    uploaded_bom = st.file_uploader(
        "Chọn file BOM mới (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="uploader_bom"
    )
    if uploaded_bom is not None:
        file_bytes = uploaded_bom.getvalue()
        if st.session_state.get('custom_bom_name') != uploaded_bom.name or st.session_state.get('custom_bom_bytes') != file_bytes:
            try:
                bio_test = io.BytesIO(file_bytes)
                bio_test.name = uploaded_bom.name
                df_test_bom = core_engine.load_bom_data(bio_test)
                st.session_state.custom_bom_bytes = file_bytes
                st.session_state.custom_bom_name = uploaded_bom.name
                with open(BOM_PERSIST_FILE, "wb") as f:
                    f.write(file_bytes)
                save_meta({'bom_name': uploaded_bom.name})
                st.cache_data.clear()
                st.toast(f"✅ Đã nạp thành công file BOM: {uploaded_bom.name} ({len(df_test_bom):,} dòng)", icon="✅")
                st.rerun()
            except Exception as e:
                st.error(f"❌ File BOM không hợp lệ: {e}")

    st.markdown("---")

    # 3. NẠP FILE KẾ HOẠCH
    st.markdown("**3️⃣ Cập nhật File Kế Hoạch (`Kế hoạch.xlsx`):**")
    uploaded_plan = st.file_uploader(
        "Chọn file Kế hoạch (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="uploader_plan"
    )
    if uploaded_plan is not None:
        file_bytes = uploaded_plan.getvalue()
        if st.session_state.get('custom_plan_name') != uploaded_plan.name or st.session_state.get('custom_plan_bytes') != file_bytes:
            try:
                bio_test = io.BytesIO(file_bytes)
                bio_test.name = uploaded_plan.name
                df_test_plan = core_engine.load_plan_data(bio_test)
                st.session_state.custom_plan_bytes = file_bytes
                st.session_state.custom_plan_name = uploaded_plan.name
                st.session_state.plan_cleared = False
                st.session_state.queue_manually_cleared = False
                with open(PLAN_PERSIST_FILE, "wb") as f:
                    f.write(file_bytes)
                save_meta({'plan_name': uploaded_plan.name, 'plan_cleared': False})
                new_plan_orders = core_engine.extract_orders_from_plan(df_test_plan)
                save_orders_queue(new_plan_orders)
                st.cache_data.clear()
                st.toast(f"✅ Đã nạp Kế hoạch: {uploaded_plan.name} ({len(new_plan_orders)} đơn hàng)", icon="✅")
                st.rerun()
            except Exception as e:
                st.error(f"❌ File Kế hoạch không hợp lệ: {e}")

    col_pl1, col_pl2 = safe_columns(2)
    with col_pl1:
        if not st.session_state.get('plan_cleared', False):
            if st.button("🗑️ Xóa Kế hoạch", key="btn_clear_plan_sidebar", **_STRETCH_KW):
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = True
                if os.path.exists(PLAN_PERSIST_FILE):
                    try:
                        os.remove(PLAN_PERSIST_FILE)
                    except Exception:
                        pass
                save_meta({'plan_name': None, 'plan_cleared': True})
                cleaned_q = [
                    item for item in st.session_state.orders_queue
                    if item.get('source') != 'Kế hoạch' and not str(item.get('id', '')).startswith('plan_')
                ]
                save_orders_queue(cleaned_q)
                st.cache_data.clear()
                st.toast("✅ Đã xóa dữ liệu file Kế hoạch!", icon="🗑️")
                st.rerun()
        else:
            if st.button("🔄 Nạp lại KH", key="btn_restore_plan_sidebar", **_STRETCH_KW):
                st.session_state.plan_cleared = False
                st.session_state.queue_manually_cleared = False
                save_meta({'plan_cleared': False})
                default_plan_f = find_data_file(['kế hoạch']) or find_data_file(['ke hoach']) or find_data_file(['plan'])
                if default_plan_f:
                    try:
                        df_p_def = core_engine.load_plan_data(default_plan_f)
                        save_orders_queue(core_engine.extract_orders_from_plan(df_p_def))
                    except Exception:
                        pass
                st.cache_data.clear()
                st.toast("✅ Đã khôi phục file Kế hoạch mặc định!", icon="✅")
                st.rerun()

    with col_pl2:
        if st.session_state.custom_bom_bytes is not None or st.session_state.custom_stock_bytes is not None or st.session_state.custom_plan_bytes is not None:
            if st.button("♻️ Mặc định gốc", key="btn_reset_all_default", **_STRETCH_KW):
                st.session_state.custom_bom_bytes = None
                st.session_state.custom_bom_name = None
                st.session_state.custom_stock_bytes = None
                st.session_state.custom_stock_name = None
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = False
                st.session_state.queue_manually_cleared = False
                for p in [BOM_PERSIST_FILE, STOCK_PERSIST_FILE, PLAN_PERSIST_FILE, META_PERSIST_FILE]:
                    if os.path.exists(p):
                        try:
                            os.remove(p)
                        except Exception:
                            pass
                st.cache_data.clear()
                st.rerun()


# ==================== TẢI DỮ LIỆU & ĐỐI SOÁT ====================
try:
    active_bom_name = st.session_state.custom_bom_name or os.path.basename(find_data_file(['bom']) or 'Production BOM List.xlsx')
    active_stock_name = st.session_state.custom_stock_name
    default_plan_path = find_data_file(['kế hoạch']) or find_data_file(['ke hoach']) or find_data_file(['plan'])
    active_plan_name = None if st.session_state.get('plan_cleared', False) else (
        st.session_state.custom_plan_name or (os.path.basename(default_plan_path) if default_plan_path else None)
    )
    repo_sig = get_repo_stock_signature()

    df_bom, df_stock, df_plan, orders_info, readiness_data, resolved_stock_name = get_loaded_data(
        st.session_state.custom_bom_bytes,
        st.session_state.custom_stock_bytes,
        st.session_state.custom_plan_bytes,
        bom_name=active_bom_name,
        stock_name=active_stock_name,
        plan_name=active_plan_name,
        plan_cleared=st.session_state.get('plan_cleared', False),
        repo_stock_sig=repo_sig
    )
    active_stock_name = resolved_stock_name or active_stock_name or 'Stock 61.xlsx'
except Exception as e:
    st.error(f"❌ Lỗi đọc dữ liệu: {e}")
    st.stop()

# Tự động nạp danh sách đơn hàng từ Kế hoạch vào hàng chờ FIFO lần đầu nếu hàng chờ đang trống
if len(st.session_state.orders_queue) == 0 and not st.session_state.get('queue_manually_cleared', False):
    if df_plan is not None and len(df_plan) > 0 and not st.session_state.get('plan_cleared', False):
        initial_plan_orders = core_engine.extract_orders_from_plan(df_plan)
        if initial_plan_orders:
            save_orders_queue(initial_plan_orders)

# Hiển thị thông tin tóm tắt dữ liệu đang hoạt động ở Sidebar
with st.sidebar:
    st.markdown("---")
    st.markdown("### 📊 Trạng Thái Dữ Liệu")
    wh_breakdown = {}
    if 'Warehouse' in df_stock.columns:
        for wh_c, grp_s in df_stock.groupby('Warehouse'):
            wh_lbl = wh_c if wh_c else "Chưa gắn mã"
            wh_breakdown[wh_lbl] = (len(grp_s['Stock_Code'].unique()), int(grp_s['Qty'].sum()))
    wh_lines_html = "".join([
        f"<div style='margin-left:8px; font-size:12px; color:#047857;'>• Kho <b>{k}</b>: {v[0]:,} mã ({v[1]:,} cái)</div>"
        for k, v in sorted(wh_breakdown.items())
    ])
    st.markdown(f"""
    <div style="background:#ECFDF5; border:1px solid #6EE7B7; border-radius:12px; padding:12px 14px; font-size:12.5px; color:#065F46; line-height:1.6;">
        ✅ <b>File Tồn kho:</b> <code>{active_stock_name}</code><br>
        {wh_lines_html}
        ✅ <b>File BOM:</b> <code>{active_bom_name}</code><br>
        <span style="font-size:12px; color:#047857;">• Tổng đơn BOM: <b>{len(orders_info):,}</b> đơn ({len(df_bom):,} dòng)</span><br>
        ✅ <b>File Kế hoạch:</b> <code>{active_plan_name if df_plan is not None else 'Đang tắt / Không dùng'}</code>
    </div>
    """, unsafe_allow_html=True)


# ==================== BANNER CHÍNH ====================
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">🏭 HỆ THỐNG ĐỐI SOÁT TỒN KHO & ĐỊNH MỨC BOM SẢN XUẤT (FIFO)</div>
    <div class="hero-subtitle">
        • <b>Quy tắc phân tích Product Code</b>: Mã bắt đầu bằng <b>1</b> (4 số đầu: Đơn hàng | <b>19</b>: Trái, <b>20</b>: Phải) & Mã bắt đầu bằng <b>8</b> (bỏ số 8, 4 số tiếp: Đơn hàng | <b>1</b>: Trái, <b>2</b>: Phải).<br>
        • <b>Quy tắc Kho (WH)</b>: Chọn kho nào đối soát chính xác kho đó. Mặc định khi không chọn kho chỉ lấy dữ liệu ở <b>2 kho 01 & 61</b> (tuyệt đối không lấy kho khác).<br>
        • <b>Cơ chế trừ kho chuẩn xác</b>: Ưu tiên đơn hàng từ trên xuống (FIFO), tự động tính <b>Số bộ đồng bộ tối đa chạy được ngay</b> tại mỗi bước và xuất báo cáo Excel tô màu <b>XANH (Đủ)</b> / <b>ĐỎ (Thiếu)</b>.
    </div>
</div>
""", unsafe_allow_html=True)


# ==================== TÍNH TOÁN FIFO HÀNG CHỜ HIỆN TẠI ====================
calculation_results = []
if len(st.session_state.orders_queue) > 0:
    calculation_results = core_engine.calculate_inventory_allocation(
        st.session_state.orders_queue, df_bom, df_stock
    )

actual_ok_orders = sum(1 for r in calculation_results if r.get('is_fully_ok', False))
runnable_now_orders = sum(1 for r in calculation_results if int(r.get('max_runnable_fifo', 0)) > 0)
queue_len = len(st.session_state.orders_queue)
df_stock_default_wh = core_engine.filter_stock_by_warehouses(df_stock, None)

# ==================== 5 THẺ KPI TỔNG QUAN ====================
col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
with col_m1:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-blue">
        <div class="kpi-label">📁 Tổng Đơn Trong BOM</div>
        <div class="kpi-value">{len(orders_info):,}<span class="kpi-unit">đơn</span></div>
        <div style="font-size:11.5px; color:#64748B; margin-top:2px;">Đủ 100% mã độc lập: <b>{readiness_data['ready_count']}</b> đơn</div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    if queue_len > 0:
        st.markdown(f"""
        <div class="kpi-card kpi-card-border-emerald">
            <div class="kpi-label">✅ Đơn Đủ 100% (Theo SL Đặt)</div>
            <div class="kpi-value" style="color:#059669;">{actual_ok_orders}<span class="kpi-unit" style="color:#059669;">/ {queue_len} đơn</span></div>
            <div style="font-size:11.5px; font-weight:700; color:#059669; margin-top:2px;">✨ Đáp ứng 100% SL yêu cầu</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="kpi-card kpi-card-border-emerald">
            <div class="kpi-label">🎯 Đơn Đủ 100% Mã (Kho 01&61)</div>
            <div class="kpi-value" style="color:#059669;">{readiness_data['ready_count']}<span class="kpi-unit" style="color:#059669;">/ {readiness_data['total_orders']}</span></div>
            <div style="font-size:11.5px; font-weight:700; color:#059669; margin-top:2px;">✨ {readiness_data['ready_percent']}% đơn sẵn sàng chạy</div>
        </div>
        """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-amber">
        <div class="kpi-label">⚡ Đơn Có Thể Chạy Ngay (FIFO)</div>
        <div class="kpi-value" style="color:#D97706;">{runnable_now_orders}<span class="kpi-unit" style="color:#B45309;">/ {queue_len} đơn</span></div>
        <div style="font-size:11.5px; font-weight:700; color:#B45309; margin-top:2px;">🔧 Đủ 100% mã cho >= 1 bộ đồng bộ</div>
    </div>
    """, unsafe_allow_html=True)
with col_m4:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-green">
        <div class="kpi-label">🧩 Mã Linh Kiện Kho (01&61)</div>
        <div class="kpi-value">{len(df_stock_default_wh['Stock_Code'].unique()):,}<span class="kpi-unit">mã</span></div>
        <div style="font-size:11.5px; color:#64748B; margin-top:2px;">Tổng dòng tồn: {len(df_stock_default_wh):,} dòng</div>
    </div>
    """, unsafe_allow_html=True)
with col_m5:
    st.markdown(f"""
    <div class="kpi-card kpi-card-border-purple">
        <div class="kpi-label">📦 Tổng SL Tồn Kho (01&61)</div>
        <div class="kpi-value">{int(df_stock_default_wh['Qty'].sum()):,}<span class="kpi-unit">cái</span></div>
        <div style="font-size:11.5px; color:#64748B; margin-top:2px;">Nguồn: {active_stock_name}</div>
    </div>
    """, unsafe_allow_html=True)


# ==================== 3 TABS NGHIỆP VỤ CHÍNH ====================
tabs = st.tabs([
    "📊 Đối Soát & Hàng Chờ FIFO",
    "🚀 Đơn Sẵn Sàng Chạy (Đủ 100% Mã)",
    "🧠 Đối Soát Thông Minh & Kiểm Tra Dữ Liệu (Agent)"
])

available_wh_codes = core_engine.get_available_warehouses(df_stock, df_plan)
wh_select_options = [""] + available_wh_codes
for _wh_ss_key in ("input_wh_selector", "tab2_wh_selector"):
    _ss_val = st.session_state.get(_wh_ss_key, "")
    if _ss_val and _ss_val not in wh_select_options:
        wh_select_options.append(_ss_val)


def format_wh_option(w):
    if not w or str(w).strip() in ("", "01&61"):
        return "🏬 Mặc định (Kho 01 & 61)"
    return f"🏬 Kho {core_engine.format_warehouse_code(w)}"


# ==================== TAB 1: ĐỐI SOÁT & HÀNG CHỜ FIFO ====================
with tabs[0]:
    st.markdown("""
    <div class="floating-box">
        <div class="box-header">
            <span>🎯 NHẬP ĐƠN HÀNG, SỐ LƯỢNG CẦN CHẠY & CHỌN KHO ĐỐI SOÁT (WH)</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    curr_wh_input_raw = st.session_state.get("input_wh_selector", "")
    active_input_wh_code = core_engine.format_warehouse_code(curr_wh_input_raw)
    tab1_readiness_key = (
        active_bom_name or "default_bom",
        active_stock_name or "default_stock",
        repo_sig,
        len(df_bom),
        len(df_stock),
        active_input_wh_code
    )
    tab1_readiness = (
        readiness_data if active_input_wh_code == "01&61"
        else get_cached_readiness_by_wh(tab1_readiness_key, df_bom, df_stock, active_input_wh_code)
    )

    sorted_order_keys = sorted(orders_info.keys())
    ready_orders_dict = {ro['order_no']: ro for ro in tab1_readiness['ready_orders']}

    # Gợi ý chọn nhanh 1 chạm
    st.markdown(f"**⚡ Chọn nhanh đơn hàng mẫu tại Kho `{active_input_wh_code}` (1 Chạm để điền mã & SL tối đa chạy được):**")
    quick_sample_orders = ['5151', '5171', '8392', '4341', '8333', '1009']
    chip_cols = safe_columns(len(quick_sample_orders))
    for c_i, ord_code in enumerate(quick_sample_orders):
        if ord_code in orders_info:
            ro_m = ready_orders_dict.get(ord_code)
            chip_label = f"🟢 Đơn {ord_code} ({ro_m['max_runnable_qty']:,} bộ)" if ro_m else f"⚠️ Đơn {ord_code}"
            if chip_cols[c_i].button(chip_label, key=f"quick_{ord_code}", **_STRETCH_KW):
                st.session_state.quick_selected_order = ord_code
                if ro_m:
                    st.session_state.quick_selected_qty = str(int(ro_m['max_runnable_qty']))
                st.rerun()

    def format_order_item(k):
        ro = ready_orders_dict.get(k)
        if ro:
            return f"🟢 Đơn {k} - {orders_info[k]['description']} (✅ ĐỦ 100% MÃ tại Kho {active_input_wh_code} - Tối đa: {ro['max_runnable_qty']:,} bộ)"
        return f"⚪ Đơn {k} - {orders_info[k]['description']} (⚠️ Thiếu linh kiện tại Kho {active_input_wh_code})"

    col_input1, col_input2, col_input3, col_input4 = safe_columns([4.2, 2.0, 2.6, 2.0], vertical_alignment="bottom")
    with col_input1:
        curr_target = st.session_state.get('quick_selected_order', '5151')
        default_idx = sorted_order_keys.index(curr_target) if curr_target in sorted_order_keys else 0
        selected_code = st.selectbox(
            "1. Chọn hoặc gõ tìm mã đơn hàng (Ví dụ: 5151, 5171, 8392...):",
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
        _wh_sb_kwargs = {"accept_new_options": True} if _SELECTBOX_SUPPORTS_NEW else {}
        selected_wh_raw = st.selectbox(
            "3. Chọn hoặc điền Kho (WH):",
            options=wh_select_options,
            format_func=format_wh_option,
            index=0,
            key="input_wh_selector",
            help="Điền/chọn kho nào sẽ lấy đúng tồn kho đó ra đối soát. Nếu không chọn kho nào thì mặc định lấy 2 kho 01 & 61.",
            **_wh_sb_kwargs
        )
    with col_input4:
        btn_add = st.button("➕ BẤM OK ĐỂ TÍNH", type="primary", **_STRETCH_KW)

    selected_wh_code = core_engine.format_warehouse_code(selected_wh_raw)
    selected_wh_display = f"Kho {selected_wh_code}" if (selected_wh_raw and str(selected_wh_raw).strip() not in ("", "01&61")) else "Kho 01 & 61 (Mặc định)"

    if selected_code:
        curr_info = orders_info.get(selected_code, {})
        ready_match = ready_orders_dict.get(selected_code)
        if ready_match:
            ready_badge_str = f"<span class='badge-ok'>✅ ĐỦ 100% MÃ ĐỒNG BỘ (Tối đa chạy được {ready_match['max_runnable_qty']:,} bộ)</span>"
            bottleneck_str = f" | <span style='color:#D97706; font-weight:700;'>⚠️ Điểm nghẽn: <b title='{ready_match['bottleneck_desc']}'>{ready_match['bottleneck_code']}</b> (Tồn: {int(ready_match['bottleneck_stock']):,} cái)</span>"
        else:
            ready_badge_str = f"<span class='badge-missing'>⚠️ Đang thiếu linh kiện tại {selected_wh_display}</span>"
            bottleneck_str = ""
        st.markdown(
            f"<div style='background:#F1F5F9; border-radius:10px; padding:10px 14px; font-size:13.5px; color:#334155; margin-top:8px; border-left:4px solid #2563EB;'>"
            f"ℹ️ <b>Đơn hàng {selected_code}</b>: {curr_info.get('description', '')} | <b>Nguyên đơn</b>: Trái + Phải | "
            f"<b>Kho đối soát</b>: <span class='badge-wh'>🏬 {selected_wh_display}</span> | "
            f"Độ sẵn sàng độc lập: {ready_badge_str}{bottleneck_str}"
            f"</div>",
            unsafe_allow_html=True
        )

    if btn_add:
        try:
            val = float(str(qty_input_str).replace(',', '').strip())
            if val > 0 and selected_code:
                new_queue = list(st.session_state.orders_queue)
                new_queue.append({
                    'id': f"{selected_code}_{selected_wh_code}_{len(new_queue)+1}_{uuid.uuid4().hex[:8]}",
                    'order_no': selected_code,
                    'wh': selected_wh_code,
                    'description': orders_info[selected_code]['description'],
                    'qty': val,
                    'qty_left': val,
                    'qty_right': val,
                    'created_at': datetime.datetime.now().strftime("%H:%M:%S")
                })
                st.session_state.queue_manually_cleared = False
                save_orders_queue(new_queue)
                st.toast(f"✅ Đã thêm đơn {selected_code} ({val:,g} bộ | Kho {selected_wh_code}) vào hàng chờ!", icon="✅")
                st.rerun()
            else:
                st.error("⚠️ Số lượng cần chạy phải lớn hơn 0.")
        except ValueError:
            st.error("⚠️ Vui lòng nhập số lượng hợp lệ.")

    # ==================== KHUNG QUẢN LÝ HÀNG CHỜ FIFO & TỐI ƯU ĐỐI SOÁT ====================
    st.markdown("---")
    col_q_title, col_q_opt1, col_q_opt2, col_q_clear = safe_columns([4.4, 2.5, 2.1, 1.3], vertical_alignment="center")
    with col_q_title:
        st.markdown(f"### ⏳ Hàng Chờ Đối Soát FIFO ({len(st.session_state.orders_queue)} đơn)")
    with col_q_opt1:
        if len(st.session_state.orders_queue) > 1:
            if st.button("⚡ Ưu tiên đơn ĐỦ 100% lên đầu", key="btn_sort_fifo_optimal", help="Tự động đưa các đơn hàng đủ 100% vật tư lên đầu hàng chờ FIFO để không bị các đơn thiếu phía trên chiếm mất linh kiện chung", **_STRETCH_KW):
                opt_queue = core_engine.optimize_fifo_queue(st.session_state.orders_queue, df_bom, df_stock)
                save_orders_queue(opt_queue)
                st.toast("✅ Đã sắp xếp tối ưu hàng chờ: Ưu tiên các đơn ĐỦ 100% VẬT TƯ lên đầu!", icon="⚡")
                st.rerun()
    with col_q_opt2:
        has_adjustable = any(
            (not r.get('is_fully_ok', False)) and int(r.get('max_runnable_fifo', 0)) > 0
            for r in calculation_results
        )
        if has_adjustable:
            if st.button("📐 Chốt SL theo mức Tối đa", key="btn_cap_runnable_qty", help="Tự động điều chỉnh số lượng các đơn đang thiếu về đúng số bộ đồng bộ tối đa có thể chạy ngay từ tồn kho hiện có", **_STRETCH_KW):
                updated_q, adjusted_cnt = core_engine.cap_queue_to_runnable_fifo(st.session_state.orders_queue, df_bom, df_stock)
                save_orders_queue(updated_q)
                st.toast(f"✅ Đã chốt số lượng {adjusted_cnt} đơn theo đúng mức bộ đồng bộ tối đa chạy được!", icon="📐")
                st.rerun()
    with col_q_clear:
        if len(st.session_state.orders_queue) > 0:
            if st.button("🗑️ Xóa hàng chờ", type="secondary", **_STRETCH_KW):
                st.session_state.queue_manually_cleared = True
                save_orders_queue([])
                st.rerun()

    # Nạp từ file kế hoạch nếu có
    if df_plan is not None and len(df_plan) > 0:
        plan_orders_extracted = core_engine.extract_orders_from_plan(df_plan)
        plan_wh_counts = {}
        for po in plan_orders_extracted:
            w_k = po.get('wh', '01&61')
            plan_wh_counts[w_k] = plan_wh_counts.get(w_k, 0) + 1
        wh_summary_str = ", ".join([f"Kho {k}: {v} đơn" for k, v in sorted(plan_wh_counts.items())])
        col_p_btn, col_p_del, col_p_info = safe_columns([4.8, 2.2, 5.0], vertical_alignment="center")
        with col_p_btn:
            if st.button(f"📥 Nạp / Đặt lại toàn bộ {len(plan_orders_extracted)} đơn từ File Kế hoạch", type="primary", **_STRETCH_KW):
                st.session_state.queue_manually_cleared = False
                save_orders_queue(plan_orders_extracted)
                st.toast(f"✅ Đã nạp {len(plan_orders_extracted)} đơn hàng từ file Kế hoạch!", icon="✅")
                st.rerun()
        with col_p_del:
            if st.button("🗑️ Tắt Kế hoạch", key="btn_clear_plan_tab1", type="secondary", help="Tắt file kế hoạch sản xuất để chỉ đối soát theo đơn tự chọn", **_STRETCH_KW):
                st.session_state.custom_plan_bytes = None
                st.session_state.custom_plan_name = None
                st.session_state.plan_cleared = True
                st.session_state.queue_manually_cleared = True
                if os.path.exists(PLAN_PERSIST_FILE):
                    try:
                        os.remove(PLAN_PERSIST_FILE)
                    except Exception:
                        pass
                save_meta({'plan_name': None, 'plan_cleared': True})
                cleaned_q = [
                    item for item in st.session_state.orders_queue
                    if item.get('source') != 'Kế hoạch' and not str(item.get('id', '')).startswith('plan_')
                ]
                save_orders_queue(cleaned_q)
                st.cache_data.clear()
                st.rerun()
        with col_p_info:
            st.caption(f"📋 File Kế hoạch: **{len(df_plan)}** dòng Model → **{len(plan_orders_extracted)}** đơn ({wh_summary_str}).")
        with st.expander(f"📋 Xem bảng dữ liệu gốc File Kế hoạch sản xuất ({len(df_plan)} dòng Model)", expanded=False):
            display_plan_cols = [c for c in ['MODEL', 'Order_No', 'Side', 'Qty', 'WH', 'Due_Date'] if c in df_plan.columns]
            safe_dataframe(
                df_plan[display_plan_cols].rename(columns={
                    'Order_No': 'Mã đơn',
                    'Side': 'Vế',
                    'Qty': 'Số lượng (Qty)',
                    'WH': 'Kho (WH)',
                    'Due_Date': 'Ngày giao (Due date)'
                }),
                height=250
            )
    elif st.session_state.get('plan_cleared', False):
        col_p_nt1, col_p_nt2 = safe_columns([8, 2], vertical_alignment="center")
        with col_p_nt1:
            st.caption("ℹ️ Chế độ Không dùng File Kế hoạch đang bật. Bạn có thể tự chọn đơn hàng bất kỳ hoặc nạp lại Kế hoạch mặc định.")
        with col_p_nt2:
            if st.button("🔄 Nạp lại Kế hoạch", key="btn_restore_plan_tab1", **_STRETCH_KW):
                st.session_state.plan_cleared = False
                st.session_state.queue_manually_cleared = False
                save_meta({'plan_cleared': False})
                default_plan_f = find_data_file(['kế hoạch']) or find_data_file(['ke hoach']) or find_data_file(['plan'])
                if default_plan_f:
                    try:
                        df_p_def = core_engine.load_plan_data(default_plan_f)
                        save_orders_queue(core_engine.extract_orders_from_plan(df_p_def))
                    except Exception:
                        pass
                st.cache_data.clear()
                st.rerun()

    # Hiển thị gợi ý thông minh nếu có đơn ở dưới bị đơn thiếu phía trên chiếm linh kiện hoặc đặt vượt mức tồn kho
    if len(calculation_results) > 0:
        partial_runnable = [
            r for r in calculation_results
            if not r.get('is_fully_ok', False) and int(r.get('max_runnable_fifo', 0)) > 0
        ]
        if partial_runnable:
            sample_txt = ", ".join([f"<b>Đơn {r['order_no']}</b> (chạy ngay được <b>{r['max_runnable_fifo']:,}</b>/{int(r['order_qty']):,} bộ)" for r in partial_runnable[:5]])
            st.markdown(
                f"<div style='background:#FFFBEB; border:1px solid #FCD34D; border-left:5px solid #F59E0B; border-radius:10px; padding:10px 14px; font-size:13px; color:#92400E; margin:8px 0;'>"
                f"💡 <b>Phát hiện đối soát thông minh:</b> Có <b>{len(partial_runnable)} đơn hàng</b> chưa đủ 100% cho toàn bộ số lượng đặt cao, nhưng <b>đủ 100% mã để chạy ngay một phần</b>: {sample_txt}... "
                f"Bạn có thể bấm <b>⚡ Ưu tiên đơn ĐỦ 100% lên đầu</b> hoặc <b>📐 Chốt SL theo mức Tối đa</b> ở góc phải phía trên để tối ưu ngay!"
                f"</div>",
                unsafe_allow_html=True
            )

    if len(st.session_state.orders_queue) == 0:
        st.info("💡 Hàng chờ đang trống. Hãy chọn mã đơn hàng ở trên rồi bấm **➕ BẤM OK ĐỂ TÍNH** hoặc bấm **📥 Nạp từ File Kế hoạch**.")
    else:
        # Bảng hiển thị hàng chờ FIFO
        cols_header = safe_columns([0.7, 1.2, 2.5, 1.3, 1.5, 2.2, 2.0, 1.1, 0.7])
        cols_header[0].markdown("**TT**")
        cols_header[1].markdown("**Mã đơn**")
        cols_header[2].markdown("**Tên sản phẩm**")
        cols_header[3].markdown("**Kho (WH)**")
        cols_header[4].markdown("**SL Đặt**")
        cols_header[5].markdown("**Kết quả theo SL Đặt**")
        cols_header[6].markdown("**SL Chạy Tối Đa (FIFO)**")
        cols_header[7].markdown("**Lịch / Giờ**")
        cols_header[8].markdown("**Xóa**")

        to_delete_idx = None
        for idx, item in enumerate(st.session_state.orders_queue):
            c = safe_columns([0.7, 1.2, 2.5, 1.3, 1.5, 2.2, 2.0, 1.1, 0.7], vertical_alignment="center")
            c[0].markdown(f"**#{idx + 1}**")
            c[1].markdown(f"<span class='badge-order'>{item['order_no']}</span>", unsafe_allow_html=True)
            c[2].write(item.get('description') or orders_info.get(item['order_no'], {}).get('description', ''))

            item_wh_code = core_engine.format_warehouse_code(item.get('wh'))
            c[3].markdown(f"<span class='badge-wh'>🏬 Kho {item_wh_code}</span>", unsafe_allow_html=True)

            raw_qty = item.get('qty', 0)
            try:
                qty_num = float(raw_qty)
            except Exception:
                qty_num = 0.0
            qty_l = float(item.get('qty_left', qty_num))
            qty_r = float(item.get('qty_right', qty_num))
            if qty_l == qty_r:
                c[4].markdown(f"**{qty_num:,g}** bộ")
            else:
                c[4].markdown(f"T: **{qty_l:,g}** | P: **{qty_r:,g}**")

            if idx < len(calculation_results):
                res_item = calculation_results[idx]
                max_run_fifo = int(res_item.get('max_runnable_fifo', 0))
                if res_item.get('is_fully_ok', False):
                    c[5].markdown("<span class='badge-ok'>✅ ĐỦ 100% VẬT TƯ</span>", unsafe_allow_html=True)
                else:
                    m_count = res_item.get('missing_count', 0)
                    c[5].markdown(f"<span class='badge-missing'>⚠️ THIẾU {m_count}/{res_item.get('total_items', 0)} MÃ</span>", unsafe_allow_html=True)

                if max_run_fifo >= qty_num and qty_num > 0:
                    c[6].markdown(f"<span class='badge-ok'>🟢 Tối đa: {max_run_fifo:,} bộ</span>", unsafe_allow_html=True)
                elif max_run_fifo > 0:
                    c[6].markdown(f"<span class='badge-runnable'>⚡ Chạy ngay: {max_run_fifo:,} bộ</span>", unsafe_allow_html=True)
                else:
                    c[6].markdown("<span style='color:#94A3B8; font-size:12.5px; font-weight:600;'>🔴 0 bộ (Hết mã nghẽn)</span>", unsafe_allow_html=True)
            else:
                c[5].markdown("<span>⏳ Đang tính...</span>", unsafe_allow_html=True)
                c[6].write("-")

            c[7].write(item.get('due_date') or item.get('created_at', ''))
            if c[8].button("❌", key=f"del_{item.get('id', idx)}_{idx}"):
                to_delete_idx = idx

        if to_delete_idx is not None:
            new_q = list(st.session_state.orders_queue)
            new_q.pop(to_delete_idx)
            if len(new_q) == 0:
                st.session_state.queue_manually_cleared = True
            save_orders_queue(new_q)
            st.rerun()

    # ==================== BÁO CÁO CHI TIẾT ĐỐI SOÁT TỪNG ĐƠN & XUẤT EXCEL ====================
    if len(st.session_state.orders_queue) > 0 and calculation_results:
        st.markdown("---")
        st.markdown(f"### 📊 Báo Cáo Chi Tiết Đối Soát Từng Đơn ({len(calculation_results)} đơn)")

        queue_cache_key = tuple(
            (item.get('id', ''), item['order_no'], item.get('wh', '01&61'), item['order_qty'], item.get('ok_count', 0), item.get('missing_count', 0), item.get('max_runnable_fifo', 0))
            for item in calculation_results
        )
        all_excel_bytes = get_cached_all_excel(queue_cache_key, calculation_results)
        st.download_button(
            label=f"📥 Tải Trọn Bộ Excel Đối Soát ({len(calculation_results)} đơn hàng - Kèm cột SL Bộ Đáp Ứng & Chi Tiết Trái/Phải)",
            data=all_excel_bytes,
            file_name=f"Doi_Soat_Tong_Hop_{len(calculation_results)}_Don_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            **_STRETCH_KW
        )

        st.markdown("<br>", unsafe_allow_html=True)
        for q_idx, r in enumerate(calculation_results):
            order_no = r['order_no']
            order_qty = r['order_qty']
            order_wh = r.get('wh', '01&61')
            is_fully_ok = r.get('is_fully_ok', False)
            max_run_fifo = int(r.get('max_runnable_fifo', 0))
            stt = r['stt']

            if is_fully_ok:
                status_badge = "<span class='badge-ok'>✅ ĐỦ VẬT TƯ (100% OK)</span>"
                card_class = "order-card-ok"
            elif max_run_fifo > 0:
                status_badge = f"<span class='badge-missing'>⚠️ THIẾU {r['missing_count']} MÃ (CHO {order_qty:g} BỘ)</span> <span class='badge-runnable'>⚡ CHẠY NGAY ĐƯỢC: {max_run_fifo:,} BỘ</span>"
                card_class = "order-card-partial"
            else:
                status_badge = f"<span class='badge-missing'>⚠️ THIẾU {r['missing_count']} ITEM CODE (0 BỘ)</span>"
                card_class = "order-card-missing"

            qty_l = float(r.get('qty_left', order_qty))
            qty_r = float(r.get('qty_right', order_qty))
            qty_str = f"<b>{order_qty:,g} bộ</b> (Trái + Phải)" if qty_l == qty_r else f"Trái: <b>{qty_l:,g}</b> | Phải: <b>{qty_r:,g}</b>"
            bn_code = r.get('fifo_bottleneck_code', '')
            bn_stock = float(r.get('fifo_bottleneck_stock', 0.0))
            bn_info_html = f" | Linh kiện giới hạn: <b>{bn_code}</b> (Tồn khả dụng: {bn_stock:g})" if bn_code else ""

            st.markdown(f"""
            <div class="order-card {card_class}">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                    <div>
                        <span style="font-size: 17px; font-weight: 850; color: #1E3A8A;">Ưu tiên #{stt} - Đơn hàng {order_no}</span>
                        <span style="font-size: 14.5px; color: #475569; margin-left: 6px;">({r.get('product_title', '')} | Variant: {r.get('variant', '')})</span>
                        <div style="margin-top: 4px; font-size: 13.5px; color: #64748B;">
                            Kho đối soát: <span class="badge-wh">🏬 Kho {order_wh}</span> | Số lượng đặt: {qty_str} | Trạng thái: {status_badge}{bn_info_html}
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_ck1, col_ck2 = safe_columns([6.2, 3.8], vertical_alignment="center")
            with col_ck1:
                k1, k2, k3, k4 = safe_columns(4)
                k1.metric("Tổng mã LK", f"{r['total_items']} mã")
                k2.metric("Mã ĐỦ (theo SL đặt)", f"{r['ok_count']} mã")
                k3.metric("Mã THIẾU", f"{r['missing_count']} mã", delta=f"-{r['missing_count']}" if r['missing_count'] > 0 else "0", delta_color="inverse")
                k4.metric("SL Bộ Chạy Ngay", f"{max_run_fifo:,} bộ")
            with col_ck2:
                if (not is_fully_ok) and (0 < max_run_fifo < order_qty):
                    if st.button(f"📐 Chốt đơn #{order_no} về {max_run_fifo:,} bộ (Đủ 100% VT)", key=f"cap_single_{r.get('id', stt)}_{stt}", **_STRETCH_KW):
                        new_q = list(st.session_state.orders_queue)
                        if q_idx < len(new_q):
                            new_q[q_idx] = dict(new_q[q_idx])
                            new_q[q_idx]['qty'] = float(max_run_fifo)
                            new_q[q_idx]['qty_left'] = float(max_run_fifo)
                            new_q[q_idx]['qty_right'] = float(max_run_fifo)
                            save_orders_queue(new_q)
                            st.toast(f"✅ Đã chốt đơn {order_no} về {max_run_fifo:,} bộ (Đủ 100% vật tư)!", icon="✅")
                            st.rerun()
                single_cache_key = f"{r.get('id', '')}_{order_no}_{order_wh}_{order_qty}_{r.get('ok_count', 0)}_{r.get('missing_count', 0)}_{max_run_fifo}"
                single_excel = get_cached_single_order_excel(single_cache_key, r)
                st.download_button(
                    label=f"📥 Tải Excel Đơn {order_no} (Kho {order_wh})",
                    data=single_excel,
                    file_name=f"Doi_Soat_Don_{order_no}_Kho_{order_wh.replace('&', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_{r.get('id', stt)}_{stt}",
                    **_STRETCH_KW
                )

            if len(r.get('summary_table', pd.DataFrame())) > 0:
                with st.expander(f"🔍 Xem bảng đối soát vật tư Đơn {order_no} tại Kho {order_wh} ({r['ok_count']} ĐỦ / {r['missing_count']} THIẾU | Chạy tối đa: {max_run_fifo:,} bộ)", expanded=False):
                    sub_t1, sub_t2 = st.tabs(["📋 Tổng Hợp Linh Kiện Nguyên Đơn", "↔️ Chi Tiết Định Mức Trái / Phải"])
                    with sub_t1:
                        safe_dataframe(r['summary_table'], height=280)
                    with sub_t2:
                        safe_dataframe(r.get('detail_table', pd.DataFrame()), height=280)
            st.markdown("<hr style='margin: 12px 0; border: none; border-top: 1px dashed #CBD5E1;'>", unsafe_allow_html=True)


# ==================== TAB 2: ĐƠN SẴN SÀNG CHẠY (ĐỦ 100% MÃ) ====================
with tabs[1]:
    col_t2_wh, col_t2_info = safe_columns([4, 6], vertical_alignment="bottom")
    with col_t2_wh:
        tab2_wh_raw = st.selectbox(
            "🏬 Chọn Kho đối soát đơn sẵn sàng (WH):",
            options=wh_select_options,
            format_func=format_wh_option,
            index=0,
            key="tab2_wh_selector",
            help="Mặc định lấy dữ liệu ở 2 kho 01 & 61 ra đối soát (không lấy kho khác). Khi chọn kho cụ thể sẽ đối soát riêng kho đó."
        )
    tab2_wh_code = core_engine.format_warehouse_code(tab2_wh_raw)
    tab2_readiness_key = (
        active_bom_name or "default_bom",
        active_stock_name or "default_stock",
        repo_sig,
        len(df_bom),
        len(df_stock),
        tab2_wh_code
    )
    tab2_readiness_data = (
        readiness_data if tab2_wh_code == "01&61"
        else get_cached_readiness_by_wh(tab2_readiness_key, df_bom, df_stock, tab2_wh_code)
    )
    with col_t2_info:
        st.caption(f"Đang đối soát tồn kho tại: **Kho {tab2_wh_code}** (Tuyệt đối không lấy dữ liệu ở kho khác).")

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #064E3B 0%, #059669 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
        <div style="font-size: 20px; font-weight: 850;">
            🚀 TỔNG HỢP {tab2_readiness_data['ready_count']} / {tab2_readiness_data['total_orders']} ĐƠN HÀNG ĐỦ 100% MÃ ĐỒNG BỘ TẠI KHO {tab2_wh_code}
        </div>
        <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
            Hệ thống tự động tính toán số lượng tối đa có thể sản xuất ngay tại Kho {tab2_wh_code} và chỉ rõ linh kiện nghẽn giới hạn số bộ.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_dl_r1, col_dl_r2, col_dl_r3 = safe_columns([4.5, 3.0, 2.5], vertical_alignment="center")
    with col_dl_r1:
        readiness_key = (tab2_wh_code, tab2_readiness_data['ready_count'], tab2_readiness_data['not_ready_count'], tab2_readiness_data['total_orders'])
        readiness_excel = get_cached_readiness_excel(readiness_key, tab2_readiness_data)
        st.download_button(
            label=f"📥 Tải Báo Cáo Excel {tab2_readiness_data['ready_count']} Đơn Đủ Mã Tại Kho {tab2_wh_code} (Kèm Đơn Thiếu)",
            data=readiness_excel,
            file_name=f"Bao_Cao_Don_Du_Ma_Kho_{tab2_wh_code.replace('&', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            **_STRETCH_KW
        )
    with col_dl_r2:
        if len(tab2_readiness_data['ready_orders']) > 0:
            ready_opts = [ro['order_no'] for ro in tab2_readiness_data['ready_orders']]
            ready_map_t2 = {ro['order_no']: ro for ro in tab2_readiness_data['ready_orders']}
            sel_ready_to_add = st.selectbox(
                "Chọn nhanh đơn đủ mã để đưa vào hàng chờ FIFO:",
                options=ready_opts,
                format_func=lambda x: f"Đơn {x} - Tối đa {ready_map_t2[x]['max_runnable_qty']:,} bộ",
                key="sel_ready_tab2"
            )
        else:
            sel_ready_to_add = None
    with col_dl_r3:
        if sel_ready_to_add:
            ro_picked = ready_map_t2[sel_ready_to_add]
            if st.button(f"➕ Đưa Đơn {sel_ready_to_add} ({ro_picked['max_runnable_qty']:,} bộ) vào FIFO", key="btn_add_ready_tab2", **_STRETCH_KW):
                new_q = list(st.session_state.orders_queue)
                new_q.append({
                    'id': f"{sel_ready_to_add}_{tab2_wh_code}_{len(new_q)+1}_{uuid.uuid4().hex[:8]}",
                    'order_no': sel_ready_to_add,
                    'wh': tab2_wh_code,
                    'description': ro_picked['description'],
                    'qty': float(ro_picked['max_runnable_qty']),
                    'qty_left': float(ro_picked['max_runnable_qty']),
                    'qty_right': float(ro_picked['max_runnable_qty']),
                    'created_at': datetime.datetime.now().strftime("%H:%M:%S")
                })
                st.session_state.queue_manually_cleared = False
                save_orders_queue(new_q)
                st.toast(f"✅ Đã thêm đơn {sel_ready_to_add} ({ro_picked['max_runnable_qty']:,} bộ) vào hàng chờ FIFO!", icon="✅")
                st.rerun()

    # Tìm kiếm & phân trang bảng đơn đủ mã
    col_s1, col_s2, col_s3 = safe_columns([5, 3, 2], vertical_alignment="bottom")
    with col_s1:
        search_ready = st.text_input("🔍 Tìm kiếm đơn hàng:", placeholder="Gõ mã đơn hoặc tên sản phẩm...")
    with col_s2:
        sort_mode = st.selectbox("Sắp xếp:", ["🔥 SL chạy tối đa (Giảm dần)", "🔢 SL chạy tối đa (Tăng dần)", "🏷️ Mã đơn (A-Z)"])
    with col_s3:
        page_size = st.selectbox("Số dòng / trang:", [15, 25, 50, 100], index=0)

    filtered_ready = list(tab2_readiness_data['ready_orders'])
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

    ready_html = """
    <div style="overflow-x: auto;">
    <table class="styled-table">
        <thead>
            <tr>
                <th style="width: 50px; text-align:center;">STT</th>
                <th style="width: 110px; text-align:center;">Mã Đơn</th>
                <th style="width: 120px; text-align:center;">Kho (WH)</th>
                <th>Tên Sản Phẩm</th>
                <th style="width: 110px; text-align:center;">Tổng Mã LK</th>
                <th style="text-align:center; width: 200px; background:#059669; color:#FFFFFF;">SL TỐI ĐA CHẠY ĐƯỢC</th>
                <th style="text-align:center; width: 190px;">Linh Kiện Giới Hạn</th>
            </tr>
        </thead>
        <tbody>
    """
    for idx_row, r_item in enumerate(page_items, start_idx + 1):
        hover_info = f"Điểm nghẽn: {r_item['bottleneck_code']} ({r_item['bottleneck_desc']}) - Kho còn: {int(r_item['bottleneck_stock']):,} cái"
        ready_html += f"""
            <tr class="row-ok">
                <td style="text-align:center;"><b>#{idx_row}</b></td>
                <td style="text-align:center;"><span class="badge-order">{r_item['order_no']}</span></td>
                <td style="text-align:center;"><span class="badge-wh">🏬 Kho {r_item.get('wh', tab2_wh_code)}</span></td>
                <td><b>{r_item['description']}</b></td>
                <td style="text-align:center;"><b>{r_item['total_components']}</b> mã</td>
                <td style="text-align:center; font-size:16px; font-weight:900; color:#059669; background:#ECFDF5;" title="{hover_info}">
                    {r_item['max_runnable_qty']:,} bộ
                </td>
                <td style="text-align:center;">
                    <span class="badge-ok">✅ ĐỦ 100% MÃ</span>
                    <div style="font-size:11.5px; color:#D97706; margin-top:3px; font-weight:700;" title="{r_item['bottleneck_desc']}">⚠️ Nghẽn: <b>{r_item['bottleneck_code']}</b> ({int(r_item['bottleneck_stock']):,})</div>
                </td>
            </tr>
        """
    ready_html += "</tbody></table></div>"
    st.markdown(ready_html, unsafe_allow_html=True)

    col_p1, col_p2, col_p3 = safe_columns([3, 4, 3], vertical_alignment="center")
    with col_p1:
        st.caption(f"Hiển thị {len(page_items)} / {len(filtered_ready)} đơn (Trang {st.session_state.ready_page_idx} / {total_pages})")
    with col_p2:
        c_prev, c_txt, c_next = safe_columns([1, 2, 1], vertical_alignment="center")
        if c_prev.button("◀ Trang trước", disabled=st.session_state.ready_page_idx <= 1, **_STRETCH_KW):
            st.session_state.ready_page_idx -= 1
            st.rerun()
        c_txt.markdown(f"<div style='text-align:center; font-weight:700;'>Trang {st.session_state.ready_page_idx}</div>", unsafe_allow_html=True)
        if c_next.button("Trang sau ▶", disabled=st.session_state.ready_page_idx >= total_pages, **_STRETCH_KW):
            st.session_state.ready_page_idx += 1
            st.rerun()

    if len(tab2_readiness_data.get('not_ready_orders', [])) > 0:
        with st.expander(f"📋 Xem danh sách {len(tab2_readiness_data['not_ready_orders'])} đơn hàng đang thiếu linh kiện tại Kho {tab2_wh_code}", expanded=False):
            df_not_ready_view = pd.DataFrame(tab2_readiness_data['not_ready_orders']).rename(columns={
                'order_no': 'Mã đơn',
                'wh': 'Kho (WH)',
                'variant': 'Variant',
                'description': 'Tên sản phẩm',
                'total_components': 'Tổng mã LK',
                'components_in_stock': 'Số mã có tồn',
                'missing_count': 'Số mã thiếu',
                'readiness_pct': 'Tỷ lệ đủ (%)',
                'sample_missing': 'Các mã thiếu tiêu biểu'
            })
            cols_nr = [c for c in ['Mã đơn', 'Kho (WH)', 'Tên sản phẩm', 'Tổng mã LK', 'Số mã có tồn', 'Số mã thiếu', 'Tỷ lệ đủ (%)', 'Các mã thiếu tiêu biểu'] if c in df_not_ready_view.columns]
            safe_dataframe(df_not_ready_view[cols_nr], height=320)


# ==================== TAB 3: ĐỐI SOÁT THÔNG MINH & KIỂM TRA DỮ LIỆU (AGENT) ====================
with tabs[2]:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); border-radius: 14px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 16px;">
        <div style="font-size: 20px; font-weight: 850;">
            🧠 ĐỐI SOÁT THÔNG MINH, LỌC DỮ LIỆU CHO AGENT & KIỂM TRA TOÀN VẸN
        </div>
        <div style="font-size: 14px; opacity: 0.95; margin-top: 4px;">
            Tự động lọc dữ liệu quan trọng, phát hiện linh kiện nút cổ chai, đơn cận sẵn sàng (chỉ thiếu 1-3 mã), gợi ý mã tương đương trong kho và xuất gói dữ liệu chuẩn cho các AI Agent.
        </div>
    </div>
    """, unsafe_allow_html=True)

    smart_cache_key = (
        active_bom_name or "default_bom",
        active_stock_name or "default_stock",
        active_plan_name or "default_plan",
        repo_sig,
        st.session_state.get('plan_cleared', False),
        len(df_bom),
        len(df_stock),
        len(df_plan) if df_plan is not None else 0,
        queue_len,
        actual_ok_orders
    )
    integrity_report, smart_insights = get_cached_integrity_and_smart(
        smart_cache_key, df_bom, df_stock, df_plan, calculation_results, readiness_data
    )

    if integrity_report and smart_insights:
        c_h1, c_h2, c_h3, c_h4, c_h5 = safe_columns(5)
        c_h1.metric("Điểm Sức Khỏe Dữ Liệu", f"{integrity_report['overall_health']}%")
        c_h2.metric("Mã LK Khớp Trong Kho", f"{integrity_report['matched_comps']} / {integrity_report['total_bom_comps']} mã")
        c_h3.metric("Đơn Cận Sẵn Sàng (Thiếu 1-3 mã)", f"{readiness_data.get('near_ready_count', 0)} đơn")
        c_h4.metric("Mã Có Linh Kiện Tương Đương", f"{integrity_report.get('substitutes_count', 0)} mã")
        c_h5.metric("Sai Lệch Kế Hoạch Sản Xuất", f"{integrity_report['plan_issues_count']} đơn")

        st.markdown("---")
        sub_tabs = st.tabs([
            "🤖 Gói Dữ Liệu Quan Trọng Cho Agent",
            "⚡ Đơn Cận Sẵn Sàng & Mã Tương Đương",
            "🔍 Rà Soát Chi Tiết Toàn Vẹn Dữ Liệu"
        ])

        with sub_tabs[0]:
            st.markdown("#### 🤖 Báo Cáo Lọc Dữ Liệu Trọng Tâm & Xuất Ngữ Cảnh Cho AI Agent")
            st.caption("Tự động chắt lọc những thông tin đối soát quan trọng nhất từ các nguồn dữ liệu (BOM, Tồn kho, Kế hoạch) theo định dạng chuẩn hóa JSON Schema & Markdown.")
            col_ag1, col_ag2 = safe_columns(2)
            with col_ag1:
                agent_json_str = json.dumps(smart_insights['agent_payload'], ensure_ascii=False, indent=2)
                st.download_button(
                    label="📥 Tải Gói Dữ Liệu Chuẩn Hóa Cho Agent (.JSON)",
                    data=agent_json_str.encode('utf-8'),
                    file_name=f"Agent_Reconciliation_Context_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    type="primary",
                    **_STRETCH_KW
                )
            with col_ag2:
                st.download_button(
                    label="📥 Tải Bản Tóm Tắt Đối Soát Thông Minh (.MD)",
                    data=smart_insights['agent_markdown_brief'].encode('utf-8'),
                    file_name=f"Smart_Reconciliation_Brief_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                    mime="text/markdown",
                    **_STRETCH_KW
                )

            st.markdown(
                f"<div style='background:#F8FAFC; border:1.5px solid #CBD5E1; border-radius:12px; padding:16px 20px; margin-top:10px;'>"
                f"<b>💡 Tóm tắt nhanh cho Điều độ & Agent:</b> Hiện có <b>{readiness_data['ready_count']}/{readiness_data['total_orders']} đơn ({readiness_data['ready_percent']}%)</b> đủ 100% vật tư chạy ngay; "
                f"<b>{readiness_data.get('near_ready_count', 0)} đơn cận sẵn sàng</b> (chỉ thiếu 1-3 mã linh kiện) và <b>{integrity_report.get('substitutes_count', 0)} mã thiếu</b> có mã tương đương cùng họ đang sẵn tồn kho lớn."
                f"</div>",
                unsafe_allow_html=True
            )

            if len(smart_insights['df_plan_smart']) > 0:
                st.markdown("##### 📅 Phân Tích Thông Minh File Kế Hoạch Sản Xuất (FIFO Lũy Kế vs Độc Lập):")
                safe_dataframe(smart_insights['df_plan_smart'], height=320)

            with st.expander("📄 Xem trực tiếp nội dung Báo cáo Agent Brief & JSON Schema", expanded=False):
                st.markdown(smart_insights['agent_markdown_brief'])

        with sub_tabs[1]:
            st.markdown(f"#### ⚡ 1. Danh sách {len(smart_insights['df_near_ready'])} Đơn Hàng Cận Sẵn Sàng (Chỉ thiếu từ 1 đến 3 mã linh kiện):")
            st.caption("Ưu tiên xử lý hoặc thúc nhập kho các mã thiếu dưới đây sẽ mở khóa ngay các đơn hàng có độ đồng bộ đạt 95% - 99%:")
            if len(smart_insights['df_near_ready']) > 0:
                safe_dataframe(smart_insights['df_near_ready'], height=300)
            else:
                st.info("Hiện không có đơn hàng nào ở trạng thái cận sẵn sàng (thiếu 1-3 mã).")

            st.markdown(f"#### 🔄 2. Gợi ý {len(smart_insights['df_substitutes'])} Mã Linh Kiện Tương Đương / Cùng Họ Sẵn Có Trong Kho:")
            st.caption("Phát hiện thông minh các mã linh kiện thiếu trong BOM nhưng trong kho đang có tồn kho mã cùng họ (khác hậu tố mạ/đóng gói PA, PD, PE, TI, 01 vs 51...):")
            if len(smart_insights['df_substitutes']) > 0:
                safe_dataframe(smart_insights['df_substitutes'], height=280)
            else:
                st.success("Không phát hiện mã thiếu nào có biến thể tương đương cần đối chiếu.")

        with sub_tabs[2]:
            st.markdown(f"#### ⚠️ 1. Danh sách {integrity_report['unmatched_comps_count']} mã linh kiện trong BOM chưa có trong kho (Kèm mức độ ảnh hưởng):")
            st.caption("Sắp xếp theo số lượng đơn hàng BOM bị chặn giảm dần để bộ phận Mua hàng / Kho ưu tiên xử lý:")
            if len(integrity_report['df_unmatched_comps']) > 0:
                safe_dataframe(integrity_report['df_unmatched_comps'], height=280)
            else:
                st.success("✅ Tuyệt vời: 100% linh kiện trong BOM đều tồn tại trong danh mục tồn kho.")

            st.markdown("#### 📅 2. Kiểm tra tính đồng bộ của File Kế hoạch với BOM:")
            if df_plan is not None:
                if integrity_report['plan_issues_count'] > 0:
                    safe_dataframe(integrity_report['df_plan_issues'])
                else:
                    st.success("✅ 100% các đơn hàng trong File Kế hoạch đều có cấu trúc định mức BOM đầy đủ, sẵn sàng chạy đối soát.")
            else:
                st.info("ℹ️ Chưa nạp File Kế hoạch (hoặc dữ liệu Kế hoạch đã được xóa). Hệ thống tính tồn kho đối soát theo toàn bộ đơn hàng trong BOM.")

            st.markdown("#### 📦 3. Kiểm tra các dòng tồn kho âm hoặc bất thường:")
            if integrity_report['zero_stock_count'] > 0:
                safe_dataframe(integrity_report['df_zero_stock'])
            else:
                st.success("✅ Dữ liệu số lượng tồn kho hợp lệ, không có mã nào bị âm.")


if __name__ == "__main__":
    from streamlit.runtime import exists
    if not exists():
        import sys
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", __file__]
        sys.exit(stcli.main())
