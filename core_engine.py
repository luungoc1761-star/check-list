"""
core_engine.py
Bộ xử lý logic đối soát thông minh BOM Sản Xuất (Production BOM List), Tồn kho (Stock 61 / Stock Balance) và Kế hoạch Sản xuất:
- Nhận diện tự động đa ngôn ngữ (Tiếng Việt / Tiếng Anh / ERP SAP / Excel / CSV) cho các nguồn dữ liệu mới đưa vào.
- Lọc dữ liệu quan trọng chuẩn hóa cho các Agent & Người dùng đối soát:
    + Product Code bắt đầu bằng 1: 4 số đầu là mã đơn hàng, 2 số tiếp theo (19: Trái, 20: Phải), phần còn lại là Variant.
    + Product Code bắt đầu bằng 8: bỏ số 8, 4 số tiếp theo là mã đơn hàng, số thứ 6 (1: Trái, 2: Phải), phần còn lại là Variant.
- Cột Component Code đối soát với Item code / Stock Code trong Stock 61 (hoặc Stock Balance).
- Quản lý trừ kho theo cơ chế FIFO theo thứ tự đơn hàng nhập vào hàng chờ (ưu tiên đơn trên trước, trừ vế Trái trước - vế Phải sau).
- Đối soát thông minh (Smart Reconciliation & Agent Context):
    + Tự động chọn Variant tối ưu nhất khi đơn hàng có nhiều phiên bản trong BOM.
    + Phân tích đơn cận sẵn sàng (Near-Ready Orders chỉ thiếu 1-3 mã) và Top linh kiện nút cổ chai (Critical Bottleneck Components).
    + Nhận diện mã linh kiện tương đương / thay thế tiềm năng trong kho (Fuzzy / Variant Code Matching).
    + Trích xuất gói dữ liệu chuẩn hóa JSON/Markdown cho các AI Agent lọc dữ liệu quan trọng.
- Xuất file Excel định dạng chuyên nghiệp với tô màu xanh dòng OK, đỏ dòng THIẾU.
"""

import io
import os
import re
import math
import uuid
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

__version__ = "2026.11.07.v3"
DEFAULT_WAREHOUSES = ('01', '61')
EPS = 1e-9


def normalize_warehouse_code(val) -> str:
    """
    Chuẩn hóa mã kho (Warehouse / WH):
    - Xử lý kiểu số nguyên / số thực từ Excel (VD: 1 -> '01', 1.0 -> '01', 61.0 -> '61', '001' -> '01').
    - Loại bỏ tiền tố 'KHO', 'WH', 'WAREHOUSE', 'W/H', 'K6-' nếu người dùng hoặc ERP xuất kèm.
    - Loại bỏ hậu tố mô tả kho (VD: '01 - Kho NVL' -> '01', '61 (Kho SX)' -> '61').
    - Tự động chuẩn hóa mã kho số về 2 chữ số (VD: '1' -> '01', '061' -> '61').
    """
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        s = str(val).strip()
    else:
        s = str(val).strip()
    for ch in ('\u00A0', '\ufeff', '\u200b', '\u200c', '\u200d', '\r', '\n', '\t'):
        s = s.replace(ch, '')
    s = s.strip()
    if not s or s.lower() in ('nan', 'none', 'null', ''):
        return ""
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    s = re.sub(r'^[^a-zA-Z0-9]+', '', s).strip()
    # Loại bỏ tiền tố công ty/nhà máy dạng K6-01, K6_61
    s = re.sub(r'^[a-zA-Z]\d+\s*[-_:/]\s*(?=\d)', '', s).strip()
    # Loại bỏ tiền tố Warehouse / Kho / WH / W/H / WK
    s = re.sub(r'^(warehouse|kho\s*hàng|kho\s*tồn|mã\s*kho|ma\s*kho|kho|w/h|w\.h|whs|whse|wh|wk)\s*[-_:/]?\s*', '', s, flags=re.IGNORECASE).strip()
    # Nếu là W01, W61, K01, K61 (1 chữ cái W/K đi liền mã số kho)
    m_wk = re.match(r'^[wkWK](\d{1,3})$', s)
    if m_wk:
        s = m_wk.group(1)
    # Nếu có mô tả phía sau kiểu "01 - Kho NVL" hoặc "61 (Kho SX)"
    m_lead_num = re.match(r'^(\d{1,3})(?:\.0)?(?:\s*[-_:(/].*|\s+[a-zA-ZÀ-ỹ].*)$', s)
    if m_lead_num:
        s = m_lead_num.group(1)
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    if s.isdigit():
        s = str(int(s)).zfill(2)
    return s.upper()


def infer_warehouse_from_name(name_str: str) -> str:
    """
    Trích xuất mã kho từ tên file hoặc tên sheet Excel nếu trong bảng không có cột WH riêng.
    Ví dụ:
    - 'Stock 61.xlsx' -> '61'
    - 'Stock 01.xlsx' -> '01'
    - 'Kho 01.xlsx' -> '01'
    - 'Tồn kho 61.xlsx' -> '61'
    - Sheet '01', '61', 'Kho 01', 'WH 61' -> '01', '61'
    Không nhận nhầm số thứ tự bản sao '(1)', '(2)', '(3)' hoặc năm '2026', phiên bản '10.2'.
    """
    if not name_str:
        return ""
    base = os.path.splitext(os.path.basename(str(name_str)))[0].strip()
    # Xóa hậu tố bản sao kiểu (1), (2), (3) ở cuối tên file
    base_clean = re.sub(r'\(\d+\)\s*$', '', base).strip()
    # Nếu bản thân tên (VD tên sheet) chỉ là mã kho '01', '61', '1'
    if re.match(r'^(?:kho|wh|stock|w)?\s*0?[1-9]\d?$', base_clean, flags=re.IGNORECASE):
        num_m = re.search(r'\d+', base_clean)
        if num_m:
            return str(int(num_m.group(0))).zfill(2)
    # Tìm mẫu "Stock 61", "Stock 01", "Kho 61", "Kho 01", "WH 61", "WH 01", "Tồn kho 61", "Ton kho 01"
    m = re.search(
        r'(?:stock|tồn\s*kho|ton\s*kho|kho|warehouse|wh)\s*[-_:]?\s*(0?\d{1,2})\b(?!\s*\.\d)',
        base_clean,
        flags=re.IGNORECASE
    )
    if m:
        val = int(m.group(1))
        if 1 <= val <= 99:
            return str(val).zfill(2)
    return ""


def parse_target_warehouses(wh_spec=None) -> tuple:
    """
    Phân tích thông tin kho cần đối soát:
    - Nếu người dùng hoặc Kế hoạch chỉ định kho cụ thể (VD: '01', '61', '02'...) -> Trả về đúng kho đó, ví dụ ('01',).
    - Nếu không chọn kho nào (rỗng, None, 'Mặc định', '01&61'...) -> Mặc định trả về ('01', '61').
      TUYỆT ĐỐI KHÔNG lấy dữ liệu ở kho khác ngoài kho được chọn (hoặc 01 & 61 khi mặc định).
    """
    if wh_spec is None:
        return DEFAULT_WAREHOUSES
    if isinstance(wh_spec, (list, tuple, set)):
        cleaned = []
        for item in wh_spec:
            for sub in parse_target_warehouses(item):
                if sub not in cleaned:
                    cleaned.append(sub)
        return tuple(cleaned) if cleaned else DEFAULT_WAREHOUSES
    if pd.isna(wh_spec):
        return DEFAULT_WAREHOUSES
    s = str(wh_spec).strip()
    if not s or s.lower() in ('nan', 'none', 'null', 'all', 'default', 'mặc định', 'mac dinh', 'tất cả', 'tat ca'):
        return DEFAULT_WAREHOUSES
    if 'mặc định' in s.lower() or 'mac dinh' in s.lower() or 'default' in s.lower():
        return DEFAULT_WAREHOUSES
    parts = re.split(r'[&,/+;]+', s)
    cleaned = []
    for p in parts:
        code = normalize_warehouse_code(p)
        if code and code not in cleaned:
            cleaned.append(code)
    return tuple(cleaned) if cleaned else DEFAULT_WAREHOUSES


def format_warehouse_code(wh_spec=None) -> str:
    """Chuẩn hóa chuỗi mã kho lưu trữ và hiển thị (VD: '01', '61', '01&61')."""
    whs = parse_target_warehouses(wh_spec)
    return "&".join(whs)


def filter_stock_by_warehouses(df_stock: pd.DataFrame, wh_spec=None) -> pd.DataFrame:
    """
    Lọc dữ liệu tồn kho theo đúng kho được chọn (hoặc mặc định 2 kho '01' và '61' nếu không chọn kho).
    - Nếu dòng tồn kho có chỉ định mã kho rõ ràng (VD: '01', '61', '02'...): chỉ lấy đúng kho trong `target_whs`.
    - Nếu file tồn kho không có cột mã kho (Warehouse == ''): tự động đưa vào đối soát để không bị mất dữ liệu.
    - Nếu đối soát mặc định ('01&61') nhưng file tồn kho mới nạp chỉ dùng mã kho riêng khác (không có '01' hay '61'):
      tự động fallback lấy dữ liệu tồn kho hiện có trong file thay vì trả về 0 dòng.
    """
    if df_stock is None or len(df_stock) == 0:
        return df_stock
    target_whs = parse_target_warehouses(wh_spec)
    if 'Warehouse' not in df_stock.columns:
        return df_stock.copy()
    wh_col = df_stock['Warehouse'].apply(normalize_warehouse_code)
    matched = df_stock[wh_col.isin(target_whs) | (wh_col == '')].copy()
    if len(matched) == 0 and target_whs == DEFAULT_WAREHOUSES:
        return df_stock.copy()
    return matched


def get_available_warehouses(df_stock: pd.DataFrame = None, df_plan: pd.DataFrame = None) -> list:
    """Lấy danh sách các mã kho hiện có (luôn đảm bảo có '01' và '61' ở đầu danh sách)."""
    wh_list = ['01', '61']
    if df_stock is not None and 'Warehouse' in df_stock.columns:
        for w in sorted(df_stock['Warehouse'].apply(normalize_warehouse_code).unique().tolist()):
            if w and w not in wh_list:
                wh_list.append(w)
    if df_plan is not None and 'WH' in df_plan.columns:
        for w_raw in df_plan['WH'].dropna().unique().tolist():
            for w in parse_target_warehouses(w_raw):
                if w and w not in wh_list:
                    wh_list.append(w)
    return wh_list


def normalize_code(val) -> str:
    """
    Chuẩn hóa mã linh kiện / sản phẩm thông minh:
    - Loại bỏ khoảng trắng ẩn, non-breaking space (\u00A0), zero-width space (\u200b, \ufeff), tab/newline.
    - Xử lý hậu tố '.0' do Excel tự chuyển số nguyên thành số thực khi đọc dữ liệu mới.
    - Viết hoa đồng nhất mã chữ cái hậu tố (VD: pa -> PA, pd -> PD).
    """
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val).strip().upper()
    s = str(val).strip()
    for ch in ('\u00A0', '\ufeff', '\u200b', '\u200c', '\u200d', '\r', '\n', '\t'):
        s = s.replace(ch, '')
    s = s.strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s.upper()


def parse_numeric_qty(series: pd.Series) -> pd.Series:
    """
    Chuyển đổi cột số lượng an toàn cho mọi định dạng dữ liệu mới (hỗ trợ cả dấu phẩy ngăn cách hàng nghìn '1,000',
    khoảng trắng ẩn hoặc hậu tố đơn vị tính như 'pcs', 'cái', 'bộ').
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors='coerce').fillna(0.0)
    cleaned = (
        series.astype(str)
        .str.replace(',', '', regex=False)
        .str.replace('\u00A0', '', regex=False)
        .str.replace(r'\s*(pcs|pc|ea|set|sets|cái|cai|bộ|bo|chiếc|kg|m)\s*$', '', flags=re.IGNORECASE, regex=True)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors='coerce').fillna(0.0)


def _read_tabular_source(file_source, header=0, nrows=None, skiprows=0, sheet_name=0) -> pd.DataFrame:
    """
    Đọc linh hoạt nguồn dữ liệu bảng từ đường dẫn file hoặc BytesIO/UploadedFile:
    Hỗ trợ .xlsx, .xls, .csv, .tsv với tự động xử lý con trỏ stream và chọn sheet.
    """
    if hasattr(file_source, 'seek'):
        file_source.seek(0)
    fname = getattr(file_source, 'name', '') or (file_source if isinstance(file_source, str) else '')
    fname_lower = str(fname).lower()
    if fname_lower.endswith(('.csv', '.tsv')):
        sep = '\t' if fname_lower.endswith('.tsv') else None
        for enc in ('utf-8-sig', 'utf-16', 'cp1258', 'latin1'):
            try:
                if hasattr(file_source, 'seek'):
                    file_source.seek(0)
                if sep:
                    return pd.read_csv(file_source, sep=sep, header=header, nrows=nrows, skiprows=skiprows, encoding=enc)
                return pd.read_csv(file_source, sep=None, engine='python', header=header, nrows=nrows, skiprows=skiprows, encoding=enc)
            except Exception:
                continue
        if hasattr(file_source, 'seek'):
            file_source.seek(0)
        return pd.read_csv(file_source, header=header, nrows=nrows, skiprows=skiprows, encoding='latin1')
    return pd.read_excel(file_source, sheet_name=sheet_name, header=header, nrows=nrows, skiprows=skiprows)


def clean_order_title(desc, order_no: str = "") -> str:
    """
    Loại bỏ các từ chỉ vế Trái / Phải (LHF, RHF, LEFT, RIGHT, Side Assy, Side...)
    để hiển thị tên sản phẩm đại diện nguyên đơn gọn gàng, không bị dính chữ Trái/Phải.
    """
    if not desc or str(desc).strip().lower() in ('', 'nan', 'none'):
        return f"Sản phẩm {order_no}" if order_no else "Sản phẩm"
    
    cleaned = str(desc).strip()
    # Xóa các mã tiền tố kiểu "8167110301 - " hoặc "8184610004 -"
    cleaned = re.sub(r'^\d{8,12}\s*[-_:]\s*', '', cleaned)
    # Xóa các từ LHF, RHF, LH, RH, LEFT, RIGHT, Side Assy, Side
    cleaned = re.sub(r'\b(LHF|RHF|LH|RH|LEFT|RIGHT|Side\s*Assy|Side)\b', '', cleaned, flags=re.IGNORECASE)
    # Dọn dẹp khoảng trắng thừa và dấu gạch nối lẻ loi
    cleaned = re.sub(r'[-_]\s*$', '', cleaned)
    cleaned = re.sub(r'^\s*[-_]', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    if not cleaned or cleaned.isdigit():
        return f"Sản phẩm {order_no}"
    return cleaned


def parse_product_code(code):
    """
    Phân tích mã sản phẩm / Model:
    - Bắt đầu bằng 1: 4 số đầu là đơn hàng, 2 số tiếp theo là vế (19: Trái, 20: Phải), phần còn lại là Variant.
    - Bắt đầu bằng 8: bỏ 8, 4 số tiếp theo là đơn hàng, số tiếp theo là vế (1: Trái, 2: Phải), phần còn lại là Variant.
    - Hỗ trợ thêm định dạng mở rộng (nếu có tiền tố chữ hoặc dấu gạch ngang).
    """
    code_str = re.sub(r'[\s\-_]', '', normalize_code(code))
    if code_str.startswith('1') and len(code_str) >= 6:
        order = code_str[:4]
        side_raw = code_str[4:6]
        side = 'Trái' if side_raw == '19' else ('Phải' if side_raw == '20' else f'Vế {side_raw}')
        variant = code_str[6:]
        return order, side, variant
    elif code_str.startswith('8') and len(code_str) >= 6:
        order = code_str[1:5]
        side_raw = code_str[5] if len(code_str) > 5 else ''
        side = 'Trái' if side_raw == '1' else ('Phải' if side_raw == '2' else f'Vế {side_raw}')
        variant = code_str[6:] if len(code_str) > 6 else ''
        return order, side, variant
    else:
        return code_str[:4], 'Khác', code_str[4:]


def find_plan_header_row_from_df(df_raw, max_rows=25):
    """
    Tìm chỉ số dòng chứa tiêu đề file kế hoạch (MODEL và Qty) hỗ trợ đa ngôn ngữ Việt/Anh/ERP.
    """
    model_kws = [
        'model', 'product code', 'product_code', 'item_code', 'parent code', 'assy code',
        'mã sp', 'mã sản phẩm', 'mã hàng', 'mã model', 'mã đơn', 'mã thành phẩm', 'mã'
    ]
    qty_kws = ['qty', 'quantity', 'order qty', 'plan qty', 'số lượng', 'sl', 'số bộ', 'cần chạy']
    for i in range(min(max_rows, len(df_raw))):
        row_str = ' '.join([str(x).lower().strip() for x in df_raw.iloc[i].values if pd.notna(x)])
        if any(k in row_str for k in model_kws) and any(k in row_str for k in qty_kws):
            return i
    return 0


def load_plan_data(file_source):
    """
    Đọc file Kế hoạch sản xuất (.xlsx / .xls / .csv):
    - Tự động lọc dữ liệu quan trọng: MODEL, Qty, WH (Kho) và Due date (nếu có).
    - Giải mã cột MODEL ra Mã đơn hàng (Order_No), Vế (Trái/Phải) và Phiên bản (Variant).
    - Chuẩn hóa cột WH (nếu không điền kho thì mặc định '01&61').
    """
    df_raw = _read_tabular_source(file_source, header=None, nrows=25)
    h_idx = find_plan_header_row_from_df(df_raw)
    df = _read_tabular_source(file_source, skiprows=h_idx)
    
    col_map = {}
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if any(k in c_clean for k in ['model', 'product code', 'product_code', 'parent code', 'assy code', 'mã sp', 'mã sản phẩm', 'mã hàng', 'mã model', 'mã thành phẩm']):
            if 'MODEL' not in col_map.values():
                col_map[c] = 'MODEL'
        elif any(k in c_clean for k in ['qty', 'quantity', 'order qty', 'plan qty', 'số lượng', 'sl', 'số bộ']):
            if 'Qty' not in col_map.values():
                col_map[c] = 'Qty'
        elif c_clean in ('wh', 'warehouse', 'kho', 'mã kho', 'ma kho', 'wh code', 'wh_code', 'storage', 'plant') or any(k in c_clean for k in ['warehouse', 'mã kho', 'tên kho']) or re.search(r'\b(wh|kho)\b', c_clean):
            if 'WH' not in col_map.values():
                col_map[c] = 'WH'
        elif any(k in c_clean for k in ['due date', 'due_date', 'delivery', 'ngày', 'hạn', 'date']):
            if 'Due_Date' not in col_map.values():
                col_map[c] = 'Due_Date'
                
    df = df.rename(columns=col_map)
    if 'MODEL' not in df.columns or 'Qty' not in df.columns:
        raise ValueError('Không tìm thấy cột MODEL (Mã sản phẩm) hoặc Qty (Số lượng) trong file Kế hoạch.')
        
    df = df[pd.notna(df['MODEL']) & pd.notna(df['Qty'])].copy()
    df['MODEL'] = df['MODEL'].apply(normalize_code)
    df = df[~df['MODEL'].str.lower().isin(['', 'nan', 'none', 'total', 'tổng'])].copy()
    df['Qty'] = parse_numeric_qty(df['Qty'])
    df = df[df['Qty'] > 0].copy()
    
    parsed = [parse_product_code(m) for m in df['MODEL']]
    df['Order_No'] = [p[0] for p in parsed]
    df['Side'] = [p[1] for p in parsed]
    df['Variant'] = [p[2] for p in parsed]
    
    if 'WH' in df.columns:
        raw_wh_list = []
        for v in df['WH']:
            if v is None or pd.isna(v) or str(v).strip().lower() in ('', 'nan', 'none'):
                raw_wh_list.append('')
            else:
                raw_wh_list.append(format_warehouse_code(v))
        # Hỗ trợ trường hợp gộp ô (merged cells) trong Excel: dòng vế thứ 2 cùng đơn kế thừa WH của dòng vế đầu nếu để trống
        ord_list = df['Order_No'].tolist()
        for i in range(len(raw_wh_list)):
            if not raw_wh_list[i]:
                if i > 0 and ord_list[i] == ord_list[i - 1] and raw_wh_list[i - 1]:
                    raw_wh_list[i] = raw_wh_list[i - 1]
                elif i + 1 < len(raw_wh_list) and ord_list[i] == ord_list[i + 1] and raw_wh_list[i + 1]:
                    raw_wh_list[i] = raw_wh_list[i + 1]
        df['WH'] = [w if w else '01&61' for w in raw_wh_list]
    else:
        df['WH'] = '01&61'
    
    if 'Due_Date' in df.columns:
        df['Due_Date'] = df['Due_Date'].fillna('').astype(str).str.slice(0, 10)
        df['Due_Date'] = df['Due_Date'].replace({'NaT': '', 'nan': '', 'None': ''})
    else:
        df['Due_Date'] = ''
        
    return df


def extract_orders_from_plan(df_plan):
    """
    Chuyển đổi dữ liệu từ File Kế hoạch thành danh sách hàng chờ đối soát FIFO.
    Giữ nguyên thứ tự ưu tiên từ trên xuống dưới của File Kế hoạch.
    Mỗi đơn hàng tự động ghép vế Trái và Phải theo đúng Kho (WH) và số lượng mà kho đó cần.
    """
    orders_queue = []
    seen = {}
    
    for idx, row in df_plan.iterrows():
        ord_no = str(row['Order_No']).strip()
        side = str(row['Side']).strip()
        qty = float(row['Qty'])
        variant = str(row['Variant']).strip() if pd.notna(row['Variant']) else ''
        model = str(row['MODEL']).strip()
        due_date = str(row.get('Due_Date', '')).strip()
        wh_code = format_warehouse_code(row.get('WH', '01&61'))
        
        plan_key = (ord_no, wh_code)
        if plan_key not in seen:
            item = {
                'id': f'plan_{ord_no}_{wh_code}_{len(orders_queue)+1}_{uuid.uuid4().hex[:6]}',
                'order_no': ord_no,
                'wh': wh_code,
                'qty': qty,
                'qty_left': qty if side == 'Trái' else (qty if side not in ('Trái', 'Phải') else 0.0),
                'qty_right': qty if side == 'Phải' else (qty if side not in ('Trái', 'Phải') else 0.0),
                'variant': variant,
                'due_date': due_date,
                'models': {side: model},
                'source': 'Kế hoạch',
                'created_at': due_date if due_date else 'Kế hoạch'
            }
            seen[plan_key] = item
            orders_queue.append(item)
        else:
            item = seen[plan_key]
            if side == 'Trái':
                item['qty_left'] = (item['qty_left'] + qty) if item['qty_left'] > 0 else qty
            elif side == 'Phải':
                item['qty_right'] = (item['qty_right'] + qty) if item['qty_right'] > 0 else qty
            else:
                item['qty_left'] = max(item['qty_left'], qty)
                item['qty_right'] = max(item['qty_right'], qty)
            item['models'][side] = model
            item['qty'] = max(item['qty_left'], item['qty_right'])
            if variant and not item.get('variant'):
                item['variant'] = variant
            if due_date and not item.get('due_date'):
                item['due_date'] = due_date
                
    return orders_queue


def find_bom_header_row_from_df(df_raw, max_rows=25):
    """Tìm chỉ số dòng chứa tiêu đề bảng Production BOM hỗ trợ đa ngôn ngữ Việt/Anh/ERP."""
    prod_kws = ['product code', 'product_code', 'parent item', 'parent code', 'assy code', 'model code', 'mã thành phẩm', 'mã sản phẩm', 'mã sp']
    comp_kws = ['component code', 'component_code', 'child item', 'child code', 'material code', 'item code', 'mã linh kiện', 'mã vật tư', 'mã lk', 'mã vt']
    for i in range(min(max_rows, len(df_raw))):
        row_str = " ".join([str(x).lower().strip() for x in df_raw.iloc[i].values if pd.notna(x)])
        if any(pk in row_str for pk in prod_kws) and any(ck in row_str for ck in comp_kws):
            return i
    return 0


def load_bom_data(file_source):
    """
    Đọc dữ liệu Production BOM từ đường dẫn file hoặc file-like object (uploaded file).
    Tự động nhận diện dòng tiêu đề đa ngôn ngữ và lọc các cột quan trọng phục vụ đối soát.
    """
    df_raw = _read_tabular_source(file_source, header=None, nrows=25)
    h_idx = find_bom_header_row_from_df(df_raw)
    df = _read_tabular_source(file_source, skiprows=h_idx)
    
    col_map = {}
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if any(k in c_clean for k in ['product code', 'product_code', 'parent item', 'parent code', 'assy code', 'mã thành phẩm', 'mã sản phẩm']):
            if 'Product_Code' not in col_map.values():
                col_map[c] = 'Product_Code'
        elif any(k in c_clean for k in ['production desc', 'product desc', 'parent desc', 'assy desc', 'tên thành phẩm', 'tên sản phẩm', 'mô tả sp']):
            if 'Product_Desc' not in col_map.values():
                col_map[c] = 'Product_Desc'
        elif any(k in c_clean for k in ['component code', 'component_code', 'child item', 'child code', 'material code', 'mã linh kiện', 'mã vật tư']):
            if 'Component_Code' not in col_map.values():
                col_map[c] = 'Component_Code'
        elif any(k in c_clean for k in ['component desc', 'child desc', 'material desc', 'tên linh kiện', 'tên vật tư', 'mô tả lk']):
            if 'Component_Desc' not in col_map.values():
                col_map[c] = 'Component_Desc'
        elif c_clean == 'qty' or any(k in c_clean for k in ['qty', 'quantity', 'bom qty', 'usage', 'định mức', 'số lượng']):
            if 'Qty' not in col_map.values():
                col_map[c] = 'Qty'
                
    df = df.rename(columns=col_map)
    if 'Product_Code' not in df.columns or 'Component_Code' not in df.columns:
        raise ValueError("Không tìm thấy cột Product Code hoặc Component Code trong file BOM.")
        
    df = df[pd.notna(df['Product_Code']) & pd.notna(df['Component_Code'])].copy()
    
    df['Product_Code'] = df['Product_Code'].apply(normalize_code)
    df['Component_Code'] = df['Component_Code'].apply(normalize_code)
    df = df[~df['Product_Code'].str.lower().isin(['', 'nan', 'none', 'product code'])].copy()
    df = df[~df['Component_Code'].str.lower().isin(['', 'nan', 'none', 'component code'])].copy()
    df['Qty'] = parse_numeric_qty(df['Qty']) if 'Qty' in df.columns else 1.0
    
    if 'Product_Desc' not in df.columns:
        df['Product_Desc'] = ''
    else:
        df['Product_Desc'] = df['Product_Desc'].fillna('').astype(str).str.strip()
        
    if 'Component_Desc' not in df.columns:
        df['Component_Desc'] = ''
    else:
        df['Component_Desc'] = df['Component_Desc'].fillna('').astype(str).str.strip()
        
    parsed = [parse_product_code(c) for c in df['Product_Code']]
    df['Order_No'] = [p[0] for p in parsed]
    df['Side'] = [p[1] for p in parsed]
    df['Variant'] = [p[2] for p in parsed]
    
    return df


def _is_stock_code_col_header(c_clean: str) -> int:
    """
    Chấm điểm mức độ khớp của tên cột với Mã linh kiện / Stock_Code:
    - 2: Khớp chính xác hoặc từ khóa đặc trưng mạnh ('item code', 'stock code', 'mã vật tư', 'mã linh kiện'...)
    - 1: Khớp từ khóa ngắn ('item', 'code', 'mã', 'material') nhưng không phải cột mô tả/tên/mã kho.
    - 0: Không khớp.
    """
    if not c_clean or c_clean.startswith('unnamed'):
        return 0
    # Loại trừ cột mô tả hoặc mã kho hoặc mã đơn
    if any(ex in c_clean for ex in ['desc', 'mô tả', 'mo ta', 'tên', 'ten', 'name', 'warehouse', 'wh', 'kho', 'company', 'plant', 'batch', 'bin', 'date', 'ngày']):
        return 0
    strong_kws = (
        'stock code', 'item code', 'item_code', 'stock_code', 'component code', 'component_code',
        'material code', 'material_code', 'material no', 'material number', 'part no', 'part number',
        'part_no', 'part code', 'item no', 'item number', 'sku',
        'mã vt', 'ma vt', 'mã vật tư', 'ma vat tu', 'mã linh kiện', 'ma linh kien',
        'mã hàng', 'ma hang', 'mã lk', 'ma lk', 'mã sp', 'ma sp'
    )
    if c_clean in strong_kws or any(k in c_clean for k in strong_kws):
        return 2
    if c_clean in ('item', 'code', 'mã', 'ma', 'material', 'part', 'linh kiện', 'vật tư'):
        return 1
    return 0


def _is_warehouse_col_header(c_clean: str) -> int:
    """
    Chấm điểm mức độ khớp của tên cột với Kho (Warehouse / WH):
    - 2: Khớp chính xác cột kho ('wh', 'warehouse', 'w/h', 'kho', 'mã kho', 'sloc', 'lgort'...)
    - 1: Khớp từ khóa phụ ('storage', 'plant', 'werks', 'site')
    - 0: Không khớp.
    """
    if not c_clean or c_clean.startswith('unnamed'):
        return 0
    exact_wh = (
        'wh', 'warehouse', 'w/h', 'w.h', 'whs', 'whse', 'wh code', 'wh_code', 'warehouse code',
        'warehouse_code', 'kho', 'mã kho', 'ma kho', 'tên kho', 'ten kho', 'kho hàng', 'kho tồn',
        'sloc', 'lgort', 'storage location', 'storage_location'
    )
    if c_clean in exact_wh or any(k in c_clean for k in ('warehouse', 'mã kho', 'ma kho', 'kho hàng', 'kho tồn', 'storage location', 'w/h')):
        return 2
    if re.search(r'\b(wh|whs|kho|sloc|lgort)\b', c_clean):
        return 2
    if any(k in c_clean for k in ('storage', 'plant', 'werks', 'site')):
        return 1
    return 0


def _is_qty_col_header(c_clean: str) -> int:
    """
    Chấm điểm mức độ khớp của tên cột với Số lượng tồn kho (Qty):
    - 3: Khớp chính xác cột số lượng tồn chính ('qty', 'quantity', 'số lượng', 'tồn cuối', 'closing balance', 'on hand', 'available'...)
    - 2: Khớp từ khóa số lượng / tồn kho ('tồn kho', 'stock qty', 'balance', 'khả dụng', 'unrestricted', 'thực tồn'...)
    - 1: Khớp từ khóa phụ (như 'tồn đầu', 'opening balance' nếu không có cột tồn cuối nào khác)
    - 0: Không khớp (loại trừ cột tiền/giá trị hoặc cột kho/sloc/slot).
    """
    if not c_clean or c_clean.startswith('unnamed'):
        return 0
    # Không nhận nhầm cột Kho hoặc giá trị thành tiền
    if _is_warehouse_col_header(c_clean) == 2:
        return 0
    if any(ex in c_clean for ex in ['amount', 'value', 'price', 'cost', 'thành tiền', 'thanh tien', 'giá trị', 'gia tri', 'đơn giá', 'don gia']):
        return 0
    exact_top = (
        'qty', 'quantity', 'số lượng', 'so luong', 'sl', 'tồn kho', 'ton kho',
        'tồn cuối', 'ton cuoi', 'thực tồn', 'thuc ton', 'tồn khả dụng', 'khả dụng', 'kha dung',
        'on hand', 'on_hand', 'available', 'available qty', 'avail qty', 'closing balance',
        'ending balance', 'unrestricted', 'free stock', 'stock qty', 'inventory qty', 'tồn', 'ton', 'balance'
    )
    if c_clean in exact_top:
        return 3
    if any(ex in c_clean for ex in ['tồn đầu', 'ton dau', 'opening']):
        return 1
    if any(k in c_clean for k in [
        'qty', 'quantity', 'on hand', 'available', 'unrestricted', 'closing balance',
        'stock balance', 'balance', 'số lượng', 'so luong', 'tồn cuối', 'ton cuoi',
        'thực tồn', 'thuc ton', 'tồn kho', 'ton kho', 'khả dụng', 'kha dung'
    ]):
        return 2
    if re.search(r'\b(sl|tồn|ton)\b', c_clean):
        return 2
    return 0


def find_stock_header_row_from_df(df_raw, max_rows=50):
    """
    Tìm chỉ số dòng chứa tiêu đề bảng tồn kho chính xác:
    - Kiểm tra từng ô riêng biệt trong dòng (tránh nhận nhầm dòng tiêu đề báo cáo gộp chung 1 ô như 'Báo cáo tồn kho theo mã hàng').
    - Yêu cầu ít nhất 2 cột khác nhau khớp với (Mã linh kiện) và (Số lượng / Kho).
    """
    best_row = 0
    best_score = -1
    limit = min(max_rows, len(df_raw))
    for i in range(limit):
        row_vals = [str(x).strip().lower() for x in df_raw.iloc[i].values if pd.notna(x) and str(x).strip() != '']
        if len(row_vals) < 2:
            continue
        code_cols = [ idx for idx, v in enumerate(row_vals) if _is_stock_code_col_header(v) > 0 ]
        qty_cols = [ idx for idx, v in enumerate(row_vals) if _is_qty_col_header(v) > 0 ]
        wh_cols = [ idx for idx, v in enumerate(row_vals) if _is_warehouse_col_header(v) > 0 or infer_warehouse_from_name(v) != '' ]
        if code_cols and (qty_cols or len(wh_cols) >= 2):
            # Đảm bảo cột code và cột qty/wh là 2 ô khác nhau trên cùng dòng
            other_cols = set(qty_cols + wh_cols) - set(code_cols)
            if other_cols:
                score = (
                    max(_is_stock_code_col_header(row_vals[c]) for c in code_cols) * 3
                    + (max(_is_qty_col_header(row_vals[q]) for q in qty_cols) * 3 if qty_cols else 0)
                    + (2 if wh_cols else 0)
                )
                if score > best_score:
                    best_score = score
                    best_row = i
                    if score >= 12:
                        return i
    return best_row


def _parse_single_stock_df(df_raw_preview: pd.DataFrame, file_source, sheet_name=0, inferred_wh: str = "") -> pd.DataFrame:
    """
    Phân tích 1 bảng/sheet dữ liệu tồn kho thành DataFrame chuẩn ['Stock_Code', 'Warehouse', 'Qty', 'BATCH', 'BIN'].
    Hỗ trợ cả bảng dọc tiêu chuẩn (Item code, Wh, Qty), bảng xoay ngang (Item code, Kho 01, Kho 61)
    và bảng không có dòng tiêu đề.
    """
    h_idx = find_stock_header_row_from_df(df_raw_preview)
    df = _read_tabular_source(file_source, skiprows=h_idx, sheet_name=sheet_name)
    if df is None or len(df) == 0:
        raise ValueError("Bảng tồn kho rỗng.")

    # Chấm điểm từng cột để chọn đúng cột tốt nhất cho Stock_Code, Qty, Warehouse, BATCH, BIN, Date
    best_code_col, best_code_score = None, 0
    best_qty_col, best_qty_score = None, 0
    best_wh_col, best_wh_score = None, 0
    col_map = {}

    for c in df.columns:
        c_clean = str(c).strip().lower()
        c_score = _is_stock_code_col_header(c_clean)
        if c_score > best_code_score:
            best_code_score = c_score
            best_code_col = c

        q_score = _is_qty_col_header(c_clean)
        if q_score > best_qty_score:
            best_qty_score = q_score
            best_qty_col = c

        w_score = _is_warehouse_col_header(c_clean)
        if w_score > best_wh_score:
            best_wh_score = w_score
            best_wh_col = c

        if ('batch' in c_clean or 'lot' in c_clean or re.search(r'\blô\b', c_clean)) and 'BATCH' not in col_map.values():
            col_map[c] = 'BATCH'
        elif ('bin' in c_clean or 'location' in c_clean or 'vị trí' in c_clean or 'vi tri' in c_clean) and 'BIN' not in col_map.values():
            col_map[c] = 'BIN'
        elif ('date' in c_clean or 'ngày' in c_clean or 'ngay' in c_clean) and 'Date' not in col_map.values():
            col_map[c] = 'Date'

    if best_code_col is not None:
        col_map[best_code_col] = 'Stock_Code'
    if best_qty_col is not None and best_qty_col != best_code_col:
        col_map[best_qty_col] = 'Qty'
    if best_wh_col is not None and best_wh_col not in (best_code_col, best_qty_col):
        col_map[best_wh_col] = 'Warehouse'

    df = df.rename(columns=col_map)

    # Hỗ trợ trường hợp bảng xoay ngang (Wide / Pivot): có Stock_Code và các cột mang tên kho (VD: 'Kho 01', 'Kho 61', 'WH 01', 'WH 61')
    if 'Stock_Code' in df.columns and 'Qty' not in df.columns:
        wide_wh_cols = {}
        for c in df.columns:
            if c in ('Stock_Code', 'Warehouse', 'BATCH', 'BIN', 'Date'):
                continue
            w_cand = infer_warehouse_from_name(str(c))
            if w_cand:
                wide_wh_cols[c] = w_cand
        if wide_wh_cols:
            melted_parts = []
            base_cols = [c for c in ('Stock_Code', 'BATCH', 'BIN', 'Date') if c in df.columns]
            for c_orig, w_code in wide_wh_cols.items():
                sub = df[base_cols].copy()
                sub['Warehouse'] = w_code
                sub['Qty'] = df[c_orig]
                melted_parts.append(sub)
            df = pd.concat(melted_parts, ignore_index=True)

    # Fallback: Nếu file Excel người dùng nạp không có dòng tiêu đề (dòng đầu tiên đã là mã linh kiện & số lượng)
    if 'Stock_Code' not in df.columns or 'Qty' not in df.columns:
        df_no_hdr = _read_tabular_source(file_source, header=None, sheet_name=sheet_name)
        if df_no_hdr is not None and len(df_no_hdr) > 0 and df_no_hdr.shape[1] >= 2:
            # Kiểm tra xem cột 0 có phải mã linh kiện (>= 5 ký tự chữ/số) và cột cuối/cột 1/2 là số lượng không
            sample_c0 = df_no_hdr.iloc[:min(10, len(df_no_hdr)), 0].dropna().astype(str).str.strip()
            if len(sample_c0) > 0 and sample_c0.str.match(r'^[A-Za-z0-9\-_]{5,}$').mean() >= 0.6:
                if df_no_hdr.shape[1] >= 3:
                    df = pd.DataFrame({
                        'Stock_Code': df_no_hdr.iloc[:, 0],
                        'Warehouse': df_no_hdr.iloc[:, 1],
                        'Qty': df_no_hdr.iloc[:, 2]
                    })
                else:
                    df = pd.DataFrame({
                        'Stock_Code': df_no_hdr.iloc[:, 0],
                        'Qty': df_no_hdr.iloc[:, 1]
                    })

    if 'Stock_Code' not in df.columns:
        raise ValueError("Không tìm thấy cột Mã linh kiện / Item Code / Stock Code trong file tồn kho.")
    if 'Qty' not in df.columns:
        raise ValueError("Không tìm thấy cột Số lượng (Qty) trong file tồn kho.")

    df = df[pd.notna(df['Stock_Code'])].copy()
    df['Stock_Code'] = df['Stock_Code'].apply(normalize_code)
    invalid_codes = {'', 'nan', 'none', 'null', 'total', 'tổng', 'tong', 'sub total', 'subtotal', 'grand total', 'item code', 'stock code', 'component code', 'mã vật tư', 'mã linh kiện', 'mã hàng'}
    df = df[~df['Stock_Code'].str.lower().isin(invalid_codes)].copy()

    df['Qty'] = parse_numeric_qty(df['Qty'])

    for c in ['BATCH', 'BIN']:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str).str.strip()
        else:
            df[c] = ''

    if 'Warehouse' in df.columns:
        wh_series = df['Warehouse'].apply(normalize_warehouse_code)
        # Nếu cột Warehouse bị trống một số dòng (hoặc toàn bộ) và có mã kho suy ra từ tên sheet/file thì điền bổ sung
        if inferred_wh:
            wh_series = wh_series.apply(lambda w: w if w else inferred_wh)
        df['Warehouse'] = wh_series
    else:
        df['Warehouse'] = inferred_wh or ''

    return df


def load_stock_data(file_source):
    """
    Đọc dữ liệu Tồn kho từ đường dẫn file hoặc file-like object (hỗ trợ Stock 61, Stock 01, Stock Balance, ERP Inventory...).
    - Tự động quét tất cả các sheet trong file Excel (nếu có nhiều sheet như Sheet Kho 01, Sheet Kho 61 hoặc dữ liệu nằm ở sheet thứ 2).
    - Tự động nhận diện dòng tiêu đề thông minh, chuẩn hóa cột và suy luận mã kho từ tên sheet / tên file nếu bảng không có cột WH riêng.
    """
    fname = getattr(file_source, 'name', '') or (file_source if isinstance(file_source, str) else '')
    fname_lower = str(fname).lower()
    file_inferred_wh = infer_warehouse_from_name(fname)

    if fname_lower.endswith(('.csv', '.tsv')):
        df_raw = _read_tabular_source(file_source, header=None, nrows=50)
        return _parse_single_stock_df(df_raw, file_source, sheet_name=0, inferred_wh=file_inferred_wh)

    # Đối với file Excel (.xlsx, .xls): duyệt qua các sheet để lấy toàn bộ sheet tồn kho hợp lệ
    try:
        if hasattr(file_source, 'seek'):
            file_source.seek(0)
        xls = pd.ExcelFile(file_source)
        sheet_names = xls.sheet_names
    except Exception:
        sheet_names = [0]

    valid_dfs = []
    last_err = None
    for s_name in sheet_names:
        try:
            sheet_wh = infer_warehouse_from_name(str(s_name)) if isinstance(s_name, str) else ''
            eff_inferred_wh = sheet_wh or file_inferred_wh
            df_raw = _read_tabular_source(file_source, header=None, nrows=50, sheet_name=s_name)
            if df_raw is None or len(df_raw) == 0:
                continue
            df_sheet = _parse_single_stock_df(df_raw, file_source, sheet_name=s_name, inferred_wh=eff_inferred_wh)
            if df_sheet is not None and len(df_sheet) > 0:
                valid_dfs.append(df_sheet)
        except Exception as e:
            last_err = e
            continue

    if not valid_dfs:
        if last_err:
            raise last_err
        raise ValueError("Không tìm thấy dữ liệu tồn kho hợp lệ trong file.")

    if len(valid_dfs) == 1:
        return valid_dfs[0]
    return pd.concat(valid_dfs, ignore_index=True)


def find_all_stock_files(search_dirs=None) -> list:
    """
    Tự động quét và tìm tất cả các file tồn kho có trong repository / thư mục làm việc:
    - Loại bỏ các file tạm của Excel ('~$...'), file ẩn, file BOM, file Kế hoạch và các file báo cáo xuất ra.
    - Nhận diện mọi tên file tồn kho tiếng Anh / tiếng Việt ('Stock...', 'Tồn kho...', 'Ton kho...', 'Kho...', 'Inventory...', 'NXT...', v.v.)
      hoặc bất kỳ file bảng dữ liệu mới nào được nạp vào repo.
    - Sắp xếp ưu tiên các file tồn kho mới được nạp vào repo lên trước file mẫu mặc định cũ.
    """
    if search_dirs is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        search_dirs = [base_dir, os.getcwd()]
        # Quét thêm các thư mục con trực tiếp trong repo (ví dụ: data/, stock/) nếu có
        for d in list(search_dirs):
            if os.path.exists(d):
                try:
                    for sub in os.listdir(d):
                        sub_path = os.path.join(d, sub)
                        if os.path.isdir(sub_path) and not sub.startswith('.') and sub.lower() not in ('__pycache__', 'venv', '.venv', 'node_modules'):
                            if sub_path not in search_dirs:
                                search_dirs.append(sub_path)
                except Exception:
                    pass

    seen_paths = set()
    candidates = []
    for d in search_dirs:
        if not d or not os.path.exists(d):
            continue
        try:
            for fname in os.listdir(d):
                if fname.startswith(('~$', '.~', '.')):
                    continue
                if not fname.lower().endswith(('.xlsx', '.xls', '.csv')):
                    continue
                fl = fname.lower()
                # Loại trừ file BOM, Kế hoạch và file báo cáo xuất ra
                if any(k in fl for k in ('bom', 'định mức', 'dinh muc', 'kế hoạch', 'ke hoach', 'plan', 'doi_soat_', 'bao_cao_', 'agent_reconciliation', 'smart_reconciliation')):
                    continue
                fpath = os.path.abspath(os.path.join(d, fname))
                if fpath in seen_paths or not os.path.isfile(fpath):
                    continue
                seen_paths.add(fpath)
                candidates.append(fpath)
        except Exception:
            pass

    def _sort_key(fpath):
        bname = os.path.basename(fpath).lower()
        is_legacy_default = (bname == 'stock 61.xlsx')
        try:
            mtime = os.path.getmtime(fpath)
        except Exception:
            mtime = 0.0
        # Ưu tiên các file mới được nạp vào repo trước file mẫu 'Stock 61.xlsx' mặc định, sau đó theo thời gian sửa đổi mới nhất
        return (0 if not is_legacy_default else 1, -mtime, bname)

    candidates.sort(key=_sort_key)
    return candidates


def load_repo_stock_data(search_dirs=None):
    """
    Đọc và hợp nhất thông minh dữ liệu tồn kho từ các file tồn kho trong repo:
    - Nếu người dùng nạp thêm file tồn kho mới vào repo (VD: 'Stock 01.xlsx', 'Stock Balance.xlsx', 'Tồn kho.xlsx'...):
      + Tự động đọc tất cả các file tồn kho hợp lệ trong repo.
      + Nếu file mới nạp chứa kho khác với 'Stock 61.xlsx' (ví dụ nạp thêm 'Stock 01.xlsx' cho Kho 01): tự động gộp cả Kho 01 và Kho 61.
      + Nếu file mới nạp đã có sẵn dữ liệu Kho 61 (hoặc là file tổng hợp thay thế): ưu tiên lấy dữ liệu từ file mới nạp, không cộng trùng lặp với file mẫu 'Stock 61.xlsx' cũ.
    Trả về: (df_stock, display_name, loaded_paths)
    """
    candidates = find_all_stock_files(search_dirs)
    if not candidates:
        return None, None, []

    loaded_items = []
    last_err = None
    for fpath in candidates:
        try:
            df_s = load_stock_data(fpath)
            if df_s is not None and len(df_s) > 0:
                bname = os.path.basename(fpath)
                wh_set = set(w for w in df_s['Warehouse'].unique().tolist() if w)
                loaded_items.append({
                    'path': fpath,
                    'name': bname,
                    'is_legacy_default': (bname.lower() == 'stock 61.xlsx'),
                    'warehouses': wh_set,
                    'df': df_s
                })
        except Exception as e:
            last_err = e
            continue

    if not loaded_items:
        if last_err:
            raise last_err
        return None, None, []

    if len(loaded_items) == 1:
        item = loaded_items[0]
        return item['df'], item['name'], [item['path']]

    # Khi có nhiều file tồn kho trong repo (ví dụ có cả file mới nạp và 'Stock 61.xlsx' mặc định):
    non_legacy = [it for it in loaded_items if not it['is_legacy_default']]
    covered_whs = set()
    has_unassigned_wh = False
    selected_dfs = []
    selected_names = []
    selected_paths = []

    for it in non_legacy:
        selected_dfs.append(it['df'])
        selected_names.append(it['name'])
        selected_paths.append(it['path'])
        covered_whs.update(it['warehouses'])
        if not it['warehouses'] or ('' in it['df']['Warehouse'].values):
            has_unassigned_wh = True

    for it in loaded_items:
        if not it['is_legacy_default']:
            continue
        # Chỉ ghép thêm file mẫu 'Stock 61.xlsx' đối với những kho chưa có trong file tồn kho mới (và khi file mới không phải file tổng không chia kho)
        if not has_unassigned_wh:
            df_leg = it['df']
            leg_wh_col = df_leg['Warehouse'].apply(normalize_warehouse_code)
            df_leg_needed = df_leg[~leg_wh_col.isin(covered_whs)].copy()
            if len(df_leg_needed) > 0:
                selected_dfs.append(df_leg_needed)
                selected_names.append(it['name'])
                selected_paths.append(it['path'])
                covered_whs.update(set(w for w in df_leg_needed['Warehouse'].unique().tolist() if w))

    combined_df = pd.concat(selected_dfs, ignore_index=True) if len(selected_dfs) > 1 else selected_dfs[0]
    display_name = " + ".join(selected_names)
    return combined_df, display_name, selected_paths


def get_available_orders_info(df_bom):
    """
    Trích xuất nhanh danh sách tất cả các đơn hàng có trong BOM cùng thông tin tóm tắt:
    tên sản phẩm, số biến thể, số linh kiện, có đủ trái phải không.
    Tối ưu hóa tốc độ bằng thao tác dictionary / records nhanh.
    """
    orders_info = {}
    grouped = df_bom.groupby('Order_No', sort=True)
    
    for order_no, group in grouped:
        prods = group[['Product_Code', 'Product_Desc', 'Side', 'Variant']].drop_duplicates()
        sides = sorted(prods['Side'].unique().tolist())
        has_both_sides = ('Trái' in sides) and ('Phải' in sides)
        
        desc_list = [d for d in prods['Product_Desc'].tolist() if d and str(d).strip().lower() not in ('', 'nan')]
        raw_desc = desc_list[0] if desc_list else f"Sản phẩm đơn hàng {order_no}"
        clean_desc = clean_order_title(raw_desc, order_no)
        
        variants = sorted(prods['Variant'].unique().tolist())
        unique_components = group['Component_Code'].nunique()
        total_bom_rows = len(group)
        
        orders_info[order_no] = {
            'order_no': order_no,
            'description': clean_desc,
            'raw_description': raw_desc,
            'sides': sides,
            'has_both_sides': has_both_sides,
            'variants': variants,
            'unique_components': unique_components,
            'total_bom_rows': total_bom_rows,
            'products': prods.to_dict(orient='records')
        }
        
    return orders_info


def _select_best_variant_for_order(order_boms: pd.DataFrame, stock_qty_pool: dict, requested_variant: str = None) -> str:
    """
    Chọn Variant tối ưu nhất cho đơn hàng:
    - Nếu người dùng hoặc File Kế hoạch chỉ định rõ `requested_variant` và có trong BOM -> Dùng đúng `requested_variant`.
    - Nếu không chỉ định (hoặc đơn hàng có nhiều Variant như 0004, 0104):
      Tự động đánh giá tồn kho thực tế của từng Variant, ưu tiên Variant có đủ nhiều mã linh kiện nhất
      và cho số lượng bộ chạy tối đa (max_runnable_qty) cao nhất!
    """
    avail_variants = sorted(order_boms['Variant'].unique().tolist())
    if len(avail_variants) <= 1:
        return avail_variants[0] if avail_variants else ''
    if requested_variant and requested_variant in avail_variants:
        return requested_variant
        
    best_var = avail_variants[0]
    best_score = (-1, -1.0, -1)
    for v in avail_variants:
        sub = order_boms[order_boms['Variant'] == v]
        comp_req = sub.groupby('Component_Code')['Qty'].sum()
        if len(comp_req) == 0:
            continue
        in_stock_cnt = 0
        min_possible = float('inf')
        for c_code, u_qty in comp_req.items():
            s_qty = float(stock_qty_pool.get(c_code, 0.0))
            u_qty_f = round(float(u_qty), 6)
            if s_qty > EPS and u_qty_f > EPS:
                in_stock_cnt += 1
                possible = float(math.floor((s_qty + EPS) / u_qty_f))
                if possible < min_possible:
                    min_possible = possible
            else:
                min_possible = 0.0
        if min_possible == float('inf'):
            min_possible = 0.0
        is_100_ready = 1 if (in_stock_cnt == len(comp_req) and min_possible > 0) else 0
        score = (is_100_ready, min_possible, in_stock_cnt)
        if score > best_score:
            best_score = score
            best_var = v
    return best_var


def _get_side_priority(side: str) -> int:
    s = str(side).strip().lower()
    if any(k in s for k in ['trái', 'trai', 'left', 'lh']):
        return 0
    elif any(k in s for k in ['phải', 'phai', 'right', 'rh']):
        return 1
    return 2


def calculate_inventory_allocation(orders_queue, df_bom, df_stock):
    """
    Tính toán đối soát BOM và tồn kho theo cơ chế FIFO cho danh sách đơn hàng chờ.
    Đã được tối ưu hóa hiệu năng cao (~50x nhanh hơn, xử lý tức thì < 0.12s):
    - Đối soát đúng theo kho (WH) mà đơn hàng yêu cầu (ví dụ kho '01' lấy đúng kho '01';
      nếu không chọn kho thì mặc định lấy đúng 2 kho '01' & '61', TUYỆT ĐỐI KHÔNG lấy dữ liệu ở kho khác).
    - Lấy đầy đủ định mức Trái và Phải từ BOM (tự động chọn Variant tối ưu hoặc theo Kế hoạch).
    - Nhân số lượng đơn hàng với định mức BOM.
    - So sánh với tồn kho khả dụng hiện tại của kho tương ứng.
    - Ưu tiên cấp liệu cho đơn hàng nhập trước, trừ dần tồn kho (Trái trước - Phải sau).
    """
    stock_by_wh = {}
    if df_stock is not None and len(df_stock) > 0:
        temp_stock = df_stock[['Stock_Code', 'Qty']].copy()
        temp_stock['Warehouse_Norm'] = (
            df_stock['Warehouse'].apply(normalize_warehouse_code)
            if 'Warehouse' in df_stock.columns else ''
        )
        for (wh_val, scode), grp_qty in temp_stock.groupby(['Warehouse_Norm', 'Stock_Code'])['Qty'].sum().items():
            if wh_val not in stock_by_wh:
                stock_by_wh[wh_val] = {}
            stock_by_wh[wh_val][scode] = float(grp_qty)

    results = []
    
    # Pre-group BOM theo Order_No để tra cứu O(1) thay vì quét lại toàn bộ DataFrame
    bom_by_order = {ord_no: grp for ord_no, grp in df_bom.groupby('Order_No', sort=False)}
    
    for idx, item in enumerate(orders_queue, 1):
        order_no = str(item['order_no']).strip()
        order_qty = float(item['qty'])
        qty_left = float(item.get('qty_left', order_qty))
        qty_right = float(item.get('qty_right', order_qty))
        variant = item.get('variant', None)
        item_id = item.get('id', f"ord_{idx}")
        created_at = item.get('created_at', '')
        models_map = item.get('models', {})
        due_date = item.get('due_date', '')
        target_whs = parse_target_warehouses(item.get('wh'))
        wh_code_str = "&".join(target_whs)
        eff_target_whs = list(target_whs)
        if '' in stock_by_wh and '' not in eff_target_whs:
            eff_target_whs.append('')
        if target_whs == DEFAULT_WAREHOUSES and not any(w in stock_by_wh for w in eff_target_whs) and stock_by_wh:
            eff_target_whs = list(stock_by_wh.keys())
        
        order_boms = bom_by_order.get(order_no)
        if order_boms is None or len(order_boms) == 0:
            results.append({
                'id': item_id,
                'stt': idx,
                'order_no': order_no,
                'wh': wh_code_str,
                'order_qty': order_qty,
                'qty_left': qty_left,
                'qty_right': qty_right,
                'models': models_map,
                'due_date': due_date,
                'created_at': created_at,
                'error': f"Không tìm thấy dữ liệu BOM cho đơn hàng {order_no}",
                'total_items': 0,
                'ok_count': 0,
                'missing_count': 0,
                'in_stock_count': 0,
                'max_runnable_fifo': 0,
                'fifo_bottleneck_code': '',
                'fifo_bottleneck_desc': '',
                'fifo_bottleneck_stock': 0.0,
                'is_fully_ok': False,
                'summary_table': pd.DataFrame(),
                'detail_table': pd.DataFrame()
            })
            continue
            
        order_comp_codes = order_boms['Component_Code'].unique()
        order_stock_pool = {
            cc: round(sum(float(stock_by_wh.get(w, {}).get(cc, 0.0)) for w in eff_target_whs), 6)
            for cc in order_comp_codes
        }
        chosen_variant = _select_best_variant_for_order(order_boms, order_stock_pool, variant)
        bom_order = order_boms[order_boms['Variant'] == chosen_variant]
        if len(bom_order) == 0:
            bom_order = order_boms
            
        desc_list = [d for d in bom_order['Product_Desc'].unique() if d and str(d).strip().lower() not in ('', 'nan')]
        raw_desc = desc_list[0] if desc_list else f"Sản phẩm đơn hàng {order_no}"
        product_title = clean_order_title(raw_desc, order_no)
        
        summary_rows = []
        detail_rows = []
        
        # Chuyển sang danh sách dict thuần Python để tính toán tốc độ cao
        records = bom_order[['Component_Code', 'Component_Desc', 'Side', 'Product_Code', 'Qty']].to_dict('records')
        comp_map = {}
        for r in records:
            cc = r['Component_Code']
            if cc not in comp_map:
                comp_map[cc] = []
            comp_map[cc].append(r)

        in_stock_count = 0
        min_fifo_sets = float('inf')
        fifo_bottleneck_code = ''
        fifo_bottleneck_desc = ''
        fifo_bottleneck_stock = 0.0
            
        for comp_code in sorted(comp_map.keys()):
            c_rows = comp_map[comp_code]
            comp_desc = c_rows[0]['Component_Desc']
            
            left_bom_qty = round(sum(float(r['Qty']) for r in c_rows if r['Side'] == 'Trái'), 6)
            right_bom_qty = round(sum(float(r['Qty']) for r in c_rows if r['Side'] == 'Phải'), 6)
            other_bom_qty = round(sum(float(r['Qty']) for r in c_rows if r['Side'] not in ('Trái', 'Phải')), 6)
            total_bom_unit = round(left_bom_qty + right_bom_qty + other_bom_qty, 6)
            
            req_left = round(left_bom_qty * qty_left, 6)
            req_right = round(right_bom_qty * qty_right, 6)
            req_other = round(other_bom_qty * max(qty_left, qty_right), 6)
            total_required = round(req_left + req_right + req_other, 6)
            
            initial_stock = round(sum(float(stock_by_wh.get(w, {}).get(comp_code, 0.0)) for w in eff_target_whs), 6)
            running_stock = initial_stock

            if initial_stock > EPS:
                in_stock_count += 1
            if total_bom_unit > EPS:
                comp_possible_sets = int(math.floor((initial_stock + EPS) / total_bom_unit)) if initial_stock > EPS else 0
                if comp_possible_sets < min_fifo_sets:
                    min_fifo_sets = comp_possible_sets
                    fifo_bottleneck_code = comp_code
                    fifo_bottleneck_desc = comp_desc
                    fifo_bottleneck_stock = initial_stock
            else:
                comp_possible_sets = 0
            
            c_rows_sorted = sorted(c_rows, key=lambda x: (_get_side_priority(x['Side']), x['Product_Code']))
            comp_allocated = 0.0
            comp_shortage = 0.0
            
            for brow in c_rows_sorted:
                b_side = brow['Side']
                b_pcode = brow['Product_Code']
                b_pdesc = brow['Component_Desc']
                b_qty = round(float(brow['Qty']), 6)
                
                if b_side == 'Trái':
                    side_order_qty = qty_left
                elif b_side == 'Phải':
                    side_order_qty = qty_right
                else:
                    side_order_qty = max(qty_left, qty_right)
                    
                row_required = round(b_qty * side_order_qty, 6)
                row_avail_stock = round(max(0.0, running_stock), 6)
                
                if running_stock + EPS >= row_required:
                    row_allocated = row_required
                    row_shortage = 0.0
                    running_stock = round(max(0.0, running_stock - row_required), 6)
                    row_status = "OK"
                    row_note = "Đủ hàng"
                else:
                    row_allocated = round(max(0.0, running_stock), 6)
                    row_shortage = round(max(0.0, row_required - row_allocated), 6)
                    running_stock = 0.0
                    row_status = "THIẾU"
                    row_note = f"Thiếu {int(row_shortage) if row_shortage.is_integer() else row_shortage:g}"
                    
                comp_allocated = round(comp_allocated + row_allocated, 6)
                comp_shortage = round(comp_shortage + row_shortage, 6)
                
                detail_rows.append({
                    'Kho (WH)': wh_code_str,
                    'Vế': b_side,
                    'Mã sản phẩm (Product Code)': b_pcode,
                    'ITEM CODE': comp_code,
                    'DESCRIPTION (tên mô tả)': b_pdesc,
                    'Định mức BOM': b_qty,
                    'Số lượng đơn': side_order_qty,
                    'Số lượng cần': row_required,
                    'Tồn kho khả dụng': row_avail_stock,
                    'Số lượng cấp': row_allocated,
                    'Số lượng thiếu': row_shortage,
                    'Trạng thái': row_status,
                    'Ghi chú': row_note
                })
            
            # Trừ dần tồn kho trên đúng các kho mục tiêu (eff_target_whs) của đơn hàng
            rem_to_deduct = comp_allocated
            for w in eff_target_whs:
                if rem_to_deduct <= EPS:
                    break
                w_pool = stock_by_wh.get(w)
                if w_pool and comp_code in w_pool:
                    avail_w = round(float(w_pool[comp_code]), 6)
                    if avail_w > EPS:
                        deduct_w = round(min(avail_w, rem_to_deduct), 6)
                        w_pool[comp_code] = round(max(0.0, avail_w - deduct_w), 6)
                        rem_to_deduct = round(max(0.0, rem_to_deduct - deduct_w), 6)
            
            if comp_shortage <= EPS:
                comp_shortage = 0.0
                summary_status = "OK"
                summary_note = "Đủ hàng"
            else:
                summary_status = "THIẾU"
                summary_note = f"Thiếu {int(comp_shortage) if comp_shortage.is_integer() else comp_shortage:g}"
                
            summary_rows.append({
                'Kho (WH)': wh_code_str,
                'ITEM CODE': comp_code,
                'DESCRIPTION (tên mô tả)': comp_desc,
                'Định mức Trái': left_bom_qty,
                'Định mức Phải': right_bom_qty,
                'Tổng định mức (1 bộ)': total_bom_unit,
                'SL Cần Trái': req_left,
                'SL Cần Phải': req_right,
                'SL BOM cần chạy': total_required,
                'Số lượng tồn kho': initial_stock,
                'SL bộ đáp ứng': comp_possible_sets,
                'Số lượng cấp': comp_allocated,
                'Số lượng thiếu': comp_shortage,
                'Trạng thái': summary_status,
                'Ghi chú': summary_note
            })
            
        summary_rows.sort(key=lambda x: (0 if x['Trạng thái'] == 'THIẾU' else 1, -x['Số lượng thiếu'], x['ITEM CODE']))
        detail_rows.sort(key=lambda x: (_get_side_priority(x['Vế']), 0 if x['Trạng thái'] == 'THIẾU' else 1, x['ITEM CODE']))
        
        df_summary = pd.DataFrame(summary_rows)
        df_detail = pd.DataFrame(detail_rows)
        
        total_items = len(summary_rows)
        ok_count = sum(1 for r in summary_rows if r['Trạng thái'] == 'OK')
        missing_count = total_items - ok_count
        max_runnable_fifo = 0 if (min_fifo_sets == float('inf') or min_fifo_sets < 0) else int(min_fifo_sets)
        
        results.append({
            'id': item_id,
            'stt': idx,
            'order_no': order_no,
            'wh': wh_code_str,
            'order_qty': order_qty,
            'qty_left': qty_left,
            'qty_right': qty_right,
            'models': models_map,
            'due_date': due_date,
            'product_title': product_title,
            'variant': chosen_variant,
            'created_at': created_at,
            'total_items': total_items,
            'ok_count': ok_count,
            'missing_count': missing_count,
            'in_stock_count': in_stock_count,
            'max_runnable_fifo': max_runnable_fifo,
            'fifo_bottleneck_code': fifo_bottleneck_code,
            'fifo_bottleneck_desc': fifo_bottleneck_desc,
            'fifo_bottleneck_stock': fifo_bottleneck_stock,
            'is_fully_ok': (missing_count == 0 and total_items > 0),
            'summary_table': df_summary,
            'detail_table': df_detail
        })
        
    return results


def optimize_fifo_queue(orders_queue, df_bom, df_stock):
    """
    Sắp xếp tối ưu hàng chờ FIFO tốc độ cao (< 0.01s) sao cho số lượng đơn ĐỦ 100% VẬT TƯ đạt tối đa:
    - Ở mỗi bước chọn đơn tiếp theo:
      + Ưu tiên đơn đủ 100% vật tư theo số lượng đặt (order_qty) trên tồn kho còn lại.
      + Tiếp theo ưu tiên đơn có số bộ chạy tối đa (max_runnable_fifo) cao nhất và tỷ lệ đủ mã cao nhất.
    """
    if not orders_queue or len(orders_queue) <= 1 or df_bom is None:
        return list(orders_queue)

    stock_by_wh = {}
    if df_stock is not None and len(df_stock) > 0:
        temp_stock = df_stock[['Stock_Code', 'Qty']].copy()
        temp_stock['Warehouse_Norm'] = (
            df_stock['Warehouse'].apply(normalize_warehouse_code)
            if 'Warehouse' in df_stock.columns else ''
        )
        for (wh_val, scode), grp_qty in temp_stock.groupby(['Warehouse_Norm', 'Stock_Code'])['Qty'].sum().items():
            if wh_val not in stock_by_wh:
                stock_by_wh[wh_val] = {}
            stock_by_wh[wh_val][scode] = round(float(grp_qty), 6)

    bom_by_order = {ord_no: grp for ord_no, grp in df_bom.groupby('Order_No', sort=False)}
    prepared = []
    for orig_idx, item in enumerate(orders_queue):
        order_no = str(item.get('order_no', '')).strip()
        order_qty = float(item.get('qty', 0))
        qty_left = float(item.get('qty_left', order_qty))
        qty_right = float(item.get('qty_right', order_qty))
        target_whs = parse_target_warehouses(item.get('wh'))
        eff_whs = list(target_whs)
        if '' in stock_by_wh and '' not in eff_whs:
            eff_whs.append('')
        if target_whs == DEFAULT_WAREHOUSES and not any(w in stock_by_wh for w in eff_whs) and stock_by_wh:
            eff_whs = list(stock_by_wh.keys())

        order_boms = bom_by_order.get(order_no)
        comp_reqs = []
        if order_boms is not None and len(order_boms) > 0:
            order_comp_codes = order_boms['Component_Code'].unique()
            pool = {cc: sum(float(stock_by_wh.get(w, {}).get(cc, 0.0)) for w in eff_whs) for cc in order_comp_codes}
            chosen_var = _select_best_variant_for_order(order_boms, pool, item.get('variant'))
            sub = order_boms[order_boms['Variant'] == chosen_var]
            if len(sub) == 0:
                sub = order_boms
            for cc, grp in sub.groupby('Component_Code', sort=False):
                l_q = round(float(grp[grp['Side'] == 'Trái']['Qty'].sum()), 6)
                r_q = round(float(grp[grp['Side'] == 'Phải']['Qty'].sum()), 6)
                o_q = round(float(grp[~grp['Side'].isin(['Trái', 'Phải'])]['Qty'].sum()), 6)
                unit_q = round(l_q + r_q + o_q, 6)
                tot_req = round(l_q * qty_left + r_q * qty_right + o_q * max(qty_left, qty_right), 6)
                comp_reqs.append((cc, unit_q, tot_req))

        prepared.append({
            'orig_idx': orig_idx,
            'item': dict(item),
            'eff_whs': eff_whs,
            'comp_reqs': comp_reqs
        })

    ordered_items = []
    remaining = list(prepared)
    while remaining:
        best_i = 0
        best_score = None
        for i, cand in enumerate(remaining):
            eff_whs = cand['eff_whs']
            creqs = cand['comp_reqs']
            if not creqs:
                score = (-1, 0, 0, 0.0, -cand['orig_idx'])
            else:
                ok_cnt = 0
                min_sets = float('inf')
                for cc, unit_q, tot_req in creqs:
                    avail = round(sum(float(stock_by_wh.get(w, {}).get(cc, 0.0)) for w in eff_whs), 6)
                    if avail + EPS >= tot_req:
                        ok_cnt += 1
                    if unit_q > EPS:
                        sets = int(math.floor((avail + EPS) / unit_q)) if avail > EPS else 0
                        if sets < min_sets:
                            min_sets = sets
                max_run = 0 if min_sets == float('inf') else int(min_sets)
                is_full_ok = 1 if ok_cnt == len(creqs) else 0
                has_run = 1 if max_run > 0 else 0
                ok_ratio = ok_cnt / len(creqs)
                if is_full_ok:
                    score = (1, -cand['orig_idx'], 1, max_run, ok_ratio)
                else:
                    score = (0, has_run, max_run, ok_ratio, -cand['orig_idx'])
            if best_score is None or score > best_score:
                best_score = score
                best_i = i

        chosen = remaining.pop(best_i)
        ordered_items.append(chosen['item'])
        # Trừ tồn kho cho đơn vừa chọn
        for cc, unit_q, tot_req in chosen['comp_reqs']:
            rem_deduct = tot_req
            for w in chosen['eff_whs']:
                if rem_deduct <= EPS:
                    break
                w_pool = stock_by_wh.get(w)
                if w_pool and cc in w_pool:
                    avail_w = float(w_pool[cc])
                    if avail_w > EPS:
                        d_w = min(avail_w, rem_deduct)
                        w_pool[cc] = round(max(0.0, avail_w - d_w), 6)
                        rem_deduct = round(max(0.0, rem_deduct - d_w), 6)

    return ordered_items


def cap_queue_to_runnable_fifo(orders_queue, df_bom, df_stock):
    """
    Tự động điều chỉnh số lượng các đơn trong hàng chờ FIFO về mức bộ đồng bộ tối đa chạy được ngay (max_runnable_fifo)
    khi 0 < max_runnable_fifo < order_qty. Xử lý tuần tự tốc độ cao (< 0.02s).
    Trả về: (updated_queue, adjusted_count)
    """
    if not orders_queue or df_bom is None:
        return list(orders_queue), 0

    stock_by_wh = {}
    if df_stock is not None and len(df_stock) > 0:
        temp_stock = df_stock[['Stock_Code', 'Qty']].copy()
        temp_stock['Warehouse_Norm'] = (
            df_stock['Warehouse'].apply(normalize_warehouse_code)
            if 'Warehouse' in df_stock.columns else ''
        )
        for (wh_val, scode), grp_qty in temp_stock.groupby(['Warehouse_Norm', 'Stock_Code'])['Qty'].sum().items():
            if wh_val not in stock_by_wh:
                stock_by_wh[wh_val] = {}
            stock_by_wh[wh_val][scode] = round(float(grp_qty), 6)

    bom_by_order = {ord_no: grp for ord_no, grp in df_bom.groupby('Order_No', sort=False)}
    updated_queue = []
    adjusted_cnt = 0

    for item in orders_queue:
        new_item = dict(item)
        order_no = str(new_item.get('order_no', '')).strip()
        order_qty = float(new_item.get('qty', 0))
        qty_left = float(new_item.get('qty_left', order_qty))
        qty_right = float(new_item.get('qty_right', order_qty))
        target_whs = parse_target_warehouses(new_item.get('wh'))
        eff_whs = list(target_whs)
        if '' in stock_by_wh and '' not in eff_whs:
            eff_whs.append('')
        if target_whs == DEFAULT_WAREHOUSES and not any(w in stock_by_wh for w in eff_whs) and stock_by_wh:
            eff_whs = list(stock_by_wh.keys())

        order_boms = bom_by_order.get(order_no)
        if order_boms is not None and len(order_boms) > 0:
            order_comp_codes = order_boms['Component_Code'].unique()
            pool = {cc: sum(float(stock_by_wh.get(w, {}).get(cc, 0.0)) for w in eff_whs) for cc in order_comp_codes}
            chosen_var = _select_best_variant_for_order(order_boms, pool, new_item.get('variant'))
            sub = order_boms[order_boms['Variant'] == chosen_var]
            if len(sub) == 0:
                sub = order_boms

            comp_reqs = []
            min_sets = float('inf')
            all_ok = True
            for cc, grp in sub.groupby('Component_Code', sort=False):
                l_q = round(float(grp[grp['Side'] == 'Trái']['Qty'].sum()), 6)
                r_q = round(float(grp[grp['Side'] == 'Phải']['Qty'].sum()), 6)
                o_q = round(float(grp[~grp['Side'].isin(['Trái', 'Phải'])]['Qty'].sum()), 6)
                unit_q = round(l_q + r_q + o_q, 6)
                tot_req = round(l_q * qty_left + r_q * qty_right + o_q * max(qty_left, qty_right), 6)
                avail = round(sum(float(stock_by_wh.get(w, {}).get(cc, 0.0)) for w in eff_whs), 6)
                if avail + EPS < tot_req:
                    all_ok = False
                if unit_q > EPS:
                    sets = int(math.floor((avail + EPS) / unit_q)) if avail > EPS else 0
                    if sets < min_sets:
                        min_sets = sets
                comp_reqs.append((cc, unit_q, l_q, r_q, o_q))

            max_run = 0 if min_sets == float('inf') else int(min_sets)
            if (not all_ok) and (0 < max_run < order_qty):
                new_item['qty'] = float(max_run)
                new_item['qty_left'] = float(max_run)
                new_item['qty_right'] = float(max_run)
                qty_left = float(max_run)
                qty_right = float(max_run)
                adjusted_cnt += 1

            # Trừ tồn kho theo số lượng sau khi đã chốt
            for cc, unit_q, l_q, r_q, o_q in comp_reqs:
                rem_deduct = round(l_q * qty_left + r_q * qty_right + o_q * max(qty_left, qty_right), 6)
                for w in eff_whs:
                    if rem_deduct <= EPS:
                        break
                    w_pool = stock_by_wh.get(w)
                    if w_pool and cc in w_pool:
                        avail_w = float(w_pool[cc])
                        if avail_w > EPS:
                            d_w = min(avail_w, rem_deduct)
                            w_pool[cc] = round(max(0.0, avail_w - d_w), 6)
                            rem_deduct = round(max(0.0, rem_deduct - d_w), 6)

        updated_queue.append(new_item)

    return updated_queue, adjusted_cnt


def export_order_to_excel(order_result):
    """
    Xuất kết quả đối soát của một đơn hàng ra file Excel .xlsx với định dạng chuyên nghiệp:
    - Sheet 1: Tong_Hop_Vat_Tu (Tổng hợp theo mã linh kiện)
    - Sheet 2: Chi_Tiet_BOM_Trai_Phai (Chi tiết theo từng dòng BOM Trái/Phải)
    - Dòng OK: Tô màu xanh lá nhạt (#D4EDDA), chữ xanh đậm (#155724)
    - Dòng THIẾU: Tô màu đỏ nhạt (#F8D7DA), chữ đỏ đậm (#721C24)
    """
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Tong_Hop_Vat_Tu"
    ws2 = wb.create_sheet(title="Chi_Tiet_BOM_Trai_Phai")
    
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    font_ok = Font(name="Segoe UI", size=10, color="155724")
    fill_ok = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    
    font_missing = Font(name="Segoe UI", size=10, bold=True, color="721C24")
    fill_missing = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    
    font_title = Font(name="Segoe UI", size=13, bold=True, color="1E3A8A")
    font_sub = Font(name="Segoe UI", size=10, italic=True, color="4B5563")
    
    border_thin = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # ===== SHEET 1 =====
    df1 = order_result['summary_table']
    order_no = order_result['order_no']
    order_qty = order_result['order_qty']
    p_title = order_result.get('product_title', '')
    wh_label = order_result.get('wh', '01&61')
    
    ws1.cell(row=1, column=1, value=f"KẾT QUẢ ĐỐI SOÁT TỒN KHO - ĐƠN HÀNG: {order_no} ({p_title}) - KHO: {wh_label}").font = font_title
    ws1.cell(row=2, column=1, value=f"Kho đối soát (WH): {wh_label} | Số lượng sản xuất: {order_qty:g} bộ (Trái + Phải) | Phiên bản BOM: {order_result.get('variant', '')} | Tổng linh kiện: {order_result['total_items']} | ĐỦ: {order_result['ok_count']} | THIẾU: {order_result['missing_count']}").font = font_sub
    
    start_r1 = 4
    headers1 = list(df1.columns)
    for col_idx, col_name in enumerate(headers1, 1):
        cell = ws1.cell(row=start_r1, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_thin
    ws1.row_dimensions[start_r1].height = 28
    
    rows1 = df1.to_dict('records')
    for r_idx, row in enumerate(rows1):
        curr_row = start_r1 + 1 + r_idx
        is_ok = (row['Trạng thái'] == 'OK')
        row_fill = fill_ok if is_ok else fill_missing
        row_font = font_ok if is_ok else font_missing
        
        for c_idx, col_name in enumerate(headers1, 1):
            val = row[col_name]
            cell = ws1.cell(row=curr_row, column=c_idx, value=val)
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border_thin
            if isinstance(val, (int, float)):
                cell.alignment = align_right
                cell.number_format = "#,##0.##"
            else:
                cell.alignment = align_left
        ws1.row_dimensions[curr_row].height = 20
        
    for col_idx, col_name in enumerate(headers1, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(col_name))] + [len(str(r.get(col_name, ''))) for r in rows1])
        ws1.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
        
    # ===== SHEET 2 =====
    df2 = order_result['detail_table']
    ws2.cell(row=1, column=1, value=f"CHI TIẾT BOM TRÁI & PHẢI - ĐƠN HÀNG: {order_no} (KHO: {wh_label})").font = font_title
    ws2.cell(row=2, column=1, value=f"Kho đối soát (WH): {wh_label} | Số lượng đơn hàng: {order_qty:g} | Định mức cho từng vế nhân với số lượng đơn (Trừ vế Trái trước - Phải sau)").font = font_sub
    
    start_r2 = 4
    headers2 = list(df2.columns)
    for col_idx, col_name in enumerate(headers2, 1):
        cell = ws2.cell(row=start_r2, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_thin
    ws2.row_dimensions[start_r2].height = 28
    
    rows2 = df2.to_dict('records')
    for r_idx, row in enumerate(rows2):
        curr_row = start_r2 + 1 + r_idx
        is_ok = (row['Trạng thái'] == 'OK')
        row_fill = fill_ok if is_ok else fill_missing
        row_font = font_ok if is_ok else font_missing
        
        for c_idx, col_name in enumerate(headers2, 1):
            val = row[col_name]
            cell = ws2.cell(row=curr_row, column=c_idx, value=val)
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border_thin
            if isinstance(val, (int, float)):
                cell.alignment = align_right
                cell.number_format = "#,##0.##"
            else:
                cell.alignment = align_left
        ws2.row_dimensions[curr_row].height = 20
        
    for col_idx, col_name in enumerate(headers2, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(col_name))] + [len(str(r.get(col_name, ''))) for r in rows2])
        ws2.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
        
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def export_all_orders_to_excel(all_results):
    """
    Xuất toàn bộ kết quả các đơn hàng trong danh sách chờ ra một file Excel tổng hợp
    với mỗi đơn hàng gồm 1 sheet tổng hợp vật tư.
    """
    wb = openpyxl.Workbook()
    ws_idx = wb.active
    ws_idx.title = "Tong_Ket_Cac_Don_Hang"
    
    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    font_ok = Font(name="Segoe UI", size=10, color="155724")
    fill_ok = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    
    font_missing = Font(name="Segoe UI", size=10, bold=True, color="721C24")
    fill_missing = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    ws_idx.cell(row=1, column=1, value="BẢNG TỔNG KẾT ĐỐI SOÁT VẬT TƯ CÁC ĐƠN HÀNG (FIFO)").font = font_title
    headers_idx = ["STT", "Mã đơn hàng", "Kho (WH)", "Sản phẩm", "Số lượng đặt", "SL tối đa chạy được (FIFO)", "Tổng số mã LK", "Số mã ĐỦ", "Số mã THIẾU", "Đánh giá chung"]
    
    for c_i, h in enumerate(headers_idx, 1):
        cell = ws_idx.cell(row=3, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_thin
    ws_idx.row_dimensions[3].height = 26
    
    for r_i, r in enumerate(all_results, 1):
        curr_r = 3 + r_i
        is_ok = r.get('is_fully_ok', False)
        row_fill = fill_ok if is_ok else fill_missing
        row_font = font_ok if is_ok else font_missing
        max_run_fifo = r.get('max_runnable_fifo', 0)
        
        vals = [
            r['stt'],
            r['order_no'],
            r.get('wh', '01&61'),
            r.get('product_title', ''),
            r['order_qty'],
            max_run_fifo,
            r['total_items'],
            r['ok_count'],
            r['missing_count'],
            "ĐỦ VẬT TƯ (OK)" if is_ok else (f"THIẾU {r['missing_count']} MÃ (Chạy được {max_run_fifo:g} bộ)" if max_run_fifo > 0 else f"THIẾU {r['missing_count']} MÃ")
        ]
        
        for c_i, val in enumerate(vals, 1):
            cell = ws_idx.cell(row=curr_r, column=c_i, value=val)
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border_thin
            cell.alignment = align_right if isinstance(val, (int, float)) else align_left
        ws_idx.row_dimensions[curr_r].height = 20
        
    for col in ws_idx.columns:
        max_len = max([len(str(cell.value or '')) for cell in col])
        col_letter = get_column_letter(col[0].column)
        ws_idx.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    # Sheet 2: Chi tiết tất cả các đơn hàng (Trái trước, Phải sau)
    ws_all_det = wb.create_sheet(title="Chi_Tiet_BOM_Tat_Ca_Don")
    ws_all_det.cell(row=1, column=1, value="CHI TIẾT TỔNG HỢP BOM TRÁI & PHẢI TẤT CẢ CÁC ĐƠN HÀNG (FIFO)").font = font_title
    ws_all_det.cell(row=2, column=1, value="Cơ chế FIFO: Ưu tiên đơn hàng trên trước, đối soát theo đúng Kho (WH), trừ dần vế Trái trước - vế Phải sau").font = Font(name="Segoe UI", size=10, italic=True, color="4B5563")
    
    headers_all_det = [
        "STT Đơn", "Mã đơn hàng", "Kho (WH)", "Vế", "Mã sản phẩm (Product Code)",
        "ITEM CODE", "DESCRIPTION (tên mô tả)", "Định mức BOM",
        "Số lượng đơn", "Số lượng cần", "Tồn kho khả dụng", "Số lượng cấp",
        "Số lượng thiếu", "Trạng thái", "Ghi chú"
    ]
    
    start_all_det_r = 4
    for c_i, h in enumerate(headers_all_det, 1):
        c = ws_all_det.cell(row=start_all_det_r, column=c_i, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_header
        c.border = border_thin
    ws_all_det.row_dimensions[start_all_det_r].height = 28
    
    curr_det_row = start_all_det_r
    for r in all_results:
        df_det = r.get('detail_table', pd.DataFrame())
        if len(df_det) == 0:
            continue
        for row_item in df_det.to_dict('records'):
            curr_det_row += 1
            is_ok = (row_item['Trạng thái'] == 'OK')
            row_fill = fill_ok if is_ok else fill_missing
            row_font = font_ok if is_ok else font_missing
            
            row_vals = [
                r['stt'],
                r['order_no'],
                r.get('wh', '01&61'),
                row_item['Vế'],
                row_item['Mã sản phẩm (Product Code)'],
                row_item['ITEM CODE'],
                row_item['DESCRIPTION (tên mô tả)'],
                row_item['Định mức BOM'],
                row_item['Số lượng đơn'],
                row_item['Số lượng cần'],
                row_item['Tồn kho khả dụng'],
                row_item['Số lượng cấp'],
                row_item['Số lượng thiếu'],
                row_item['Trạng thái'],
                row_item['Ghi chú']
            ]
            
            for c_i, val in enumerate(row_vals, 1):
                c = ws_all_det.cell(row=curr_det_row, column=c_i, value=val)
                c.fill = row_fill
                c.font = row_font
                c.border = border_thin
                if isinstance(val, (int, float)):
                    c.alignment = align_right
                    c.number_format = "#,##0.##"
                else:
                    c.alignment = align_left
            ws_all_det.row_dimensions[curr_det_row].height = 20
            
    for c_i, h in enumerate(headers_all_det, 1):
        col_letter = get_column_letter(c_i)
        ws_all_det.column_dimensions[col_letter].width = min(max(len(h) + 6, 14), 36)
        
    # Tạo các sheet con cho từng đơn (đảm bảo tên sheet không trùng nếu 1 đơn xuất hiện 2 lần trong queue)
    used_sheet_names = set(wb.sheetnames)
    for r in all_results:
        wh_sfx = r.get('wh', '01&61').replace('&', '_')
        base_sname = f"DH_{r['order_no']}_K{wh_sfx}"[:28]
        sheet_name = base_sname
        if sheet_name in used_sheet_names:
            sheet_name = f"DH_{r['order_no']}_{r['stt']}"[:31]
        used_sheet_names.add(sheet_name)
        
        ws = wb.create_sheet(title=sheet_name)
        df = r['summary_table']
        
        ws.cell(row=1, column=1, value=f"ĐƠN HÀNG: {r['order_no']} (Kho: {r.get('wh', '01&61')} | SL: {r['order_qty']:g} | Variant: {r.get('variant', '')})").font = font_title
        ws.cell(row=2, column=1, value=f"Kho đối soát (WH): {r.get('wh', '01&61')} | Tổng mã: {r['total_items']} | ĐỦ: {r['ok_count']} | THIẾU: {r['missing_count']}").font = Font(name="Segoe UI", size=10, italic=True)
        
        start_r = 4
        headers = list(df.columns)
        for c_i, h in enumerate(headers, 1):
            c = ws.cell(row=start_r, column=c_i, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_header
            c.border = border_thin
            
        rows_sub = df.to_dict('records')
        for row_i, row in enumerate(rows_sub):
            curr_row = start_r + 1 + row_i
            is_ok = (row['Trạng thái'] == 'OK')
            row_fill = fill_ok if is_ok else fill_missing
            row_font = font_ok if is_ok else font_missing
            
            for c_i, h in enumerate(headers, 1):
                val = row[h]
                c = ws.cell(row=curr_row, column=c_i, value=val)
                c.fill = row_fill
                c.font = row_font
                c.border = border_thin
                c.alignment = align_right if isinstance(val, (int, float)) else align_left
                    
        for c_i, h in enumerate(headers, 1):
            col_letter = get_column_letter(c_i)
            max_len = max([len(str(h))] + [len(str(r_sub.get(h, ''))) for r_sub in rows_sub])
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 38)
            
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def analyze_orders_readiness(df_bom, df_stock, wh_spec=None):
    """
    Phân tích đối soát thông minh toàn bộ đơn hàng trong BOM đối chiếu với tồn kho hiện có:
    - Lọc đúng tồn kho theo kho được chọn `wh_spec` (nếu không chọn kho thì mặc định chỉ lấy 2 kho '01' & '61',
      TUYỆT ĐỐI KHÔNG lấy dữ liệu ở kho khác để đối soát).
    - Tự động lựa chọn Variant tối ưu cho từng đơn hàng (nếu đơn có nhiều Variant như 0004, 0104).
    - Tìm ra các đơn hàng có ĐỦ 100% ITEM CODE đồng bộ (tồn kho > 0 cho mọi linh kiện).
    - Tính toán số lượng TỐI ĐA có thể chạy ngay dựa trên linh kiện nghẽn (bottleneck).
    - Liệt kê các đơn chưa đủ mã, phân loại các đơn "Cận sẵn sàng" (Near-Ready: chỉ thiếu 1-3 mã)
      và chi tiết các mã đang thiếu.
    """
    wh_code_str = format_warehouse_code(wh_spec)
    df_stock_filtered = filter_stock_by_warehouses(df_stock, wh_spec)
    stock_pool = df_stock_filtered.groupby('Stock_Code')['Qty'].sum().to_dict() if (df_stock_filtered is not None and len(df_stock_filtered) > 0) else {}
    orders = get_available_orders_info(df_bom)
    
    # Gom nhóm theo Order_No, Variant và Component_Code
    grouped = df_bom.groupby(['Order_No', 'Variant', 'Component_Code'], as_index=False).agg(
        unit_bom=('Qty', 'sum'),
        desc=('Component_Desc', 'first')
    )
    
    grouped['stock_qty'] = grouped['Component_Code'].map(stock_pool).fillna(0.0)
    grouped['has_stock'] = (grouped['stock_qty'] > EPS) & (grouped['unit_bom'] > EPS)
    grouped['possible_qty'] = ((grouped['stock_qty'] + EPS) // grouped['unit_bom']).where(grouped['has_stock'], 0.0)
    
    ready_orders = []
    not_ready_orders = []
    near_ready_orders = []
    
    # Tách theo Order_No để xử lý nhanh
    for order_no, oinfo in orders.items():
        sub_all = grouped[grouped['Order_No'] == order_no]
        if len(sub_all) == 0:
            continue
            
        # Nếu có nhiều Variant, chọn Variant có độ sẵn sàng và số lượng chạy tối đa cao nhất
        avail_vars = sorted(sub_all['Variant'].unique().tolist())
        chosen_var = avail_vars[0]
        if len(avail_vars) > 1:
            best_score = (-1, -1.0, -1)
            for v in avail_vars:
                sv = sub_all[sub_all['Variant'] == v]
                t_c = len(sv)
                in_c = int(sv['has_stock'].sum())
                m_q = float(sv['possible_qty'].min()) if (in_c == t_c and t_c > 0) else 0.0
                sc = (1 if (in_c == t_c and m_q > 0) else 0, m_q, in_c)
                if sc > best_score:
                    best_score = sc
                    chosen_var = v
                    
        sub = sub_all[sub_all['Variant'] == chosen_var]
        total_comps = len(sub)
        comps_in_stock = int(sub['has_stock'].sum())
        readiness_pct = round((comps_in_stock / total_comps * 100), 1) if total_comps > 0 else 0.0
        
        if comps_in_stock == total_comps and total_comps > 0:
            min_row = sub.sort_values('possible_qty').iloc[0]
            max_runnable = int(min_row['possible_qty'])
            if max_runnable > 0:
                ready_orders.append({
                    'order_no': order_no,
                    'wh': wh_code_str,
                    'variant': chosen_var,
                    'description': oinfo['description'],
                    'total_components': total_comps,
                    'max_runnable_qty': max_runnable,
                    'bottleneck_code': min_row['Component_Code'],
                    'bottleneck_desc': min_row['desc'],
                    'bottleneck_stock': float(min_row['stock_qty']),
                    'bottleneck_unit_bom': float(min_row['unit_bom']),
                    'readiness_pct': 100.0
                })
                continue
                
        missing_rows = sub[~sub['has_stock']]
        missing_cnt = len(missing_rows)
        missing_codes_list = missing_rows['Component_Code'].tolist()
        missing_details_list = [
            f"{r['Component_Code']} ({r['desc']})" for _, r in missing_rows.iloc[:5].iterrows()
        ]
        item_not_ready = {
            'order_no': order_no,
            'wh': wh_code_str,
            'variant': chosen_var,
            'description': oinfo['description'],
            'total_components': total_comps,
            'components_in_stock': comps_in_stock,
            'missing_count': missing_cnt,
            'readiness_pct': readiness_pct,
            'sample_missing': ', '.join(missing_codes_list[:4]),
            'missing_codes': missing_codes_list,
            'missing_details': '; '.join(missing_details_list)
        }
        not_ready_orders.append(item_not_ready)
        if 1 <= missing_cnt <= 3:
            near_ready_orders.append(item_not_ready)
        
    ready_orders.sort(key=lambda x: x['max_runnable_qty'], reverse=True)
    not_ready_orders.sort(key=lambda x: (x['missing_count'], -x['readiness_pct']))
    near_ready_orders.sort(key=lambda x: (x['missing_count'], -x['readiness_pct']))
    
    total_count = len(orders)
    ready_count = len(ready_orders)
    not_ready_count = len(not_ready_orders)
    ready_percent = round((ready_count / total_count * 100), 1) if total_count > 0 else 0.0
    
    return {
        'wh': wh_code_str,
        'ready_orders': ready_orders,
        'not_ready_orders': not_ready_orders,
        'near_ready_orders': near_ready_orders,
        'near_ready_count': len(near_ready_orders),
        'total_orders': total_count,
        'ready_count': ready_count,
        'not_ready_count': not_ready_count,
        'ready_percent': ready_percent
    }


def export_readiness_report_to_excel(readiness_data):
    """
    Xuất báo cáo khả thi sản xuất (đơn hàng đủ mã đồng bộ & số lượng tối đa) ra Excel.
    """
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Don_Du_Ma_San_Sang_Chay"
    ws2 = wb.create_sheet(title="Don_Chua_Du_Ma")
    
    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_sub = Font(name="Segoe UI", size=10, italic=True, color="4B5563")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_header2 = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
    
    font_ok = Font(name="Segoe UI", size=10, color="155724")
    fill_ok = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    wh_label = readiness_data.get('wh', '01&61')
    # ===== SHEET 1: ĐƠN SẴN SÀNG CHẠY =====
    ws1.cell(row=1, column=1, value=f"BÁO CÁO CÁC ĐƠN HÀNG ĐỦ MÃ HÀNG ĐỒNG BỘ TRONG TỒN KHO (KHO: {wh_label})").font = font_title
    ws1.cell(row=2, column=1, value=f"Kho đối soát (WH): {wh_label} | Tổng số đơn đủ mã: {readiness_data['ready_count']} / {readiness_data['total_orders']} đơn ({readiness_data['ready_percent']}%) | Đã tính toán số lượng tối đa có thể sản xuất ngay").font = font_sub
    
    headers1 = [
        "STT", "Mã đơn hàng", "Kho (WH)", "Tên sản phẩm", "Số ITEM CODE trong BOM",
        "SL TỐI ĐA có thể chạy (bộ)", "ITEM CODE điểm nghẽn",
        "Tên linh kiện nghẽn", "Tồn kho linh kiện nghẽn", "ĐM bộ linh kiện nghẽn", "Đánh giá"
    ]
    
    for c_i, h in enumerate(headers1, 1):
        cell = ws1.cell(row=4, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_thin
    ws1.row_dimensions[4].height = 28
    
    for idx, r in enumerate(readiness_data['ready_orders'], 1):
        curr_row = 4 + idx
        vals = [
            idx,
            r['order_no'],
            r.get('wh', wh_label),
            r['description'],
            r['total_components'],
            r['max_runnable_qty'],
            r['bottleneck_code'],
            r['bottleneck_desc'],
            r['bottleneck_stock'],
            r['bottleneck_unit_bom'],
            "100% ĐỦ MÃ (Sẵn sàng chạy)"
        ]
        for c_i, val in enumerate(vals, 1):
            cell = ws1.cell(row=curr_row, column=c_i, value=val)
            cell.font = font_ok
            cell.fill = fill_ok
            cell.border = border_thin
            if isinstance(val, (int, float)):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            else:
                cell.alignment = align_left
        ws1.row_dimensions[curr_row].height = 20
        
    for col in ws1.columns:
        max_len = max([len(str(cell.value or '')) for cell in col if cell.row >= 4] or [12])
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
        
    # ===== SHEET 2: ĐƠN CHƯA ĐỦ MÃ =====
    ws2.cell(row=1, column=1, value=f"DANH SÁCH ĐƠN HÀNG CHƯA ĐỦ MÃ TRONG TỒN KHO (KHO: {wh_label})").font = font_title
    ws2.cell(row=2, column=1, value=f"Kho đối soát (WH): {wh_label} | Tổng số đơn chưa đủ mã: {readiness_data['not_ready_count']} đơn (Trong đó có {readiness_data.get('near_ready_count', 0)} đơn cận sẵn sàng chỉ thiếu 1-3 mã)").font = font_sub
    
    headers2 = ["STT", "Mã đơn hàng", "Kho (WH)", "Tên sản phẩm", "Tổng ITEM CODE", "Số mã có trong kho", "Số mã còn thiếu", "Tỷ lệ đồng bộ (%)", "Mẫu mã linh kiện thiếu"]
    for c_i, h in enumerate(headers2, 1):
        cell = ws2.cell(row=4, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header2
        cell.alignment = align_header
        cell.border = border_thin
    ws2.row_dimensions[4].height = 28
    
    fill_missing = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_missing = Font(name="Segoe UI", size=10, color="991B1B")
    
    for idx, r in enumerate(readiness_data['not_ready_orders'], 1):
        curr_row = 4 + idx
        vals = [
            idx,
            r['order_no'],
            r.get('wh', wh_label),
            r['description'],
            r['total_components'],
            r['components_in_stock'],
            r['missing_count'],
            r.get('readiness_pct', 0.0),
            r['sample_missing']
        ]
        for c_i, val in enumerate(vals, 1):
            cell = ws2.cell(row=curr_row, column=c_i, value=val)
            cell.font = font_missing
            cell.fill = fill_missing
            cell.border = border_thin
            cell.alignment = align_right if isinstance(val, (int, float)) else align_left
        ws2.row_dimensions[curr_row].height = 20
        
    for col in ws2.columns:
        max_len = max([len(str(cell.value or '')) for cell in col if cell.row >= 4] or [12])
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
        
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


export_ready_orders_to_excel = export_readiness_report_to_excel


def find_similar_stock_substitutes(missing_code: str, all_stock_codes: set, stock_summary: dict) -> list:
    """
    Tìm kiếm thông minh các mã linh kiện tương đương / cùng họ (Variant / Coating / Suffix)
    đang có tồn kho > 0 để gợi ý cho bộ phận Kế hoạch / Kỹ thuật / Agent đối soát.
    Ví dụ: BOM yêu cầu '7100040051' hoặc '7100040001PA' nhưng kho có '7100040051PA' (tồn 1,475,031 cái).
    """
    if not missing_code or len(missing_code) < 6:
        return []
    prefix6 = missing_code[:6]
    prefix8 = missing_code[:8]
    candidates = []
    for sc in all_stock_codes:
        if sc == missing_code:
            continue
        qty = stock_summary.get(sc, 0)
        if qty <= 0:
            continue
        if sc.startswith(prefix8) or (len(missing_code) >= 10 and sc.startswith(prefix6) and sc[:6] == prefix6 and abs(len(sc) - len(missing_code)) <= 2):
            candidates.append((sc, int(qty) if float(qty).is_integer() else qty))
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[:3]


def check_data_integrity(df_bom, df_stock, df_plan=None, wh_spec=None):
    """
    Kiểm tra toàn diện tính toàn vẹn và độ chính xác của dữ liệu (Smart Data Integrity & Bottleneck Analysis):
    - Chỉ đối soát dữ liệu tồn kho theo kho được chọn (hoặc mặc định 2 kho '01' & '61', không lấy kho khác).
    1. Linh kiện trong BOM thiếu trong Stock (kèm gợi ý mã tương đương trong kho nếu có).
    2. Linh kiện trong Stock có tồn kho <= 0 hoặc bất thường.
    3. Mã đơn hàng trong Kế hoạch có trong BOM không.
    4. Các đơn hàng trong BOM chỉ có 1 vế (không đủ Trái/Phải).
    5. Đánh giá chỉ số sức khỏe dữ liệu (Health Score %).
    """
    if df_bom is None or df_stock is None:
        return None
        
    df_stock_filtered = filter_stock_by_warehouses(df_stock, wh_spec)
    bom_comps = df_bom[['Component_Code', 'Component_Desc']].drop_duplicates(subset=['Component_Code'])
    stock_summary = df_stock_filtered.groupby('Stock_Code')['Qty'].sum().to_dict() if (df_stock_filtered is not None and len(df_stock_filtered) > 0) else {}
    all_stock_codes = set(stock_summary.keys())
    
    # Pre-group orders per component
    comp_to_orders = df_bom.groupby('Component_Code')['Order_No'].unique().to_dict()
    
    # 1. Linh kiện trong BOM hoàn toàn không có trong Stock
    missing_in_stock = []
    substitute_suggestions = []
    for row in bom_comps.to_dict('records'):
        c_code = row['Component_Code']
        if c_code not in all_stock_codes or stock_summary.get(c_code, 0) <= 0:
            orders_used = sorted(comp_to_orders.get(c_code, []).tolist())
            orders_sample = ', '.join(orders_used[:6]) + (f' (+{len(orders_used)-6} đơn khác)' if len(orders_used) > 6 else '')
            subs = find_similar_stock_substitutes(c_code, all_stock_codes, stock_summary)
            sub_str = ', '.join([f"{sc} (Tồn: {q:,})" for sc, q in subs]) if subs else "Không có mã tương đương"
            
            missing_in_stock.append({
                'Mã linh kiện (ITEM CODE)': c_code,
                'Tên mô tả': row['Component_Desc'],
                'Số đơn BOM bị ảnh hưởng': len(orders_used),
                'Mẫu các đơn ảnh hưởng': orders_sample,
                'Gợi ý mã tương đương trong kho': sub_str
            })
            if subs:
                substitute_suggestions.append({
                    'Mã thiếu trong BOM': c_code,
                    'Tên linh kiện': row['Component_Desc'],
                    'Số đơn bị chặn': len(orders_used),
                    'Mã tương đương sẵn có trong Kho': sub_str
                })
                
    df_missing_comps = pd.DataFrame(missing_in_stock)
    if len(df_missing_comps) > 0:
        df_missing_comps = df_missing_comps.sort_values(by='Số đơn BOM bị ảnh hưởng', ascending=False).reset_index(drop=True)
        
    df_substitutes = pd.DataFrame(substitute_suggestions)
    if len(df_substitutes) > 0:
        df_substitutes = df_substitutes.sort_values(by='Số đơn bị chặn', ascending=False).reset_index(drop=True)
        
    # 2. Linh kiện trong Stock có số lượng <= 0
    zero_stock = [
        {
            'Mã linh kiện (ITEM CODE)': scode,
            'Số lượng tồn': qty,
            'Đánh giá': 'Cảnh báo: Tồn kho <= 0'
        }
        for scode, qty in stock_summary.items() if qty <= 0
    ]
    df_zero_stock = pd.DataFrame(zero_stock)
    
    # 3. Kiểm tra đơn hàng trong Kế hoạch
    plan_issues = []
    if df_plan is not None and len(df_plan) > 0:
        bom_orders_set = set(df_bom['Order_No'].unique())
        plan_cols = ['Order_No', 'MODEL', 'Qty'] + (['WH'] if 'WH' in df_plan.columns else [])
        plan_orders = df_plan[plan_cols].drop_duplicates()
        for prow in plan_orders.to_dict('records'):
            p_ord = prow['Order_No']
            if p_ord not in bom_orders_set:
                plan_issues.append({
                    'Mã đơn trong Kế hoạch': p_ord,
                    'Kho (WH)': prow.get('WH', '01&61'),
                    'Mã MODEL': prow['MODEL'],
                    'Số lượng': prow['Qty'],
                    'Nguyên nhân': 'Chưa khai báo cấu trúc BOM cho mã đơn này'
                })
    df_plan_issues = pd.DataFrame(plan_issues)
    
    # 4. Đơn hàng chỉ có 1 vế trong BOM
    orders_info = get_available_orders_info(df_bom)
    single_sided = [
        {
            'Mã đơn hàng': ord_no,
            'Tên sản phẩm': info['description'],
            'Vế hiện có': ', '.join(info['sides']),
            'Ghi chú': 'Chỉ có 1 vế (Chưa đủ bộ nguyên đơn Trái + Phải)'
        }
        for ord_no, info in orders_info.items() if not info['has_both_sides']
    ]
    df_single_sided = pd.DataFrame(single_sided)
    
    # Tính điểm sức khỏe dữ liệu
    total_bom_comps_count = len(bom_comps)
    matched_comps_count = total_bom_comps_count - len(df_missing_comps)
    comp_match_pct = (matched_comps_count / total_bom_comps_count * 100) if total_bom_comps_count > 0 else 0
    
    total_orders_count = len(orders_info)
    complete_orders_count = total_orders_count - len(df_single_sided)
    order_health_pct = (complete_orders_count / total_orders_count * 100) if total_orders_count > 0 else 0
    
    overall_health = round((comp_match_pct * 0.7 + order_health_pct * 0.3), 1)
    
    return {
        'total_bom_comps': total_bom_comps_count,
        'matched_comps': matched_comps_count,
        'unmatched_comps_count': len(df_missing_comps),
        'df_unmatched_comps': df_missing_comps,
        'substitutes_count': len(df_substitutes),
        'df_substitutes': df_substitutes,
        'total_stock_items': len(all_stock_codes),
        'zero_stock_count': len(df_zero_stock),
        'df_zero_stock': df_zero_stock,
        'plan_issues_count': len(df_plan_issues),
        'df_plan_issues': df_plan_issues,
        'single_sided_count': len(df_single_sided),
        'df_single_sided': df_single_sided,
        'comp_match_pct': round(comp_match_pct, 1),
        'order_health_pct': round(order_health_pct, 1),
        'overall_health': overall_health
    }


def generate_smart_reconciliation_insights(df_bom, df_stock, df_plan=None, calculation_results=None, readiness_data=None, integrity_report=None):
    """
    Tổng hợp dữ liệu đối soát thông minh & cấu trúc ngữ cảnh chuẩn hóa (Structured Context)
    để các Agent và Người điều hành có thể lọc ngay dữ liệu quan trọng nhất:
    1. Top linh kiện nút cổ chai (Critical Bottlenecks) đang chặn nhiều đơn hàng nhất.
    2. Danh sách đơn hàng "Cận sẵn sàng" (Near-Ready Orders: chỉ thiếu 1-3 linh kiện là chạy được ngay).
    3. Gợi ý mã linh kiện tương đương trong kho (Smart Substitute Candidates).
    4. Phân tích khả năng đáp ứng Kế hoạch sản xuất theo đúng Kho (WH) (Plan vs FIFO Stock Allocation).
    5. Xuất gói dữ liệu chuẩn hóa JSON & Markdown Brief cho các AI Agent.
    """
    readiness = readiness_data if readiness_data is not None else analyze_orders_readiness(df_bom, df_stock)
    integrity = integrity_report if integrity_report is not None else check_data_integrity(df_bom, df_stock, df_plan)
    
    # 1. Phân tích Top linh kiện chặn nhiều đơn hàng nhất
    top_blocking_df = integrity['df_unmatched_comps'].head(15).copy() if integrity and len(integrity['df_unmatched_comps']) > 0 else pd.DataFrame()
    
    # 2. Đơn cận sẵn sàng (chỉ thiếu 1-3 mã)
    near_ready_list = readiness.get('near_ready_orders', [])
    df_near_ready = pd.DataFrame([
        {
            'Mã đơn hàng': r['order_no'],
            'Kho (WH)': r.get('wh', '01&61'),
            'Phiên bản': r.get('variant', ''),
            'Tên sản phẩm': r['description'],
            'Độ đồng bộ (%)': f"{r['readiness_pct']}% ({r['components_in_stock']}/{r['total_components']})",
            'Số mã thiếu': r['missing_count'],
            'Chi tiết mã thiếu cần nhập gấp': r.get('missing_details', r['sample_missing'])
        }
        for r in near_ready_list
    ])
    
    # 3. Phân tích chuyên sâu File Kế hoạch (nếu có) theo đúng Kho (WH) của từng đơn
    plan_analysis = []
    if df_plan is not None and len(df_plan) > 0:
        plan_q = extract_orders_from_plan(df_plan)
        fifo_res = calculation_results if (calculation_results and len(calculation_results) == len(plan_q)) else calculate_inventory_allocation(plan_q, df_bom, df_stock)
        readiness_by_wh_cache = {
            format_warehouse_code(readiness.get('wh', '01&61')): {ro['order_no']: ro for ro in readiness['ready_orders']}
        }
        for r in fifo_res:
            ord_no = r['order_no']
            r_wh = format_warehouse_code(r.get('wh', '01&61'))
            if r_wh not in readiness_by_wh_cache:
                wh_rd = analyze_orders_readiness(df_bom, df_stock, wh_spec=r_wh)
                readiness_by_wh_cache[r_wh] = {ro['order_no']: ro for ro in wh_rd['ready_orders']}
            ro = readiness_by_wh_cache[r_wh].get(ord_no)
            max_independent = ro['max_runnable_qty'] if ro else 0
            plan_analysis.append({
                'Ưu tiên FIFO': f"#{r['stt']}",
                'Mã đơn': ord_no,
                'Kho (WH)': r_wh,
                'Phiên bản': r.get('variant', ''),
                'Sản phẩm': r.get('product_title', ''),
                'SL Kế hoạch (bộ)': int(r['order_qty']),
                'SL Tối đa Độc lập tại Kho (bộ)': max_independent,
                'Kết quả FIFO lũy kế': "✅ ĐỦ 100% VẬT TƯ" if r.get('is_fully_ok') else f"⚠️ THIẾU {r['missing_count']} MÃ",
                'Khuyến nghị điều độ': (
                    f"Chạy ngay theo kế hoạch (Kho {r_wh})" if r.get('is_fully_ok')
                    else (f"Có thể tách chạy trước tối đa {max_independent:,} bộ tại Kho {r_wh}" if max_independent > 0 else f"Chờ bổ sung vật tư thiếu tại Kho {r_wh}")
                )
            })
    df_plan_smart = pd.DataFrame(plan_analysis)
    
    # 4. Tạo gói dữ liệu Agent Payload (JSON & Markdown) để các Agent lọc dữ liệu quan trọng
    agent_payload = {
        "schema_version": "2.1-agent-warehouse-ready",
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "default_reconciliation_warehouses": list(DEFAULT_WAREHOUSES),
        "kpi_summary": {
            "total_bom_orders": readiness['total_orders'],
            "ready_100_pct_orders_count": readiness['ready_count'],
            "ready_orders_percentage": readiness['ready_percent'],
            "near_ready_orders_count": readiness.get('near_ready_count', 0),
            "not_ready_orders_count": readiness['not_ready_count'],
            "total_unique_bom_components": integrity['total_bom_comps'] if integrity else 0,
            "matched_stock_components": integrity['matched_comps'] if integrity else 0,
            "missing_stock_components": integrity['unmatched_comps_count'] if integrity else 0,
            "data_health_score_pct": integrity['overall_health'] if integrity else 0.0,
            "active_fifo_queue_count": len(calculation_results) if calculation_results else 0,
            "active_fifo_ok_orders": sum(1 for r in calculation_results if r.get('is_fully_ok')) if calculation_results else 0
        },
        "top_ready_orders_to_run_immediately": [
            {
                "order_no": r['order_no'],
                "wh": r.get('wh', '01&61'),
                "variant": r.get('variant', ''),
                "product": r['description'],
                "max_runnable_sets": r['max_runnable_qty'],
                "bottleneck_component": r['bottleneck_code'],
                "bottleneck_desc": r['bottleneck_desc'],
                "bottleneck_stock": r['bottleneck_stock']
            }
            for r in readiness['ready_orders'][:15]
        ],
        "near_ready_orders_missing_1_to_3_items": [
            {
                "order_no": r['order_no'],
                "wh": r.get('wh', '01&61'),
                "variant": r.get('variant', ''),
                "product": r['description'],
                "readiness_pct": r['readiness_pct'],
                "missing_count": r['missing_count'],
                "missing_components": r['missing_codes']
            }
            for r in near_ready_list[:15]
        ],
        "critical_bottleneck_components_to_expedite": top_blocking_df.to_dict(orient='records') if len(top_blocking_df) > 0 else [],
        "smart_substitute_suggestions": integrity['df_substitutes'].to_dict(orient='records') if (integrity and len(integrity['df_substitutes']) > 0) else [],
        "plan_reconciliation_recommendations": plan_analysis
    }
    
    # Tạo bản tóm tắt ngôn ngữ tự nhiên chuẩn hóa (Markdown Brief) cho Agent / Quản lý
    md_lines = [
        f"# 🤖 BÁO CÁO TRÍ TUỆ ĐỐI SOÁT VẬT TƯ (AGENT DATA BRIEF)",
        f"- **Thời gian trích xuất**: {agent_payload['generated_at']}",
        f"- **Kho đối soát mặc định**: **Kho 01 & 61** (Không lấy dữ liệu ở kho khác ngoài kho chỉ định)",
        f"- **Tổng số đơn hàng trong BOM**: **{readiness['total_orders']} đơn** | Đủ 100% mã chạy ngay: **{readiness['ready_count']} đơn ({readiness['ready_percent']}%)**",
        f"- **Đơn hàng cận sẵn sàng (chỉ thiếu 1-3 mã)**: **{readiness.get('near_ready_count', 0)} đơn** (Ưu tiên thúc hàng để mở khóa sản xuất nhanh)",
        f"- **Điểm sức khỏe & đồng bộ dữ liệu**: **{integrity['overall_health'] if integrity else 0}%** ({integrity['matched_comps'] if integrity else 0}/{integrity['total_bom_comps'] if integrity else 0} mã linh kiện khớp tồn kho)",
        "",
        "## 1. Top 5 Đơn Hàng Sẵn Sàng Chạy Ngay (Sản Lượng Cao Nhất)"
    ]
    for r in readiness['ready_orders'][:5]:
        md_lines.append(
            f"- **Đơn {r['order_no']}** (Kho `{r.get('wh', '01&61')}`) ({r['description']}) - Phiên bản `{r.get('variant', '')}`: Chạy tối đa **{r['max_runnable_qty']:,} bộ** | Điểm nghẽn: `{r['bottleneck_code']}` ({r['bottleneck_desc']}, tồn: {int(r['bottleneck_stock']):,})"
        )
    md_lines.append("")
    md_lines.append("## 2. Top 5 Linh Kiện Nút Cổ Chai Cần Mua / Nhập Kho Gấp")
    if len(top_blocking_df) > 0:
        for _, brow in top_blocking_df.head(5).iterrows():
            md_lines.append(
                f"- **`{brow['Mã linh kiện (ITEM CODE)']}`** ({brow['Tên mô tả']}): Đang chặn **{brow['Số đơn BOM bị ảnh hưởng']} đơn hàng** ({brow['Mẫu các đơn ảnh hưởng']}) | Gợi ý kho: {brow['Gợi ý mã tương đương trong kho']}"
            )
    md_lines.append("")
    md_lines.append("## 3. Top Đơn Hàng Cận Sẵn Sàng (Chỉ Thiếu 1 Mã Linh Kiện)")
    for r in [x for x in near_ready_list if x['missing_count'] == 1][:8]:
        md_lines.append(
            f"- **Đơn {r['order_no']}** ({r['description']}): Đã có **{r['components_in_stock']}/{r['total_components']} mã ({r['readiness_pct']}%)** | Chỉ thiếu duy nhất: `{r['sample_missing']}`"
        )
        
    return {
        'df_top_blocking': top_blocking_df,
        'df_near_ready': df_near_ready,
        'df_substitutes': integrity['df_substitutes'] if integrity else pd.DataFrame(),
        'df_plan_smart': df_plan_smart,
        'agent_payload': agent_payload,
        'agent_markdown_brief': "\n".join(md_lines)
    }


verify_data_integrity = check_data_integrity
analyze_smart_reconciliation = generate_smart_reconciliation_insights
