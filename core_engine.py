"""
core_engine.py
Xử lý logic đối soát BOM (Production BOM List) và Tồn kho (Stock Balance).
- Nhận diện Product Code:
    + Bắt đầu bằng 1: 4 số đầu là mã đơn hàng, 2 số tiếp theo (19: Trái, 20: Phải).
    + Bắt đầu bằng 8: bỏ số 8, 4 số tiếp theo là mã đơn hàng, số thứ 6 (1: Trái, 2: Phải).
- Cột Component Code đối soát với Stock Code trong Stock Balance.
- Quản lý trừ kho theo cơ chế FIFO theo thứ tự đơn hàng nhập vào hàng chờ.
- Xuất file Excel định dạng chuyên nghiệp với tô màu xanh dòng OK, đỏ dòng THIẾU.
"""

import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def find_header_row_from_df(df_raw, keywords, max_rows=20):
    """Tìm chỉ số dòng chứa các từ khóa tiêu đề."""
    for i in range(min(max_rows, len(df_raw))):
        row_str = " ".join([str(x).lower().strip() for x in df_raw.iloc[i].values if pd.notna(x)])
        if all(kw.lower() in row_str for kw in keywords):
            return i
    return 0


def load_bom_data(file_source):
    """
    Đọc dữ liệu Production BOM từ đường dẫn file hoặc file-like object (uploaded file).
    Tự động nhận diện dòng tiêu đề và chuẩn hóa cột.
    """
    df_raw = pd.read_excel(file_source, header=None, nrows=20)
    h_idx = find_header_row_from_df(df_raw, ['product code', 'component code'])
    
    if hasattr(file_source, 'seek'):
        file_source.seek(0)
        
    df = pd.read_excel(file_source, skiprows=h_idx)
    
    col_map = {}
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if 'product code' in c_clean:
            col_map[c] = 'Product_Code'
        elif 'production desc' in c_clean or 'product desc' in c_clean:
            if 'Product_Desc' not in col_map.values():
                col_map[c] = 'Product_Desc'
        elif 'component code' in c_clean:
            col_map[c] = 'Component_Code'
        elif 'component desc' in c_clean:
            if 'Component_Desc' not in col_map.values():
                col_map[c] = 'Component_Desc'
        elif c_clean == 'qty' or 'qty' in c_clean:
            if 'Qty' not in col_map.values():
                col_map[c] = 'Qty'
                
    df = df.rename(columns=col_map)
    df = df[pd.notna(df['Product_Code']) & pd.notna(df['Component_Code'])].copy()
    
    df['Product_Code'] = df['Product_Code'].astype(str).str.strip()
    df['Component_Code'] = df['Component_Code'].astype(str).str.strip()
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    
    if 'Product_Desc' not in df.columns:
        df['Product_Desc'] = ''
    else:
        df['Product_Desc'] = df['Product_Desc'].fillna('').astype(str).str.strip()
        
    if 'Component_Desc' not in df.columns:
        df['Component_Desc'] = ''
    else:
        df['Component_Desc'] = df['Component_Desc'].fillna('').astype(str).str.strip()
        
    def parse_product_code(code):
        """
        - Bắt đầu bằng 1: 4 số đầu là đơn hàng, 2 số tiếp theo là vế (19: Trái, 20: Phải)
        - Bắt đầu bằng 8: bỏ 8, 4 số tiếp theo là đơn hàng, số tiếp theo là vế (1: Trái, 2: Phải)
        """
        code_str = str(code).strip()
        if code_str.startswith('1'):
            order = code_str[:4]
            side_raw = code_str[4:6]
            side = 'Trái' if side_raw == '19' else ('Phải' if side_raw == '20' else f'Vế {side_raw}')
            variant = code_str[6:]
            return order, side, variant
        elif code_str.startswith('8'):
            order = code_str[1:5]
            side_raw = code_str[5] if len(code_str) > 5 else ''
            side = 'Trái' if side_raw == '1' else ('Phải' if side_raw == '2' else f'Vế {side_raw}')
            variant = code_str[6:] if len(code_str) > 6 else ''
            return order, side, variant
        else:
            return code_str[:4], 'Khác', code_str[4:]

    parsed = [parse_product_code(c) for c in df['Product_Code']]
    df['Order_No'] = [p[0] for p in parsed]
    df['Side'] = [p[1] for p in parsed]
    df['Variant'] = [p[2] for p in parsed]
    
    return df


def load_stock_data(file_source):
    """
    Đọc dữ liệu Stock Balance từ đường dẫn file hoặc file-like object.
    Tự động nhận diện dòng tiêu đề và chuẩn hóa cột.
    """
    df_raw = pd.read_excel(file_source, header=None, nrows=20)
    h_idx = find_header_row_from_df(df_raw, ['stock code', 'qty'])
    
    if hasattr(file_source, 'seek'):
        file_source.seek(0)
        
    df = pd.read_excel(file_source, skiprows=h_idx)
    
    col_map = {}
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if 'stock code' in c_clean:
            col_map[c] = 'Stock_Code'
        elif 'batch' in c_clean:
            col_map[c] = 'BATCH'
        elif 'bin' in c_clean:
            col_map[c] = 'BIN'
        elif c_clean == 'qty' or 'qty' in c_clean:
            if 'Qty' not in col_map.values():
                col_map[c] = 'Qty'
        elif 'warehouse' in c_clean:
            col_map[c] = 'Warehouse'
        elif 'date' in c_clean:
            col_map[c] = 'Date'
            
    df = df.rename(columns=col_map)
    df = df[pd.notna(df['Stock_Code'])].copy()
    
    df['Stock_Code'] = df['Stock_Code'].astype(str).str.strip()
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    
    for c in ['BATCH', 'BIN', 'Warehouse']:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str).str.strip()
        else:
            df[c] = ''
            
    return df


def get_available_orders_info(df_bom):
    """
    Trích xuất danh sách tất cả các đơn hàng có trong BOM cùng thông tin tóm tắt:
    tên sản phẩm, số biến thể, số linh kiện, có đủ trái phải không.
    """
    orders_info = {}
    grouped = df_bom.groupby('Order_No')
    
    for order_no, group in grouped:
        prods = group[['Product_Code', 'Product_Desc', 'Side', 'Variant']].drop_duplicates()
        sides = sorted(group['Side'].unique())
        has_both_sides = ('Trái' in sides) and ('Phải' in sides)
        
        # Lấy mô tả sản phẩm đại diện
        desc_list = [d for d in group['Product_Desc'].unique() if d and d != 'nan']
        main_desc = desc_list[0] if desc_list else f"Sản phẩm đơn hàng {order_no}"
        
        variants = sorted(group['Variant'].unique())
        unique_components = len(group['Component_Code'].unique())
        total_bom_rows = len(group)
        
        orders_info[order_no] = {
            'order_no': order_no,
            'description': main_desc,
            'sides': sides,
            'has_both_sides': has_both_sides,
            'variants': variants,
            'unique_components': unique_components,
            'total_bom_rows': total_bom_rows,
            'products': prods.to_dict(orient='records')
        }
        
    return orders_info


def calculate_inventory_allocation(orders_queue, df_bom, df_stock):
    """
    Tính toán đối soát BOM và tồn kho theo cơ chế FIFO cho danh sách đơn hàng chờ.
    Mỗi đơn hàng được tính toán:
    - Lấy đầy đủ định mức Trái và Phải từ BOM.
    - Nhân số lượng đơn hàng với định mức BOM.
    - So sánh với tồn kho khả dụng hiện tại.
    - Ưu tiên cấp liệu cho đơn hàng nhập trước, trừ dần tồn kho.
    """
    # Khởi tạo quỹ tồn kho độc lập để trừ dần
    stock_qty_pool = df_stock.groupby('Stock_Code')['Qty'].sum().to_dict()
    
    # Gom thông tin BIN và BATCH theo từng Stock_Code
    stock_meta = {}
    for code, group in df_stock.groupby('Stock_Code'):
        bins = sorted(set(b for b in group['BIN'].unique() if b and b != 'nan'))
        batches = sorted(set(b for b in group['BATCH'].unique() if b and b != 'nan'))
        
        bin_breakdown = []
        for b, bg in group.groupby('BIN'):
            if b and b != 'nan':
                q = bg['Qty'].sum()
                bin_breakdown.append(f"{b} ({int(q):,})")
                
        stock_meta[code] = {
            'total_initial': group['Qty'].sum(),
            'bins': ", ".join(bins),
            'batches': ", ".join(batches),
            'bin_detail': ", ".join(bin_breakdown)
        }
        
    results = []
    
    for idx, item in enumerate(orders_queue, 1):
        order_no = str(item['order_no']).strip()
        order_qty = float(item['qty'])
        variant = item.get('variant', None)
        item_id = item.get('id', f"ord_{idx}")
        created_at = item.get('created_at', '')
        
        # Lọc dữ liệu BOM cho đơn hàng này
        mask = df_bom['Order_No'] == order_no
        if variant:
            mask = mask & (df_bom['Variant'] == variant)
        bom_order = df_bom[mask].copy()
        
        if len(bom_order) == 0:
            results.append({
                'id': item_id,
                'stt': idx,
                'order_no': order_no,
                'order_qty': order_qty,
                'created_at': created_at,
                'error': f"Không tìm thấy dữ liệu BOM cho đơn hàng {order_no}",
                'total_items': 0,
                'ok_count': 0,
                'missing_count': 0,
                'summary_table': pd.DataFrame(),
                'detail_table': pd.DataFrame()
            })
            continue
            
        # Lấy tên sản phẩm đại diện
        desc_list = [d for d in bom_order['Product_Desc'].unique() if d and d != 'nan']
        product_title = desc_list[0] if desc_list else f"Đơn hàng {order_no}"
        
        # Gom nhóm theo từng linh kiện (Component Code)
        comp_groups = bom_order.groupby('Component_Code')
        summary_rows = []
        
        for comp_code, cgroup in comp_groups:
            comp_desc = cgroup['Component_Desc'].iloc[0]
            
            # Định mức bên Trái
            left_rows = cgroup[cgroup['Side'] == 'Trái']
            left_bom_qty = left_rows['Qty'].sum() if len(left_rows) > 0 else 0.0
            
            # Định mức bên Phải
            right_rows = cgroup[cgroup['Side'] == 'Phải']
            right_bom_qty = right_rows['Qty'].sum() if len(right_rows) > 0 else 0.0
            
            # Nếu có dòng khác ngoài Trái/Phải
            other_rows = cgroup[~cgroup['Side'].isin(['Trái', 'Phải'])]
            other_bom_qty = other_rows['Qty'].sum() if len(other_rows) > 0 else 0.0
            
            total_bom_unit = left_bom_qty + right_bom_qty + other_bom_qty
            total_required = total_bom_unit * order_qty
            
            # Tồn kho khả dụng trước khi trừ
            curr_stock = stock_qty_pool.get(comp_code, 0.0)
            
            # Đối soát số lượng
            if curr_stock >= total_required:
                allocated = total_required
                shortage = 0.0
                status = "OK"
                note = "OK"
                stock_qty_pool[comp_code] = curr_stock - total_required
            else:
                allocated = max(0.0, curr_stock)
                shortage = total_required - allocated
                status = "THIẾU"
                note = f"Thiếu {int(shortage) if shortage.is_integer() else shortage:g}"
                stock_qty_pool[comp_code] = 0.0
                
            info = stock_meta.get(comp_code, {'bins': '', 'batches': '', 'bin_detail': ''})
            
            summary_rows.append({
                'Mã linh kiện': comp_code,
                'Tên linh kiện': comp_desc,
                'Định mức Trái': left_bom_qty,
                'Định mức Phải': right_bom_qty,
                'Tổng định mức (1 bộ)': total_bom_unit,
                'Số lượng đơn': order_qty,
                'Số lượng cần': total_required,
                'Tồn kho trước trừ': curr_stock,
                'Số lượng cấp': allocated,
                'Số lượng thiếu': shortage,
                'Trạng thái': status,
                'Ghi chú': note,
                'Vị trí BIN': info['bins'],
                'Lô BATCH': info['batches'],
                'Chi tiết BIN tồn': info['bin_detail']
            })
            
        df_summary = pd.DataFrame(summary_rows)
        # Sắp xếp ưu tiên hiển thị các mã THIẾU lên đầu để dễ nhận biết
        if len(df_summary) > 0:
            df_summary = df_summary.sort_values(
                by=['Trạng thái', 'Số lượng thiếu', 'Mã linh kiện'],
                ascending=[False, False, True]
            ).reset_index(drop=True)
            
        # Tạo bảng chi tiết theo từng dòng của BOM (Trái & Phải)
        detail_rows = []
        summary_map = df_summary.set_index('Mã linh kiện').to_dict(orient='index') if len(df_summary) > 0 else {}
        
        for _, brow in bom_order.iterrows():
            ccode = brow['Component_Code']
            comp_info = summary_map.get(ccode, {
                'Trạng thái': 'THIẾU',
                'Ghi chú': f"Thiếu {brow['Qty'] * order_qty:g}",
                'Vị trí BIN': '',
                'Lô BATCH': '',
                'Tồn kho trước trừ': 0.0
            })
            
            row_required = brow['Qty'] * order_qty
            detail_rows.append({
                'Vế': brow['Side'],
                'Mã sản phẩm (Product Code)': brow['Product_Code'],
                'Mã linh kiện (Component Code)': ccode,
                'Tên linh kiện': brow['Component_Desc'],
                'Định mức BOM': brow['Qty'],
                'Số lượng đơn': order_qty,
                'Số lượng cần': row_required,
                'Tồn kho khả dụng': comp_info.get('Tồn kho trước trừ', 0.0),
                'Trạng thái': comp_info.get('Trạng thái', 'THIẾU'),
                'Ghi chú': comp_info.get('Ghi chú', 'THIẾU'),
                'Vị trí BIN': comp_info.get('Vị trí BIN', ''),
                'Lô BATCH': comp_info.get('Lô BATCH', '')
            })
            
        df_detail = pd.DataFrame(detail_rows)
        if len(df_detail) > 0:
            df_detail = df_detail.sort_values(
                by=['Vế', 'Trạng thái', 'Mã linh kiện (Component Code)'],
                ascending=[True, False, True]
            ).reset_index(drop=True)
            
        total_items = len(df_summary)
        ok_count = int((df_summary['Trạng thái'] == 'OK').sum()) if total_items > 0 else 0
        missing_count = int((df_summary['Trạng thái'] == 'THIẾU').sum()) if total_items > 0 else 0
        
        results.append({
            'id': item_id,
            'stt': idx,
            'order_no': order_no,
            'order_qty': order_qty,
            'product_title': product_title,
            'created_at': created_at,
            'total_items': total_items,
            'ok_count': ok_count,
            'missing_count': missing_count,
            'is_fully_ok': (missing_count == 0 and total_items > 0),
            'summary_table': df_summary,
            'detail_table': df_detail
        })
        
    return results


def export_order_to_excel(order_result):
    """
    Xuất kết quả đối soát của một đơn hàng ra file Excel .xlsx với định dạng chuyên nghiệp:
    - Sheet 1: Tong_Hop_Vat_Tu (Tổng hợp theo mã linh kiện)
    - Sheet 2: Chi_Tiet_BOM_Trai_Phai (Chi tiết theo từng dòng BOM Trái/Phải)
    - Dòng OK: Tô màu xanh lá nhạt (#D4EDDA), chữ xanh đậm (#155724)
    - Dòng THIẾU: Tô màu đỏ nhạt (#F8D7DA), chữ đỏ đậm (#721C24)
    - Tiêu đề header màu xanh dương sang trọng, border thanh mảnh, auto fit độ rộng cột.
    """
    wb = openpyxl.Workbook()
    
    # Sheet 1: Tổng hợp
    ws1 = wb.active
    ws1.title = "Tong_Hop_Vat_Tu"
    
    # Sheet 2: Chi tiết
    ws2 = wb.create_sheet(title="Chi_Tiet_BOM_Trai_Phai")
    
    # Định nghĩa Styles
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
    
    # ===== SHEET 1 =====
    df1 = order_result['summary_table']
    order_no = order_result['order_no']
    order_qty = order_result['order_qty']
    p_title = order_result.get('product_title', '')
    
    ws1.cell(row=1, column=1, value=f"KẾT QUẢ ĐỐI SOÁT TỒN KHO - ĐƠN HÀNG: {order_no} ({p_title})").font = font_title
    ws1.cell(row=2, column=1, value=f"Số lượng sản xuất: {order_qty:g} bộ (Trái + Phải) | Tổng linh kiện: {order_result['total_items']} | ĐỦ: {order_result['ok_count']} | THIẾU: {order_result['missing_count']}").font = font_sub
    
    start_r1 = 4
    headers1 = list(df1.columns)
    for col_idx, col_name in enumerate(headers1, 1):
        cell = ws1.cell(row=start_r1, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
    ws1.row_dimensions[start_r1].height = 28
    
    for r_idx, row in df1.iterrows():
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
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.##"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws1.row_dimensions[curr_row].height = 20
        
    for col in ws1.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < start_r1:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws1.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
        
    # ===== SHEET 2 =====
    df2 = order_result['detail_table']
    ws2.cell(row=1, column=1, value=f"CHI TIẾT BOM TRÁI & PHẢI - ĐƠN HÀNG: {order_no}").font = font_title
    ws2.cell(row=2, column=1, value=f"Số lượng đơn hàng: {order_qty:g} | Định mức cho từng vế nhân với số lượng đơn").font = font_sub
    
    start_r2 = 4
    headers2 = list(df2.columns)
    for col_idx, col_name in enumerate(headers2, 1):
        cell = ws2.cell(row=start_r2, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
    ws2.row_dimensions[start_r2].height = 28
    
    for r_idx, row in df2.iterrows():
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
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.##"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws2.row_dimensions[curr_row].height = 20
        
    for col in ws2.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < start_r2:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
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
    # Sheet index tóm tắt
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
    
    ws_idx.cell(row=1, column=1, value="BẢNG TỔNG KẾT ĐỐI SOÁT VẬT TƯ CÁC ĐƠN HÀNG (FIFO)").font = font_title
    headers_idx = ["STT", "Mã đơn hàng", "Sản phẩm", "Số lượng đặt", "Tổng số mã LK", "Số mã ĐỦ", "Số mã THIẾU", "Đánh giá chung"]
    
    for c_i, h in enumerate(headers_idx, 1):
        cell = ws_idx.cell(row=3, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
    ws_idx.row_dimensions[3].height = 26
    
    for r_i, r in enumerate(all_results, 1):
        curr_r = 3 + r_i
        is_ok = r.get('is_fully_ok', False)
        row_fill = fill_ok if is_ok else fill_missing
        row_font = font_ok if is_ok else font_missing
        
        vals = [
            r['stt'],
            r['order_no'],
            r.get('product_title', ''),
            r['order_qty'],
            r['total_items'],
            r['ok_count'],
            r['missing_count'],
            "ĐỦ VẬT TƯ (OK)" if is_ok else f"THIẾU {r['missing_count']} MÃ"
        ]
        
        for c_i, val in enumerate(vals, 1):
            cell = ws_idx.cell(row=curr_r, column=c_i, value=val)
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border_thin
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws_idx.row_dimensions[curr_r].height = 20
        
    for col in ws_idx.columns:
        max_len = max([len(str(cell.value or '')) for cell in col])
        col_letter = get_column_letter(col[0].column)
        ws_idx.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    # Tạo các sheet con cho từng đơn
    for r in all_results:
        sheet_name = f"DH_{r['order_no']}"[:31]
        ws = wb.create_sheet(title=sheet_name)
        df = r['summary_table']
        
        ws.cell(row=1, column=1, value=f"ĐƠN HÀNG: {r['order_no']} (SL: {r['order_qty']})").font = font_title
        ws.cell(row=2, column=1, value=f"Tổng mã: {r['total_items']} | ĐỦ: {r['ok_count']} | THIẾU: {r['missing_count']}").font = Font(name="Segoe UI", size=10, italic=True)
        
        start_r = 4
        headers = list(df.columns)
        for c_i, h in enumerate(headers, 1):
            c = ws.cell(row=start_r, column=c_i, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_thin
            
        for row_i, row in df.iterrows():
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
                if isinstance(val, (int, float)):
                    c.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row < start_r:
                    continue
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)
            
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
