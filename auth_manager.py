"""
auth_manager.py
Hệ thống Quản lý Người dùng, Xác thực & Phân quyền (Role-Based Access Control - RBAC)
và Ghi nhận Nhật ký hoạt động (Audit Logging).

Hỗ trợ 4 vai trò chính:
1. ADMIN (Quản trị viên): Toàn quyền cấu hình hệ thống, quản lý tài khoản, upload và xóa dữ liệu.
2. PLANNER (Kế hoạch & Sản xuất): Quản lý hàng chờ FIFO, nạp file Kế hoạch, đối soát và xuất Excel.
3. WAREHOUSE (Thủ kho / Tồn kho): Cập nhật file Tồn kho, tra cứu tồn và danh mục vật tư thiếu.
4. VIEWER (Người xem / Công nhân): Chỉ xem báo cáo, theo dõi đối soát và tải Excel (Read-only).
"""

import os
import json
import hashlib
import secrets
import datetime
from typing import Dict, List, Optional, Tuple

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PERSIST_DIR = os.path.join(APP_DIR, ".persistent_data")
os.makedirs(DATA_PERSIST_DIR, exist_ok=True)

USERS_FILE = os.path.join(DATA_PERSIST_DIR, "users.json")
AUDIT_LOG_FILE = os.path.join(DATA_PERSIST_DIR, "audit_log.json")

# Danh sách vai trò và định nghĩa quyền hạn
ROLE_ADMIN = "admin"
ROLE_PLANNER = "planner"
ROLE_WAREHOUSE = "warehouse"
ROLE_VIEWER = "viewer"

ROLES_METADATA = {
    ROLE_ADMIN: {
        "name": "Quản trị viên (Admin)",
        "badge_color": "#DC2626",
        "bg_color": "#FEE2E2",
        "description": "Toàn quyền: Quản lý tài khoản, phân quyền, cấu hình hệ thống, upload & xóa dữ liệu",
        "permissions": {
            "can_upload_stock": True,
            "can_upload_bom": True,
            "can_upload_plan": True,
            "can_manage_queue": True,
            "can_export_excel": True,
            "can_manage_users": True,
            "can_reset_data": True,
            "can_view_audit": True,
            "can_switch_workspace": True,
        }
    },
    ROLE_PLANNER: {
        "name": "Kế hoạch & Sản xuất (Planner)",
        "badge_color": "#2563EB",
        "bg_color": "#EFF6FF",
        "description": "Quyền kế hoạch: Nạp kế hoạch, quản lý hàng chờ FIFO, đối soát vật tư & xuất Excel",
        "permissions": {
            "can_upload_stock": False,
            "can_upload_bom": True,
            "can_upload_plan": True,
            "can_manage_queue": True,
            "can_export_excel": True,
            "can_manage_users": False,
            "can_reset_data": False,
            "can_view_audit": True,
            "can_switch_workspace": True,
        }
    },
    ROLE_WAREHOUSE: {
        "name": "Thủ kho & Tồn kho (Warehouse)",
        "badge_color": "#059669",
        "bg_color": "#ECFDF5",
        "description": "Quyền kho: Cập nhật file tồn kho Stock 61, kiểm tra vật tư thiếu & xuất Excel",
        "permissions": {
            "can_upload_stock": True,
            "can_upload_bom": False,
            "can_upload_plan": False,
            "can_manage_queue": False,
            "can_export_excel": True,
            "can_manage_users": False,
            "can_reset_data": False,
            "can_view_audit": False,
            "can_switch_workspace": False,
        }
    },
    ROLE_VIEWER: {
        "name": "Người xem / Báo cáo (Viewer)",
        "badge_color": "#64748B",
        "bg_color": "#F1F5F9",
        "description": "Chỉ xem (Read-only): Xem tình hình sẵn sàng đơn, xem kết quả FIFO & tải Excel",
        "permissions": {
            "can_upload_stock": False,
            "can_upload_bom": False,
            "can_upload_plan": False,
            "can_manage_queue": False,
            "can_export_excel": True,
            "can_manage_users": False,
            "can_reset_data": False,
            "can_view_audit": False,
            "can_switch_workspace": False,
        }
    }
}


def _hash_password(password: str, salt: Optional[str] = None) -> Tuple[str, str]:
    """Tạo salt và băm mật khẩu bằng SHA-256 an toàn."""
    if not salt:
        salt = secrets.token_hex(16)
    hashed = hashlib.sha256((salt + password).encode('utf-8')).hexdigest()
    return hashed, salt


def _verify_password(password: str, hashed: str, salt: str) -> bool:
    """Xác thực mật khẩu với hash và salt đã lưu."""
    calc_hash, _ = _hash_password(password, salt)
    return secrets.compare_digest(calc_hash, hashed)


def _safe_write_json(file_path: str, data) -> bool:
    """Ghi file an toàn atomic thông qua file tạm để tránh corrupt khi có nhiều truy cập cùng lúc."""
    tmp_path = f"{file_path}.tmp"
    try:
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, file_path)
        return True
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return False


def _safe_read_json(file_path: str, default=None):
    """Đọc file JSON an toàn."""
    if default is None:
        default = {}
    if not os.path.exists(file_path):
        return default
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def get_default_users() -> Dict[str, dict]:
    """Danh sách 4 tài khoản mặc định được khởi tạo sẵn sàng sử dụng."""
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    defaults = [
        ("admin", "admin123", "Quản Trị Viên Hệ Thống", ROLE_ADMIN, "Ban Quản Trị / Trưởng Phòng"),
        ("planner", "plan123", "Nguyễn Văn Kế Hoạch", ROLE_PLANNER, "Phòng Kế Hoạch & Điều Độ Sản Xuất"),
        ("warehouse", "kho123", "Trần Thị Thủ Kho", ROLE_WAREHOUSE, "Phòng Quản Lý Kho & Vật Tư"),
        ("viewer", "view123", "Công Nhân / Khách Xem", ROLE_VIEWER, "Bộ Phận Sản Xuất / Khách")
    ]
    
    users = {}
    for username, raw_pw, fullname, role, dept in defaults:
        pwd_hash, salt = _hash_password(raw_pw)
        users[username] = {
            "username": username,
            "password_hash": pwd_hash,
            "salt": salt,
            "full_name": fullname,
            "role": role,
            "department": dept,
            "is_active": True,
            "created_at": now_iso,
            "last_login": None
        }
    return users


def init_auth_system():
    """Khởi tạo file users.json nếu chưa tồn tại."""
    if not os.path.exists(USERS_FILE):
        default_users = get_default_users()
        _safe_write_json(USERS_FILE, default_users)
        log_activity("SYSTEM", "INIT", "Khởi tạo hệ thống quản lý tài khoản & phân quyền")


def get_all_users() -> Dict[str, dict]:
    """Lấy danh sách tất cả người dùng trong hệ thống."""
    init_auth_system()
    users = _safe_read_json(USERS_FILE, default={})
    if not users:
        users = get_default_users()
        _safe_write_json(USERS_FILE, users)
    return users


def authenticate(username: str, password: str) -> Tuple[bool, Optional[dict], str]:
    """
    Xác thực đăng nhập người dùng.
    Trả về: (thành công, thông tin user, thông điệp)
    """
    users = get_all_users()
    uname = username.strip().lower()
    
    if uname not in users:
        return False, None, "Tên đăng nhập không tồn tại trên hệ thống."
        
    user = users[uname]
    if not user.get("is_active", True):
        return False, None, "Tài khoản này đã bị tạm khóa. Vui lòng liên hệ Quản trị viên."
        
    if _verify_password(password, user["password_hash"], user["salt"]):
        # Cập nhật thời gian đăng nhập
        user["last_login"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        users[uname] = user
        _safe_write_json(USERS_FILE, users)
        
        log_activity(uname, "LOGIN", f"Người dùng {user['full_name']} đăng nhập thành công")
        
        # Trả về thông tin an toàn (bỏ password_hash và salt)
        safe_user = {k: v for k, v in user.items() if k not in ("password_hash", "salt")}
        safe_user["permissions"] = ROLES_METADATA.get(user["role"], {}).get("permissions", {})
        return True, safe_user, "Đăng nhập thành công!"
    else:
        log_activity(uname, "LOGIN_FAILED", "Đăng nhập thất bại (sai mật khẩu)")
        return False, None, "Mật khẩu không chính xác. Vui lòng thử lại."


def has_permission(user_dict: Optional[dict], permission_key: str) -> bool:
    """Kiểm tra quyền hạn của người dùng đối với một tính năng."""
    if not user_dict:
        return False
    role = user_dict.get("role", ROLE_VIEWER)
    perms = ROLES_METADATA.get(role, {}).get("permissions", {})
    return bool(perms.get(permission_key, False))


def create_user(
    current_admin_username: str,
    username: str,
    password: str,
    full_name: str,
    role: str,
    department: str = ""
) -> Tuple[bool, str]:
    """Admin tạo tài khoản mới."""
    uname = username.strip().lower()
    if not uname or len(uname) < 3:
        return False, "Tên đăng nhập phải có ít nhất 3 ký tự."
    if not password or len(password) < 4:
        return False, "Mật khẩu phải có ít nhất 4 ký tự."
    if role not in ROLES_METADATA:
        return False, "Vai trò không hợp lệ."
        
    users = get_all_users()
    if uname in users:
        return False, f"Tên đăng nhập '{uname}' đã tồn tại."
        
    pwd_hash, salt = _hash_password(password)
    users[uname] = {
        "username": uname,
        "password_hash": pwd_hash,
        "salt": salt,
        "full_name": full_name.strip() or uname,
        "role": role,
        "department": department.strip(),
        "is_active": True,
        "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "last_login": None
    }
    
    if _safe_write_json(USERS_FILE, users):
        log_activity(current_admin_username, "CREATE_USER", f"Tạo tài khoản mới: {uname} ({role})")
        return True, f"Tạo tài khoản '{uname}' thành công!"
    return False, "Lỗi khi lưu dữ liệu người dùng."


def update_user_role(
    current_admin_username: str,
    target_username: str,
    new_role: str
) -> Tuple[bool, str]:
    """Admin đổi vai trò người dùng."""
    uname = target_username.strip().lower()
    users = get_all_users()
    
    if uname not in users:
        return False, "Người dùng không tồn tại."
    if new_role not in ROLES_METADATA:
        return False, "Vai trò không hợp lệ."
    if uname == "admin" and new_role != ROLE_ADMIN:
        return False, "Không thể hạ quyền của tài khoản Admin gốc."
        
    old_role = users[uname]["role"]
    users[uname]["role"] = new_role
    if _safe_write_json(USERS_FILE, users):
        log_activity(current_admin_username, "UPDATE_ROLE", f"Đổi vai trò {uname}: {old_role} -> {new_role}")
        return True, f"Đã cập nhật vai trò của '{uname}' thành {ROLES_METADATA[new_role]['name']}."
    return False, "Lỗi khi lưu dữ liệu."


def toggle_user_active(
    current_admin_username: str,
    target_username: str
) -> Tuple[bool, str]:
    """Admin khóa hoặc mở khóa tài khoản."""
    uname = target_username.strip().lower()
    users = get_all_users()
    
    if uname not in users:
        return False, "Người dùng không tồn tại."
    if uname == "admin":
        return False, "Không thể khóa tài khoản Admin gốc."
    if uname == current_admin_username.lower():
        return False, "Không thể tự khóa tài khoản của chính mình."
        
    current_status = users[uname].get("is_active", True)
    users[uname]["is_active"] = not current_status
    action_text = "mở khóa" if not current_status else "khóa"
    
    if _safe_write_json(USERS_FILE, users):
        log_activity(current_admin_username, "TOGGLE_USER", f"Đã {action_text} tài khoản {uname}")
        return True, f"Đã {action_text} tài khoản '{uname}' thành công."
    return False, "Lỗi khi lưu dữ liệu."


def reset_user_password(
    current_admin_username: str,
    target_username: str,
    new_password: str
) -> Tuple[bool, str]:
    """Admin đặt lại mật khẩu cho người dùng."""
    uname = target_username.strip().lower()
    if not new_password or len(new_password) < 4:
        return False, "Mật khẩu mới phải có ít nhất 4 ký tự."
        
    users = get_all_users()
    if uname not in users:
        return False, "Người dùng không tồn tại."
        
    pwd_hash, salt = _hash_password(new_password)
    users[uname]["password_hash"] = pwd_hash
    users[uname]["salt"] = salt
    
    if _safe_write_json(USERS_FILE, users):
        log_activity(current_admin_username, "RESET_PASSWORD", f"Đặt lại mật khẩu cho tài khoản {uname}")
        return True, f"Đã đặt lại mật khẩu cho tài khoản '{uname}' thành công."
    return False, "Lỗi khi lưu dữ liệu."


def change_own_password(
    username: str,
    old_password: str,
    new_password: str
) -> Tuple[bool, str]:
    """Người dùng tự đổi mật khẩu của mình."""
    uname = username.strip().lower()
    users = get_all_users()
    
    if uname not in users:
        return False, "Người dùng không tồn tại."
        
    user = users[uname]
    if not _verify_password(old_password, user["password_hash"], user["salt"]):
        return False, "Mật khẩu hiện tại không đúng."
        
    if not new_password or len(new_password) < 4:
        return False, "Mật khẩu mới phải có ít nhất 4 ký tự."
        
    pwd_hash, salt = _hash_password(new_password)
    user["password_hash"] = pwd_hash
    user["salt"] = salt
    users[uname] = user
    
    if _safe_write_json(USERS_FILE, users):
        log_activity(uname, "CHANGE_PASSWORD", "Người dùng tự đổi mật khẩu thành công")
        return True, "Đổi mật khẩu thành công!"
    return False, "Lỗi khi lưu dữ liệu."


# ==================== NHẬT KÝ HOẠT ĐỘNG (AUDIT LOGGING) ====================

def log_activity(username: str, action: str, details: str):
    """Ghi nhận một thao tác vào sổ nhật ký hoạt động của hệ thống."""
    logs = _safe_read_json(AUDIT_LOG_FILE, default=[])
    if not isinstance(logs, list):
        logs = []
        
    log_entry = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "username": username or "ANONYMOUS",
        "action": action,
        "details": details
    }
    logs.insert(0, log_entry)  # Đưa log mới nhất lên đầu
    # Giữ tối đa 500 logs gần nhất để đảm bảo hiệu năng
    logs = logs[:500]
    _safe_write_json(AUDIT_LOG_FILE, logs)


def get_audit_logs(limit: int = 100) -> List[dict]:
    """Lấy danh sách nhật ký hoạt động gần nhất."""
    logs = _safe_read_json(AUDIT_LOG_FILE, default=[])
    if not isinstance(logs, list):
        return []
    return logs[:limit]


def clear_audit_logs(admin_username: str) -> bool:
    """Xóa nhật ký (chỉ Admin)."""
    success = _safe_write_json(AUDIT_LOG_FILE, [])
    if success:
        log_activity(admin_username, "CLEAR_LOGS", "Đã dọn sạch nhật ký hoạt động hệ thống")
    return success
