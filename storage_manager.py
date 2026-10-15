"""
storage_manager.py
Quản lý lưu trữ kiên cố, an toàn giao dịch file (Atomic Operations)
và Quản lý Hàng chờ Đa Người Dùng (Multi-User Workspaces).

Giải quyết triệt để vấn đề xung đột khi chia sẻ cho nhiều người:
- Không gian Cá Nhân (Private Workspace): Mỗi người dùng có một hàng chờ đối soát riêng biệt.
- Không gian Đội Nhóm (Shared Team Workspace): Hàng chờ đồng bộ chung cho toàn nhà máy / phòng kế hoạch.
"""

import os
import json
import shutil
import datetime
from typing import List, Dict, Tuple, Optional

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PERSIST_DIR = os.path.join(APP_DIR, ".persistent_data")
USER_QUEUES_DIR = os.path.join(DATA_PERSIST_DIR, "user_queues")
BACKUP_DIR = os.path.join(DATA_PERSIST_DIR, "backups")

os.makedirs(DATA_PERSIST_DIR, exist_ok=True)
os.makedirs(USER_QUEUES_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

SHARED_QUEUE_FILE = os.path.join(DATA_PERSIST_DIR, "orders_queue.json")
SHARED_META_FILE = os.path.join(DATA_PERSIST_DIR, "metadata.json")

WORKSPACE_PRIVATE = "private"
WORKSPACE_SHARED = "shared"


def safe_write_json(file_path: str, data) -> bool:
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


def safe_read_json(file_path: str, default=None):
    """Đọc file JSON an toàn, fallback về giá trị mặc định nếu có lỗi."""
    if default is None:
        default = {}
    if not os.path.exists(file_path):
        return default
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def get_user_queue_file(username: str) -> str:
    """Đường dẫn file hàng chờ riêng của từng người dùng."""
    clean_name = "".join([c for c in username if c.isalnum() or c in ('_', '-')]).lower()
    if not clean_name:
        clean_name = "guest"
    return os.path.join(USER_QUEUES_DIR, f"{clean_name}_queue.json")


def load_queue(username: str, workspace_mode: str = WORKSPACE_PRIVATE) -> List[dict]:
    """
    Tải hàng chờ tính toán:
    - Nếu mode == 'private': Tải hàng chờ cá nhân của người dùng.
    - Nếu mode == 'shared': Tải hàng chờ chung của nhà máy / đội nhóm.
    """
    if workspace_mode == WORKSPACE_SHARED:
        data = safe_read_json(SHARED_QUEUE_FILE, default=[])
        return data if isinstance(data, list) else []
    else:
        user_file = get_user_queue_file(username)
        data = safe_read_json(user_file, default=None)
        if data is None:
            # Nếu người dùng mới chưa có hàng chờ riêng, nạp từ hàng chờ chung làm mẫu ban đầu
            shared_data = safe_read_json(SHARED_QUEUE_FILE, default=[])
            if isinstance(shared_data, list) and len(shared_data) > 0:
                safe_write_json(user_file, shared_data)
                return shared_data
            return []
        return data if isinstance(data, list) else []


def save_queue(queue: List[dict], username: str, workspace_mode: str = WORKSPACE_PRIVATE) -> bool:
    """
    Lưu hàng chờ tính toán:
    - Nếu mode == 'private': Lưu vào file cá nhân của người dùng.
    - Nếu mode == 'shared': Lưu vào file chung và cập nhật metadata ai là người sửa đổi cuối.
    """
    if workspace_mode == WORKSPACE_SHARED:
        success = safe_write_json(SHARED_QUEUE_FILE, queue)
        if success:
            update_shared_meta({
                "queue_last_updated_by": username,
                "queue_last_updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "queue_item_count": len(queue)
            })
        return success
    else:
        user_file = get_user_queue_file(username)
        return safe_write_json(user_file, queue)


def copy_queue(from_mode: str, to_mode: str, username: str) -> Tuple[bool, str]:
    """
    Sao chép hàng chờ giữa Không gian Cá nhân và Không gian Chung:
    - 'private_to_shared': Xuất bản hàng chờ cá nhân thành hàng chờ chung của nhóm.
    - 'shared_to_private': Đồng bộ hàng chờ chung về làm việc riêng.
    """
    if from_mode == WORKSPACE_PRIVATE and to_mode == WORKSPACE_SHARED:
        user_q = load_queue(username, WORKSPACE_PRIVATE)
        success = save_queue(user_q, username, WORKSPACE_SHARED)
        if success:
            return True, f"Đã xuất bản thành công {len(user_q)} đơn hàng từ Không gian Cá nhân lên Không gian Chung!"
        return False, "Không thể lưu dữ liệu hàng chờ chung."
        
    elif from_mode == WORKSPACE_SHARED and to_mode == WORKSPACE_PRIVATE:
        shared_q = load_queue(username, WORKSPACE_SHARED)
        success = save_queue(shared_q, username, WORKSPACE_PRIVATE)
        if success:
            return True, f"Đã sao chép thành công {len(shared_q)} đơn hàng từ Không gian Chung về Không gian Cá nhân!"
        return False, "Không thể lưu dữ liệu hàng chờ cá nhân."
        
    return False, "Thao tác sao chép không hợp lệ."


def load_shared_meta() -> dict:
    """Tải siêu dữ liệu dùng chung (tên file upload, thời gian nạp,...)."""
    return safe_read_json(SHARED_META_FILE, default={})


def update_shared_meta(new_meta: dict) -> bool:
    """Cập nhật siêu dữ liệu dùng chung."""
    curr = load_shared_meta()
    curr.update(new_meta)
    return safe_write_json(SHARED_META_FILE, curr)


def create_data_backup(label: str = "auto") -> Optional[str]:
    """Tạo bản sao lưu dữ liệu hệ thống đề phòng sự cố."""
    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_subdir = os.path.join(BACKUP_DIR, f"backup_{label}_{now_str}")
    try:
        os.makedirs(backup_subdir, exist_ok=True)
        for fname in os.listdir(DATA_PERSIST_DIR):
            fpath = os.path.join(DATA_PERSIST_DIR, fname)
            if os.path.isfile(fpath) and not fname.endswith('.tmp'):
                shutil.copy2(fpath, os.path.join(backup_subdir, fname))
        return backup_subdir
    except Exception:
        return None
