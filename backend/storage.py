"""资料库存储管理：基于 JSON 的元数据存储 + 文件系统"""
import json
import os
import threading
import time
import uuid

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
LIBRARY_DIR = os.path.join(DATA_DIR, "library")
TEMPLATES_DIR = os.path.join(DATA_DIR, "templates")
OUTPUT_DIR = os.path.join(DATA_DIR, "output")
DB_PATH = os.path.join(DATA_DIR, "db.json")

def ensure_dirs():
    """确保数据目录存在。

    数据目录可能在服务运行期间被外部删除（清理脚本、杀毒软件、手动整理等）。
    如果只在模块导入时建一次，之后写入就会抛 FileNotFoundError，
    而 FastAPI 会把它变成一个没有任何信息量的 500。
    所以每次写入前都兜一次底。
    """
    for d in (DATA_DIR, LIBRARY_DIR, TEMPLATES_DIR, OUTPUT_DIR):
        os.makedirs(d, exist_ok=True)


ensure_dirs()

_lock = threading.Lock()

_EMPTY_DB = {
    "library_files": [],   # 资料文件（文档/表格/PDF/txt）
    "images": [],          # 图片素材
    "entries": {},         # 标准化键 -> {key, value, source}
    "templates": [],       # 模板
    "records": [],         # 生成记录
}


def _load():
    if not os.path.exists(DB_PATH):
        return json.loads(json.dumps(_EMPTY_DB))
    try:
        with open(DB_PATH, "r", encoding="utf-8") as f:
            db = json.load(f)
    except (json.JSONDecodeError, OSError):
        return json.loads(json.dumps(_EMPTY_DB))
    for k, v in _EMPTY_DB.items():
        if k not in db:
            db[k] = json.loads(json.dumps(v)) if isinstance(v, (dict, list)) else v
    return db


def _save(db):
    ensure_dirs()
    with open(DB_PATH, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


def db_read():
    with _lock:
        return _load()


def db_write(fn):
    """线程安全地读写数据库，fn(db) 修改 db 后返回结果"""
    with _lock:
        db = _load()
        result = fn(db)
        _save(db)
        return result


def new_id():
    return uuid.uuid4().hex[:12]


def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def save_upload(directory, filename, content_bytes):
    """保存上传文件，重名自动加序号，返回 (stored_name, path)"""
    ensure_dirs()
    safe = os.path.basename(filename).replace("\\", "/").split("/")[-1]
    base, ext = os.path.splitext(safe)
    stored = safe
    i = 1
    while os.path.exists(os.path.join(directory, stored)):
        stored = f"{base}_{i}{ext}"
        i += 1
    path = os.path.join(directory, stored)
    with open(path, "wb") as f:
        f.write(content_bytes)
    return stored, path


def remove_quiet(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass
