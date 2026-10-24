"""
历史生成记录存储
每次成功生成实验报告后写入记录，供反复下载，避免重复消耗 API token。
"""
import json
import time
from pathlib import Path
from config import OUTPUT_DIR

HISTORY_FILE = OUTPUT_DIR / "_history.json"
MAX_RECORDS = 50


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def load_history() -> list:
    """读取历史记录，最新的在最前"""
    try:
        if HISTORY_FILE.exists():
            data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
    except Exception:
        pass
    return []


def _valid_path(p) -> bool:
    """记录的文件路径是否仍存在"""
    if not p:
        return False
    try:
        return Path(p).exists()
    except Exception:
        return False


def add_record(record: dict) -> list:
    """新增一条记录，返回更新后的完整历史。record 至少含 path/docx 之一"""
    items = load_history()
    docx = record.get("docx") or record.get("path") or ""
    meta = {
        "time": record.get("time", _now()),
        "title": record.get("title") or Path(docx).stem if docx else "未命名",
        "model": record.get("model", ""),
        "docx": docx if _valid_path(docx) else "",
        "pdf": record.get("pdf", "") if _valid_path(record.get("pdf", "")) else "",
        "preview": record.get("preview", ""),
    }
    # 去重：同样的 docx 不重复加
    items = [i for i in items if i.get("docx") != meta["docx"]]
    items.insert(0, meta)
    items = items[:MAX_RECORDS]
    try:
        HISTORY_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return items


def delete_record(docx_path: str) -> list:
    """按 docx 路径删除一条记录（不会删文件本体），返回更新后历史"""
    items = load_history()
    items = [i for i in items if i.get("docx") not in (docx_path, "")]
    # 同时删除同类记录
    items = [i for i in items if Path(str(i.get("docx",""))).resolve() != Path(docx_path).resolve()]
    try:
        HISTORY_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return items


def render_history_html() -> str:
    """把历史渲染成带下载链接的 HTML 列表。
    用 /gradio_api/file= URL 前缀 + 正确编码，保证中文文件名可下载。
    """
    items = [i for i in load_history() if i.get("docx") and _valid_path(i["docx"])]
    if not items:
        return ("<p style='color:#999'>📭 还没有生成记录。生成成功后会自动保存，"
                "以后可随时回来下载，无需再次消耗 Token。</p>")

    try:
        import urllib.parse
        from gradio_client.utils import encode_file_path

        def _file_url(path):
            # 统一为绝对路径（gradio /file= 需要绝对路径才能 serve）
            p = Path(path)
            if not p.is_absolute():
                p = (OUTPUT_DIR / p.name if p.parent == Path(".") else Path.cwd() / p)
            p = p.resolve()
            return f"/gradio_api/file={encode_file_path(str(p).replace(chr(92),'/'))}"
    except Exception:
        def _file_url(path):
            return f"/file={str(path).replace('\\\\', '/')}"

    cards = []
    for i in items:
        url = _file_url(i["docx"])
        portions = [
            "<div style='background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px 14px;margin-bottom:10px'>",
            f"<div style='font-weight:600;color:#1f2937'>📄 {_htmlesc(i.get('title','未命名'))}</div>",
            f"<div style='color:#6b7280;font-size:12px;margin:2px 0 8px'>🕐 {i.get('time','')} · 🤖 {_htmlesc(i.get('model','') or '—')}</div>",
            f"<a href='{url}' download style='display:inline-block;background:#4f46e5;color:white;"
            f"padding:6px 14px;border-radius:6px;text-decoration:none;font-size:13px'>⬇️ 下载 .docx</a>",
            "</div>",
        ]
        cards.append("".join(portions))
    return "".join(cards)


def _htmlesc(s):
    import html
    return html.escape(str(s))