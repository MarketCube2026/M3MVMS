from __future__ import annotations

import json
import mimetypes
import os
import posixpath
import re
from datetime import date, datetime, time, timedelta
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote, urlparse

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


BASE_DIR = Path(__file__).resolve().parent
SM_DIR = BASE_DIR / "SM"
MASTER_EXCEL = SM_DIR / "汇总表-市场部市场活动计划汇总.xlsx"
HOST = "127.0.0.1"
PORT = int(os.environ.get("MEETING_DASHBOARD_PORT", "8765"))
LOG_FILE = BASE_DIR / "server-runtime.log"

ATTACHMENT_EXTENSIONS = {
    ".doc",
    ".docx",
    ".pdf",
    ".ppt",
    ".pptx",
    ".xls",
    ".xlsx",
    ".csv",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".txt",
    ".md",
}


def normalize_header(value: Any) -> str:
    return str(value or "").strip().replace("\n", "")


def clean_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def excel_date_to_iso(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return from_excel(value).date().isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    match = re.search(r"(20\d{2})[.\-/年]?\s*(\d{1,2})[.\-/月]?\s*(\d{1,2})", text)
    if match:
        year, month, day = map(int, match.groups())
        return date(year, month, day).isoformat()
    return text


def infer_date_from_text(text: str) -> str:
    match = re.search(r"(20\d{2})[.\-/年]?\s*(\d{1,2})[.\-/月]?\s*(\d{1,2})", text or "")
    if not match:
        return ""
    year, month, day = map(int, match.groups())
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return ""


def date_key(iso_date: str) -> str:
    match = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso_date or "")
    return "".join(match.groups()) if match else ""


def split_people(value: str) -> list[str]:
    if not value:
        return []
    parts = re.split(r"[、,，;/；\s]+", value)
    return [part for part in (p.strip() for p in parts) if part]


def safe_relative_path(path: Path) -> str:
    relative = path.resolve().relative_to(SM_DIR.resolve())
    return relative.as_posix()


def file_url(path: Path) -> str:
    return "/files/" + quote(safe_relative_path(path), safe="/")


def folder_mtime(path: Path) -> float:
    latest = path.stat().st_mtime if path.exists() else 0
    if path.exists():
        for item in path.rglob("*"):
            try:
                latest = max(latest, item.stat().st_mtime)
            except OSError:
                continue
    return latest


def scan_folders() -> list[Path]:
    if not SM_DIR.exists():
        return []
    return [item for item in SM_DIR.iterdir() if item.is_dir()]


def score_folder(folder: Path, event_date_key: str, title: str) -> int:
    name = folder.name
    score = 0
    if event_date_key and event_date_key in name:
        score += 100
    title_tokens = [token for token in re.split(r"[\s\-（）()_/]+", title) if len(token) >= 2]
    for token in title_tokens:
        if token in name:
            score += min(len(token), 12)
    return score


def match_folder(event_date_key: str, title: str, folders: list[Path]) -> Path | None:
    candidates = sorted(((score_folder(folder, event_date_key, title), folder) for folder in folders), reverse=True)
    if candidates and candidates[0][0] >= 20:
        return candidates[0][1]
    if len(candidates) == 1 and candidates[0][0] >= 8:
        return candidates[0][1]
    if len(candidates) > 1 and candidates[0][0] >= 8 and candidates[0][0] >= candidates[1][0] + 4:
        return candidates[0][1]
    return None


def scan_attachments(folder: Path | None) -> list[dict[str, Any]]:
    if folder is None or not folder.exists():
        return []
    files: list[dict[str, Any]] = []
    for item in folder.rglob("*"):
        if not item.is_file() or item.suffix.lower() not in ATTACHMENT_EXTENSIONS:
            continue
        if item.name.startswith("~$"):
            continue
        try:
            relative_parts = item.relative_to(folder).parts
            category = relative_parts[0] if len(relative_parts) > 1 else "根目录"
            files.append(
                {
                    "name": item.name,
                    "category": category,
                    "relativePath": safe_relative_path(item),
                    "url": file_url(item),
                    "extension": item.suffix.lower(),
                    "size": item.stat().st_size,
                    "modifiedAt": datetime.fromtimestamp(item.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
                }
            )
        except OSError:
            continue
    return sorted(files, key=lambda f: (f["category"], f["name"]))


def extract_local_excel_tables(folder: Path | None) -> list[dict[str, Any]]:
    if folder is None:
        return []
    tables: list[dict[str, Any]] = []
    for excel_path in folder.rglob("*.xlsx"):
        if excel_path.name.startswith("~$") or excel_path.resolve() == MASTER_EXCEL.resolve():
            continue
        try:
            workbook = load_workbook(excel_path, data_only=True, read_only=True)
            worksheet = workbook[workbook.sheetnames[0]]
            rows = []
            for row in worksheet.iter_rows(min_row=1, max_row=min(worksheet.max_row, 12), values_only=True):
                values = [clean_value(cell) for cell in row]
                if any(values):
                    rows.append(values)
            if rows:
                tables.append(
                    {
                        "fileName": excel_path.name,
                        "sheetName": worksheet.title,
                        "url": file_url(excel_path),
                        "rows": rows,
                    }
                )
            workbook.close()
        except Exception as exc:
            tables.append({"fileName": excel_path.name, "sheetName": "", "url": file_url(excel_path), "error": str(exc)})
    return tables[:3]


def read_master_events() -> dict[str, Any]:
    if not MASTER_EXCEL.exists():
        raise FileNotFoundError(f"未找到总表：{MASTER_EXCEL}")

    workbook = load_workbook(MASTER_EXCEL, data_only=True, read_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    header_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
    headers = [normalize_header(cell) for cell in header_row]
    folders = scan_folders()
    events: list[dict[str, Any]] = []

    for row_index, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
        values = {headers[index]: row[index] if index < len(row) else None for index in range(len(headers))}
        if not any(cell is not None and str(cell).strip() for cell in row):
            continue

        title = clean_value(values.get("活动/会议名称")) or clean_value(values.get("项目名称")) or f"第 {row_index} 行会议"
        event_date = excel_date_to_iso(values.get("日期")) or infer_date_from_text(title)
        event_date_key = date_key(event_date)
        folder = match_folder(event_date_key, title, folders)
        attachments = scan_attachments(folder)
        local_tables = extract_local_excel_tables(folder)
        project_id = clean_value(values.get("项目编号")) or f"row-{row_index}"

        events.append(
            {
                "id": project_id,
                "rowIndex": row_index,
                "serial": clean_value(values.get("序号")),
                "projectId": project_id,
                "type": clean_value(values.get("类型")),
                "projectType": clean_value(values.get("项目类型")),
                "projectName": clean_value(values.get("项目名称")),
                "title": title,
                "region": clean_value(values.get("区域")),
                "date": event_date,
                "dateKey": event_date_key,
                "location": clean_value(values.get("地理位置")),
                "owner": clean_value(values.get("项目负责人")),
                "participation": clean_value(values.get("参与环节")),
                "topic": clean_value(values.get("讲题/内容")),
                "guests": split_people(clean_value(values.get("参与嘉宾"))),
                "guestText": clean_value(values.get("参与嘉宾")),
                "status": clean_value(values.get("项目状态")),
                "progress": clean_value(values.get("关键进度/问题")),
                "notes": clean_value(values.get("备注")),
                "folderName": folder.name if folder else "",
                "folderPath": safe_relative_path(folder) if folder else "",
                "folderMatched": bool(folder),
                "attachments": attachments,
                "localTables": local_tables,
            }
        )

    workbook.close()
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mtimes = [MASTER_EXCEL.stat().st_mtime, *(folder_mtime(folder) for folder in folders)]
    latest_mtime = max(mtimes)
    return {
        "generatedAt": generated_at,
        "sourceFile": MASTER_EXCEL.name,
        "sourceModifiedAt": datetime.fromtimestamp(MASTER_EXCEL.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "dataVersion": latest_mtime,
        "count": len(events),
        "events": events,
    }


class MeetingDashboardHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        return

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/events":
            self.serve_events()
            return
        if parsed.path.startswith("/files/"):
            self.serve_sm_file(parsed.path.removeprefix("/files/"))
            return
        if parsed.path == "/":
            self.path = "/index.html"
        super().do_GET()

    def serve_events(self) -> None:
        try:
            payload = read_master_events()
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except Exception as exc:
            body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def serve_sm_file(self, encoded_relative_path: str) -> None:
        relative = unquote(encoded_relative_path)
        relative = posixpath.normpath(relative).lstrip("/")
        target = (SM_DIR / Path(relative)).resolve()
        try:
            target.relative_to(SM_DIR.resolve())
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN, "文件访问被限制在 SM 目录内")
            return
        if not target.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "文件不存在")
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        try:
            data = target.read_bytes()
        except OSError:
            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR, "文件读取失败")
            return
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f"inline; filename*=UTF-8''{quote(target.name)}")
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    os.chdir(BASE_DIR)
    server = ThreadingHTTPServer((HOST, PORT), MeetingDashboardHandler)
    LOG_FILE.write_text(
        f"会议可视化页面已启动：http://{HOST}:{PORT}\n主数据源：{MASTER_EXCEL}\n",
        encoding="utf-8",
    )
    server.serve_forever()


if __name__ == "__main__":
    main()
