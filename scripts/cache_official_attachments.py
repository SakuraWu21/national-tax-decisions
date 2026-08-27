#!/usr/bin/env python3
"""下载并保留已核验的官方文书附件，供官网直链失效时稳定访问。"""

from __future__ import annotations

import hashlib
import io
import json
import os
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "public" / "data" / "tax-decisions.json"
MANIFEST_PATH = ROOT / "public" / "data" / "link-fallbacks.json"
CACHE_DIR = ROOT / "public" / "official-attachments"
MAX_BYTES = 25 * 1024 * 1024
TIMEOUT = (10, 45)
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0 Safari/537.36"
)
TZ_NAME = "Asia/Shanghai"


def now_in_project_timezone() -> datetime:
    """所有环境统一使用项目约定的北京时间。"""
    return datetime.now(ZoneInfo(TZ_NAME))


def new_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=2,
        connect=2,
        read=2,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
    )
    session.mount("http://", HTTPAdapter(max_retries=retry))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,*/*;q=0.8",
    })
    return session


def document_extension(content: bytes) -> str:
    if content.startswith(b"%PDF"):
        return ".pdf"
    if content.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return ".doc"
    if content.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                if any(name.startswith("word/") for name in archive.namelist()):
                    return ".docx"
        except zipfile.BadZipFile:
            return ""
    return ""


def candidate_for(record: dict) -> tuple[str, str, str] | None:
    attachment_url = str(record.get("attachmentUrl") or "").strip()
    official_url = str(record.get("officialUrl") or "").strip()
    if attachment_url:
        return attachment_url, official_url, "attachment"
    if official_url and Path(urlparse(official_url).path).suffix.lower() in {".pdf", ".doc", ".docx"}:
        return official_url, "", "official_document"
    return None


def cached_file_from_entry(entry: dict) -> Path | None:
    cached_url = str(entry.get("cachedUrl") or "")
    if not cached_url.startswith("/official-attachments/"):
        return None
    path = ROOT / "public" / cached_url.removeprefix("/")
    return path if path.is_file() else None


def download_document(original_url: str, source_page_url: str) -> dict:
    checked_at = now_in_project_timezone().isoformat(timespec="seconds")
    headers = {"Referer": source_page_url} if source_page_url else {}
    try:
        with new_session() as session:
            with session.get(original_url, headers=headers, timeout=TIMEOUT, allow_redirects=True, stream=True) as response:
                response.raise_for_status()
                declared_size = int(response.headers.get("Content-Length") or 0)
                if declared_size > MAX_BYTES:
                    raise ValueError(f"附件超过 {MAX_BYTES // 1024 // 1024}MB 上限")
                chunks: list[bytes] = []
                size = 0
                for chunk in response.iter_content(128 * 1024):
                    if not chunk:
                        continue
                    size += len(chunk)
                    if size > MAX_BYTES:
                        raise ValueError(f"附件超过 {MAX_BYTES // 1024 // 1024}MB 上限")
                    chunks.append(chunk)
                content = b"".join(chunks)
                extension = document_extension(content)
                if not extension:
                    raise ValueError("官方地址返回的不是可验证的 PDF/Word 文件")
                url_hash = hashlib.sha256(original_url.encode("utf-8")).hexdigest()[:20]
                target = CACHE_DIR / f"{url_hash}{extension}"
                temp = target.with_suffix(f".tmp{extension}")
                temp.write_bytes(content)
                os.replace(temp, target)
                return {
                    "status": "cached",
                    "checkedAt": checked_at,
                    "statusCode": response.status_code,
                    "contentType": response.headers.get("Content-Type", "").split(";")[0],
                    "cachedUrl": f"/official-attachments/{target.name}",
                    "sha256": hashlib.sha256(content).hexdigest(),
                    "bytes": len(content),
                    "message": "已从官方地址核验并保留原始附件。",
                }
    except Exception as exc:
        return {
            "status": "unavailable",
            "checkedAt": checked_at,
            "statusCode": None,
            "contentType": None,
            "cachedUrl": None,
            "sha256": None,
            "bytes": None,
            "message": f"{type(exc).__name__}: {exc}",
        }


def load_previous_entries() -> dict[str, dict]:
    if not MANIFEST_PATH.exists():
        return {}
    try:
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
        return manifest.get("attachments", {}) if isinstance(manifest, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def main() -> int:
    records = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    previous_entries = load_previous_entries()

    candidates: dict[str, tuple[str, str]] = {}
    record_candidates: dict[str, tuple[str, str, str]] = {}
    for record in records:
        candidate = candidate_for(record)
        if not candidate:
            continue
        original_url, source_page_url, link_type = candidate
        record_id = str(record["id"])
        record_candidates[record_id] = candidate
        candidates.setdefault(original_url, (source_page_url, link_type))

    fetched: dict[str, dict] = {}
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {
            executor.submit(download_document, original_url, source_page_url): original_url
            for original_url, (source_page_url, _link_type) in candidates.items()
        }
        for future in as_completed(futures):
            original_url = futures[future]
            fetched[original_url] = future.result()

    entries: dict[str, dict] = {}
    for record_id, (original_url, source_page_url, link_type) in record_candidates.items():
        result = dict(fetched[original_url])
        if not result.get("cachedUrl"):
            previous = previous_entries.get(record_id, {})
            previous_file = cached_file_from_entry(previous)
            if previous.get("originalUrl") == original_url and previous_file:
                result.update({
                    "status": "cached_previous",
                    "cachedUrl": previous.get("cachedUrl"),
                    "sha256": previous.get("sha256"),
                    "bytes": previous.get("bytes"),
                    "message": "本次官网暂时不可访问，继续使用上次核验保存的官方附件。",
                })
        entries[record_id] = {
            "originalUrl": original_url,
            "sourcePageUrl": source_page_url or None,
            "linkType": link_type,
            **result,
        }

    manifest = {
        "generatedAt": now_in_project_timezone().isoformat(timespec="seconds"),
        "attachments": dict(sorted(entries.items())),
    }
    temp_manifest = MANIFEST_PATH.with_suffix(".tmp.json")
    temp_manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_manifest, MANIFEST_PATH)

    cached = sum(1 for entry in entries.values() if entry.get("cachedUrl"))
    unavailable = len(entries) - cached
    print(json.dumps({
        "attachment_records": len(entries),
        "unique_official_files": len(candidates),
        "cached_records": cached,
        "unavailable_records": unavailable,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
