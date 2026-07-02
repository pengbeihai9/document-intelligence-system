"""根据文献元数据下载合法开放获取 PDF，并生成下载清单。"""

import argparse
import csv
import json
import re
import time
from pathlib import Path
from urllib.parse import urlparse

import requests


def safe_filename(text: str, max_len: int = 120) -> str:
    """把论文标题转换成适合作为文件名的安全字符串。"""
    text = re.sub(r"[\\/:*?\"<>|\r\n\t]+", "_", text or "")
    text = re.sub(r"\s+", " ", text).strip(" ._")
    return (text[:max_len].rstrip(" ._") or "paper")


def read_rows(path: Path) -> list[dict]:
    """读取 CSV 或 JSON 格式的文献元数据。"""
    if path.suffix.lower() == ".jsonl":
        rows = []
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip():
                rows.append(json.loads(line))
        return rows

    with path.open("r", encoding="utf-8-sig", newline="") as reader_file:
        return list(csv.DictReader(reader_file))


def choose_url(row: dict) -> str:
    """从一条元数据中选择最可能可下载 PDF 的 URL。"""
    for key in ["pdf_url", "oa_url", "open_access_pdf", "url_for_pdf"]:
        value = (row.get(key) or "").strip()
        if value:
            return value
    return ""


def looks_like_pdf(response: requests.Response, content: bytes) -> bool:
    """根据响应头和文件头判断下载内容是否确实是 PDF。"""
    content_type = response.headers.get("content-type", "").lower()
    if "application/pdf" in content_type:
        return True
    return content.startswith(b"%PDF")


def download_pdf(url: str, out_path: Path, timeout: int, retries: int, sleep_seconds: float) -> tuple[bool, str]:
    """带重试地下载 PDF，返回是否成功及状态说明。"""
    headers = {
        "User-Agent": "doc-summary-tool/1.0 (+open-access-pdf-downloader)",
        "Accept": "application/pdf,text/html;q=0.8,*/*;q=0.5",
    }
    last_error = ""
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
            if response.status_code in {429, 500, 502, 503, 504}:
                last_error = f"HTTP {response.status_code}"
                time.sleep(sleep_seconds * attempt)
                continue
            response.raise_for_status()
            content = response.content
            if not looks_like_pdf(response, content):
                return False, f"not a PDF: content-type={response.headers.get('content-type', '')}"
            out_path.write_bytes(content)
            return True, "downloaded"
        except Exception as exc:
            last_error = str(exc)
            time.sleep(sleep_seconds * attempt)
    return False, last_error


def main() -> None:
    """命令行入口：批量下载 PDF 并记录 manifest。"""
    parser = argparse.ArgumentParser(description="Download legal open-access PDFs from collected paper metadata.")
    parser.add_argument("--input", required=True, help="CSV or JSONL from collect_openalex_top_papers.py or similar.")
    parser.add_argument("--out-dir", default="data/papers/pdfs")
    parser.add_argument("--manifest", default="data/papers/download_manifest.csv")
    parser.add_argument("--limit", type=int, default=50)
    parser.add_argument("--timeout", type=int, default=60)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--sleep", type=float, default=1.0)
    parser.add_argument("--skip-existing", action="store_true")
    args = parser.parse_args()

    input_path = Path(args.input)
    out_dir = Path(args.out_dir)
    manifest_path = Path(args.manifest)
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)

    rows = read_rows(input_path)
    manifest_rows = []
    downloaded = 0
    attempted = 0

    for index, row in enumerate(rows[: args.limit], start=1):
        title = row.get("title") or f"paper_{index}"
        year = row.get("year") or row.get("publication_year") or ""
        venue = row.get("venue") or ""
        doi = row.get("doi") or ""
        url = choose_url(row)
        filename = f"{index:03d}_{safe_filename(year)}_{safe_filename(title)}.pdf"
        out_path = out_dir / filename

        status = "skipped"
        message = ""
        if not url:
            message = "missing pdf_url"
        elif args.skip_existing and out_path.exists() and out_path.stat().st_size > 0:
            status = "exists"
            message = "already exists"
            downloaded += 1
        else:
            attempted += 1
            ok, message = download_pdf(url, out_path, args.timeout, args.retries, args.sleep)
            status = "downloaded" if ok else "failed"
            if ok:
                downloaded += 1

        manifest_rows.append(
            {
                "index": index,
                "status": status,
                "message": message,
                "title": title,
                "year": year,
                "venue": venue,
                "doi": doi,
                "pdf_url": url,
                "file_path": str(out_path) if status in {"downloaded", "exists"} else "",
            }
        )
        print(f"[{index}/{min(len(rows), args.limit)}] {status}: {title}")

    fieldnames = ["index", "status", "message", "title", "year", "venue", "doi", "pdf_url", "file_path"]
    with manifest_path.open("w", encoding="utf-8-sig", newline="") as writer_file:
        writer = csv.DictWriter(writer_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(manifest_rows)

    print(f"Attempted downloads: {attempted}")
    print(f"Available PDFs: {downloaded}")
    print(f"PDF directory: {out_dir}")
    print(f"Manifest: {manifest_path}")


if __name__ == "__main__":
    main()
