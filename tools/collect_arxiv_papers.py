"""从 arXiv API 采集论文元数据并导出 CSV。"""

import argparse
import csv
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote

import requests


ARXIV_API = "http://export.arxiv.org/api/query"
NS = {"atom": "http://www.w3.org/2005/Atom", "arxiv": "http://arxiv.org/schemas/atom"}


def entry_text(entry: ET.Element, name: str) -> str:
    """从 arXiv XML 条目中读取指定字段文本。"""
    node = entry.find(f"atom:{name}", NS)
    return " ".join(node.text.split()) if node is not None and node.text else ""


def collect(query: str, limit: int, sleep_seconds: float) -> list[dict]:
    """按查询词分页采集 arXiv 论文元数据。"""
    rows = []
    start = 0
    batch_size = min(100, limit)
    while len(rows) < limit:
        search_query = quote(f"all:{query}")
        url = (
            f"{ARXIV_API}?search_query={search_query}"
            f"&start={start}&max_results={batch_size}&sortBy=submittedDate&sortOrder=descending"
        )
        response = requests.get(url, timeout=60)
        response.raise_for_status()
        root = ET.fromstring(response.text)
        entries = root.findall("atom:entry", NS)
        if not entries:
            break
        for entry in entries:
            arxiv_id = entry_text(entry, "id")
            pdf_url = arxiv_id.replace("/abs/", "/pdf/") + ".pdf" if arxiv_id else ""
            categories = [node.attrib.get("term", "") for node in entry.findall("atom:category", NS)]
            authors = [
                entry_text(author, "name")
                for author in entry.findall("atom:author", NS)
            ]
            rows.append(
                {
                    "title": entry_text(entry, "title"),
                    "year": entry_text(entry, "published")[:4],
                    "venue": "arXiv",
                    "arxiv_id": arxiv_id,
                    "pdf_url": pdf_url,
                    "landing_page_url": arxiv_id,
                    "categories": ";".join(categories),
                    "authors": "; ".join(authors[:12]),
                    "abstract": entry_text(entry, "summary"),
                }
            )
            if len(rows) >= limit:
                break
        start += len(entries)
        time.sleep(sleep_seconds)
    return rows


def main() -> None:
    """命令行入口：执行 arXiv 检索并保存 CSV。"""
    parser = argparse.ArgumentParser(description="Collect open arXiv paper metadata.")
    parser.add_argument("--query", required=True)
    parser.add_argument("--out", default="data/papers/arxiv_papers.csv")
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--sleep", type=float, default=3.0)
    args = parser.parse_args()

    rows = collect(args.query, args.limit, args.sleep)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["title", "year", "venue", "arxiv_id", "pdf_url", "landing_page_url", "categories", "authors", "abstract"]
    with out_path.open("w", encoding="utf-8-sig", newline="") as writer_file:
        writer = csv.DictWriter(writer_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Saved {len(rows)} arXiv papers to {out_path}")


if __name__ == "__main__":
    main()
