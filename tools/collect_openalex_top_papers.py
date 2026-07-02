"""从 OpenAlex 采集高水平期刊/会议论文元数据。"""

import argparse
import csv
import json
import time
from pathlib import Path
from urllib.parse import quote

import requests


DEFAULT_TOP_VENUES = [
    "Nature",
    "Science",
    "Cell",
    "Proceedings of the National Academy of Sciences",
    "The Lancet",
    "New England Journal of Medicine",
    "Journal of Finance",
    "Journal of Financial Economics",
    "Review of Financial Studies",
    "Management Science",
    "Journal of Management",
    "Strategic Management Journal",
    "Academy of Management Journal",
    "American Economic Review",
    "Quarterly Journal of Economics",
    "Journal of Political Economy",
    "Econometrica",
    "Review of Economic Studies",
    "NeurIPS",
    "ICML",
    "ICLR",
    "ACL",
    "EMNLP",
    "NAACL",
    "CVPR",
    "ICCV",
    "ECCV",
    "AAAI",
    "IJCAI",
    "KDD",
    "WWW",
    "SIGIR",
    "SIGMOD",
    "VLDB",
    "USENIX Security",
    "IEEE Symposium on Security and Privacy",
    "ACM CCS",
]


def inverted_index_to_text(index: dict | None) -> str:
    """把 OpenAlex 的 abstract inverted index 还原为普通摘要文本。"""
    if not index:
        return ""
    positions = []
    for word, offsets in index.items():
        for offset in offsets:
            positions.append((offset, word))
    return " ".join(word for _, word in sorted(positions))


def normalize(text: str) -> str:
    """规范化字符串，便于期刊/会议名称匹配。"""
    return " ".join((text or "").lower().replace("&", "and").split())


def venue_allowed(work: dict, venues: list[str]) -> bool:
    """判断论文来源是否属于允许的期刊或会议列表。"""
    source = work.get("primary_location", {}).get("source") or {}
    names = [
        source.get("display_name", ""),
        work.get("primary_location", {}).get("source", {}).get("host_organization_name", ""),
    ]
    host_venue = work.get("host_venue") or {}
    names.append(host_venue.get("display_name", ""))
    normalized_names = [normalize(name) for name in names if name]
    for venue in venues:
        v = normalize(venue)
        if any(v and v in name for name in normalized_names):
            return True
    return False


def collect(query: str, venues: list[str], limit: int, mailto: str | None, sleep_seconds: float) -> list[dict]:
    """调用 OpenAlex API 检索论文并整理为统一字段。"""
    rows = []
    cursor = "*"
    per_page = min(200, max(25, limit))
    headers = {"User-Agent": f"doc-summary-tool/1.0 ({mailto or 'no-email'})"}

    while len(rows) < limit:
        params = [
            f"search={quote(query)}",
            "filter=type:article|proceedings-article,from_publication_date:2018-01-01",
            f"per-page={per_page}",
            f"cursor={quote(cursor)}",
            "select=id,doi,title,publication_year,type,cited_by_count,abstract_inverted_index,primary_location,open_access,authorships",
            "sort=cited_by_count:desc",
        ]
        if mailto:
            params.append(f"mailto={quote(mailto)}")
        url = "https://api.openalex.org/works?" + "&".join(params)
        response = None
        retryable_failed = False
        for attempt in range(1, 5):
            response = requests.get(url, headers=headers, timeout=60)
            if response.status_code in {429, 500, 502, 503, 504}:
                wait = attempt * 3
                print(f"[WARN] OpenAlex returned {response.status_code}; retrying in {wait}s")
                time.sleep(wait)
                retryable_failed = True
                continue
            response.raise_for_status()
            retryable_failed = False
            break
        if response is None:
            break
        if retryable_failed and response.status_code in {429, 500, 502, 503, 504}:
            print(f"[WARN] OpenAlex unavailable after retries; saved {len(rows)} collected rows")
            break
        response.raise_for_status()
        data = response.json()
        results = data.get("results") or []
        if not results:
            break

        for work in results:
            if not venue_allowed(work, venues):
                continue
            primary = work.get("primary_location") or {}
            source = primary.get("source") or {}
            open_access = work.get("open_access") or {}
            rows.append(
                {
                    "title": work.get("title") or "",
                    "year": work.get("publication_year") or "",
                    "type": work.get("type") or "",
                    "venue": source.get("display_name") or "",
                    "doi": work.get("doi") or "",
                    "openalex_id": work.get("id") or "",
                    "cited_by_count": work.get("cited_by_count") or 0,
                    "is_oa": open_access.get("is_oa", False),
                    "oa_status": open_access.get("oa_status") or "",
                    "pdf_url": primary.get("pdf_url") or open_access.get("oa_url") or "",
                    "landing_page_url": primary.get("landing_page_url") or "",
                    "abstract": inverted_index_to_text(work.get("abstract_inverted_index")),
                }
            )
            if len(rows) >= limit:
                break

        cursor = data.get("meta", {}).get("next_cursor")
        if not cursor:
            break
        time.sleep(sleep_seconds)

    return rows


def main() -> None:
    """命令行入口：采集 OpenAlex 元数据并写出 CSV/JSON。"""
    parser = argparse.ArgumentParser(description="Collect top venue paper metadata from OpenAlex.")
    parser.add_argument("--query", required=True, help="Search query, e.g. transformer bankruptcy prediction.")
    parser.add_argument("--out", default="data/papers/openalex_top_papers.csv")
    parser.add_argument("--jsonl", default="")
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--venues-file", default="", help="Optional text file with one venue per line.")
    parser.add_argument("--mailto", default="", help="Email for polite OpenAlex API usage.")
    parser.add_argument("--sleep", type=float, default=0.2)
    args = parser.parse_args()

    venues = DEFAULT_TOP_VENUES
    if args.venues_file:
        venues = [
            line.strip()
            for line in Path(args.venues_file).read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.strip().startswith("#")
        ]

    rows = collect(args.query, venues, args.limit, args.mailto or None, args.sleep)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "title",
        "year",
        "type",
        "venue",
        "doi",
        "openalex_id",
        "cited_by_count",
        "is_oa",
        "oa_status",
        "pdf_url",
        "landing_page_url",
        "abstract",
    ]
    with out_path.open("w", encoding="utf-8-sig", newline="") as writer_file:
        writer = csv.DictWriter(writer_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    if args.jsonl:
        jsonl_path = Path(args.jsonl)
        jsonl_path.parent.mkdir(parents=True, exist_ok=True)
        with jsonl_path.open("w", encoding="utf-8") as writer:
            for row in rows:
                writer.write(json.dumps(row, ensure_ascii=False) + "\n")

    print(f"Saved {len(rows)} papers to {out_path}")
    if args.jsonl:
        print(f"Saved jsonl to {args.jsonl}")


if __name__ == "__main__":
    main()
