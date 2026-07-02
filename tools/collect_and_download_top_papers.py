"""串联文献检索和开放获取 PDF 下载的总控脚本。"""

import argparse
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]


def run(command: list[str]) -> None:
    """执行子命令，失败时直接抛出异常终止流程。"""
    print(" ".join(command))
    subprocess.run(command, cwd=PROJECT_DIR, check=True)


def main() -> None:
    """命令行入口：先采集元数据，再下载可获取的 PDF。"""
    parser = argparse.ArgumentParser(description="Collect top-venue metadata and download legal open-access PDFs.")
    parser.add_argument("--query", required=True)
    parser.add_argument("--limit", type=int, default=50)
    parser.add_argument("--download-limit", type=int, default=30)
    parser.add_argument("--mailto", default="")
    parser.add_argument("--out-root", default="data/papers")
    parser.add_argument("--skip-existing", action="store_true")
    args = parser.parse_args()

    out_root = Path(args.out_root)
    metadata_csv = out_root / "openalex_top_papers.csv"
    metadata_jsonl = out_root / "openalex_top_papers.jsonl"
    pdf_dir = out_root / "pdfs"
    manifest = out_root / "download_manifest.csv"

    collect_cmd = [
        sys.executable,
        "tools/collect_openalex_top_papers.py",
        "--query",
        args.query,
        "--out",
        str(metadata_csv),
        "--jsonl",
        str(metadata_jsonl),
        "--limit",
        str(args.limit),
    ]
    if args.mailto:
        collect_cmd.extend(["--mailto", args.mailto])
    run(collect_cmd)

    download_cmd = [
        sys.executable,
        "tools/download_open_access_pdfs.py",
        "--input",
        str(metadata_csv),
        "--out-dir",
        str(pdf_dir),
        "--manifest",
        str(manifest),
        "--limit",
        str(args.download_limit),
    ]
    if args.skip_existing:
        download_cmd.append("--skip-existing")
    run(download_cmd)

    print(f"Metadata CSV: {metadata_csv}")
    print(f"Metadata JSONL: {metadata_jsonl}")
    print(f"PDF directory: {pdf_dir}")
    print(f"Download manifest: {manifest}")


if __name__ == "__main__":
    main()
