"""从本地 PDF 论文构建高质量中文摘要 SFT 数据集。"""

import argparse
import json
import os
import re
import time
from pathlib import Path

import fitz
import requests


SYSTEM_PROMPT = (
    "你是严谨的中英文学术论文摘要助手。只依据原文总结，不编造数据、方法或结论。"
    "输出应是一段自然流畅的中文学术摘要，不写成关键词清单、课堂笔记、机械提纲或分标题结构。"
)

SECTION_PATTERNS = {
    "abstract": [
        r"(?is)\babstract\b\s*[:.\-]?\s*(.{300,3000}?)(?=\n\s*(?:keywords|index terms|1\.?\s*introduction|introduction)\b)",
        r"(?s)摘要\s*[:：]?\s*(.{120,1800}?)(?=\n\s*(?:关键词|引言|一、|1[.、]))",
    ],
    "introduction": [
        r"(?is)(?:^|\n)\s*(?:1\.?\s*)?introduction\s*\n(.{800,5000}?)(?=\n\s*(?:2\.|related work|background|method|methodology|preliminaries)\b)",
        r"(?s)(?:^|\n)\s*(?:一、|1[.、])?\s*引言\s*\n(.{400,3000}?)(?=\n\s*(?:二、|2[.、]|相关工作|方法|研究设计))",
    ],
    "conclusion": [
        r"(?is)(?:^|\n)\s*(?:\d+\.?\s*)?(?:conclusion|conclusions|discussion and conclusion)\s*\n(.{400,4000}?)(?=\n\s*(?:references|acknowledg|appendix)\b|$)",
        r"(?s)(?:^|\n)\s*(?:六、|七、|\d+[.、])?\s*(?:结论|总结|讨论与结论)\s*\n(.{200,2500}?)(?=\n\s*(?:参考文献|附录|致谢)|$)",
    ],
}


def clean_text(text: str) -> str:
    """清理 PDF 提取文本中的空字符、断词换行和多余空白。"""
    text = text.replace("\x00", " ")
    text = re.sub(r"-\s*\n\s*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_summary(text: str) -> str:
    """清理大模型生成的训练标签，去掉标题、箭头、加号等坏格式。"""
    text = re.sub(r"^#+\s*", "", text.strip())
    text = re.sub(r"^\s*(摘要|核心概述|关键信息|重要结论|建议行动)[:：]\s*", "", text)
    text = re.sub(r"^本文围绕([^。]{2,80}?)(?:展开研究|进行系统综述|进行研究)[，。]?", r"本文以\1为研究对象，", text)
    text = re.sub(r"[→➜⇒]", "，", text)
    text = re.sub(r"\s+\+\s+", "，", text)
    text = re.sub(r"\n+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" ：:;-")


def extract_pdf_text(path: Path, max_pages: int = 24) -> str:
    """读取 PDF 前若干页文本，用于构建论文摘要训练样本。"""
    doc = fitz.open(path)
    parts = []
    for page in doc[:max_pages]:
        parts.append(page.get_text("text"))
    return clean_text("\n".join(parts))


def first_match(text: str, patterns: list[str]) -> str:
    """按正则列表提取第一个命中的章节内容。"""
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return clean_text(match.group(1))
    return ""


def title_from_filename(path: Path) -> str:
    """根据 PDF 文件名生成论文标题兜底值。"""
    title = re.sub(r"[_\-]+", " ", path.stem)
    return re.sub(r"\s+", " ", title).strip()


def is_collection_or_index_pdf(path: Path, text: str) -> bool:
    """识别论文合集目录页，避免把索引文件误当成单篇论文。"""
    title = title_from_filename(path).lower()
    if any(token in title for token in ["合集", "collection", "代码合集", "论文&代码"]):
        return True
    url_count = len(re.findall(r"https?://", text))
    venue_count = sum(text.count(token) for token in ["会议：", "期刊：", "CVPR", "ICML", "ICCV", "ECCV", "AAAI"])
    abstract_signal = re.search(r"(?i)\babstract\b|摘要", text[:5000]) is not None
    return url_count >= 20 and venue_count >= 20 and not abstract_signal


def detect_language(text: str) -> str:
    """粗略判断论文主体语言是中文还是英文。"""
    chinese = len(re.findall(r"[\u4e00-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    return "zh" if chinese >= latin * 0.35 else "en"


def build_source_record(path: Path, text: str) -> dict:
    """提取论文核心字段，组装给大模型生成摘要标签的源记录。"""
    abstract = first_match(text, SECTION_PATTERNS["abstract"])
    introduction = first_match(text, SECTION_PATTERNS["introduction"])
    conclusion = first_match(text, SECTION_PATTERNS["conclusion"])
    core_text = "\n\n".join(
        part
        for part in [
            f"Abstract:\n{abstract}" if abstract else "",
            f"Introduction:\n{introduction}" if introduction else "",
            f"Conclusion:\n{conclusion}" if conclusion else "",
        ]
        if part
    )
    if len(core_text) < 800:
        core_text = text[:9000]
    return {
        "title": title_from_filename(path),
        "file": str(path),
        "language": detect_language(text),
        "abstract": abstract,
        "introduction": introduction[:5000],
        "conclusion": conclusion[:4000],
        "core_text": core_text[:12000],
        "full_text_chars": len(text),
    }


def get_llm_config() -> tuple[str | None, str, str]:
    """从环境变量读取大模型 API 配置。"""
    api_key = os.getenv("LLM_API_KEY")
    api_base = os.getenv("LLM_API_BASE", "https://api.openai.com/v1").rstrip("/")
    model = os.getenv("LLM_MODEL", "gpt-5.4-mini")
    return api_key, api_base, model


def call_llm(prompt: str, timeout: int = 90) -> str:
    """调用大模型生成高质量摘要标签。"""
    api_key, api_base, model = get_llm_config()
    if not api_key:
        raise RuntimeError("LLM_API_KEY is required for quality SFT target generation")
    response = requests.post(
        f"{api_base}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.1,
            "max_tokens": 750,
        },
        timeout=timeout,
    )
    response.raise_for_status()
    return clean_summary(response.json()["choices"][0]["message"]["content"])


def build_target_summary(record: dict) -> str:
    """为单篇论文构建提示词并生成目标摘要。"""
    prompt = f"""请根据以下论文核心内容生成一段中文学术摘要。
要求：
1. 300-500 个中文字符。
2. 覆盖研究背景或问题、方法或模型、数据或实验、主要发现、贡献和意义。
3. 语言自然流畅，像人工写作，不要写成关键词清单、课堂笔记、机械提纲或分标题结构。
4. 保留必要英文术语。
5. 不要编造原文没有的数据、指标、模型名称或结论。
6. 不要使用箭头、加号、斜杠、编号符号、括号式解释或“核心概述/关键信息/建议行动”等标题。
7. 不要用“本文围绕……”作为开头，优先用具体研究问题、研究对象或方法切入。
8. 只输出摘要正文。

论文标题：{record["title"]}
论文语言：{record["language"]}
论文核心内容：
{record["core_text"]}
"""
    return call_llm(prompt)


def to_chat_sample(record: dict, target_summary: str) -> dict:
    """把论文源内容和目标摘要转换成 chat SFT JSONL 样本。"""
    user_content = f"""请根据以下论文内容生成一段自然流畅的中文学术摘要，必要时保留英文术语。
要求只输出摘要正文，不要分标题，不要使用箭头、加号、列表或括号式解释。
不要用“本文围绕……”作为开头。

标题：{record["title"]}
语言：{record["language"]}
核心内容：
{record["core_text"]}
"""
    return {
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
            {"role": "assistant", "content": target_summary},
        ],
        "meta": {
            "title": record["title"],
            "file": record["file"],
            "language": record["language"],
            "full_text_chars": record["full_text_chars"],
            "target_source": "llm_generated_needs_review",
        },
    }


def main() -> None:
    """命令行入口：遍历 PDF、生成摘要标签并写出 JSONL。"""
    parser = argparse.ArgumentParser(description="Build high-quality paper summarization SFT data from local PDFs.")
    parser.add_argument("--pdf-dir", required=True)
    parser.add_argument("--out", default="artifacts/sft/quality_paper_summary_sft.jsonl")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--max-pages", type=int, default=24)
    parser.add_argument("--sleep", type=float, default=0.5)
    args = parser.parse_args()

    pdf_dir = Path(args.pdf_dir)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(pdf_dir.glob("*.pdf"))
    if args.limit:
        pdfs = pdfs[: args.limit]
    if not pdfs:
        raise SystemExit(f"No PDF files found in {pdf_dir}")

    count = 0
    with out_path.open("w", encoding="utf-8") as writer:
        for index, pdf in enumerate(pdfs, start=1):
            try:
                text = extract_pdf_text(pdf, max_pages=args.max_pages)
                if len(text) < 500:
                    print(f"[SKIP] too little text: {pdf.name}")
                    continue
                if is_collection_or_index_pdf(pdf, text):
                    print(f"[SKIP] collection/index pdf: {pdf.name}")
                    continue
                record = build_source_record(pdf, text)
                summary = build_target_summary(record)
                writer.write(json.dumps(to_chat_sample(record, summary), ensure_ascii=False) + "\n")
                count += 1
                print(f"[OK] {index}/{len(pdfs)} {pdf.name}")
                if args.sleep > 0:
                    time.sleep(args.sleep)
            except Exception as exc:
                print(f"[WARN] failed: {pdf.name}: {exc}")

    print(f"Saved {count} samples to {out_path}")


if __name__ == "__main__":
    main()
