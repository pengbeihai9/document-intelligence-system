"""对本地文献样例运行摘要函数，快速检查摘要质量。"""

import json
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = PROJECT_DIR / "app.py"
REPORT_PATH = PROJECT_DIR / "docs" / "十篇文献本地摘要测试结果.md"


def load_summary_functions() -> dict:
    """只加载 app.py 中摘要相关函数，避免启动 Streamlit 页面。"""
    source = APP_PATH.read_text(encoding="utf-8-sig")
    head = source.split("# ---------- Streamlit 页面入口 ----------")[0]
    namespace = {"__file__": str(APP_PATH)}
    exec(compile(head, "app_head.py", "exec"), namespace)
    namespace["get_llm_config"] = lambda: (None, "unused", "local-only")
    return namespace


def read_items(path: Path, limit: int = 10) -> list[tuple[str, str, dict]]:
    """读取待测试的文献文本或训练样本。"""
    items = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            continue
        meta = item.get("meta") or {}
        messages = item.get("messages") or []
        user_content = next((msg.get("content") or "" for msg in messages if msg.get("role") == "user"), "")
        title = meta.get("title") or f"paper_{len(items) + 1}"
        if user_content:
            items.append((title, user_content, meta))
        if len(items) >= limit:
            break
    return items


def main() -> None:
    """命令行入口：运行本地摘要测试并输出结果。"""
    ns = load_summary_functions()
    sample_path = PROJECT_DIR / "data" / "sft" / "paper_summary_sft_sample.jsonl"
    full_path = PROJECT_DIR / "data" / "sft" / "paper_summary_sft.jsonl"
    items = read_items(sample_path, 10)
    if len(items) < 10:
        items = read_items(full_path, 10)

    lines = [
        "# 十篇文献本地摘要测试结果",
        "",
        "测试模式：强制关闭大模型 API，只测试本地规则兜底摘要。",
        "",
        f"测试数量：{len(items)}",
        "",
    ]

    quality_flags = 0
    for index, (title, text, meta) in enumerate(items, 1):
        test_text = f"标题：{title}\n文献类型：论文/研究报告\n{text}"
        summary, source_name = ns["summarize_text_with_llm"](test_text)
        low_quality = ns["is_low_quality_summary"](summary, academic=True)
        if low_quality:
            quality_flags += 1

        lines.extend(
            [
                f"## {index}. {title}",
                "",
                f"- 输入长度：{len(text)} 字符",
                f"- 原文长度记录：{meta.get('full_text_chars', '未知')}",
                f"- 语言：{meta.get('language', '未知')}",
                f"- 摘要来源：{source_name}",
                f"- 摘要长度：{len(summary)} 字符",
                f"- 低质量检测：{'是' if low_quality else '否'}",
                "",
                summary,
                "",
            ]
        )

    lines.extend(
        [
            "## 汇总",
            "",
            f"- 共测试 {len(items)} 篇文献。",
            f"- 低质量检测标记 {quality_flags} 篇。",
            "- 本测试用于验证大模型掉线时的本地兜底稳定性，不代表在线大模型摘要上限。",
            "",
        ]
    )
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(str(REPORT_PATH))


if __name__ == "__main__":
    main()
