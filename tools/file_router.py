from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


CONFIG_PATH = Path(__file__).resolve().with_name("file_router_config.json")


@dataclass
class RouteResult:
    file_type: str
    file_type_label: str
    parser: str
    task_type: str
    task_label: str
    confidence: float
    task_scores: dict[str, int]
    recommended_outputs: list[str]
    reason: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def load_config(path: Path = CONFIG_PATH) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def detect_file_type(filename: str, config: dict[str, Any] | None = None) -> tuple[str, dict[str, Any]]:
    config = config or load_config()
    suffix = Path(filename).suffix.lower()
    for file_type, rule in config["file_type_rules"].items():
        if suffix in {ext.lower() for ext in rule.get("extensions", [])}:
            return file_type, rule
    return "unknown", {"label": "未知文件", "parser": "unsupported"}


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def score_task(text: str, rule: dict[str, Any]) -> int:
    normalized = normalize_text(text).lower()
    score = 0
    for keyword in rule.get("keywords", []):
        keyword_lower = keyword.lower()
        if not keyword_lower:
            continue
        if re.search(r"[\u4e00-\u9fff]", keyword):
            score += normalized.count(keyword_lower) * 2
        else:
            score += len(re.findall(rf"\b{re.escape(keyword_lower)}\b", normalized))
    return score


def infer_task_type(text: str, config: dict[str, Any] | None = None) -> tuple[str, dict[str, int]]:
    config = config or load_config()
    task_rules = config["task_rules"]
    scores = {task: score_task(text, rule) for task, rule in task_rules.items()}
    non_general = {key: value for key, value in scores.items() if key != "general_document"}
    best_task = max(non_general, key=non_general.get) if non_general else "general_document"
    if not non_general or non_general[best_task] <= 0:
        best_task = "general_document"
    return best_task, scores


def route_file(filename: str, text: str = "", config: dict[str, Any] | None = None) -> RouteResult:
    config = config or load_config()
    file_type, file_rule = detect_file_type(filename, config)
    task_type, scores = infer_task_type(text, config)
    task_rule = config["task_rules"][task_type]
    best_score = scores.get(task_type, 0)
    total_positive = sum(score for key, score in scores.items() if key != "general_document" and score > 0)
    confidence = 0.2 if task_type == "general_document" else min(0.95, 0.45 + best_score / max(total_positive, 1) * 0.5)
    reason = "未命中明确业务关键词，按通用文档处理。"
    if task_type != "general_document":
        reason = f"命中{best_score}个与“{task_rule['label']}”相关的关键词信号。"
    return RouteResult(
        file_type=file_type,
        file_type_label=file_rule.get("label", file_type),
        parser=file_rule.get("parser", "unsupported"),
        task_type=task_type,
        task_label=task_rule["label"],
        confidence=round(confidence, 3),
        task_scores=scores,
        recommended_outputs=task_rule.get("outputs", []),
        reason=reason,
    )


def route_path(path: str | Path) -> RouteResult:
    path = Path(path)
    text = ""
    if path.suffix.lower() in {".txt", ".md", ".json", ".jsonl", ".csv"}:
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")[:20000]
        except Exception:
            text = ""
    return route_file(path.name, text)


def main() -> None:
    import argparse
    import csv

    parser = argparse.ArgumentParser(description="Detect file type and route document task.")
    parser.add_argument("paths", nargs="+", help="Files or directories to classify.")
    parser.add_argument("--output", type=Path, help="Optional CSV output path.")
    args = parser.parse_args()

    files: list[Path] = []
    for item in args.paths:
        path = Path(item)
        if path.is_dir():
            files.extend(p for p in path.rglob("*") if p.is_file())
        elif path.is_file():
            files.append(path)

    rows = []
    for path in files:
        result = route_path(path).to_dict()
        result["path"] = str(path)
        result["recommended_outputs"] = "；".join(result["recommended_outputs"])
        result["task_scores"] = json.dumps(result["task_scores"], ensure_ascii=False)
        rows.append(result)

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        fieldnames = ["path", "file_type_label", "parser", "task_label", "confidence", "reason", "recommended_outputs", "task_scores"]
        with args.output.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows({key: row.get(key, "") for key in fieldnames} for row in rows)
    else:
        print(json.dumps(rows, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
