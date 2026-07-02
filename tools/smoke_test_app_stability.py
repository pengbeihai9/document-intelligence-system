"""对 app.py 的核心函数做轻量冒烟测试，确认主流程没有被改坏。"""

import os
import sqlite3
import sys
import tempfile
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = PROJECT_DIR / "app.py"
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


def load_app_head() -> dict:
    """加载 app.py 中 Streamlit 页面入口之前的函数定义。"""
    source = APP_PATH.read_text(encoding="utf-8-sig")
    head = source.split("# ---------- Streamlit 页面入口 ----------")[0]
    namespace = {"__file__": str(APP_PATH)}
    exec(compile(head, "app_head.py", "exec"), namespace)
    namespace["get_llm_config"] = lambda: (None, "unused", ["local-only"])
    return namespace


def main() -> None:
    """执行存储、解析和摘要相关的稳定性检查。"""
    os.environ.pop("LLM_API_KEY", None)
    ns = load_app_head()

    text = (
        "摘要：本文研究上市公司财务风险预警问题，使用CatBoost模型进行风险排序。"
        "方法上构建财务趋势、行业偏离、治理审计和风险持续性特征。"
        "实验采用时间阻断测试，结果用于有限审查资源下的高风险企业优先识别。"
    )

    summary_result = ns["summarize_text_with_llm"](text)
    assert isinstance(summary_result, tuple) and len(summary_result) == 3
    assert "PDCA" in summary_result[2]

    analysis_result = ns["analyze_text"](text)
    assert isinstance(analysis_result, tuple) and len(analysis_result) == 7
    category, scores, summary, source, pdca_report, freq_df, route = analysis_result
    assert isinstance(category, str) and isinstance(scores, dict)
    assert isinstance(summary, str) and isinstance(source, str)
    assert isinstance(pdca_report, str)
    assert list(freq_df.columns) == ["word", "count"]
    assert freq_df.columns.is_unique
    assert isinstance(route, dict)

    old_db = ns["DB_PATH"]
    old_data_dir = ns["DATA_DIR"]
    old_upload_dir = ns["UPLOAD_DIR"]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        tmp_path = Path(tmp)
        ns["DATA_DIR"] = tmp_path / "data"
        ns["UPLOAD_DIR"] = ns["DATA_DIR"] / "uploads"
        ns["DB_PATH"] = ns["DATA_DIR"] / "app.db"
        ns["init_storage"]()

        with sqlite3.connect(ns["DB_PATH"]) as conn:
            columns = {row[1] for row in conn.execute("PRAGMA table_info(documents)").fetchall()}
        assert "pdca_report" in columns

        ok, message = ns["create_user"]("smoke_user", "123456")
        assert ok, message
        user = ns["authenticate"]("smoke_user", "123456")
        assert user is not None

        ns["save_document"](
            user_id=user["id"],
            filename="sample.txt",
            file_bytes=b"sample",
            category=category,
            summary=summary,
            pdca_report=pdca_report,
            extracted_text=text,
            word_count=len(freq_df),
        )
        docs = ns["list_user_documents"](user["id"])
        assert len(docs) == 1
        doc = ns["get_user_document"](user["id"], docs[0]["id"])
        assert doc is not None
        assert doc["pdca_report"] == pdca_report

    ns["DB_PATH"] = old_db
    ns["DATA_DIR"] = old_data_dir
    ns["UPLOAD_DIR"] = old_upload_dir
    print("smoke test passed")


if __name__ == "__main__":
    main()
