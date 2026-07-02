# tools 目录说明

## OCR 运行环境

图片和扫描版 PDF 识别依赖 Tesseract OCR。为了让项目拷贝到其他 Windows 电脑后也能自动识别，推荐把 Tesseract 放在这里：

```text
tools/Tesseract-OCR/tesseract.exe
tools/Tesseract-OCR/tessdata/chi_sim.traineddata
tools/Tesseract-OCR/tessdata/eng.traineddata
```

程序启动时会自动按以下顺序查找 OCR：

1. 环境变量 `TESSERACT_CMD`
2. 项目内 `tools/Tesseract-OCR/tesseract.exe`
3. 项目内 `tesseract/tesseract.exe`
4. Windows 默认安装目录
5. 系统 `PATH` 中的 `tesseract`

## 论文摘要微调工具

`build_paper_sft_dataset.py` 用于把本地 PDF 论文转换为 chat SFT 训练数据。

示例：

```powershell
python tools\build_paper_sft_dataset.py `
  --pdf-dir "H:\transformer对破产影响记录\100篇顶会顶刊Transformer论文&代码合集" `
  --out artifacts\sft\paper_summary_sft.jsonl `
  --limit 20
```

`train_lora_cloud.sh` 用于在云端 GPU 机器上进行 Qwen 7B LoRA 微调。

## 文件路由工具

`file_router.py` 是应用运行时会调用的工具，用于识别文件类型和任务类别。

`file_router_config.json` 是对应配置文件。新增 PPT、图片、表格或业务分类规则时，优先修改这个 JSON，而不是直接改主程序。

## 提示词模板

`prompt_templates.py` 存放摘要提示词工程中相对稳定的模板函数，包括：

- 材料类型标签
- 中文/英文/双语输出约束
- 证据约束
- 顶刊/顶会摘要写作合同
- 不同材料类型的摘要输出规格

如果只是改摘要风格、长度要求或顶刊/顶会写作标准，优先改这个文件。

## 文献采集工具

`collect_arxiv_papers.py`、`collect_openalex_top_papers.py`、`collect_and_download_top_papers.py` 和 `download_open_access_pdfs.py` 用于顶刊/顶会文献元数据采集与开放获取 PDF 下载。

## 测试工具

`smoke_test_app_stability.py` 用于基础稳定性检查。

`run_local_literature_summary_test.py` 用于本地文献摘要效果测试。
