# 文档智能处理系统 V2

这是一个基于 Streamlit 的文档智能处理 Web 工具，支持多用户登录、文档上传、文本提取、图片/扫描 PDF OCR、智能摘要、词频可视化和文档自动分类。

项目适合用于课程作业展示、文档分析演示和小规模内部试用。正式多人长期使用时，建议将本地 SQLite 数据库和上传文件迁移到外部数据库与对象存储。

## 主要功能

- 多用户注册、登录、退出登录。
- 用户文件隔离：每个用户只能查看自己的上传记录。
- 支持 PDF、Word、图片和扫描版 PDF。
- 普通 PDF/Word 自动提取文本。
- 图片和扫描版 PDF 支持 OCR 识别。
- 支持大模型摘要；API 不可用时自动回退到本地规则摘要。
- 支持中英文文献和普通中文材料摘要。
- 支持词频统计、柱状图和词云展示。
- 支持基于关键词规则的文档分类。
- 支持 Streamlit Cloud 部署。

## 项目结构

```text
app.py                         主程序
requirements.txt               Python 依赖
packages.txt                   Streamlit Cloud 系统依赖
install_ocr_windows.ps1        Windows OCR 安装辅助脚本
academic_abstract_examples.jsonl
                               摘要风格示例库

.streamlit/                    Streamlit 配置，secrets.toml 不提交
data/                          运行数据
  app.db                       用户、文档记录数据库
  uploads/                     用户上传文件

artifacts/                     训练产物和模型文件，不提交
  models/                      LoRA 等模型参数包
  sft/                         摘要微调训练集

experiments/                   测试样例和实验文件，不提交
backup/                        备份文件，不提交
docs/                          项目说明和优化记录
tools/                         OCR、文献采集、训练集构建等工具脚本
literature_training_data/      文献采集和训练资料
```

更详细的代码导航见：

```text
docs/代码结构与维护说明.md
```

## 本地运行

进入项目目录：

```powershell
cd /d D:\DESTOP\老师的题目_V2
```

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

启动应用：

```powershell
python -m streamlit run app.py
```

浏览器访问：

```text
http://localhost:8501
```

## 大模型摘要配置

系统支持 OpenAI 兼容接口。配置后会优先调用大模型生成摘要；如果 API 超时、失败或未配置，会自动使用本地规则摘要兜底。

本地可在 `.streamlit/secrets.toml` 中配置：

```toml
LLM_API_KEY = "your_api_key"
LLM_MODEL = "gpt-5.4-mini"
LLM_API_BASE = "https://api.tokensmart.ai/v1"
LLM_BACKUP_MODELS = "backup_model_1,backup_model_2"
```

不要把真实 API Key 提交到 GitHub。

### 可选本地英文/双语摘要模型

如果需要在没有在线 API 时生成英文摘要或双语摘要，可以额外准备一个本地 `transformers` 文本生成模型，并配置：

```toml
LOCAL_SUMMARY_MODEL_PATH = "D:/models/your-local-summary-model"
```

本地模型依赖不默认安装，避免应用变重。需要时再安装：

```powershell
python -m pip install transformers torch
```

未配置本地模型时，英文/双语模式会回退到中文规则摘要，并提示本地模型不可用。

## OCR 说明

图片和扫描版 PDF 需要 OCR。系统会按以下顺序查找 OCR 能力：

1. 环境变量 `TESSERACT_CMD`
2. 项目内 `tools/Tesseract-OCR/tesseract.exe`
3. 项目内 `tesseract/tesseract.exe`
4. Windows 默认安装目录
5. 系统 `PATH` 中的 `tesseract`
6. Windows OCR 兜底能力

Windows 本地可运行：

```powershell
.\install_ocr_windows.ps1
```

Streamlit Cloud 部署时，`packages.txt` 已包含：

```text
tesseract-ocr
tesseract-ocr-chi-sim
poppler-utils
```

## Streamlit Cloud 部署

1. 将项目上传到 GitHub。
2. 打开 Streamlit Cloud：

```text
https://share.streamlit.io/
```

3. 选择 GitHub 仓库。
4. Main file path 填写：

```text
app.py
```

5. 在 Secrets 中填写大模型配置。
6. 点击 Deploy。

部署后，用户只需要访问公网链接即可使用，不需要在自己电脑上安装 Python、Tesseract 或 Poppler。

## 数据保存说明

本地运行时，数据保存在：

```text
data/app.db
data/uploads/
```

Streamlit Cloud 免费环境不适合长期保存数据，重启或重新部署后可能丢失。正式多人长期使用时，建议改为：

- 外部数据库：PostgreSQL、MySQL 或云数据库。
- 对象存储：OSS、COS、S3 等。

## 摘要训练产物

当前项目已经包含一份云端训练得到的论文摘要 LoRA 参数包：

```text
artifacts/models/qwen2_5_7b_quality100_lora.tar.gz
```

以及对应训练集：

```text
artifacts/sft/quality_paper_summary_sft.jsonl
```

这部分主要用于后续模型微调和摘要风格优化，不是运行 Streamlit 应用的必需文件。

## 验收要点

- 可以注册和登录用户。
- 不同用户只能看到自己的上传历史。
- PDF、Word、图片可以上传并解析。
- 扫描版 PDF 可以 OCR。
- 摘要结果应避免箭头、加号、机械小标题和乱码。
- 大模型失败时，本地摘要兜底仍可工作。
- 页面有处理进度提示。
