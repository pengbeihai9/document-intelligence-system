# OCR 安装与扫描件识别说明

## 当前问题

图片和扫描版 PDF 需要 OCR。项目使用 Tesseract OCR：

- 普通 Word：不需要 OCR
- 可复制文字的 PDF：不需要 OCR
- 图片：需要 Tesseract
- 扫描版 PDF：需要 Tesseract + Poppler

当前项目已增加 Windows 本机兜底方案：如果 Windows 10/11 内置 OCR 可用，即使没有 Tesseract，本机也可以识别图片和扫描版 PDF。扫描版 PDF 会优先用 Poppler，失败后使用 PyMuPDF 渲染页面再调用 Windows OCR。

## 一键检测

在项目根目录运行：

```powershell
python tools\check_ocr.py
```

结果含义：

- `OCR_READY`：图片 OCR 可用
- `OCR_READY_WINDOWS_FALLBACK`：未安装 Tesseract，但 Windows 内置 OCR 可用
- `OCR_NOT_READY`：未找到 `tesseract.exe`
- `OCR_PARTIAL`：找到 Tesseract，但缺少 `chi_sim` 或 `eng` 语言包

## 推荐解决方式

### 方式一：运行安装脚本

```powershell
.\install_ocr_windows.ps1
```

脚本会先检查已有环境，然后尝试通过 `winget` 安装 UB-Mannheim Tesseract OCR。

### 方式二：放置便携版

如果当前网络无法下载安装器，可以从其他电脑复制已安装的 `Tesseract-OCR` 文件夹到：

```text
tools/Tesseract-OCR/
```

最终结构必须是：

```text
tools/Tesseract-OCR/tesseract.exe
tools/Tesseract-OCR/tessdata/chi_sim.traineddata
tools/Tesseract-OCR/tessdata/eng.traineddata
```

复制完成后运行：

```powershell
python tools\check_ocr.py
```

## Poppler 说明

扫描版 PDF 需要先把 PDF 页面转成图片，因此还需要 Poppler。

Streamlit Cloud 部署时，`packages.txt` 已包含：

```text
poppler-utils
```

Windows 本机如果只测试图片 OCR，不需要 Poppler；如果要测试扫描版 PDF OCR，需要额外安装 Poppler 并加入 PATH。

## 本次处理结果

已新增：

- `tools/check_ocr.py`：OCR 环境检测脚本
- `tools/windows_ocr.ps1`：Windows 内置 OCR 兜底脚本
- `install_ocr_windows.ps1`：增强版安装/检测脚本
- `tools/Tesseract-OCR/README.md`：便携版放置说明

当前本机由于网络无法连接 GitHub 安装器，Tesseract 尚未实际安装；但 Windows 内置 OCR 已验证可用，本地图片 OCR 和扫描版 PDF OCR 已可通过兜底方案运行。放入便携版或恢复网络后运行安装脚本，可切换到 Tesseract 路径。
