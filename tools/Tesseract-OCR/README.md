# 便携版 Tesseract 放置说明

图片和扫描版 PDF 的 OCR 依赖 Tesseract。普通可复制文字的 PDF 和 Word 不需要 Tesseract。

如果希望项目拷贝到另一台 Windows 电脑后也能直接识别图片和扫描版 PDF，可以把便携版 Tesseract 放在本目录，结构如下：

```text
tools/Tesseract-OCR/tesseract.exe
tools/Tesseract-OCR/tessdata/chi_sim.traineddata
tools/Tesseract-OCR/tessdata/eng.traineddata
```

程序启动时会自动查找：

1. 环境变量 `TESSERACT_CMD`
2. `tools/Tesseract-OCR/tesseract.exe`
3. `tesseract/tesseract.exe`
4. Windows 默认安装目录
5. 系统 `PATH` 中的 `tesseract`

## 验证 OCR

在项目根目录运行：

```powershell
python tools\check_ocr.py
```

看到 `OCR_READY` 表示图片 OCR 可用。扫描版 PDF 还需要 Poppler；云端部署时 `packages.txt` 会安装 `poppler-utils`。

## 本机安装

如果不使用便携版，可以运行项目根目录的：

```powershell
.\install_ocr_windows.ps1
```

如果网络无法下载，请手动下载 UB-Mannheim Tesseract Windows 安装包，或把已安装电脑上的 `Tesseract-OCR` 目录复制到本目录。
