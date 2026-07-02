"""检查本机 OCR 环境是否可用，并给出 Tesseract/Windows OCR/Poppler 状态。"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]


def find_tesseract() -> Path | None:
    """按项目约定路径、系统安装路径和 PATH 查找 tesseract.exe。"""
    candidates = [
        os.getenv("TESSERACT_CMD"),
        PROJECT_DIR / "tools" / "Tesseract-OCR" / "tesseract.exe",
        PROJECT_DIR / "tesseract" / "tesseract.exe",
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        shutil.which("tesseract"),
    ]
    for candidate in candidates:
        if not candidate:
            continue
        path = Path(candidate)
        if path.exists():
            return path
    return None


def run(command: list[str]) -> tuple[int, str]:
    """执行外部命令并返回退出码与合并后的输出文本。"""
    try:
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20)
        return result.returncode, (result.stdout + result.stderr).strip()
    except Exception as exc:
        return 1, str(exc)


def windows_ocr_available() -> bool:
    """检查 Windows 内置 OCR 是否可作为 Tesseract 不可用时的兜底。"""
    if os.name != "nt":
        return False
    script = PROJECT_DIR / "tools" / "windows_ocr.ps1"
    if not script.exists() or not shutil.which("powershell"):
        return False
    command = (
        "try {"
        "[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null;"
        "$engine=[Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages();"
        "if ($null -eq $engine) { exit 1 } else { exit 0 }"
        "} catch { exit 1 }"
    )
    code, _ = run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command])
    return code == 0


def main() -> int:
    """输出 OCR 检查结果，返回 0 表示 OCR 能力可用。"""
    tesseract = find_tesseract()
    if not tesseract:
        if windows_ocr_available():
            print("OCR_READY_WINDOWS_FALLBACK")
            print("未找到 Tesseract，但 Windows 内置 OCR 可用。")
            print("图片 OCR 可用；扫描版 PDF 将使用 PyMuPDF 渲染页面后调用 Windows OCR。")
            print("云端部署仍会优先使用 packages.txt 中的 Tesseract/Poppler。")
            return 0
        print("OCR_NOT_READY")
        print("未找到 tesseract.exe。")
        print("请安装 Tesseract，或放置便携版到 tools/Tesseract-OCR/tesseract.exe。")
        print("需要语言包：tools/Tesseract-OCR/tessdata/chi_sim.traineddata 和 eng.traineddata")
        return 1

    tessdata = tesseract.parent / "tessdata"
    chi = tessdata / "chi_sim.traineddata"
    eng = tessdata / "eng.traineddata"
    print(f"TESSERACT={tesseract}")
    code, version = run([str(tesseract), "--version"])
    print(version.splitlines()[0] if version else "version unknown")

    code, languages = run([str(tesseract), "--list-langs"])
    has_chi = "chi_sim" in languages
    has_eng = "eng" in languages
    print(f"LANG_CHI_SIM={'yes' if has_chi else 'no'}")
    print(f"LANG_ENG={'yes' if has_eng else 'no'}")

    if not has_chi or not has_eng:
        print("OCR_PARTIAL")
        print(f"Expected tessdata dir: {tessdata}")
        print(f"chi_sim exists: {chi.exists()}")
        print(f"eng exists: {eng.exists()}")
        return 2

    poppler = shutil.which("pdftoppm")
    print(f"POPPLER={'yes: ' + poppler if poppler else 'no'}")
    if not poppler:
        print("提示：图片 OCR 可用；扫描版 PDF 转图片还需要 Poppler。Streamlit Cloud 可通过 packages.txt 安装 poppler-utils。")

    print("OCR_READY")
    return 0


if __name__ == "__main__":
    sys.exit(main())
