from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
DIST_ROOT = ROOT / "dist"
BUILD_ROOT = ROOT / "build" / "pyinstaller"
APP_NAME = "HuolalaQuoteTool"
PORTABLE_NAME = "HuolalaQuoteToolPortable"


def copy_tree(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def main() -> None:
    dist_app = DIST_ROOT / APP_NAME
    portable_dir = DIST_ROOT / PORTABLE_NAME
    spec_file = ROOT / f"{APP_NAME}.spec"

    for path in (dist_app, portable_dir, BUILD_ROOT):
        if path.exists():
            shutil.rmtree(path)
    if spec_file.exists():
        spec_file.unlink()

    data_sep = ";" if os.name == "nt" else ":"
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(DIST_ROOT),
        "--workpath",
        str(BUILD_ROOT),
        "--specpath",
        str(ROOT),
        "--collect-all",
        "playwright",
        "--hidden-import",
        "playwright.sync_api",
        "--add-data",
        f"{ROOT / 'config'}{data_sep}config",
        str(ROOT / "huolala_quote_tool_launcher.py"),
    ]
    subprocess.run(command, cwd=ROOT, check=True)

    dist_app.rename(portable_dir)
    copy_tree(ROOT / "config", portable_dir / "config")
    (portable_dir / "logs").mkdir(exist_ok=True)

    for doc_name in ("README.md", "使用说明.md", "真实测试记录.md"):
        src = ROOT / doc_name
        if src.exists():
            shutil.copy2(src, portable_dir / doc_name)

    launcher = portable_dir / "启动货拉拉报价工具.bat"
    launcher.write_text('@echo off\nstart "" "%~dp0HuolalaQuoteTool.exe"\n', encoding="utf-8")

    print(f"portable_dir={portable_dir}")


if __name__ == "__main__":
    main()
