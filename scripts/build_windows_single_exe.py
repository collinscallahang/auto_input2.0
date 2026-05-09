from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
BUILD_ROOT = ROOT / "build" / "pyinstaller_single"
RELEASE_ROOT = ROOT / "release"
APP_NAME = "auto_input2.0"


def main() -> None:
    RELEASE_ROOT.mkdir(exist_ok=True)
    exe_path = RELEASE_ROOT / f"{APP_NAME}.exe"
    spec_file = ROOT / f"{APP_NAME}.spec"

    if BUILD_ROOT.exists():
        shutil.rmtree(BUILD_ROOT)
    if exe_path.exists():
        exe_path.unlink()
    if spec_file.exists():
        spec_file.unlink()

    data_sep = ";" if os.name == "nt" else ":"
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(RELEASE_ROOT),
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
    print(f"single_exe={exe_path}")


if __name__ == "__main__":
    main()
