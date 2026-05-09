from __future__ import annotations

import sys
import shutil
from pathlib import Path


def runtime_root() -> Path:
    """Return the editable runtime folder for source and packaged builds."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundled_root() -> Path:
    bundle_path = getattr(sys, "_MEIPASS", None)
    if bundle_path:
        return Path(bundle_path)
    return runtime_root()


def config_dir() -> Path:
    return runtime_root() / "config"


def bundled_config_dir() -> Path:
    return bundled_root() / "config"


def ensure_config_files() -> None:
    dst_dir = config_dir()
    src_dir = bundled_config_dir()
    dst_dir.mkdir(parents=True, exist_ok=True)
    if not src_dir.exists() or src_dir.resolve() == dst_dir.resolve():
        return
    for name in ("vehicle_rules.csv", "vehicle_rules.json"):
        src = src_dir / name
        dst = dst_dir / name
        if src.exists() and not dst.exists():
            shutil.copy2(src, dst)


def logs_dir() -> Path:
    path = runtime_root() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path
