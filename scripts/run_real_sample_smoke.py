from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from huolala_quote_tool.browser import HuolalaClient
from huolala_quote_tool.excel_model import (
    create_output_path,
    detect_workbook,
    load_vehicle_rules,
    numeric_or_original,
    value_at,
    writable_cell,
    write_success,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run a real Huolala smoke test from one sample workbook row.")
    parser.add_argument("source", type=Path, help="Sample workbook to test.")
    parser.add_argument("--row", type=int, default=3, help="Data row to test.")
    parser.add_argument("--vehicles", type=int, default=1, help="Number of detected vehicle columns to quote.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path.cwd() / "real_test_output",
        help="Directory for the real smoke test result workbook.",
    )
    parser.add_argument("--no-login-prompt", action="store_true", help="Skip the login confirmation dialog.")
    return parser.parse_args()


def confirm_login() -> bool:
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        input("请在浏览器完成登录并停留在同城下单页后按 Enter 继续...")
        return True

    root = tk.Tk()
    root.withdraw()
    try:
        return messagebox.askokcancel(
            "登录确认",
            "浏览器已经打开货拉拉同城下单页。\n\n"
            "请在浏览器完成登录，并确认页面停留在同城下单页后点击“确定”继续。",
        )
    finally:
        root.destroy()


def main() -> None:
    args = parse_args()
    rules = load_vehicle_rules()
    detection = detect_workbook(args.source, rules=rules)
    if not detection.can_start:
        raise SystemExit(f"字段识别不完整：{'；'.join(detection.warnings)}")
    if args.row <= detection.header_row:
        raise SystemExit(f"--row 必须大于表头行 {detection.header_row}")

    origin_col = detection.required_columns["origin"]
    destination_col = detection.required_columns["destination"]
    distance_col = detection.required_columns["distance"]
    assert origin_col and destination_col and distance_col

    wb = load_workbook(args.source, data_only=False)
    ws = wb[detection.sheet_name]
    origin = str(value_at(ws, args.row, origin_col.index) or "").strip()
    destination = str(value_at(ws, args.row, destination_col.index) or "").strip()
    if not origin or not destination:
        raise SystemExit("测试行缺少发货地址或到货地址")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output_path = args.output_dir / create_output_path(args.source).name

    print(f"source={args.source}")
    print(f"output={output_path}")
    print(f"sheet={detection.sheet_name}")
    print(f"header_row={detection.header_row}")
    print(f"row={args.row}")
    print(f"origin={origin}")
    print(f"destination={destination}")

    client = HuolalaClient(log=lambda message: print(f"[browser] {message}", flush=True))
    try:
        client.start()
        client.open_order_page()
        if not args.no_login_prompt and not confirm_login():
            raise SystemExit("用户取消真实网页测试")

        client.fill_addresses(origin, destination, reuse_if_same=False)

        vehicle_columns = detection.vehicle_columns[: max(args.vehicles, 0)]
        for vehicle_col in vehicle_columns:
            assert vehicle_col.rule is not None
            price = client.quote_vehicle(vehicle_col.rule)
            write_success(writable_cell(ws, args.row, vehicle_col.index), numeric_or_original(price))
            print(
                "vehicle="
                f"{vehicle_col.header} -> {vehicle_col.rule.car_length} / "
                f"{'、'.join(vehicle_col.rule.vehicle_requirements)} price={price:g}"
            )

        distance = client.read_distance()
        write_success(writable_cell(ws, args.row, distance_col.index), numeric_or_original(distance))
        print(f"distance={distance:g}")

        wb.save(output_path)
        page = client.page
        if page is not None:
            screenshot_path = args.output_dir / "real_smoke_success.png"
            text_path = args.output_dir / "real_smoke_success_body.txt"
            page.screenshot(path=str(screenshot_path), full_page=True)
            text_path.write_text(page.locator("body").inner_text(timeout=5000), encoding="utf-8")
            print(f"success_screenshot={screenshot_path}")
            print(f"success_body={text_path}")
        print("status=ok")
    except Exception:
        page = client.page
        if page is not None:
            screenshot_path = args.output_dir / "real_smoke_error.png"
            text_path = args.output_dir / "real_smoke_error_body.txt"
            page.screenshot(path=str(screenshot_path), full_page=True)
            text_path.write_text(page.locator("body").inner_text(timeout=5000), encoding="utf-8")
            print(f"error_screenshot={screenshot_path}")
            print(f"error_body={text_path}")
        raise
    finally:
        client.stop()


if __name__ == "__main__":
    main()
