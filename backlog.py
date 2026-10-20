#!/usr/bin/env python3
"""Menu đơn giản để nhập task hoặc Daily Report lên Backlog."""

import getpass
import os
import re
import shlex
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse


BASE_DIR = Path(__file__).resolve().parent
TASK_SCRIPT = BASE_DIR / "excel_to_backlog.py"
DAILY_SCRIPT = BASE_DIR / "daily_report_to_backlog.py"
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}


def normalize_excel_path(raw_value):
    """Accept a normal path, a quoted/escaped Finder path, or a file:// URL."""
    value = str(raw_value or "").strip()
    if not value:
        raise ValueError("Đường dẫn Excel không được để trống")

    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        value = value[1:-1]

    if value.lower().startswith("file://"):
        parsed = urlparse(value)
        if parsed.netloc not in ("", "localhost"):
            raise ValueError("Chỉ hỗ trợ file Excel trên máy hiện tại")
        value = unquote(parsed.path)
    elif "\\" in value:
        try:
            pieces = shlex.split(value)
        except ValueError as exc:
            raise ValueError(f"Đường dẫn Excel không hợp lệ: {exc}") from exc
        if len(pieces) == 1:
            value = pieces[0]

    return Path(value).expanduser().resolve(strict=False)


def validate_excel_path(path):
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise ValueError("File phải có định dạng .xlsx hoặc .xlsm")
    if not path.exists():
        raise ValueError(f"Không tìm thấy file: {path}")
    if not path.is_file():
        raise ValueError(f"Đường dẫn không phải file: {path}")


def infer_milestone(path):
    match = re.search(r"(?i)(?:^|[^a-z0-9])sprint[\s_#-]*(\d+)(?:[^0-9]|$)", path.stem)
    return f"Sprint {int(match.group(1))}" if match else None


def prompt_choice(input_func):
    print("\n=== BACKLOG BULLDOZER ===")
    print("1. Đẩy task mới lên Backlog")
    print("2. Đẩy Daily Report lên Backlog")
    while True:
        choice = input_func("Chọn chức năng [1/2]: ").strip()
        if choice in {"1", "2"}:
            return choice
        print("Vui lòng nhập 1 hoặc 2.")


def prompt_excel_path(input_func):
    while True:
        raw_value = input_func("Đường dẫn file Excel trên máy (có thể kéo-thả từ Finder): ")
        try:
            path = normalize_excel_path(raw_value)
            validate_excel_path(path)
            return path
        except ValueError as exc:
            print(f"Lỗi: {exc}")


def prompt_yes_no(prompt, input_func, default=False):
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        answer = input_func(prompt + suffix).strip().lower()
        if not answer:
            return default
        if answer in {"y", "yes", "c", "co", "có"}:
            return True
        if answer in {"n", "no", "k", "khong", "không"}:
            return False
        print("Vui lòng nhập y hoặc n.")


def prompt_api_key(secret_func):
    while True:
        api_key = secret_func("Nhập Backlog API Key (sẽ được ẩn): ").strip()
        if api_key:
            return api_key
        print("API Key không được để trống.")


def task_commands(path, milestone):
    common = [
        sys.executable,
        str(TASK_SCRIPT),
        "--file",
        str(path),
        "--milestone",
        milestone,
    ]
    return common + ["--check"], common + ["--run"]


def daily_commands(path, replace_managed):
    common = [sys.executable, str(DAILY_SCRIPT), "--file", str(path)]
    run_command = common.copy()
    if replace_managed:
        run_command.append("--replace-managed-comments")
    run_command.append("--run")
    return common + ["--check"], run_command


def run_child(command, run_func, env=None):
    sys.stdout.flush()
    result = run_func(command, cwd=str(BASE_DIR), env=env, check=False)
    return int(result.returncode)


def main(input_func=input, secret_func=getpass.getpass, run_func=subprocess.run):
    upload_started = False
    try:
        choice = prompt_choice(input_func)
        excel_path = prompt_excel_path(input_func)

        if choice == "1":
            detected_milestone = infer_milestone(excel_path)
            if detected_milestone:
                milestone_input = input_func(
                    f"Tên Milestone [{detected_milestone}] (nhập - nếu không dùng): "
                ).strip()
                milestone = detected_milestone if not milestone_input else milestone_input
            else:
                milestone = input_func(
                    "Tên Milestone (ví dụ Sprint 6, Enter hoặc - nếu không dùng): "
                ).strip()
            if milestone == "-":
                milestone = ""
            check_command, run_command = task_commands(excel_path, milestone)
            action = "đẩy task mới"
            replace_managed = False
        else:
            replace_managed = prompt_yes_no(
                "Thay thế các Daily Report cũ do công cụ đã tạo?",
                input_func,
                default=False,
            )
            check_command, run_command = daily_commands(excel_path, replace_managed)
            action = "đẩy Daily Report"

        print("\nĐang kiểm tra file Excel...")
        check_code = run_child(check_command, run_func)
        if check_code != 0:
            print("\nFile chưa hợp lệ; chưa hỏi API Key và chưa thay đổi Backlog.")
            return check_code

        print("\nThông tin chạy:")
        print(f"- Chức năng: {action}")
        print(f"- File: {excel_path}")
        if choice == "1":
            print(f"- Milestone: {milestone or '(không dùng)'}")
        else:
            print(f"- Thay comment do công cụ tạo: {'Có' if replace_managed else 'Không'}")

        if not prompt_yes_no("Xác nhận chạy thật?", input_func, default=False):
            print("Đã huỷ. Backlog chưa bị thay đổi.")
            return 0

        api_key = prompt_api_key(secret_func)
        child_env = os.environ.copy()
        child_env["BACKLOG_API_KEY"] = api_key

        print("\nĐang gửi dữ liệu lên Backlog...")
        upload_started = True
        run_code = run_child(run_command, run_func, env=child_env)
        if run_code == 0:
            print("Hoàn tất.")
        else:
            print(
                f"Không hoàn tất; chương trình nhập dữ liệu trả về mã lỗi {run_code}. "
                "Backlog có thể đã nhận một phần dữ liệu, hãy kiểm tra trước khi chạy lại."
            )
        return run_code
    except OSError as exc:
        print(f"\nKhông thể khởi động chương trình nhập dữ liệu: {exc}")
        print("Backlog chưa nhận thao tác từ lần chạy này.")
        return 1
    except (EOFError, KeyboardInterrupt):
        if upload_started:
            print(
                "\nĐã dừng khi đang upload. Backlog có thể đã nhận một phần dữ liệu; "
                "hãy kiểm tra trước khi chạy lại."
            )
        else:
            print("\nĐã huỷ. Backlog chưa nhận thêm thao tác từ menu.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
