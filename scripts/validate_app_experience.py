from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import json
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
SCRIPTS_DIR = ROOT_DIR / "scripts"
OUTPUT_DIR = QUOTE_APP_DIR / "output"
DOWNLOADS_DIR = Path.home() / "Downloads"

if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))


SAMPLE_TEXT = """
活动内容规划
（一）板块一：琼韵启幕礼
设置启动仪式，安排主持人、领导致辞、签到墙、基础舞台、灯光音响和摄影摄像。

（二）板块二：滨海非遗生活市集
设置30个市集摊位，配置导视指引、摊位楣板、互动规则牌和集章卡。

（三）板块三：海风音乐会
安排民乐演出和青年乐队演唱，配置舞台背景、音响设备、灯光设备和执行人员。

宣传排期
倒计时海报
公众号软文

活动保障
医疗保障
交通引导
"""

QUICK_TESTS = [
    "test_app_flow.py",
    "test_structure_recognition.py",
    "test_generated_scheme_samples.py",
    "test_download_bug_regressions.py",
]
FULL_EXTRA_TESTS = [
    "test_deep_scheme_samples.py",
    "test_long_scheme_samples.py",
]
SCAN_SUFFIXES = {".txt", ".docx", ".pdf"}


@dataclass
class CheckResult:
    name: str
    status: str
    duration_sec: float
    detail: str = ""


def _tail(text: str, max_lines: int = 35) -> str:
    lines = [line.rstrip() for line in text.splitlines() if line.strip()]
    if len(lines) <= max_lines:
        return "\n".join(lines)
    return "\n".join(lines[-max_lines:])


def _format_duration(seconds: float) -> str:
    return f"{seconds:.2f}s"


def _run_command(name: str, command: list[str], cwd: Path, timeout: int) -> CheckResult:
    started = time.perf_counter()
    try:
        completed = subprocess.run(
            command,
            cwd=str(cwd),
            text=True,
            capture_output=True,
            timeout=timeout,
        )
        duration = time.perf_counter() - started
        output = "\n".join(part for part in [completed.stdout, completed.stderr] if part)
        if completed.returncode == 0:
            return CheckResult(name, "PASS", duration, _tail(output, 12))
        return CheckResult(
            name,
            "FAIL",
            duration,
            f"exit_code={completed.returncode}\n{_tail(output)}",
        )
    except subprocess.TimeoutExpired as exc:
        duration = time.perf_counter() - started
        output = "\n".join(part for part in [exc.stdout or "", exc.stderr or ""] if part)
        return CheckResult(name, "FAIL", duration, f"timeout={timeout}s\n{_tail(output)}")


def run_compile_check(timeout: int) -> CheckResult:
    return _run_command(
        "python compileall",
        [
            sys.executable,
            "-m",
            "compileall",
            "quote_app/app.py",
            "quote_app/app_config.py",
            "quote_app/app_services.py",
            "quote_app/core",
            "scripts",
        ],
        ROOT_DIR,
        timeout,
    )


def run_regression_tests(mode: str, timeout: int) -> list[CheckResult]:
    if mode == "smoke":
        return []
    tests = list(QUICK_TESTS)
    if mode == "full":
        tests.extend(FULL_EXTRA_TESTS)
    return [
        _run_command(f"regression {script}", [sys.executable, str(SCRIPTS_DIR / script)], ROOT_DIR, timeout)
        for script in tests
    ]


def _click_button(app, label: str) -> None:
    for button in app.button:
        if button.label == label:
            button.click()
            return
    labels = [button.label for button in app.button]
    raise AssertionError(f"button not found: {label}; buttons={labels}")


def _session_get(session_state, key: str, default=None):
    try:
        return session_state[key]
    except KeyError:
        return default


def run_app_experience_check(args: argparse.Namespace) -> CheckResult:
    started = time.perf_counter()
    old_cwd = Path.cwd()
    timings: dict[str, float] = {}
    try:
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest

        os.chdir(QUOTE_APP_DIR)
        app = AppTest.from_file("app.py", default_timeout=args.apptest_timeout)

        step_started = time.perf_counter()
        app.run()
        timings["initial_render"] = time.perf_counter() - step_started
        if app.exception:
            raise AssertionError(f"initial render exception: {app.exception}")

        app.text_area[0].set_value(SAMPLE_TEXT)
        _click_button(app, "识别报价项")
        step_started = time.perf_counter()
        app.run()
        timings["recognition"] = time.perf_counter() - step_started
        if app.exception:
            raise AssertionError(f"recognition exception: {app.exception}")

        working_df = _session_get(app.session_state, "working_quote_df")
        final_df = _session_get(app.session_state, "final_quote_df")
        if working_df is None or final_df is None or working_df.empty or final_df.empty:
            raise AssertionError("recognition did not produce editable quote data")

        before_final_count = len(final_df)
        app.session_state["review_quote_editor"] = {
            "edited_rows": {0: {"是否加入": False}},
            "added_rows": [],
            "deleted_rows": [],
        }
        step_started = time.perf_counter()
        app.run()
        timings["plain_editor_rerun"] = time.perf_counter() - step_started
        if len(app.session_state["final_quote_df"]) != before_final_count:
            raise AssertionError("review editor edit applied before submit button")

        app.session_state["final_quote_editor"] = {
            "edited_rows": {0: {"数量": 2.0, "单价": 10.0}},
            "added_rows": [],
            "deleted_rows": [],
        }
        _click_button(app, "保存报价单编辑")
        step_started = time.perf_counter()
        app.run()
        timings["save_final_edit"] = time.perf_counter() - step_started
        row = app.session_state["final_quote_df"].iloc[0]
        if float(row["数量"]) != 2.0 or float(row["单价"]) != 10.0 or float(row["合计"]) != 20.0:
            raise AssertionError(f"final quote edit did not persist: {row.to_dict()}")

        _click_button(app, "生成 Excel 报价单")
        step_started = time.perf_counter()
        app.run()
        timings["excel_export"] = time.perf_counter() - step_started
        excel_bytes = _session_get(app.session_state, "excel_bytes")
        if not isinstance(excel_bytes, bytes) or len(excel_bytes) < 1000:
            raise AssertionError("Excel export did not produce valid bytes")
        load_workbook(BytesIO(excel_bytes), data_only=False)

        warnings: list[str] = []
        thresholds = {
            "initial_render": args.initial_render_warn_sec,
            "recognition": args.recognition_warn_sec,
            "plain_editor_rerun": args.rerun_warn_sec,
            "save_final_edit": args.rerun_warn_sec,
            "excel_export": args.export_warn_sec,
        }
        for key, threshold in thresholds.items():
            if timings.get(key, 0.0) > threshold:
                warnings.append(f"{key}>{threshold}s")

        detail = "; ".join(f"{key}={_format_duration(value)}" for key, value in timings.items())
        duration = time.perf_counter() - started
        if warnings:
            return CheckResult("streamlit AppTest UX flow", "WARN", duration, f"{detail}; warnings={', '.join(warnings)}")
        return CheckResult("streamlit AppTest UX flow", "PASS", duration, detail)
    except Exception as exc:
        return CheckResult("streamlit AppTest UX flow", "FAIL", time.perf_counter() - started, str(exc))
    finally:
        os.chdir(old_cwd)


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _http_get(url: str, timeout: float) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "quote-app-validator/1.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def run_streamlit_server_smoke(args: argparse.Namespace) -> CheckResult:
    started = time.perf_counter()
    port = _free_port()
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        "--server.headless=true",
        f"--server.port={port}",
        "--server.fileWatcherType=none",
        "--browser.gatherUsageStats=false",
    ]
    proc = subprocess.Popen(
        command,
        cwd=str(QUOTE_APP_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    try:
        health_url = f"http://127.0.0.1:{port}/_stcore/health"
        root_url = f"http://127.0.0.1:{port}/"
        deadline = time.perf_counter() + args.server_start_timeout
        last_error = ""
        health_started = time.perf_counter()
        while time.perf_counter() < deadline:
            if proc.poll() is not None:
                output = proc.stdout.read() if proc.stdout else ""
                return CheckResult(
                    "streamlit server smoke",
                    "FAIL",
                    time.perf_counter() - started,
                    f"server exited early code={proc.returncode}\n{_tail(output)}",
                )
            try:
                health_body = _http_get(health_url, timeout=1.5)
                root_body = _http_get(root_url, timeout=3.0)
                duration = time.perf_counter() - started
                detail = (
                    f"port={port}; health_ms={(time.perf_counter() - health_started) * 1000:.0f}; "
                    f"health_bytes={len(health_body)}; root_bytes={len(root_body)}"
                )
                return CheckResult("streamlit server smoke", "PASS", duration, detail)
            except (urllib.error.URLError, TimeoutError, OSError) as exc:
                last_error = str(exc)
                time.sleep(0.4)
        output = proc.stdout.read() if proc.stdout else ""
        return CheckResult(
            "streamlit server smoke",
            "FAIL",
            time.perf_counter() - started,
            f"startup timeout={args.server_start_timeout}s last_error={last_error}\n{_tail(output)}",
        )
    finally:
        if proc.poll() is None:
            proc.terminate()
            with contextlib.suppress(subprocess.TimeoutExpired):
                proc.wait(timeout=5)
        if proc.poll() is None:
            proc.kill()


def _parse_after_time(raw: str | None) -> dt.datetime | None:
    if not raw:
        return None
    raw = raw.strip()
    patterns = ["%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M", "%Y-%m-%dT%H:%M", "%H:%M"]
    for pattern in patterns:
        try:
            parsed = dt.datetime.strptime(raw, pattern)
            if pattern == "%H:%M":
                today = dt.datetime.now()
                parsed = parsed.replace(year=today.year, month=today.month, day=today.day)
            return parsed
        except ValueError:
            continue
    raise ValueError(f"unsupported --downloads-after value: {raw}")


def _iter_scan_files(paths: Iterable[Path], after: dt.datetime | None, recursive: bool) -> list[Path]:
    selected: list[Path] = []
    after_ts = after.timestamp() if after else None
    for path in paths:
        if path.is_file():
            candidates = [path]
        elif path.is_dir():
            iterator = path.rglob("*") if recursive else path.iterdir()
            candidates = [child for child in iterator if child.is_file()]
        else:
            continue
        for candidate in candidates:
            if candidate.suffix.lower() not in SCAN_SUFFIXES:
                continue
            if after_ts is not None and candidate.stat().st_mtime < after_ts:
                continue
            selected.append(candidate)
    return sorted(set(selected), key=lambda item: str(item))


def run_scheme_file_scan(args: argparse.Namespace) -> CheckResult | None:
    scan_paths = [Path(path).expanduser() for path in args.scan_dir]
    after = _parse_after_time(args.downloads_after)
    if args.downloads_after and not scan_paths:
        scan_paths.append(DOWNLOADS_DIR)
    if not scan_paths:
        return None

    started = time.perf_counter()
    try:
        from core.extractor import extract_quote_items
        from core.normalizer import load_rules
        from core.quote_builder import build_quote_rows, diagnose_activity_content_ranges, extract_activity_sections
        from core.text_reader import read_text_from_path

        rules_path = QUOTE_APP_DIR / "data" / "rules_config.json"
        price_db_path = QUOTE_APP_DIR / "data" / "price_db.xlsx"
        rules = load_rules(rules_path)
        files = _iter_scan_files(scan_paths, after, args.recursive_scan)
        if not files:
            label = ", ".join(str(path) for path in scan_paths)
            return CheckResult("scheme file scan", "WARN", time.perf_counter() - started, f"no files matched: {label}")

        failed: set[str] = set()
        warned: set[str] = set()
        detail_lines: list[str] = []
        for path in files:
            try:
                text = read_text_from_path(path)
                extracted_rows = extract_quote_items(text, rules)
                sections = extract_activity_sections(text, extracted_rows)
                quote_rows = build_quote_rows(extracted_rows, price_db_path, text, activity_sections=sections)
                diagnostics = diagnose_activity_content_ranges(text)
                main_count = sum(1 for section in sections if str(section.get("section_level", "")) == "main")
                unassigned_count = sum(
                    1 for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块"
                )
                issues: list[str] = []
                text_is_too_short = len(text.strip()) < args.min_scan_chars
                if text_is_too_short:
                    warned.add(path.name)
                    issues.append("too_short_source")
                if diagnostics.get("has_short_range") and not text_is_too_short:
                    issues.append("short_activity_range")
                if diagnostics.get("directory_like"):
                    issues.append("directory_like_range")
                if main_count > args.max_main_sections:
                    issues.append(f"main_count>{args.max_main_sections}")
                if unassigned_count > args.max_unassigned:
                    issues.append(f"unassigned>{args.max_unassigned}")
                if not quote_rows:
                    warned.add(path.name)
                    issues.append("no_quote_rows")
                line = (
                    f"{path.name}: chars={len(text)} main={main_count} quote_rows={len(quote_rows)} "
                    f"unassigned={unassigned_count} ranges={diagnostics.get('range_count', 0)} "
                    f"issues={','.join(issues) if issues else 'none'}"
                )
                detail_lines.append(line)
                severe = [issue for issue in issues if issue not in {"no_quote_rows", "too_short_source"}]
                if severe:
                    failed.add(path.name)
            except Exception as exc:
                failed.add(path.name)
                detail_lines.append(f"{path.name}: read_or_scan_error={exc}")

        status = "FAIL" if failed else "WARN" if warned else "PASS"
        summary = f"files={len(files)} failed={len(failed)} warned={len(warned)}"
        return CheckResult(
            "scheme file scan",
            status,
            time.perf_counter() - started,
            summary + "\n" + "\n".join(detail_lines[: args.max_scan_detail_lines]),
        )
    except Exception as exc:
        return CheckResult("scheme file scan", "FAIL", time.perf_counter() - started, str(exc))


def write_report(results: list[CheckResult], args: argparse.Namespace) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = OUTPUT_DIR / f"internal_validation_{stamp}.json"
    payload = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "mode": args.mode,
        "python": sys.executable,
        "root_dir": str(ROOT_DIR),
        "results": [asdict(result) for result in results],
    }
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return report_path


def print_summary(results: list[CheckResult], report_path: Path) -> None:
    for result in results:
        print(f"[{result.status}] {result.name} ({_format_duration(result.duration_sec)})")
        if result.detail:
            print(_tail(result.detail, 10))
            print()
    counts = {status: sum(1 for result in results if result.status == status) for status in ("PASS", "WARN", "FAIL")}
    print(f"SUMMARY pass={counts['PASS']} warn={counts['WARN']} fail={counts['FAIL']}")
    print(f"REPORT {report_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="本地流畅度、体验链路和内部回归验证程序")
    parser.add_argument("--mode", choices=["smoke", "quick", "full"], default="quick", help="smoke最快，quick为日常验证，full包含深度长样本")
    parser.add_argument("--scan-dir", action="append", default=[], help="额外扫描方案文件目录或单个文件，可重复传入")
    parser.add_argument("--downloads-after", help="扫描下载目录中该时间后的方案文件，例如 19:00 或 2026-05-12 19:00")
    parser.add_argument("--recursive-scan", action="store_true", help="扫描目录时递归子目录")
    parser.add_argument("--max-main-sections", type=int, default=30, help="扫描方案时允许的最大 main 板块数")
    parser.add_argument("--max-unassigned", type=int, default=0, help="扫描方案时允许的最大未归属项数量")
    parser.add_argument("--min-scan-chars", type=int, default=500, help="低于该字数的扫描文件按残片/大纲警告处理")
    parser.add_argument("--max-scan-detail-lines", type=int, default=80, help="控制扫描详情输出行数")
    parser.add_argument("--skip-server-smoke", action="store_true", help="跳过真实 Streamlit 本地服务启动检查")
    parser.add_argument("--fail-on-warning", action="store_true", help="有 WARN 时也返回失败退出码")
    parser.add_argument("--timeout", type=int, default=120, help="单个回归脚本超时时间")
    parser.add_argument("--apptest-timeout", type=int, default=25, help="Streamlit AppTest 单次运行超时时间")
    parser.add_argument("--server-start-timeout", type=int, default=25, help="Streamlit 服务启动等待时间")
    parser.add_argument("--initial-render-warn-sec", type=float, default=8.0)
    parser.add_argument("--recognition-warn-sec", type=float, default=12.0)
    parser.add_argument("--rerun-warn-sec", type=float, default=8.0)
    parser.add_argument("--export-warn-sec", type=float, default=10.0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results: list[CheckResult] = []
    results.append(run_compile_check(args.timeout))
    results.extend(run_regression_tests(args.mode, args.timeout))
    results.append(run_app_experience_check(args))
    if not args.skip_server_smoke:
        results.append(run_streamlit_server_smoke(args))
    scan_result = run_scheme_file_scan(args)
    if scan_result is not None:
        results.append(scan_result)

    report_path = write_report(results, args)
    print_summary(results, report_path)

    has_failure = any(result.status == "FAIL" for result in results)
    has_warning = any(result.status == "WARN" for result in results)
    if has_failure or (args.fail_on_warning and has_warning):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
