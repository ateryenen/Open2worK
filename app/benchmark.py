import argparse
import os
import subprocess
import time
from datetime import datetime

from .main import run
from .utils import log


def _kill_notepad() -> None:
    subprocess.run(
        ["taskkill", "/IM", "notepad.exe", "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def benchmark(runs: int, prefix: str) -> int:
    success = 0

    for index in range(1, runs + 1):
        filename = f"{prefix}_{index}.txt"
        command = f"Open Notepad, type benchmark run {index}, and save it to Desktop as {filename}"

        desktop_file = os.path.expandvars(rf"%USERPROFILE%\Desktop\{filename}")
        if os.path.exists(desktop_file):
            os.remove(desktop_file)

        _kill_notepad()
        time.sleep(0.3)

        log(f"benchmark run {index}/{runs} start")
        try:
            code = run(command, dry_run=False, monitor_enabled=False)
        except Exception as exc:
            log(f"benchmark run {index}/{runs} exception: {exc}")
            code = 1
        ok = code == 0 and os.path.exists(desktop_file)
        if ok:
            success += 1
        log(f"benchmark run {index}/{runs} result: {'success' if ok else 'failed'}")

        time.sleep(0.6)

    rate = (success / runs) * 100 if runs else 0
    print("-" * 50)
    print(f"BENCHMARK RESULT: {success}/{runs} successful ({rate:.1f}%)")
    print(f"TARGET (spec): >= 80.0% => {'PASS' if rate >= 80 else 'FAIL'}")
    print("-" * 50)

    return 0 if rate >= 80 else 1


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Desktop agent benchmark")
    parser.add_argument("--runs", type=int, default=10, help="How many runs")
    parser.add_argument("--prefix", type=str, default=f"poc_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    raise SystemExit(benchmark(runs=args.runs, prefix=args.prefix))
