#!/usr/bin/env python3
from datetime import datetime
from pathlib import Path
import subprocess
import sys

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_PYTHON = BASE_DIR / ".venv" / "bin" / "python3"
LOG_DIR = BASE_DIR / "logs"
SCRIPTS = [
    "scraper.py",
    "scrape_fund_sizes.py",
    "merge_data.py",
    "send_email.py",
]


def get_python_executable():
    if DEFAULT_PYTHON.exists():
        return DEFAULT_PYTHON
    return Path(sys.executable)


def run_step(script_name):
    command = [str(get_python_executable()), str(BASE_DIR / script_name)]
    print(f"\n=== Running {script_name} ===")
    result = subprocess.run(command, cwd=BASE_DIR, check=False)
    if result.returncode != 0:
        raise RuntimeError(f"{script_name} failed with exit code {result.returncode}")



def main():
    python_executable = get_python_executable()
    if not python_executable.exists():
        print(f"Missing Python interpreter: {python_executable}")
        sys.exit(1)

    LOG_DIR.mkdir(exist_ok=True)
    print(f"Starting daily workflow at {datetime.now().isoformat(timespec='seconds')}")

    for script_name in SCRIPTS:
        run_step(script_name)

    print(f"Workflow completed at {datetime.now().isoformat(timespec='seconds')}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Workflow failed: {exc}")
        sys.exit(1)
