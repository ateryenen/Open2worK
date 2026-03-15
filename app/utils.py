from datetime import datetime
from pathlib import Path


def log(message: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {message}")


def verify_file_exists(path: str) -> bool:
    return Path(path).expanduser().exists()
