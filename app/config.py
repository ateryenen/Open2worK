from pathlib import Path
import os

DEFAULT_WAIT_OPEN_SECONDS = 1.5
DEFAULT_WAIT_TYPE_SECONDS = 0.5
DEFAULT_TEXT = "Hello Ater"
DEFAULT_FILENAME = "test.txt"
SUPPORTED_APP = "notepad"

def desktop_output_path(filename: str = DEFAULT_FILENAME) -> Path:
    desktop = Path(os.path.expandvars(r"%USERPROFILE%\Desktop"))
    return desktop / filename
