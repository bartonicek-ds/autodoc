import logging
import os
import sys
from pathlib import Path

RED = "\033[31m"
RESET = "\033[0m"
LOG_FORMAT = "%(asctime)s %(levelname)s: %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


class ColorFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        message = super().format(record)
        if record.levelno >= logging.WARNING:
            return f"{RED}{message}{RESET}"
        return message


def setup_logging() -> None:
    handler = logging.StreamHandler()
    # Only use color in a terminal, so redirected logs don't contain escape codes
    use_color = sys.stderr.isatty() and "NO_COLOR" not in os.environ
    # The classic Windows console shows escape codes as garbage; Windows Terminal
    # (the default on Windows 11, which sets WT_SESSION) handles them
    if sys.platform == "win32" and "WT_SESSION" not in os.environ:
        use_color = False
    formatter_class = ColorFormatter if use_color else logging.Formatter
    handler.setFormatter(formatter_class(LOG_FORMAT, datefmt=DATE_FORMAT))
    logging.basicConfig(level=logging.INFO, handlers=[handler])


def log_to_file(path: Path) -> None:
    """Also write all further log messages to a file, replacing it if it exists."""
    handler = logging.FileHandler(path, mode="w", encoding="utf-8")
    handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    logging.getLogger().addHandler(handler)
