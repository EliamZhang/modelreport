from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logger(output_dir: str | Path, log_file: Optional[str] = "run_log.txt", level: str = "INFO") -> logging.Logger:
    """Create a project logger with a console handler and optional file handler."""
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("model_analysis")
    logger.setLevel(getattr(logging, str(level).upper(), logging.INFO))
    logger.propagate = False

    # Avoid duplicate handlers when running in notebook / repeated local calls.
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    stream_handler.setLevel(getattr(logging, str(level).upper(), logging.INFO))
    logger.addHandler(stream_handler)

    if log_file:
        file_handler = logging.FileHandler(out / log_file, mode="w", encoding="utf-8")
        file_handler.setFormatter(fmt)
        file_handler.setLevel(logging.DEBUG)
        logger.addHandler(file_handler)

    return logger
