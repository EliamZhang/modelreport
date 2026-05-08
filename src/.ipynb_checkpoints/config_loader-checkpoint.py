from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import Any

def load_config(config_path: str | Path) -> dict[str, Any]:
    """Load Python config and resolve root/output paths."""
    path = Path(config_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    if path.suffix.lower() != ".py":
        raise ValueError(f"Only Python config files are supported now: {path}")

    spec = importlib.util.spec_from_file_location("analysis_config_module", path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load Python config module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    cfg = getattr(module, "CONFIG", None)
    if not isinstance(cfg, dict):
        raise ValueError(f"Python config must expose a dict named CONFIG: {path}")

    project = cfg.setdefault("project", {})
    root_dir = Path(project.get("root_dir", ".")).expanduser()
    if not root_dir.is_absolute():
        root_dir = (path.parent.parent / root_dir).resolve()
    project["root_dir"] = str(root_dir)

    output_dir = Path(project.get("output_dir", "./output")).expanduser()
    if not output_dir.is_absolute():
        output_dir = (root_dir / output_dir).resolve()
    project["output_dir"] = str(output_dir)

    # Resolve input paths relative to root_dir.
    for _, table_cfg in (cfg.get("input_tables") or {}).items():
        p = Path(table_cfg.get("path", "")).expanduser()
        if p and not p.is_absolute():
            p = root_dir / p
        table_cfg["path"] = str(p)

    return cfg


def get_output_dir(cfg: dict[str, Any]) -> Path:
    return Path(cfg.get("project", {}).get("output_dir", "./output")).expanduser().resolve()


def get_encoding(cfg: dict[str, Any]) -> str:
    return cfg.get("project", {}).get("encoding", "utf-8-sig")
