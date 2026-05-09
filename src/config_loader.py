from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any


def resolve_config_paths(config: dict[str, Any]) -> dict[str, Any]:
    cfg = deepcopy(config)
    project = cfg["project"]
    root_dir = Path(project.get("root_dir", ".")).expanduser()
    if not root_dir.is_absolute():
        root_dir = Path.cwd() / root_dir
    project["root_dir"] = str(root_dir.resolve())

    output_dir = Path(project["output_dir"]).expanduser()
    if not output_dir.is_absolute():
        output_dir = Path(project["root_dir"]) / output_dir
    project["output_dir"] = str(output_dir.resolve())

    for table_cfg in cfg["input_tables"].values():
        path = Path(table_cfg["path"]).expanduser()
        if not path.is_absolute():
            path = Path(project["root_dir"]) / path
        table_cfg["path"] = str(path.resolve())

    return cfg
