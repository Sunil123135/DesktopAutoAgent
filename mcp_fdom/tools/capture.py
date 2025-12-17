from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Dict

from .common import (
    add_repo_to_syspath,
    build_run_id,
    ensure_run_dirs,
    is_app_allowed,
    snapshot_state,
    write_json,
)

add_repo_to_syspath()

from utils.fdom.element_interactor import ElementInteractor  # noqa: E402


def _maybe_reset_fdom(app_name: str, force_rebuild: bool) -> None:
    app_dir = Path(__file__).resolve().parents[2] / "apps" / app_name
    fdom_path = app_dir / "fdom.json"
    if force_rebuild and fdom_path.exists():
        backup = app_dir / "backups" / f"{fdom_path.stem}.bak"
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(fdom_path, backup)
        fdom_path.unlink(missing_ok=True)


def capture_ui_state(params: Dict[str, Any]) -> Dict[str, Any]:
    app_path = params.get("app_executable_path") or params.get("app_path")
    app_name = params.get("app_name") or Path(app_path).stem if app_path else "unknown"
    target_screen = params.get("target_screen")
    force_rebuild = bool(params.get("force_rebuild", False))
    run_id = params.get("run_id") or build_run_id(app_name)

    allowed, reason = is_app_allowed(app_path or "", app_name)
    if not allowed:
        return {"success": False, "error": reason}

    _maybe_reset_fdom(app_name, force_rebuild)

    interactor = ElementInteractor(app_executable_path=app_path)
    if not interactor.state_manager.fdom_data.get("states"):
        interactor._build_initial_dom()

    snapshot = snapshot_state(interactor, run_id)
    dirs = ensure_run_dirs(run_id)
    write_json(dirs["root"] / "snapshot.json", snapshot)

    return {"success": True, "snapshot": snapshot}


tool_spec = {
    "name": "capture_ui_state",
    "description": "Launch app (allowlist), ensure fDOM exists, and return state snapshot.",
    "schema": {
        "type": "object",
        "properties": {
            "app_executable_path": {"type": "string"},
            "app_name": {"type": "string"},
            "target_screen": {"type": "integer"},
            "force_rebuild": {"type": "boolean"},
            "run_id": {"type": "string"},
        },
        "required": ["app_executable_path"],
    },
}

