from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

from .common import (
    add_repo_to_syspath,
    build_run_id,
    is_app_allowed,
    safety_check,
    sanitize_path,
    snapshot_state,
)

add_repo_to_syspath()

from utils.fdom.element_interactor import ElementInteractor  # noqa: E402


def execute_step(params: Dict[str, Any]) -> Dict[str, Any]:
    app_path = params.get("app_executable_path") or params.get("app_path")
    app_name = params.get("app_name") or Path(app_path).stem if app_path else "unknown"
    action = params.get("action", "click")
    node_id = params.get("node_id")
    payload = params.get("payload", "")
    run_id = params.get("run_id") or build_run_id(app_name)

    allowed, reason = is_app_allowed(app_path or "", app_name)
    if not allowed:
        return {"success": False, "error": reason}

    interactor = ElementInteractor(app_executable_path=app_path)

    # Safety check: basic title keyword block (best-effort)
    title = None
    try:
        if interactor.app_controller and interactor.app_controller.current_app_info:
            title = interactor.app_controller.current_app_info.get("window_title")
    except Exception:
        title = None
    safe, reason = safety_check(window_title=title)
    if not safe:
        return {"success": False, "error": reason}

    if action == "click":
        if not node_id:
            return {"success": False, "error": "missing_node_id"}
        result = interactor.click_element(node_id)
    elif action == "type_text":
        ok = interactor.app_controller.gui_api.type_text(payload or "")
        result = {"success": ok, "state_changed": False, "before_screenshot": None, "after_screenshot": None}
    elif action == "press_key":
        ok = interactor.app_controller.gui_api.send_keys(payload or "")
        result = {"success": ok, "state_changed": False, "before_screenshot": None, "after_screenshot": None}
    else:
        return {"success": False, "error": f"unsupported_action:{action}"}
    snapshot = snapshot_state(interactor, run_id)

    return {
        "success": bool(getattr(result, "success", False)) if hasattr(result, "success") else result.get("success", False),
        "state_changed": bool(getattr(result, "state_changed", False)) if hasattr(result, "state_changed") else result.get("state_changed", False),
        "previous_state": getattr(interactor, "current_state_id", snapshot.get("state_id")),
        "current_state": snapshot.get("state_id"),
        "clicked_node_id": node_id,
        "before_screenshot": sanitize_path(getattr(result, "before_screenshot", None) if hasattr(result, "before_screenshot") else result.get("before_screenshot")),
        "after_screenshot": sanitize_path(getattr(result, "after_screenshot", None) if hasattr(result, "after_screenshot") else result.get("after_screenshot")),
        "diff_image": sanitize_path(getattr(result, "screenshot_path", None) if hasattr(result, "screenshot_path") else result.get("screenshot_path")),
        "error_message": getattr(result, "error_message", None) if hasattr(result, "error_message") else result.get("error_message"),
        "snapshot": snapshot,
    }


tool_spec = {
    "name": "execute_step",
    "description": "Execute a single UI action (click/type_text/press_key) via ElementInteractor with safety checks.",
    "schema": {
        "type": "object",
        "properties": {
            "app_executable_path": {"type": "string"},
            "app_name": {"type": "string"},
            "node_id": {"type": "string"},
            "action": {"type": "string"},
            "payload": {"type": "string"},
            "run_id": {"type": "string"},
        },
        "required": ["app_executable_path"],
    },
}

