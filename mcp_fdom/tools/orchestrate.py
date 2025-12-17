from __future__ import annotations

from typing import Any, Dict, List, Optional

from .common import build_run_id, ensure_run_dirs, write_json
from .capture import capture_ui_state
from .gemini import analyze_with_gemini
from .integrate import integrate_gemini
from .execute import execute_step


def _resolve_node(snapshot: Dict[str, Any], keywords: List[str]) -> Optional[str]:
    nodes = snapshot.get("nodes", [])
    pending = snapshot.get("pending_nodes", [])
    for kw in keywords:
        for node in nodes:
            name = (node.get("g_icon_name") or "").lower()
            brief = (node.get("g_brief") or "").lower()
            if kw in name or kw in brief:
                return node.get("node_id")
    if pending:
        pending_id = pending[0]
        if "::" in pending_id:
            return pending_id.split("::", 1)[1]
        return pending_id
    return None


def _plan_steps(snapshot: Dict[str, Any], iteration: int) -> List[Dict[str, Any]]:
    return [
        {
            "label": "focus_editor",
            "action": "click",
            "node_id": _resolve_node(snapshot, ["edit", "document", "text", "pane"]),
        },
        {
            "label": "type_line",
            "action": "type_text",
            "payload": f"Iteration {iteration:02d} - Hello",
        },
        {
            "label": "open_edit_menu",
            "action": "click",
            "node_id": _resolve_node(snapshot, ["edit", "menu", "menu bar"]),
        },
        {
            "label": "select_all",
            "action": "press_key",
            "payload": "ctrl+a",
        },
        {
            "label": "delete_text",
            "action": "press_key",
            "payload": "delete",
        },
    ]


def plan_and_run_5_steps(params: Dict[str, Any]) -> Dict[str, Any]:
    app_path = params.get("app_executable_path") or params.get("app_path")
    app_name = params.get("app_name") or "notepad"
    iterations = int(params.get("iterations", 50))
    run_id = params.get("run_id") or build_run_id(app_name)

    dirs = ensure_run_dirs(run_id)

    summary = {
        "run_id": run_id,
        "app": app_name,
        "iterations_requested": iterations,
        "iterations_completed": 0,
        "failures": 0,
        "failure_reasons": [],
        "iteration_logs": [],
    }

    snapshot_result = capture_ui_state(
        {
            "app_executable_path": app_path,
            "app_name": app_name,
            "run_id": run_id,
            "force_rebuild": params.get("force_rebuild", False),
        }
    )
    if not snapshot_result.get("success"):
        return {"success": False, "error": snapshot_result.get("error"), "run_id": run_id}

    snapshot = snapshot_result["snapshot"]

    for i in range(1, iterations + 1):
        iteration_log = {"iteration": i, "steps": [], "errors": []}

        gemini_result = analyze_with_gemini({"app_name": app_name, "run_id": run_id})
        if gemini_result.get("success"):
            integrate_gemini(
                {
                    "app_name": app_name,
                    "gemini_results": gemini_result.get("gemini_results", {}),
                    "run_id": run_id,
                }
            )

        plan = _plan_steps(snapshot, i)

        for step in plan:
            if step["action"] == "click" and not step.get("node_id"):
                step_result = {"success": False, "error": "no_node_resolved", "label": step["label"]}
            else:
                step_result = execute_step(
                    {
                        "app_executable_path": app_path,
                        "app_name": app_name,
                        "action": step["action"],
                        "node_id": step.get("node_id"),
                        "payload": step.get("payload"),
                        "run_id": run_id,
                    }
                )
                step_result["label"] = step["label"]
            iteration_log["steps"].append(step_result)

        write_json(dirs["root"] / f"iteration_{i:02d}.json", iteration_log)
        summary["iteration_logs"].append(f"runs/{run_id}/iteration_{i:02d}.json")
        summary["iterations_completed"] += 1 if all(s.get("success") for s in iteration_log["steps"]) else 0
        if not all(s.get("success") for s in iteration_log["steps"]):
            summary["failures"] += 1
            summary["failure_reasons"].append(f"iteration_{i:02d}_step_failure")

        snapshot_result = capture_ui_state(
            {"app_executable_path": app_path, "app_name": app_name, "run_id": run_id}
        )
        snapshot = snapshot_result.get("snapshot", snapshot)

    summary_path = dirs["root"] / "summary.json"
    write_json(summary_path, summary)

    return {"success": True, "run_id": run_id, "summary_path": str(summary_path)}


tool_spec = {
    "name": "plan_and_run_5_steps",
    "description": "Run 50 iterations of a 5-step safe Notepad workflow with logging.",
    "schema": {
        "type": "object",
        "properties": {
            "app_executable_path": {"type": "string"},
            "app_name": {"type": "string"},
            "iterations": {"type": "integer"},
            "force_rebuild": {"type": "boolean"},
            "run_id": {"type": "string"},
        },
        "required": ["app_executable_path"],
    },
}

