from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from .common import add_repo_to_syspath, build_run_id, ensure_run_dirs, sanitize_path, write_json

add_repo_to_syspath()

from utils.fdom.state_manager import StateManager  # noqa: E402


def _load_fdom(app_name: str) -> StateManager:
    sm = StateManager(app_name)
    if not sm.fdom_data:
        sm._load_fdom_from_file()
    return sm


def _extract_mappings(gemini_results: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    mappings: Dict[str, Dict[str, Any]] = {}
    if not gemini_results:
        return mappings
    if "images" in gemini_results:
        for image in gemini_results.get("images", []):
            if not image.get("icons"):
                continue
            for icon in image["icons"]:
                icon_id = icon.get("id")
                if icon_id:
                    mappings[icon_id] = {
                        "g_icon_name": icon.get("name", "unknown"),
                        "g_brief": icon.get("usage", "Not analyzed by Gemini"),
                        "g_enabled": icon.get("enabled", True),
                        "g_interactive": icon.get("interactive", True),
                        "g_type": icon.get("type", "icon"),
                    }
    elif "seraphine_gemini_groups" in gemini_results:
        # Flatten seraphine_gemini_groups: group -> list of items with ids like H1_1
        for group_name, items in gemini_results["seraphine_gemini_groups"].items():
            if isinstance(items, dict):
                for item_id, payload in items.items():
                    mappings[item_id] = {
                        "g_icon_name": payload.get("g_icon_name", "unknown"),
                        "g_brief": payload.get("g_brief", "Not analyzed by Gemini"),
                        "g_enabled": payload.get("g_enabled", True),
                        "g_interactive": payload.get("g_interactive", True),
                        "g_type": payload.get("g_type", "icon"),
                    }
    return mappings


def integrate_gemini(params: Dict[str, Any]) -> Dict[str, Any]:
    app_name = params.get("app_name") or "unknown"
    gemini_results = params.get("gemini_results") or {}
    run_id = params.get("run_id") or build_run_id(app_name)

    mappings = _extract_mappings(gemini_results)
    if not mappings:
        return {"success": False, "error": "no_mappings_found"}

    sm = _load_fdom(app_name)
    fdom = sm.fdom_data or {}
    states = fdom.get("states", {})
    updated = 0
    for state_id, state in states.items():
        nodes = state.get("nodes", {})
        for node_id, node_data in nodes.items():
            if node_id in mappings:
                node_data.update(mappings[node_id])
                updated += 1

    sm.fdom_data = fdom
    sm.save_fdom_to_file()

    dirs = ensure_run_dirs(run_id)
    write_json(dirs["root"] / "fdom_final.json", fdom)

    return {
        "success": True,
        "updated_nodes": updated,
        "fdom_path": sanitize_path(sm.fdom_file_path) if hasattr(sm, "fdom_file_path") else None,
    }


tool_spec = {
    "name": "integrate_gemini",
    "description": "Map Gemini results into fDOM nodes and persist fdom.json.",
    "schema": {
        "type": "object",
        "properties": {
            "app_name": {"type": "string"},
            "gemini_results": {"type": "object"},
            "run_id": {"type": "string"},
        },
        "required": ["app_name", "gemini_results"],
    },
}

