from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional

from .common import add_repo_to_syspath, ensure_run_dirs, hash_file, sanitize_path, write_json

add_repo_to_syspath()

from utils.seraphine import process_image_sync  # noqa: E402


GEMINI_CACHE: Dict[str, Dict[str, Any]] = {}


def _latest_screenshot(app_name: str) -> Optional[Path]:
    shots_dir = Path(__file__).resolve().parents[2] / "apps" / app_name / "screenshots"
    if not shots_dir.exists():
        return None
    pngs = sorted(shots_dir.glob("*.png"), key=lambda p: p.stat().st_mtime, reverse=True)
    return pngs[0] if pngs else None


def analyze_with_gemini(params: Dict[str, Any]) -> Dict[str, Any]:
    app_name = params.get("app_name") or "unknown"
    image_path = params.get("image_path")
    run_id = params.get("run_id")

    if not image_path:
        latest = _latest_screenshot(app_name)
        if not latest:
            return {"success": False, "error": "no_screenshot_found"}
        image_path = str(latest)

    image_path_obj = Path(image_path)
    img_hash = hash_file(image_path_obj)
    if img_hash and img_hash in GEMINI_CACHE:
        return {"success": True, "cached": True, "gemini_results": GEMINI_CACHE[img_hash], "image_path": sanitize_path(image_path)}

    pipeline_result = process_image_sync(str(image_path_obj))
    if pipeline_result is None:
        return {"success": False, "error": "pipeline_failed"}

    gemini_results = None
    if isinstance(pipeline_result, dict):
        if "gemini_results" in pipeline_result:
            gemini_results = pipeline_result["gemini_results"]
        elif "seraphine_gemini_groups" in pipeline_result:
            gemini_results = {
                "analysis_success": True,
                "seraphine_gemini_groups": pipeline_result["seraphine_gemini_groups"],
                "total_icons_found": pipeline_result.get("total_icons_found"),
            }
        else:
            gemini_results = {"analysis_success": True, "seraphine_output": pipeline_result}
    else:
        gemini_results = {"analysis_success": True, "seraphine_output": str(pipeline_result)}

    if img_hash:
        GEMINI_CACHE[img_hash] = gemini_results

    if run_id:
        dirs = ensure_run_dirs(run_id)
        write_json(dirs["root"] / f"gemini_{img_hash or 'result'}.json", gemini_results)

    return {
        "success": True,
        "cached": False,
        "gemini_results": gemini_results,
        "image_path": sanitize_path(image_path),
    }


tool_spec = {
    "name": "analyze_with_gemini",
    "description": "Run Gemini labeling on the latest or provided screenshot (cached by hash).",
    "schema": {
        "type": "object",
        "properties": {
            "app_name": {"type": "string"},
            "image_path": {"type": "string"},
            "run_id": {"type": "string"},
        },
    },
}

