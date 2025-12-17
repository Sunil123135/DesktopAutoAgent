from __future__ import annotations

import hashlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple


PROJECT_ROOT = Path(__file__).resolve().parents[2]
UTILS_DIR = PROJECT_ROOT / "utils"
FDOM_DIR = UTILS_DIR / "fdom"


def add_repo_to_syspath() -> None:
    # Ensure project root (so `utils.seraphine` works)
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    # Ensure utils/ is importable
    if str(UTILS_DIR) not in sys.path:
        sys.path.insert(0, str(UTILS_DIR))
    # Ensure utils/fdom/ is importable for bare imports like `config_manager`
    if str(FDOM_DIR) not in sys.path:
        sys.path.insert(0, str(FDOM_DIR))


SAFE_APPS = {
    "notepad.exe",
    "notepad",
    "calc.exe",
    "calculator.exe",
    "mspaint.exe",
    "paint.exe",
}

BLOCKED_TITLE_KEYWORDS = {
    "sign in",
    "signin",
    "login",
    "password",
    "otp",
    "bank",
    "wallet",
    "payment",
    "pay",
    "invoice",
    "billing",
}

BLOCKED_OCR_KEYWORDS = BLOCKED_TITLE_KEYWORDS | {
    "2fa",
    "verification",
    "credential",
    "ssn",
    "pin",
}


def is_app_allowed(app_path: str, app_name: Optional[str]) -> Tuple[bool, str]:
    exe = Path(app_path).name.lower()
    if exe in SAFE_APPS:
        return True, "allowed"
    if app_name and app_name.lower() in SAFE_APPS:
        return True, "allowed"
    return False, f"blocked_app:{exe}"


def safety_check(window_title: Optional[str] = None, ocr_text: Optional[str] = None) -> Tuple[bool, str]:
    title = (window_title or "").lower()
    text = (ocr_text or "").lower()
    for kw in BLOCKED_TITLE_KEYWORDS:
        if kw in title:
            return False, f"blocked_title:{kw}"
    for kw in BLOCKED_OCR_KEYWORDS:
        if kw in text:
            return False, f"blocked_ocr:{kw}"
    return True, "ok"


def build_run_id(app_name: str) -> str:
    stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    safe_app = app_name.replace(" ", "_")
    return f"{stamp}__{safe_app}"


def ensure_run_dirs(run_id: str) -> Dict[str, Path]:
    runs_root = PROJECT_ROOT / "runs" / run_id
    screens_dir = runs_root / "screens"
    runs_root.mkdir(parents=True, exist_ok=True)
    screens_dir.mkdir(parents=True, exist_ok=True)
    return {"root": runs_root, "screens": screens_dir}


def hash_file(path: Path) -> Optional[str]:
    try:
        h = hashlib.sha256()
        with path.open("rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def sanitize_path(path: Optional[str]) -> Optional[str]:
    if not path:
        return None
    try:
        p = Path(path)
        return str(p.relative_to(PROJECT_ROOT))
    except Exception:
        return str(path)


def snapshot_state(interactor: Any, run_id: str) -> Dict[str, Any]:
    state_id = getattr(interactor, "current_state_id", "root")
    sm = interactor.state_manager
    fdom = sm.fdom_data or {}
    states = fdom.get("states", {})
    state_data = states.get(state_id, {})
    nodes = state_data.get("nodes", {})

    def node_to_dict(node_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "node_id": node_id,
            "bbox": data.get("bbox"),
            "g_icon_name": data.get("g_icon_name"),
            "g_brief": data.get("g_brief"),
            "g_enabled": data.get("g_enabled"),
            "g_interactive": data.get("g_interactive"),
            "g_type": data.get("g_type"),
            "source": data.get("source"),
            "group": data.get("group"),
            "state": state_id,
        }

    pending_nodes = list(getattr(sm, "pending_nodes", []))
    edges = fdom.get("edges", [])
    screenshot_path = state_data.get("image")

    return {
        "app": interactor.app_name,
        "run_id": run_id,
        "state_id": state_id,
        "fdom_path": sanitize_path(sm.fdom_file_path) if hasattr(sm, "fdom_file_path") else None,
        "screenshot_path": sanitize_path(screenshot_path),
        "nodes": [node_to_dict(nid, data) for nid, data in nodes.items()],
        "pending_nodes": pending_nodes,
        "edges": edges,
    }


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

