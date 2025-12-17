from __future__ import annotations

import asyncio
import json
import sys
from typing import Any, Callable, Dict

from mcp_fdom.tools import capture, gemini, integrate, execute, orchestrate


TOOL_HANDLERS: Dict[str, Dict[str, Any]] = {
    capture.tool_spec["name"]: {"spec": capture.tool_spec, "handler": capture.capture_ui_state},
    gemini.tool_spec["name"]: {"spec": gemini.tool_spec, "handler": gemini.analyze_with_gemini},
    integrate.tool_spec["name"]: {"spec": integrate.tool_spec, "handler": integrate.integrate_gemini},
    execute.tool_spec["name"]: {"spec": execute.tool_spec, "handler": execute.execute_step},
    orchestrate.tool_spec["name"]: {"spec": orchestrate.tool_spec, "handler": orchestrate.plan_and_run_5_steps},
}


def list_tools() -> Dict[str, Any]:
    return {"tools": [entry["spec"] for entry in TOOL_HANDLERS.values()]}


def invoke_tool(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
    if tool_name not in TOOL_HANDLERS:
        return {"success": False, "error": f"unknown_tool:{tool_name}"}
    handler: Callable[[Dict[str, Any]], Dict[str, Any]] = TOOL_HANDLERS[tool_name]["handler"]
    try:
        return handler(arguments or {})
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "error": f"exception:{exc}"}


async def _read_stdin_line() -> str:
    """Read a single line from stdin using a thread to avoid Proactor pipe issues on Windows."""
    return await asyncio.to_thread(sys.stdin.readline)


async def main() -> None:
    while True:
        line = await _read_stdin_line()
        if not line:
            break
        try:
            req = json.loads(line.strip())
        except json.JSONDecodeError:
            continue

        resp = {"jsonrpc": "2.0", "id": req.get("id")}
        method = req.get("method")

        if method == "list_tools":
            resp["result"] = list_tools()
        elif method == "invoke_tool":
            params = req.get("params", {})
            tool = params.get("tool")
            args = params.get("arguments", {})
            resp["result"] = invoke_tool(tool, args)
        else:
            resp["error"] = {"code": -32601, "message": f"Unknown method: {method}"}

        sys.stdout.write(json.dumps(resp) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    asyncio.run(main())

