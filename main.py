"""
Entry point wrapper to run the MCP fDOM server.

This exists so tools that expect `python main.py` succeed.
"""

import asyncio

from mcp_fdom.server import main as run_server


if __name__ == "__main__":
    asyncio.run(run_server())

