"""Quick sanity check: import the server module and list every tool.

This does NOT start the stdio loop and does NOT touch COM. It's the fastest
way to confirm the package is wired up correctly after editing a tools module.

    python tests/smoke/list_tools.py
"""

from __future__ import annotations

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hwpx_hwp_mcp import server


def main() -> None:
    tools = asyncio.run(server.mcp.list_tools())
    print(f"{len(tools)} tools registered:")
    for tool in sorted(tools, key=lambda t: t.name):
        first_line = (tool.description or "").splitlines()[0][:80]
        print(f"  {tool.name:<28} {first_line}")


if __name__ == "__main__":
    main()
