"""Print the number of @mcp.tool() handlers registered by the live server.

Used to keep the README "Tools NN" badge honest. Run from the repo root:

    python scripts/tool_count.py

Exits non-zero if the FastMCP registry could not be inspected.
"""

from __future__ import annotations

import sys
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parent.parent
    src_dir = repo_root / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))

    # Importing server.py triggers every wrappers/<domain>.py side-effect
    # registration, so the FastMCP tool manager ends up with the same set
    # of @mcp.tool() handlers a launched server would expose.
    import server  # noqa: F401  -- side-effect registration of all wrappers
    from mcp_core import mcp

    manager = getattr(mcp, "_tool_manager", None)
    tools_map = getattr(manager, "_tools", None)
    if not isinstance(tools_map, dict):
        print("error: FastMCP tool manager is not introspectable", file=sys.stderr)
        return 1

    pbi_tools = sorted(name for name in tools_map if name.startswith("pbi_"))
    excel_tools = sorted(name for name in tools_map if name.startswith("excel_"))
    other_tools = sorted(name for name in tools_map if not name.startswith(("pbi_", "excel_")))

    print(f"total: {len(tools_map)}")
    print(f"  pbi_*  : {len(pbi_tools)}")
    print(f"  excel_*: {len(excel_tools)}")
    if other_tools:
        print(f"  other  : {len(other_tools)} ({', '.join(other_tools)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
