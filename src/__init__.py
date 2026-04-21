"""Power BI MCP server package.

Bootstraps sys.path so flat imports (e.g. ``from pbi_connection import ...``)
work in both script mode (``python src/server.py``) and installed-package mode
(``powerbi-mcp-local`` entry point, ``python -m src.server``).
"""

from __future__ import annotations

import os
import sys

_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)
