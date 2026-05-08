"""MCP @mcp.tool() wrapper modules — split by domain.

Each module registers its wrappers as a side-effect of import.
server.py imports every wrapper module to populate the FastMCP
tool registry, then runs main().
"""
