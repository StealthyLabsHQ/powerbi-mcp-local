# Architecture

## High level

```
                 stdio / SSE
MCP client  <─────────────────>  src/server.py
                                       │
                                       ├─ src/mcp_core.py
                                       │    FastMCP instance, CONNECTION_MANAGER,
                                       │    _run, audit, profile filter,
                                       │    PID lock, parent watcher
                                       │
                                       ├─ src/wrappers/
                                       │    14 thin domain modules
                                       │    (1 register_tool() call per pbi_*_tool)
                                       │
                                       ├─ src/tools/
                                       │    business logic (the *_tool functions)
                                       │
                                       ├─ src/pbi_connection.py
                                       │    TOM + ADOMD + instance discovery,
                                       │    write/read paths, persistence warnings
                                       │
                                       └─ src/security.py
                                            path / DAX / M / payload guards,
                                            tool catalogue (READ/WRITE/DESTRUCTIVE/GRADING)
```

## Module layering

| Layer | Path | Responsibility |
|---|---|---|
| **Entry** | `src/server.py` | argparse, transport launch, `@mcp.resource()`/`@mcp.prompt()`, calls `wrappers.register_all()`. Stays thin (~320 L). |
| **Runtime** | `src/mcp_core.py` | The single FastMCP instance, `CONNECTION_MANAGER`, `_run` (audit + error normalisation), registry audit, profile filter, PID lock, parent watcher. |
| **Wrappers** | `src/wrappers/__init__.py` | `register_all()` iterates `tools.__all__` and calls `register_tool(pbi_xxx_tool)` for every exported `*_tool` — the tool list has a single source of truth. |
| **Tools** | `src/tools/<domain>.py` | Business logic — the `pbi_xxx_tool(manager, *, …)` functions called by the wrappers. The visual surface lives in the `src/tools/visuals/` package (17 submodules). |
| **Connection** | `src/pbi_connection.py` | Instance discovery, TOM / ADOMD bring-up, write helpers (`execute_write` injects the `persistence: memory_only` warning), operation history. |
| **Security** | `src/security.py` | Path traversal, allowed dirs, redaction, DAX / M sanitisers, tool category sets used by `--profile`. |

## Wrapper auto-registration

Registration is fully derived from `tools.__all__` — no per-domain wrapper files, no `@mcp.tool()` boilerplate:

```python
# src/server.py
import wrappers

wrappers.register_all()   # registers every tools.__all__ *_tool
```

Adding a tool = implement `pbi_xxx_tool` in `src/tools/<domain>.py` + export it from `src/tools/__init__.py`. Nothing else.

`register_tool()` introspects the underlying tool's signature, drops `manager` from the public schema, injects `CONNECTION_MANAGER` at call time, and pipes the result through `mcp_core._run`. Manager-injection style is auto-detected:

| Underlying signature | Injection mode |
|---|---|
| `tool(manager, *, …)` (positional) | passed positionally |
| `tool(…, *, manager=None)` (keyword-only) | passed as kwarg with `setdefault` |
| `tool(…)` (no manager) | none |

The generated wrapper has `__signature__` reassigned so FastMCP picks up the right JSON parameter schema (typed args, not `**kwargs`).

If a wrapper ever needs extra logic, call `register_tool(fn, name=…, inject_manager=…, docstring=…)` manually before `register_all()` — already-registered names are skipped by the loop.

## visuals/ package

`src/tools/visuals/` is the largest tool surface (~36 visual writers + page management + extract / compile / patch). It is split into focused submodules along the dependency graph:

```
_base.py         constants, error classes, _run helper
   ↑
_paths.py        path resolution
_refs.py         field reference normalisation
   ↑
_layout.py       load / atomic save / dry_run / page helpers
_formatting.py   literal / colour / title encoders
_home_tables.py  measure → home table resolution
   ↑
_bindings.py     prototype/select builders + live validators
   ↑
_containers.py   visual container construction + append flow
   ↑
_charts.py       6 cartesian chart tools
_cards.py        5 card-style tools
_structure.py    4 structure tools (table, slicer, matrix, map)
_pages.py        page-level tools
_ops.py          remove / move / format / convert / auto-grid / patch
_design.py       DESIGN_PRESETS + apply_theme / apply_design / build_dashboard
_repair.py       validate / repair report fields
_io.py           pbi-tools CLI + zip extraction + PowerShell graceful close
   ↑
_dispatcher.py   pbi_add_visual_tool + _VISUAL_TYPE_DISPATCH
```

`__init__.py` re-exports the public surface so historic `from tools.visuals import …` imports keep working.

## Layout writes

Every layout mutation goes through `_save_layout(folder, layout)` in `_layout.py`:

1. If a `dry_run_layout_writes()` context is active, the call is recorded into a thread-local log and the disk write is skipped.
2. Otherwise: serialise to `Layout.tmp.<pid>` next to the target, copy the previous-good Layout to `Layout.bak`, then `os.replace` the temp onto Layout. Atomic on Windows + POSIX. A crash mid-write leaves the original Layout intact.

## TOM mutations

Every TOM write (measures, relationships, columns, roles, calc groups, Power Query, TMDL) routes through `PowerBIConnectionManager.execute_write` in `pbi_connection.py`. The returned payload always carries:

```python
{
    "save_result": <TOM SaveChanges result>,
    "persistence": {
        "scope": "memory_only",
        "hint": "Change committed to the AS engine in memory. The .pbix on disk is unchanged until Power BI Desktop saves the file. …",
    },
    ...
}
```

This makes the in-memory ↔ on-disk dual-state explicit on every write.

## Profiles

`--profile` filters which `@mcp.tool()` are exposed:

| Profile | Set | Purpose |
|---|---|---|
| `all` (default) | every tool | full surface |
| `readonly` | `READ_TOOLS` (~56) | inspection only |
| `write` | `READ ∪ WRITE_TOOLS` | reads + non-destructive writes |
| `grading` | `GRADING_TOOLS` (25) | analysis + scoring tools for evaluation workflows |

The `READ_TOOLS` / `WRITE_TOOLS` / `DESTRUCTIVE_TOOLS` / `GRADING_TOOLS` sets live in `src/security.py`.

## Lifecycle

`mcp_core._acquire_single_instance_lock()` writes `%TEMP%/powerbi-mcp.pid` at startup. If a previous PID is still alive, it's killed (`psutil.Process.kill`). `atexit` + `SIGINT`/`SIGTERM` handlers remove the lock on clean exit.

`mcp_core._start_parent_watcher()` polls the parent process every 2 s. If the parent (Claude Code, Codex, …) disappears or becomes a zombie, the server releases the PID lock and `os._exit(0)` so it never lingers.

Both are stdio-only — SSE binds a port so the OS handles single-instance.

## Registry audit

`mcp_core._audit_tool_registry(strict=…)` cross-checks `tools/__all__` against the registered FastMCP tools. Orphan implementations (tools in `__all__` but not wrapped) and unknown wrappers (registered but with no matching `pbi_*_tool`) are logged. With `PBI_MCP_STRICT_REGISTRY=1` (always set in CI), drift fails the build.

## Tests

`tests/` carries 167 offline unit tests (mocks for the live PBI engine). Live-only scripts (`tests/demo_*.py`, `tests/smoke_e2e.py`, `tests/test_*_local.py`) stay gitignored. Coverage baseline: 55% enforced in CI on Windows.
