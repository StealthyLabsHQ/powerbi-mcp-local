# AGENTS.md

Operating guide for AI coding agents (Codex, Claude Code, etc.) contributing to `powerbi-mcp-local`.

---

## Repository purpose

MCP server exposing 58 tools that connect to a running Power BI Desktop local engine (Analysis Services via TOM/ADOMD.NET) and automate model, query, Power Query, Excel, and report-layout operations.

Target OS: Windows. Runtime: Python 3.11+ with `pythonnet` (.NET bridge).

---

## Project layout

```
src/
  server.py            # MCP tool registration (stdio + SSE transports)
  pbi_connection.py    # Port discovery, TOM/ADOMD connection manager
  security.py          # Path, query, and payload validation middleware
  tools/
    model.py           # Tables, columns, model info, export
    measures.py        # Measure CRUD, .dax bulk import
    relationships.py   # Relationships CRUD
    query.py           # DAX execution, RLS role, trace, refresh
    power_query.py     # Power Query (M) partition tools
    excel.py           # Excel read/write/format pipeline
    visuals.py         # Report extract/compile, pages, visuals, themes
docs/                  # Multi-platform setup guides
specs/                 # Technical specs for new tool layers
tests/                 # LOCAL-ONLY - ignored by git in this repository
CHANGELOG.md
README.md
SECURITY.md
```

---

## Hard rules

1. **Never commit anything under `tests/`.** The folder is gitignored and local-only in this repository. If a task asks for test coverage, you can create/modify tests locally for validation, but do not `git add` them.

2. **Never commit files matching `.gitignore`.** Check `.gitignore` before staging.

3. **Never push to `main` without explicit user instruction.** Create a commit and stop. Let the user review `git status` / `git diff` before pushing.

4. **Never bypass security middleware.** All file paths go through `SECURITY.validate_*`. All DAX/M strings go through the injection guards. Do not add `# noqa` or comment out these checks.

5. **Never introduce new runtime dependencies** unless the task explicitly says so. Stick to what's in `requirements.txt`: `mcp[cli]`, `openpyxl`, `pbi-pyadomd`, `pythonnet`, `psutil`.

---

## Code conventions

- **Return shape:** every tool returns `ok(message, **fields)` on success or raises a `PowerBIError` subclass (`PowerBIValidationError`, `PowerBINotFoundError`, `PowerBIConfigurationError`). The MCP framework converts raised exceptions into `err(...)` dicts.

- **Tool signatures:** keyword-only args after `manager`. Example:
  ```python
  def pbi_do_thing_tool(manager: Any, *, foo: str, bar: int = 10) -> dict[str, Any]:
  ```

- **Connection lifecycle:** read operations use `manager.run_read(name, reader_fn)`. Writes use `manager.execute_write(name, mutator_fn)`. Never open raw TOM/ADOMD connections outside these helpers unless absolutely required (and document why).

- **Error messages:** lead with the user-visible fact, include a `details={...}` dict for machine consumption. Never leak stack traces into the message string.

- **No comments describing what the code does.** Only comment *why* when the reason is non-obvious (a PBI quirk, a workaround for a specific bug, a hidden invariant).

- **Type hints:** required on all public functions. Use `dict[str, Any]`, not `Dict[str, Any]` (Python 3.11+).

---

## Registration checklist for new tools

When adding a new `pbi_*_tool` or `excel_*_tool`:

1. Define the function in the appropriate `src/tools/*.py` file.
2. Export it in `src/tools/__init__.py`.
3. Register it in `src/server.py` with the correct `@mcp.tool()` wrapper (read vs write - update `READ_TOOLS`/`WRITE_TOOLS`/`DESTRUCTIVE_TOOLS` in `security.py` as needed).
4. Update the tool count in `README.md` (badge + catalog section).
5. Add a `CHANGELOG.md` entry under `[Unreleased]`.

---

## Known Power BI / pbi-tools gotchas

- **`$Measures` entity does not work at runtime.** Measure references in `prototypeQuery.From[]` must use the measure's actual home table. Resolve the home table from `{extract_folder}/Model/tables/<Table>/measures/<Measure>.dax`.

- **Gauge projection role is `"Y"`, not `"Value"`.** Easy to miss when copy-pasting from other visual tools.

- **`prototypeQuery.Select[N].Name` must be the short column name** (e.g. `"Year"`, not `"Dim_Date.Year"`). `queryRef` in projections must match that short name.

- **`Layout` file is UTF-16-LE encoded, no BOM.** Read with `raw.decode('utf-16-le')`, write with `json.dumps(...).encode('utf-16-le')`.

- **`SecurityBindings` must be stripped** whenever rebuilding a PBIX ZIP. It is a DPAPI blob tied to the original Layout bytes and invalidates the file if left alongside a modified Layout.

- **`pbi-tools compile` fails with "Compiling a project containing a data model into a PBIX file is not supported"** whenever the extract folder has a `DataModel` entry. Current workaround: patch only the `Report/Layout` entry back into the original PBIX ZIP manually (there is no built-in `pbi_patch_layout` tool in this repo).

- **Power BI Desktop locks the PBIX file.** Writes fail with `PermissionError` until the process is closed. Tools that write PBIX should accept a `force: bool = False` argument that runs `taskkill /F /IM PBIDesktop.exe` on Windows before writing.

- **`.dax` files in the wild contain `//` and `/* */` comments.** The import parser must strip them before splitting measure blocks.

---

## Git workflow

```bash
# Before starting work
git status              # clean tree?
git log --oneline -5    # match commit-message style

# While working
# ... write code, test locally ...

# Before committing
git status              # any accidental tests/ or __pycache__/ files?
git diff --stat         # does the change match what was asked?

# Commit
git add <specific files>   # never `git add .` or `git add -A`
git commit -m "type: short subject

Body paragraph explaining why and what."

# Push — only when user says so
```

**Commit subject format:** `type: short subject` where `type` is one of `feat`, `fix`, `docs`, `refactor`, `test`, `chore`, `security`.

**Never:**
- `--amend` a pushed commit
- `git push --force` without explicit instruction
- Commit with `--no-verify` or `-c commit.gpgsign=false`

---

## When in doubt

1. Read the relevant `src/tools/*.py` file end-to-end before editing — patterns are consistent across tools, and the right answer is usually "do what the existing tools do."
2. Search memory observations for related sessions (`claude-mem` / prior commits) before inventing a new approach to a known problem.
3. If a bug reproduces only with a live Power BI Desktop process, document the reproduction steps in the commit message body so the next agent can test without guessing.

---

## Out of scope

- Windows registry edits, system-level .NET/ADOMD installs (documentation only).
- Power BI Service (cloud) APIs — this project is strictly local.
- Custom visuals marketplace management — not automatable via the local engine.
- Live visual preview — inherent to Power BI Desktop UI, not automatable.
