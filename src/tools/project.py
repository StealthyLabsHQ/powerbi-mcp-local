"""Power BI Project (PBIP/TMDL) offline helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pbi_connection import PowerBINotFoundError, PowerBIValidationError, ok
from security import resolve_local_path


PROJECT_FILE_EXTENSIONS = {".pbip"}
TMDL_FILE_EXTENSIONS = {".tmdl"}


def _resolve_project_root(project_path: str) -> Path:
    path = resolve_local_path(project_path, must_exist=True, allowed_extensions=None)
    if path.is_file():
        if path.suffix.casefold() != ".pbip":
            raise PowerBIValidationError("Project file must have a .pbip extension.", details={"path": str(path)})
        return path.parent
    if not path.is_dir():
        raise PowerBINotFoundError("Project path was not found.", details={"path": str(path)})
    return path


def _definition_folder(project_path: str) -> Path:
    root = _resolve_project_root(project_path)
    direct_candidates = [
        root / "definition",
        root / "SemanticModel" / "definition",
    ]
    direct_candidates.extend(sorted(root.glob("*.SemanticModel/definition")))
    for candidate in direct_candidates:
        if candidate.is_dir():
            return candidate.resolve()
    for candidate in sorted(root.rglob("definition")):
        if candidate.is_dir() and any(candidate.glob("*.tmdl")):
            return candidate.resolve()
    raise PowerBINotFoundError(
        "No TMDL definition folder was found under the Power BI project.",
        details={"project_path": str(root)},
    )


def _resolve_tmdl_file(project_path: str, relative_file: str, *, must_exist: bool) -> Path:
    definition = _definition_folder(project_path)
    relative = Path(str(relative_file).replace("\\", "/"))
    if relative.is_absolute() or any(part in {"..", ""} for part in relative.parts):
        raise PowerBIValidationError("relative_file must be a relative path inside the definition folder.")
    if relative.suffix.casefold() != ".tmdl":
        raise PowerBIValidationError("Only .tmdl files can be edited.", details={"relative_file": str(relative)})
    target = (definition / relative).resolve(strict=False)
    try:
        target.relative_to(definition)
    except ValueError as exc:
        raise PowerBIValidationError("TMDL path must stay inside the definition folder.") from exc
    if must_exist and not target.exists():
        raise PowerBINotFoundError("TMDL file was not found.", details={"relative_file": str(relative)})
    return target


def _basic_tmdl_issues(path: Path, content: str) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    if not content.strip():
        issues.append({"relative_file": path.name, "issue": "empty_file"})
    for line_number, line in enumerate(content.splitlines(), start=1):
        if "\t" in line:
            issues.append({"relative_file": path.name, "line": line_number, "issue": "tab_indentation"})
    return issues


def _validate_measure_name(name: str) -> str:
    value = str(name).strip()
    if not value:
        raise PowerBIValidationError("measure_name cannot be empty.")
    if any(char in value for char in "\r\n="):
        raise PowerBIValidationError("measure_name cannot contain line breaks or '='.", details={"measure_name": name})
    return value


def _measure_block(name: str, expression: str, format_string: str = "", display_folder: str = "") -> list[str]:
    if not expression.strip():
        raise PowerBIValidationError("expression cannot be empty.")
    lines = [f"    measure '{name}' = {expression.strip()}"]
    if format_string:
        lines.append(f"        formatString: {format_string!r}")
    if display_folder:
        lines.append(f"        displayFolder: {display_folder!r}")
    return lines


def _find_measure_block(lines: list[str], name: str) -> tuple[int, int] | None:
    marker = f"measure '{name.casefold()}'"
    for index, line in enumerate(lines):
        stripped = line.strip().casefold()
        if not stripped.startswith(marker):
            continue
        end = index + 1
        while end < len(lines):
            next_line = lines[end]
            if next_line.startswith("    ") and not next_line.startswith("        ") and next_line.strip():
                break
            end += 1
        return index, end
    return None


def pbi_list_tmdl_files_tool(project_path: str) -> dict[str, Any]:
    """List TMDL files in a Power BI Project semantic model definition folder."""
    definition = _definition_folder(project_path)
    files = []
    issues = []
    for path in sorted(definition.rglob("*.tmdl")):
        content = path.read_text(encoding="utf-8")
        relative = str(path.relative_to(definition)).replace("\\", "/")
        files.append({"relative_file": relative, "bytes": path.stat().st_size})
        for issue in _basic_tmdl_issues(path, content):
            issue["relative_file"] = relative
            issues.append(issue)
    return ok(
        "TMDL files listed successfully.",
        project_path=str(_resolve_project_root(project_path)),
        definition_folder=str(definition),
        files=files,
        file_count=len(files),
        issues=issues,
    )


def pbi_read_tmdl_file_tool(project_path: str, relative_file: str) -> dict[str, Any]:
    """Read one TMDL file from a Power BI Project definition folder."""
    path = _resolve_tmdl_file(project_path, relative_file, must_exist=True)
    definition = _definition_folder(project_path)
    content = path.read_text(encoding="utf-8")
    return ok(
        "TMDL file read successfully.",
        project_path=str(_resolve_project_root(project_path)),
        definition_folder=str(definition),
        relative_file=str(path.relative_to(definition)).replace("\\", "/"),
        content=content,
        issues=_basic_tmdl_issues(path, content),
    )


def pbi_write_tmdl_file_tool(
    project_path: str,
    relative_file: str,
    content: str,
    create: bool = False,
) -> dict[str, Any]:
    """Create or overwrite one TMDL file inside a Power BI Project definition folder."""
    if not content.strip():
        raise PowerBIValidationError("content cannot be empty.")
    path = _resolve_tmdl_file(project_path, relative_file, must_exist=not create)
    definition = _definition_folder(project_path)
    existed = path.exists()
    if not existed and not create:
        raise PowerBINotFoundError("TMDL file was not found.", details={"relative_file": relative_file})
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")
    return ok(
        "TMDL file written successfully.",
        project_path=str(_resolve_project_root(project_path)),
        definition_folder=str(definition),
        relative_file=str(path.relative_to(definition)).replace("\\", "/"),
        action="updated" if existed else "created",
        bytes=path.stat().st_size,
        issues=_basic_tmdl_issues(path, content),
    )


def pbi_patch_tmdl_measure_tool(
    project_path: str,
    table_file: str,
    measure_name: str,
    expression: str,
    format_string: str = "",
    display_folder: str = "",
    overwrite: bool = True,
) -> dict[str, Any]:
    """Create or replace a measure block in one table TMDL file."""
    name = _validate_measure_name(measure_name)
    path = _resolve_tmdl_file(project_path, table_file, must_exist=True)
    definition = _definition_folder(project_path)
    content = path.read_text(encoding="utf-8")
    lines = content.splitlines()
    block = _measure_block(name, expression, format_string, display_folder)
    existing = _find_measure_block(lines, name)
    if existing is None:
        if lines and lines[-1].strip():
            lines.append("")
        lines.extend(block)
        action = "created"
    else:
        if not overwrite:
            raise PowerBIValidationError("Measure already exists and overwrite=False.", details={"measure_name": name})
        start, end = existing
        lines[start:end] = block
        action = "updated"
    new_content = "\n".join(lines).rstrip() + "\n"
    path.write_text(new_content, encoding="utf-8", newline="\n")
    return ok(
        "TMDL measure patched successfully.",
        project_path=str(_resolve_project_root(project_path)),
        definition_folder=str(definition),
        relative_file=str(path.relative_to(definition)).replace("\\", "/"),
        measure=name,
        action=action,
        issues=_basic_tmdl_issues(path, new_content),
    )


__all__ = [
    "pbi_list_tmdl_files_tool",
    "pbi_patch_tmdl_measure_tool",
    "pbi_read_tmdl_file_tool",
    "pbi_write_tmdl_file_tool",
]
