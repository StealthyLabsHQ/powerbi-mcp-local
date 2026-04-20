"""Lexical validation helpers for Power Query M expressions."""

from __future__ import annotations

import re

from pbi_connection import PowerBIValidationError


_M_BLOCKED_FUNCTION_PATTERNS = (
    r"\bWeb\.Contents\b",
    r"\bWeb\.Page\b",
    r"\bWeb\.BrowserContents\b",
    r"\bExpression\.Evaluate\b",
    r"\bValue\.NativeQuery\b",
    r"\bOData\.Feed\b",
    r"\bSql\.Database\b",
    r"\bSql\.Databases\b",
    r"\bOracle\.Database\b",
    r"\bPostgreSQL\.Database\b",
    r"\bMySQL\.Database\b",
    r"\bOdbc\.DataSource\b",
    r"\bOdbc\.Query\b",
    r"\bOleDb\.DataSource\b",
    r"\bOleDb\.Query\b",
    r"\bSharePoint\.\w+",
    r"\bActiveDirectory\.\w+",
    r"\bAzureStorage\.\w+",
    r"#shared\b",
)
M_BLOCKED_FUNCTIONS = tuple(re.compile(pattern, re.IGNORECASE) for pattern in _M_BLOCKED_FUNCTION_PATTERNS)
M_ALLOWED_FUNCTIONS = {
    "Csv.Document",
    "Excel.Workbook",
    "File.Contents",
    "Folder.Contents",
    "Folder.Files",
}
M_ALLOWED_PREFIXES = (
    "Binary.",
    "Character.",
    "Comparer.",
    "Date.",
    "DateTime.",
    "DateTimeZone.",
    "Duration.",
    "Json.",
    "Lines.",
    "List.",
    "Logical.",
    "Number.",
    "Record.",
    "Splitter.",
    "Table.",
    "Text.",
    "Time.",
    "Type.",
    "Value.",
)
M_FUNCTION_CALL_RE = re.compile(r"(?<![#\w])(?P<name>[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\s*\(")
M_CUSTOM_FUNCTION_RE = re.compile(r"(?<![#\w])(?P<name>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*\(")


def strip_m_literals_and_comments(text: str) -> str:
    output: list[str] = []
    index = 0
    in_string = False
    in_line_comment = False
    in_block_comment = False
    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""

        if in_line_comment:
            if char in "\r\n":
                in_line_comment = False
                output.append(char)
            else:
                output.append(" ")
            index += 1
            continue

        if in_block_comment:
            if char == "*" and next_char == "/":
                in_block_comment = False
                output.extend("  ")
                index += 2
                continue
            output.append(char if char in "\r\n" else " ")
            index += 1
            continue

        if in_string:
            if char == '"' and next_char == '"':
                output.extend("  ")
                index += 2
                continue
            if char == '"':
                in_string = False
                output.append(" ")
                index += 1
                continue
            output.append(char if char in "\r\n" else " ")
            index += 1
            continue

        if char == "/" and next_char == "/":
            in_line_comment = True
            output.extend("  ")
            index += 2
            continue
        if char == "/" and next_char == "*":
            in_block_comment = True
            output.extend("  ")
            index += 2
            continue
        if char == '"':
            in_string = True
            output.append(" ")
            index += 1
            continue
        output.append(char)
        index += 1
    return "".join(output)


def validate_m_expression_structure(expression: str) -> None:
    text = expression.strip()
    if not text:
        raise PowerBIValidationError("m_expression cannot be empty.")

    stack: list[tuple[str, int]] = []
    pairs = {"(": ")", "[": "]", "{": "}"}
    in_string = False
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string and char in pairs:
            stack.append((char, index))
        elif not in_string and char in pairs.values():
            if not stack or pairs[stack[-1][0]] != char:
                raise PowerBIValidationError(
                    "M expression has unbalanced delimiters.",
                    details={"position": index, "character": char},
                )
            stack.pop()
        index += 1
    if in_string:
        raise PowerBIValidationError("M expression contains an unterminated string literal.")
    if stack:
        opener, position = stack[-1]
        raise PowerBIValidationError(
            "M expression has unbalanced delimiters.",
            details={"position": position, "character": opener},
        )
    if re.match(r"^\s*let\b", text, flags=re.IGNORECASE) and re.search(r"\bin\b", text, flags=re.IGNORECASE) is None:
        raise PowerBIValidationError("M expression starts with 'let' but has no matching 'in' clause.")


def _collect_custom_function_names(sanitized_expression: str) -> set[str]:
    return {match.group("name") for match in M_CUSTOM_FUNCTION_RE.finditer(sanitized_expression)}


def validate_m_expression_policy(expression: str) -> None:
    sanitized = strip_m_literals_and_comments(expression)

    for pattern in M_BLOCKED_FUNCTIONS:
        match = pattern.search(sanitized)
        if match:
            raise PowerBIValidationError(
                f"M expression contains blocked function '{match.group()}'. "
                f"Network/external database access is disabled by default. "
                f"Set PBI_MCP_ALLOW_EXTERNAL_M=1 to allow.",
                details={"blocked_function": match.group(), "position": match.start()},
            )

    custom_functions = _collect_custom_function_names(sanitized)
    disallowed: list[str] = []
    for match in M_FUNCTION_CALL_RE.finditer(sanitized):
        name = match.group("name")
        if name in custom_functions:
            continue
        if name in M_ALLOWED_FUNCTIONS or any(name.startswith(prefix) for prefix in M_ALLOWED_PREFIXES):
            continue
        disallowed.append(name)
    if disallowed:
        raise PowerBIValidationError(
            "M expression contains function calls outside the local-file allowlist.",
            details={"disallowed_functions": sorted(set(disallowed))},
        )


def validate_m_expression(expression: str, *, allow_external: bool = False) -> None:
    validate_m_expression_structure(expression)
    if not allow_external:
        validate_m_expression_policy(expression)


__all__ = [
    "M_ALLOWED_FUNCTIONS",
    "M_ALLOWED_PREFIXES",
    "M_BLOCKED_FUNCTIONS",
    "strip_m_literals_and_comments",
    "validate_m_expression",
    "validate_m_expression_policy",
    "validate_m_expression_structure",
]
