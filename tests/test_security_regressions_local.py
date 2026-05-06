from __future__ import annotations

import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from pbi_connection import PowerBIValidationError  # noqa: E402
from m_expression_security import strip_m_literals_and_comments  # noqa: E402
from tools.power_query import (  # noqa: E402
    _build_csv_m,
    _build_excel_m,
    _build_folder_m,
    _validate_m_expression,
)
from tools.query import pbi_execute_dax_as_role_tool  # noqa: E402


class DummyManager:
    pass


def test_pbi_execute_dax_as_role_blocks_connection_string_separator() -> None:
    with pytest.raises(PowerBIValidationError):
        pbi_execute_dax_as_role_tool(
            DummyManager(),
            query='EVALUATE ROW("x", 1)',
            role="Sales",
            username="user@example.com;Roles=Admin",
        )


def test_validate_m_expression_blocks_dynamic_shared_bypass() -> None:
    with pytest.raises(PowerBIValidationError):
        _validate_m_expression(
            'let Fn = Record.Field(#shared, "Web" & ".Contents") in Fn("https://example.com")'
        )


def test_validate_m_expression_ignores_blocked_tokens_inside_comments() -> None:
    _validate_m_expression(
        "let\n"
        "    /* Web.Contents(\"https://example.com\") */\n"
        "    Source = Table.FromRows({{\"x\"}}, {\"A\"})\n"
        "in\n"
        "    Source"
    )


def test_validate_m_expression_ignores_blocked_tokens_inside_escaped_strings() -> None:
    _validate_m_expression(
        'let Message = "Web.Contents(""https://example.com"")", Clean = Text.Trim(Message) in Clean'
    )


def test_validate_m_expression_allows_locally_defined_safe_function() -> None:
    _validate_m_expression(
        "let\n"
        "    Clean = (value as text) => Text.Trim(value),\n"
        "    Result = Clean(\" hello \")\n"
        "in\n"
        "    Result"
    )


def test_strip_m_literals_and_comments_redacts_comment_and_string_content() -> None:
    sanitized = strip_m_literals_and_comments(
        'let TextValue = "Web.Contents(""x"")", /* Odbc.Query("y") */ Result = Text.Trim(TextValue) in Result'
    )
    assert "Web.Contents" not in sanitized
    assert "Odbc.Query" not in sanitized
    assert "Text.Trim" in sanitized


@pytest.mark.parametrize(
    "expression",
    [
        _build_excel_m(r"C:\safe\data.xlsx", "Sheet1", True),
        _build_csv_m(r"C:\safe\data.csv"),
        _build_folder_m(r"C:\safe\folder", extension_filter=".csv"),
    ],
)
def test_validate_m_expression_allows_local_import_builders(expression: str) -> None:
    _validate_m_expression(expression)
