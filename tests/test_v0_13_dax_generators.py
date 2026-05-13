"""v0.13: DAX generator coverage.

Pins the time-intelligence template catalogue, the pattern resolution
(dependency expansion + dedup + case-folding), and the rejection of
unknown patterns. The generators emit DAX from a templated catalogue —
silently regressing a template would propagate to every downstream
measure.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIValidationError
from tools.measures import (
    _DEFAULT_TIME_INTELLIGENCE_PATTERNS,
    _TIME_INTELLIGENCE_TEMPLATES,
    _resolve_ti_patterns,
)


class TimeIntelligenceTemplateTests(unittest.TestCase):
    def test_default_pattern_set(self) -> None:
        self.assertEqual(
            _DEFAULT_TIME_INTELLIGENCE_PATTERNS,
            ["YTD", "MTD", "QTD", "SPY", "YOY", "YOY%", "MA3"],
        )

    def test_each_template_has_required_keys(self) -> None:
        for token, tmpl in _TIME_INTELLIGENCE_TEMPLATES.items():
            with self.subTest(token=token):
                self.assertIn("suffix", tmpl)
                self.assertIn("template", tmpl)
                self.assertIn("description", tmpl)

    def test_ytd_template_uses_datesytd(self) -> None:
        rendered = _TIME_INTELLIGENCE_TEMPLATES["YTD"]["template"].format(
            base="Sales", date_ref="'Date'[Date]"
        )
        self.assertIn("DATESYTD", rendered)
        self.assertIn("[Sales]", rendered)
        self.assertIn("'Date'[Date]", rendered)

    def test_spy_template_uses_sameperiodlastyear(self) -> None:
        rendered = _TIME_INTELLIGENCE_TEMPLATES["SPY"]["template"].format(
            base="Sales", date_ref="'Date'[Date]"
        )
        self.assertIn("SAMEPERIODLASTYEAR", rendered)

    def test_yoy_depends_on_spy(self) -> None:
        self.assertEqual(_TIME_INTELLIGENCE_TEMPLATES["YOY"]["depends_on"], ["SPY"])

    def test_yoy_percent_depends_on_yoy_and_spy(self) -> None:
        self.assertEqual(
            _TIME_INTELLIGENCE_TEMPLATES["YOY%"]["depends_on"],
            ["YOY", "SPY"],
        )

    def test_yoy_percent_has_format_hint(self) -> None:
        self.assertEqual(_TIME_INTELLIGENCE_TEMPLATES["YOY%"]["format_hint"], "0.00%")

    def test_ma3_uses_datesinperiod(self) -> None:
        rendered = _TIME_INTELLIGENCE_TEMPLATES["MA3"]["template"].format(
            base="Sales", date_ref="'Date'[Date]"
        )
        self.assertIn("DATESINPERIOD", rendered)
        self.assertIn("-3", rendered)


class PatternResolutionTests(unittest.TestCase):
    def test_none_returns_full_default_set(self) -> None:
        self.assertEqual(_resolve_ti_patterns(None), list(_DEFAULT_TIME_INTELLIGENCE_PATTERNS))

    def test_empty_list_rejected(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _resolve_ti_patterns([])

    def test_unknown_pattern_rejected(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _resolve_ti_patterns(["xyz"])

    def test_case_insensitive(self) -> None:
        self.assertEqual(_resolve_ti_patterns(["ytd"]), ["YTD"])

    def test_dedup_preserves_order(self) -> None:
        self.assertEqual(_resolve_ti_patterns(["MTD", "YTD", "MTD"]), ["MTD", "YTD"])

    def test_yoy_auto_expands_spy(self) -> None:
        # YOY depends on SPY → resolution should auto-add it ahead of YOY.
        resolved = _resolve_ti_patterns(["YOY"])
        self.assertEqual(resolved, ["SPY", "YOY"])

    def test_yoy_percent_auto_expands_chain(self) -> None:
        resolved = _resolve_ti_patterns(["YOY%"])
        self.assertEqual(resolved, ["SPY", "YOY", "YOY%"])

    def test_explicit_spy_then_yoy_keeps_order(self) -> None:
        self.assertEqual(_resolve_ti_patterns(["SPY", "YOY"]), ["SPY", "YOY"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
