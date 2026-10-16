"""
Tests for the shipped sanitized Excel templates.

Guards that templates/ contains two .xlsx files with 100% fake personal
data that parse cleanly through the app's own parsers, and that no real
PII leaks into the shipped artifacts.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from app.excel_parser import parse_bonus_excel, parse_meal_refund_excel
from app.form_builder import build_bonus_form_widgets, build_form_widgets

# ── Fake data pool (must match the generation spec exactly) ─────────
FAKE_NAMES = [
    "张三", "李四", "王五", "赵六", "孙七", "周八", "吴九", "郑十",
    "王小明", "冯小伟", "陈小华", "褚小红", "卫小青", "蒋小刚",
    "沈小丽", "韩小强", "杨小军", "朱小梅",
]
FAKE_NAME_SET = set(FAKE_NAMES)
FAKE_ID_PATTERN = re.compile(r"^11010119\d{9}[\dX]$")
FAKE_EMP_IDS = {"800001", "800002", "800003", "800004"}
MAKER = "张三"

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
MEAL_PATH = TEMPLATES_DIR / "返餐费明细表模板.xlsx"
BONUS_PATH = TEMPLATES_DIR / "科研业绩津贴明细表模板.xlsx"


class TestShippedTemplates:
    """Guards for the two sanitized template files shipped in-repo."""

    def test_templates_exist(self):
        """Both template files must exist under templates/."""
        assert MEAL_PATH.exists(), f"Missing: {MEAL_PATH}"
        assert BONUS_PATH.exists(), f"Missing: {BONUS_PATH}"

    # ── Meal template ───────────────────────────────────────────────

    def test_meal_parses_cleanly(self):
        """Meal template parses with expected structure and zero warnings."""
        sheet = parse_meal_refund_excel(MEAL_PATH)
        assert sheet.title == "松辽水利委员会综合服务中心餐费明细（06.15）"
        assert len(sheet.rows) == 17
        assert sheet.rows[0].serial == "1"
        assert sheet.rows[0].name == "李四"
        assert sheet.rows[0].amount == "1860"
        # At least one ID card ends with X
        x_ids = [r for r in sheet.rows if r.id_card.endswith("X")]
        assert len(x_ids) >= 1
        assert sheet.total_amount == "31620"
        assert sheet.maker == "张三"
        assert sheet.warnings == []

    # ── Bonus template ──────────────────────────────────────────────

    def test_bonus_parses_cleanly(self):
        """Bonus template parses with expected structure and zero warnings."""
        sheet = parse_bonus_excel(BONUS_PATH)
        assert sheet.title == "科研部发放吉林大学2024年科研业绩津贴明细表（08.28）"
        assert len(sheet.rows) == 4
        assert sheet.rows[0].employee_id == "800001"
        assert sheet.rows[0].bonus == "80000"
        assert sheet.rows[0].tax == "15904"
        assert sheet.rows[0].adjustment == "0"
        assert sheet.rows[0].net_amount == "64096"
        assert sheet.bonus_total == "96000"
        assert sheet.tax_total == "19504"
        assert sheet.net_total == "76496"
        assert sheet.maker == "张三"
        assert sheet.warnings == []

    # ── Sanitization guards ─────────────────────────────────────────

    def test_sanitization_meal_names_and_ids(self):
        """Every meal row name is in the fake pool; every ID matches fake pattern."""
        sheet = parse_meal_refund_excel(MEAL_PATH)
        for row in sheet.rows:
            assert row.name in FAKE_NAME_SET, f"Real name leaked: {row.name}"
            assert FAKE_ID_PATTERN.match(row.id_card), (
                f"Real ID leaked: {row.id_card}"
            )
        assert sheet.maker == "张三"

    def test_sanitization_bonus_names_ids_empids(self):
        """Every bonus row name ∈ fake pool, ID matches fake pattern, emp_ids are fake."""
        sheet = parse_bonus_excel(BONUS_PATH)
        for row in sheet.rows:
            assert row.name in FAKE_NAME_SET, f"Real name leaked: {row.name}"
            assert FAKE_ID_PATTERN.match(row.id_card), (
                f"Real ID leaked: {row.id_card}"
            )
        emp_ids = {row.employee_id for row in sheet.rows}
        assert emp_ids == FAKE_EMP_IDS
        assert sheet.maker == "张三"

    # ── Form builder integration ────────────────────────────────────

    def test_form_builders_accept_templates(self):
        """Both form builders succeed on the shipped templates."""
        meal_sheet = parse_meal_refund_excel(MEAL_PATH)
        widgets = build_form_widgets(meal_sheet, file_code="fake_code")
        assert isinstance(widgets, list)
        assert len(widgets) >= 1

        bonus_sheet = parse_bonus_excel(BONUS_PATH)
        bonus_widgets = build_bonus_form_widgets(bonus_sheet, file_code="fake_code")
        assert isinstance(bonus_widgets, list)
        # The detail-list widget contains 4 rows × 7 fields each
        detail_widget = [w for w in bonus_widgets if w.get("type") == "fieldList"]
        assert len(detail_widget) == 1
        detail_rows = detail_widget[0]["value"]
        assert len(detail_rows) == 4
        for row in detail_rows:
            assert len(row) == 7
