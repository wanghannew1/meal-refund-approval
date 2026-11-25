"""
Tests for the two-amount-column meal approval flow (派遣员工简易费用明细表审批-两列金额).

Independent of the in-use single-amount flow: new parser
``parse_meal_two_amount_excel`` and new builder ``build_meal2_form_widgets``.
"""

import json
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from app.excel_parser import MealTwoAmountRow, MealTwoAmountSheet, parse_meal_two_amount_excel
from app.form_builder import (
    MEAL2_ATTACHMENT,
    MEAL2_COLUMN_MAPPING,
    MEAL2_DETAIL_LIST,
    MEAL2_MAPPING_EXCEL_COLUMN,
    MEAL2_MAPPING_FORM_COLUMN,
    MEAL2_REMARK,
    MEAL2_ROW_AMOUNT1,
    MEAL2_ROW_AMOUNT2,
    MEAL2_ROW_ID_CARD,
    MEAL2_ROW_NAME,
    MEAL2_TITLE,
    build_meal2_form_widgets,
)


def _make_xlsx(path: Path, *, amount2_blank_row: int | None = None) -> None:
    """Write a minimal two-amount-column workbook in the observed format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet2"
    ws["A1"] = "测试中心返2026年11月单位经费、工会会费明细"
    ws.append([])
    ws.append(["序号", "姓名", "身份证号", "单位工会经费", "工会会费"])
    people = [
        ("1", "张三", "220104198302271331", "61.8", "11"),
        ("2", "李四", "220106198512270612", "57.8", "10"),
        ("3", "王五", "220182198410204311", "49.44", "9"),
    ]
    for serial, name, id_card, a1, a2 in people:
        if amount2_blank_row is not None and int(serial) == amount2_blank_row:
            a2 = None
        ws.append([int(serial), name, id_card, float(a1), None if a2 is None else float(a2)])
    ws.append(["合计", "", "", 169.04, 30])
    ws.append([])
    ws.append(["总经理签字：", "部长签字：", "财务审核：", "制表人：测试员"])
    wb.save(path)


@pytest.fixture
def meal2_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "meal2.xlsx"
    _make_xlsx(path)
    return path


class TestParseMealTwoAmount:
    def test_title(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert sheet.title == "测试中心返2026年11月单位经费、工会会费明细"

    def test_amount_columns_detected(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert sheet.amount1_column == "单位工会经费"
        assert sheet.amount2_column == "工会会费"

    def test_rows(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert len(sheet.rows) == 3
        first = sheet.rows[0]
        assert isinstance(first, MealTwoAmountRow)
        assert first.serial == "1"
        assert first.name == "张三"
        assert first.id_card == "220104198302271331"
        assert first.amount1 == "61.8"
        assert first.amount2 == "11"
        assert sheet.rows[2].amount1 == "49.44"

    def test_totals_normalized(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert sheet.amount1_total == "169.04"
        assert sheet.amount2_total == "30"

    def test_maker(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert sheet.maker == "测试员"

    def test_no_warnings(self, meal2_xlsx: Path):
        sheet = parse_meal_two_amount_excel(meal2_xlsx)
        assert sheet.warnings == []

    def test_blank_amount2_becomes_zero(self, tmp_path: Path):
        path = tmp_path / "blank.xlsx"
        _make_xlsx(path, amount2_blank_row=2)
        sheet = parse_meal_two_amount_excel(path)
        assert sheet.rows[1].amount2 == "0"
        # 合计 30 vs 明细 11+0+9=20 → 不一致告警
        assert any("工会会费明细合计" in w for w in sheet.warnings)

    def test_mismatch_warning(self, tmp_path: Path):
        path = tmp_path / "mismatch.xlsx"
        _make_xlsx(path)
        import openpyxl as _op

        wb = _op.load_workbook(path)
        ws = wb.active
        ws.cell(row=7, column=4, value=999)  # 合计行 金额1 改错
        wb.save(path)
        sheet = parse_meal_two_amount_excel(path)
        assert any("单位工会经费明细合计" in w for w in sheet.warnings)

    def test_header_not_found(self, tmp_path: Path):
        path = tmp_path / "bad.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append(["乱七八糟"])
        wb.save(path)
        with pytest.raises(ValueError, match="未找到表头行"):
            parse_meal_two_amount_excel(path)


class TestBuildMeal2FormWidgets:
    def _sample_sheet(self) -> MealTwoAmountSheet:
        return MealTwoAmountSheet(
            title="测试标题",
            rows=[
                MealTwoAmountRow(serial="1", name="张三", id_card="220104198302271331", amount1="61.8", amount2="11"),
                MealTwoAmountRow(serial="2", name="李四", id_card="220106198512270612", amount1="57.8", amount2="10"),
            ],
            amount1_column="单位工会经费",
            amount2_column="工会会费",
            amount1_total="119.6",
            amount2_total="21",
            maker="测试员",
        )

    def test_widget_count_and_ids(self):
        widgets = build_meal2_form_widgets(self._sample_sheet(), file_code="FC")
        assert [w["id"] for w in widgets] == [
            MEAL2_TITLE,
            MEAL2_COLUMN_MAPPING,
            MEAL2_DETAIL_LIST,
            MEAL2_ATTACHMENT,
            MEAL2_REMARK,
        ]

    def test_column_mapping_rows(self):
        widgets = build_meal2_form_widgets(self._sample_sheet(), file_code="FC")
        mapping = widgets[1]["value"]
        assert len(mapping) == 2
        assert mapping[0] == [
            {"id": MEAL2_MAPPING_FORM_COLUMN, "type": "input", "value": "金额1"},
            {"id": MEAL2_MAPPING_EXCEL_COLUMN, "type": "input", "value": "单位工会经费"},
        ]
        assert mapping[1][1]["value"] == "工会会费"

    def test_detail_row_children(self):
        widgets = build_meal2_form_widgets(self._sample_sheet(), file_code="FC")
        row0 = widgets[2]["value"][0]
        assert [c["id"] for c in row0] == [
            MEAL2_ROW_NAME,
            MEAL2_ROW_ID_CARD,
            MEAL2_ROW_AMOUNT1,
            MEAL2_ROW_AMOUNT2,
        ]
        assert [c["value"] for c in row0] == ["张三", "220104198302271331", "61.8", "11"]
        assert len(widgets[2]["value"]) == 2  # 序号不入表单

    def test_serial_excluded_from_form(self):
        widgets = build_meal2_form_widgets(self._sample_sheet(), file_code="FC")
        dumped = json.dumps(widgets, ensure_ascii=False)
        # 序号值不应出现在明细里（"3" 这个序号本身不会出现）
        for row in widgets[2]["value"]:
            assert all(c["id"] != "serial_widget" for c in row)
        assert dumped

    def test_title_and_remark_override(self):
        widgets = build_meal2_form_widgets(
            self._sample_sheet(), file_code="FC", title="自定义标题", remark="备注X"
        )
        assert widgets[0]["value"] == "自定义标题"
        assert widgets[4]["value"] == "备注X"
        assert widgets[3]["value"] == ["FC"]

    def test_json_serializable(self):
        widgets = build_meal2_form_widgets(self._sample_sheet(), file_code="FC")
        json.dumps(widgets, ensure_ascii=False)
