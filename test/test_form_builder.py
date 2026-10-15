"""
Tests for form builder module.
"""

import json

import pytest

from app.excel_parser import MealRefundRow, MealRefundSheet
from app.form_builder import (
    APPROVAL_CODE,
    WIDGET_ATTACHMENT,
    WIDGET_DETAIL_LIST,
    WIDGET_REMARK,
    WIDGET_ROW_AMOUNT,
    WIDGET_ROW_ID_CARD,
    WIDGET_ROW_NAME,
    WIDGET_ROW_SERIAL,
    WIDGET_TITLE,
    build_form_widgets,
)


@pytest.fixture
def sample_sheet() -> MealRefundSheet:
    """Provide a sample MealRefundSheet for testing."""
    return MealRefundSheet(
        title="测试餐费明细（01.01）",
        rows=[
            MealRefundRow(serial="1", name="张三", id_card="110101199001011234", amount="500"),
            MealRefundRow(serial="2", name="李四", id_card="11010119900202234X", amount="600"),
        ],
        total_amount="1100",
        total_chinese="（人民币大写）壹仟壹佰元",
        maker="测试员",
        warnings=[],
    )


class TestWidgetConstants:
    """Test that widget ID constants match verified values."""

    def test_widget_title_id(self):
        """Test title widget ID."""
        assert WIDGET_TITLE == "widget17811628753290001"

    def test_widget_detail_list_id(self):
        """Test detail list widget ID."""
        assert WIDGET_DETAIL_LIST == "widget16496685286700001"

    def test_widget_row_serial_id(self):
        """Test row serial widget ID."""
        assert WIDGET_ROW_SERIAL == "widget17811629865160001"

    def test_widget_row_name_id(self):
        """Test row name widget ID."""
        assert WIDGET_ROW_NAME == "widget17920477234030001"

    def test_widget_row_id_card_id(self):
        """Test row ID card widget ID."""
        assert WIDGET_ROW_ID_CARD == "widget17920477572400001"

    def test_widget_row_amount_id(self):
        """Test row amount widget ID."""
        assert WIDGET_ROW_AMOUNT == "widget17822958110070001"

    def test_widget_attachment_id(self):
        """Test attachment widget ID."""
        assert WIDGET_ATTACHMENT == "widget16496686388470001"

    def test_widget_remark_id(self):
        """Test remark widget ID."""
        assert WIDGET_REMARK == "widget17811635300840001"

    def test_approval_code(self):
        """Test approval code constant."""
        assert APPROVAL_CODE == "081B5624-B8B7-4334-AF42-D1075F8E650A"


class TestBuildFormWidgets:
    """Test suite for build_form_widgets function."""

    def test_returns_four_widgets(self, sample_sheet: MealRefundSheet):
        """Test that exactly 4 widgets are returned."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        assert len(widgets) == 4

    def test_title_widget(self, sample_sheet: MealRefundSheet):
        """Test title widget uses sheet title by default."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        title_widget = widgets[0]
        assert title_widget["id"] == WIDGET_TITLE
        assert title_widget["type"] == "input"
        assert title_widget["value"] == "测试餐费明细（01.01）"

    def test_title_override(self, sample_sheet: MealRefundSheet):
        """Test title can be overridden."""
        widgets = build_form_widgets(
            sample_sheet, "file_code_123", title="自定义标题"
        )
        assert widgets[0]["value"] == "自定义标题"

    def test_detail_list_widget(self, sample_sheet: MealRefundSheet):
        """Test detail list widget structure."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        detail_widget = widgets[1]
        assert detail_widget["id"] == WIDGET_DETAIL_LIST
        assert detail_widget["type"] == "fieldList"
        assert len(detail_widget["value"]) == 2  # 2 rows

    def test_detail_list_row_structure(self, sample_sheet: MealRefundSheet):
        """Test each detail row has 4 widgets with correct IDs."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        detail_rows = widgets[1]["value"]

        # Check first row
        row = detail_rows[0]
        assert len(row) == 4
        assert row[0]["id"] == WIDGET_ROW_SERIAL
        assert row[0]["type"] == "input"
        assert row[0]["value"] == "1"
        assert row[1]["id"] == WIDGET_ROW_NAME
        assert row[1]["value"] == "张三"
        assert row[2]["id"] == WIDGET_ROW_ID_CARD
        assert row[2]["value"] == "110101199001011234"
        assert row[3]["id"] == WIDGET_ROW_AMOUNT
        assert row[3]["type"] == "amount"
        assert row[3]["value"] == "500"

    def test_detail_list_second_row(self, sample_sheet: MealRefundSheet):
        """Test second detail row preserves ID card with X."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        row = widgets[1]["value"][1]
        assert row[0]["value"] == "2"
        assert row[1]["value"] == "李四"
        assert row[2]["value"] == "11010119900202234X"
        assert row[3]["value"] == "600"

    def test_attachment_widget(self, sample_sheet: MealRefundSheet):
        """Test attachment widget contains file code."""
        widgets = build_form_widgets(sample_sheet, "file_code_abc")
        att_widget = widgets[2]
        assert att_widget["id"] == WIDGET_ATTACHMENT
        assert att_widget["type"] == "attachmentV2"
        assert att_widget["value"] == ["file_code_abc"]

    def test_remark_widget_default(self, sample_sheet: MealRefundSheet):
        """Test remark widget defaults to empty string."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        remark_widget = widgets[3]
        assert remark_widget["id"] == WIDGET_REMARK
        assert remark_widget["type"] == "textarea"
        assert remark_widget["value"] == ""

    def test_remark_widget_custom(self, sample_sheet: MealRefundSheet):
        """Test remark widget with custom text."""
        widgets = build_form_widgets(
            sample_sheet, "file_code_123", remark="制表：测试员"
        )
        assert widgets[3]["value"] == "制表：测试员"

    def test_json_serializable(self, sample_sheet: MealRefundSheet):
        """Test that widget list is JSON serializable."""
        widgets = build_form_widgets(sample_sheet, "file_code_123")
        json_str = json.dumps(widgets, ensure_ascii=False)
        parsed = json.loads(json_str)
        assert len(parsed) == 4

    def test_empty_rows(self):
        """Test with empty rows list."""
        sheet = MealRefundSheet(
            title="Empty",
            rows=[],
            total_amount="0",
            total_chinese="",
            maker="",
        )
        widgets = build_form_widgets(sheet, "code")
        assert len(widgets[1]["value"]) == 0
