"""
Tests for form builder module.
"""

import json

import pytest

from app.excel_parser import BonusRow, BonusSheet, MealRefundRow, MealRefundSheet
from app.form_builder import (
    APPROVAL_CODE,
    BONUS_APPROVAL_CODE,
    WIDGET_ATTACHMENT,
    WIDGET_BONUS_ADJUSTMENT,
    WIDGET_BONUS_AMOUNT,
    WIDGET_BONUS_ATTACHMENT,
    WIDGET_BONUS_DETAIL_LIST,
    WIDGET_BONUS_EMPLOYEE_ID,
    WIDGET_BONUS_ID_CARD,
    WIDGET_BONUS_NAME,
    WIDGET_BONUS_NET,
    WIDGET_BONUS_REMARK,
    WIDGET_BONUS_TAX,
    WIDGET_BONUS_TITLE,
    WIDGET_DETAIL_LIST,
    WIDGET_REMARK,
    WIDGET_ROW_AMOUNT,
    WIDGET_ROW_ID_CARD,
    WIDGET_ROW_NAME,
    WIDGET_ROW_SERIAL,
    WIDGET_TITLE,
    build_bonus_form_widgets,
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

    def test_empty_id_card_passes_through_unchanged(self):
        """Regression: id_card='' reaches the widget as '' — never coerced, never dropped."""
        sheet = MealRefundSheet(
            title="测试餐费明细（无身份证列）",
            rows=[
                MealRefundRow(serial="1", name="张三", id_card="", amount="500"),
                MealRefundRow(serial="2", name="李四", id_card="110101199001011234", amount="600"),
            ],
            total_amount="1100",
            total_chinese="",
            maker="",
            warnings=[],
        )
        widgets = build_form_widgets(sheet, "file_code_123")
        detail_rows = widgets[1]["value"]

        # Row with empty id_card is kept, not skipped
        assert len(detail_rows) == 2

        empty_cell = detail_rows[0][2]
        assert empty_cell["id"] == WIDGET_ROW_ID_CARD
        assert empty_cell["type"] == "input"
        assert empty_cell["value"] == ""

        # Normal 18-digit id_card is unaffected (no regression)
        assert detail_rows[1][2]["value"] == "110101199001011234"

        # Empty value survives JSON serialization
        parsed = json.loads(json.dumps(widgets, ensure_ascii=False))
        assert parsed[1]["value"][0][2]["value"] == ""

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


@pytest.fixture
def sample_bonus_sheet() -> BonusSheet:
    """Provide a sample BonusSheet for testing."""
    return BonusSheet(
        title="科研部发放吉林大学2024年科研业绩津贴明细表（08.28）",
        rows=[
            BonusRow(
                serial="1",
                employee_id="824282",
                name="张三",
                id_card="110101199001011234",
                bonus="80000",
                tax="15000",
                net_amount="65000",
            ),
            BonusRow(
                serial="2",
                employee_id="824283",
                name="李四",
                id_card="11010119900202234X",
                bonus="16000",
                tax="4504",
                net_amount="11496",
            ),
        ],
        bonus_total="96000",
        tax_total="19504",
        net_total="76496",
        maker="张朦",
        warnings=[],
    )


class TestBonusWidgetConstants:
    """Test that bonus widget ID constants match verified values."""

    def test_bonus_approval_code(self):
        """Test bonus approval code constant."""
        assert BONUS_APPROVAL_CODE == "01BA9019-4E75-4726-8F95-38C43282F543"

    def test_widget_bonus_title_id(self):
        """Test bonus title widget ID."""
        assert WIDGET_BONUS_TITLE == "widget17811628753290001"

    def test_widget_bonus_detail_list_id(self):
        """Test bonus detail list widget ID."""
        assert WIDGET_BONUS_DETAIL_LIST == "widget16496685286700001"

    def test_widget_bonus_employee_id_id(self):
        """Test bonus employee ID widget ID."""
        assert WIDGET_BONUS_EMPLOYEE_ID == "widget17921134502450001"

    def test_widget_bonus_name_id(self):
        """Test bonus name widget ID."""
        assert WIDGET_BONUS_NAME == "widget17920477234030001"

    def test_widget_bonus_id_card_id(self):
        """Test bonus ID card widget ID."""
        assert WIDGET_BONUS_ID_CARD == "widget17920477572400001"

    def test_widget_bonus_amount_id(self):
        """Test bonus amount widget ID."""
        assert WIDGET_BONUS_AMOUNT == "widget17921134646590001"

    def test_widget_bonus_tax_id(self):
        """Test bonus tax widget ID."""
        assert WIDGET_BONUS_TAX == "widget17921134788590001"

    def test_widget_bonus_net_id(self):
        """Test bonus net amount widget ID."""
        assert WIDGET_BONUS_NET == "widget17822958110070001"

    def test_widget_bonus_adjustment_id(self):
        """Test bonus adjustment widget ID."""
        assert WIDGET_BONUS_ADJUSTMENT == "widget17921148886730001"

    def test_widget_bonus_attachment_id(self):
        """Test bonus attachment widget ID."""
        assert WIDGET_BONUS_ATTACHMENT == "widget16496686388470001"

    def test_widget_bonus_remark_id(self):
        """Test bonus remark widget ID."""
        assert WIDGET_BONUS_REMARK == "widget17811635300840001"


class TestBuildBonusFormWidgets:
    """Test suite for build_bonus_form_widgets function."""

    def test_returns_four_widgets(self, sample_bonus_sheet: BonusSheet):
        """Test that exactly 4 widgets are returned."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        assert len(widgets) == 4

    def test_widget_order(self, sample_bonus_sheet: BonusSheet):
        """Test widget order: title, detail list, attachment, remark."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        assert widgets[0]["id"] == WIDGET_BONUS_TITLE
        assert widgets[1]["id"] == WIDGET_BONUS_DETAIL_LIST
        assert widgets[2]["id"] == WIDGET_BONUS_ATTACHMENT
        assert widgets[3]["id"] == WIDGET_BONUS_REMARK

    def test_title_widget(self, sample_bonus_sheet: BonusSheet):
        """Test title widget uses sheet title by default."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        title_widget = widgets[0]
        assert title_widget["type"] == "input"
        assert title_widget["value"] == "科研部发放吉林大学2024年科研业绩津贴明细表（08.28）"

    def test_title_override(self, sample_bonus_sheet: BonusSheet):
        """Test title can be overridden."""
        widgets = build_bonus_form_widgets(
            sample_bonus_sheet, "file_code_123", title="自定义标题"
        )
        assert widgets[0]["value"] == "自定义标题"

    def test_detail_list_widget(self, sample_bonus_sheet: BonusSheet):
        """Test detail list widget structure."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        detail_widget = widgets[1]
        assert detail_widget["type"] == "fieldList"
        assert len(detail_widget["value"]) == 2

    def test_detail_row_has_seven_entries(self, sample_bonus_sheet: BonusSheet):
        """Test each detail row has exactly 7 widgets."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        for row in widgets[1]["value"]:
            assert len(row) == 7

    def test_detail_row_widget_order(self, sample_bonus_sheet: BonusSheet):
        """Test detail row widget order: 工号, 姓名, 身份证, 奖金, 个税, 调整, 实发."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        row = widgets[1]["value"][0]
        assert row[0]["id"] == WIDGET_BONUS_EMPLOYEE_ID
        assert row[1]["id"] == WIDGET_BONUS_NAME
        assert row[2]["id"] == WIDGET_BONUS_ID_CARD
        assert row[3]["id"] == WIDGET_BONUS_AMOUNT
        assert row[4]["id"] == WIDGET_BONUS_TAX
        assert row[5]["id"] == WIDGET_BONUS_ADJUSTMENT
        assert row[6]["id"] == WIDGET_BONUS_NET

    def test_detail_row_types(self, sample_bonus_sheet: BonusSheet):
        """Test input vs amount types in detail row."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        row = widgets[1]["value"][0]
        assert row[0]["type"] == "input"
        assert row[1]["type"] == "input"
        assert row[2]["type"] == "input"
        assert row[3]["type"] == "amount"
        assert row[4]["type"] == "amount"
        assert row[5]["type"] == "amount"
        assert row[6]["type"] == "amount"

    def test_detail_row_values(self, sample_bonus_sheet: BonusSheet):
        """Test detail row values match sheet data."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        row = widgets[1]["value"][0]
        assert row[0]["value"] == "824282"
        assert row[1]["value"] == "张三"
        assert row[2]["value"] == "110101199001011234"
        assert row[3]["value"] == "80000"
        assert row[4]["value"] == "15000"
        assert row[5]["value"] == "0"  # adjustment defaults to "0"
        assert row[6]["value"] == "65000"

    def test_no_serial_in_detail_row(self, sample_bonus_sheet: BonusSheet):
        """Test that serial/序号 is NOT present in detail row widgets."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        all_ids = set()
        for row in widgets[1]["value"]:
            for w in row:
                all_ids.add(w["id"])
        assert WIDGET_ROW_SERIAL not in all_ids
        assert "widget17811629865160001" not in all_ids
        for row in widgets[1]["value"]:
            adj_count = sum(1 for w in row if w["id"] == WIDGET_BONUS_ADJUSTMENT)
            assert adj_count == 1

    def test_attachment_widget(self, sample_bonus_sheet: BonusSheet):
        """Test attachment widget contains file code."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_abc")
        att_widget = widgets[2]
        assert att_widget["type"] == "attachmentV2"
        assert att_widget["value"] == ["file_code_abc"]

    def test_remark_widget_default(self, sample_bonus_sheet: BonusSheet):
        """Test remark widget defaults to empty string."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        assert widgets[3]["value"] == ""

    def test_remark_widget_custom(self, sample_bonus_sheet: BonusSheet):
        """Test remark widget with custom text."""
        widgets = build_bonus_form_widgets(
            sample_bonus_sheet, "file_code_123", remark="制表人：张朦"
        )
        assert widgets[3]["value"] == "制表人：张朦"

    def test_json_serializable(self, sample_bonus_sheet: BonusSheet):
        """Test that widget list is JSON serializable."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        json_str = json.dumps(widgets, ensure_ascii=False)
        parsed = json.loads(json_str)
        assert len(parsed) == 4

    def test_empty_rows(self):
        """Test with empty rows list."""
        sheet = BonusSheet(
            title="Empty",
            rows=[],
            bonus_total="0",
            tax_total="0",
            net_total="0",
            maker="",
        )
        widgets = build_bonus_form_widgets(sheet, "code")
        assert len(widgets[1]["value"]) == 0

    def test_adjustment_widget_at_index_5(self, sample_bonus_sheet: BonusSheet):
        """Test adjustment widget is at index 5 between tax and net."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        row = widgets[1]["value"][0]
        assert row[5]["id"] == WIDGET_BONUS_ADJUSTMENT
        assert row[5]["type"] == "amount"

    def test_adjustment_value_defaults_to_zero(self, sample_bonus_sheet: BonusSheet):
        """Test adjustment value defaults to '0' when not set."""
        widgets = build_bonus_form_widgets(sample_bonus_sheet, "file_code_123")
        for row in widgets[1]["value"]:
            assert row[5]["value"] == "0"

    def test_adjustment_value_passes_through(self):
        """Test adjustment value passes through from row.adjustment."""
        sheet = BonusSheet(
            title="Test",
            rows=[
                BonusRow(
                    serial="1",
                    employee_id="001",
                    name="张三",
                    id_card="110101199001011234",
                    bonus="1000",
                    tax="100",
                    net_amount="900",
                    adjustment="50",
                ),
            ],
            bonus_total="1000",
            tax_total="100",
            net_total="900",
            maker="测试",
            adjustment_total="50",
        )
        widgets = build_bonus_form_widgets(sheet, "code")
        row = widgets[1]["value"][0]
        assert row[5]["value"] == "50"
