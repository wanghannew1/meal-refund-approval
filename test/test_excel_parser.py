"""
Tests for Excel parser module.
"""

from pathlib import Path

import pytest

from app.excel_parser import (
    MealRefundRow,
    MealRefundSheet,
    parse_meal_refund_excel,
)


class TestRealFixture:
    """Test suite using the real template fixture file."""

    def test_parse_title(self, real_fixture_path: Path):
        """Test title extraction from row 0 col 0."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert sheet.title == "松辽水利委员会综合服务中心餐费明细（06.15）"

    def test_parse_row_count(self, real_fixture_path: Path):
        """Test that exactly 17 data rows are parsed."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert len(sheet.rows) == 17

    def test_parse_first_row(self, real_fixture_path: Path):
        """Test first data row values."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        row = sheet.rows[0]
        assert row.serial == "1"
        assert row.name == "韩东旭"
        assert row.id_card == "210319197203110478"
        assert row.amount == "1860"

    def test_parse_last_row(self, real_fixture_path: Path):
        """Test last data row values."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        row = sheet.rows[-1]
        assert row.serial == "17"
        assert row.name == "韩魁双"
        assert row.id_card == "220104197804274415"
        assert row.amount == "1860"

    def test_parse_id_card_with_x(self, real_fixture_path: Path):
        """Test that ID cards ending in X are preserved."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        # Row 2 (张世宇) has ID ending in X
        row = sheet.rows[1]
        assert row.name == "张世宇"
        assert row.id_card.endswith("X")
        assert row.id_card == "22010519830911161X"

        # Row 16 (刘洋) also has ID ending in X
        row16 = sheet.rows[15]
        assert row16.name == "刘洋"
        assert row16.id_card == "22032319880710511X"

    def test_parse_total_amount(self, real_fixture_path: Path):
        """Test total amount normalization from ￥：31620."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert sheet.total_amount == "31620"

    def test_parse_total_chinese(self, real_fixture_path: Path):
        """Test Chinese uppercase amount text."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert sheet.total_chinese == "（人民币大写）叁万壹仟陆佰贰拾元"

    def test_parse_maker(self, real_fixture_path: Path):
        """Test maker name extraction from 制表：孙世多."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert sheet.maker == "孙世多"

    def test_no_warnings_for_valid_sheet(self, real_fixture_path: Path):
        """Test that a valid sheet produces no warnings."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        assert sheet.warnings == []

    def test_serial_normalization(self, real_fixture_path: Path):
        """Test that float serials (1.0, 2.0) are normalized to strings."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        for i, row in enumerate(sheet.rows, start=1):
            assert row.serial == str(i)

    def test_amount_normalization(self, real_fixture_path: Path):
        """Test that float amounts (1860.0) are normalized to clean strings."""
        sheet = parse_meal_refund_excel(real_fixture_path)
        for row in sheet.rows:
            assert row.amount == "1860"
            assert "." not in row.amount


class TestEdgeCases:
    """Test suite for edge cases using synthetic fixtures."""

    def test_missing_header_raises(self, tmp_path: Path):
        """Test that missing header row raises ValueError."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Some title"
        ws["A2"] = "Not a header"
        filepath = tmp_path / "no_header.xlsx"
        wb.save(filepath)

        with pytest.raises(ValueError, match="未找到表头行"):
            parse_meal_refund_excel(filepath)

    def test_no_data_rows_raises(self, tmp_path: Path):
        """Test that no data rows raises ValueError."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Title"
        ws["A2"] = "序号"
        ws["B2"] = "姓名"
        ws["C2"] = "身份证号码"
        ws["D2"] = "金额"
        # No data rows, directly 总计
        ws["A3"] = "总计"
        ws["D3"] = "0"
        filepath = tmp_path / "no_data.xlsx"
        wb.save(filepath)

        with pytest.raises(ValueError, match="未找到数据行"):
            parse_meal_refund_excel(filepath)

    def test_total_mismatch_warning(self, tmp_path: Path):
        """Test that sum != total produces a warning."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Test Title"
        ws["A2"] = "序号"
        ws["B2"] = "姓名"
        ws["C2"] = "身份证号码"
        ws["D2"] = "金额"
        ws["A3"] = 1
        ws["B3"] = "张三"
        ws["C3"] = "110101199001011234"
        ws["D3"] = 500
        ws["A4"] = 2
        ws["B4"] = "李四"
        ws["C4"] = "110101199002022345"
        ws["D4"] = 600
        ws["A5"] = "总计"
        ws["B5"] = "（人民币大写）壹仟壹佰元"
        ws["D5"] = "￥：1100"  # Correct sum
        filepath = tmp_path / "mismatch.xlsx"
        wb.save(filepath)

        sheet = parse_meal_refund_excel(filepath)
        assert sheet.warnings == []  # 500+600=1100, matches

        # Now create one with wrong total
        ws["D5"] = "￥：9999"  # Wrong total
        wb.save(filepath)

        sheet2 = parse_meal_refund_excel(filepath)
        assert len(sheet2.warnings) == 1
        assert "不一致" in sheet2.warnings[0]

    def test_numeric_id_card_warning(self, tmp_path: Path):
        """Test that numeric ID card cell produces a warning."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Test"
        ws["A2"] = "序号"
        ws["B2"] = "姓名"
        ws["C2"] = "身份证号码"
        ws["D2"] = "金额"
        ws["A3"] = 1
        ws["B3"] = "张三"
        # ID card as number (will lose precision for 18 digits)
        ws["C3"] = 110101199001011234
        ws["D3"] = 500
        ws["A4"] = "总计"
        ws["D4"] = "￥：500"
        filepath = tmp_path / "numeric_id.xlsx"
        wb.save(filepath)

        sheet = parse_meal_refund_excel(filepath)
        assert len(sheet.warnings) >= 1
        assert "身份证号码" in sheet.warnings[0]

    def test_unsupported_format_raises(self, tmp_path: Path):
        """Test that unsupported file format raises ValueError."""
        filepath = tmp_path / "test.csv"
        filepath.write_text("a,b,c")

        with pytest.raises(ValueError, match="不支持的文件格式"):
            parse_meal_refund_excel(filepath)

    def test_empty_sheet_raises(self, tmp_path: Path):
        """Test that empty sheet raises ValueError."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        filepath = tmp_path / "empty.xlsx"
        wb.save(filepath)

        with pytest.raises(ValueError, match="未找到表头行"):
            parse_meal_refund_excel(filepath)

    def test_xlsx_format_works(self, tmp_path: Path):
        """Test that .xlsx format is parsed correctly."""
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "XLSX Title"
        ws["A2"] = "序号"
        ws["B2"] = "姓名"
        ws["C2"] = "身份证号码"
        ws["D2"] = "金额"
        ws["A3"] = 1
        ws["B3"] = "王五"
        ws["C3"] = "110101199003031234"
        ws["D3"] = 1000
        ws["A4"] = "总计"
        ws["B4"] = "壹仟元"
        ws["D4"] = "￥：1000"
        ws["A6"] = "制表：测试"
        filepath = tmp_path / "test.xlsx"
        wb.save(filepath)

        sheet = parse_meal_refund_excel(filepath)
        assert sheet.title == "XLSX Title"
        assert len(sheet.rows) == 1
        assert sheet.rows[0].name == "王五"
        assert sheet.total_amount == "1000"
        assert sheet.maker == "测试"
