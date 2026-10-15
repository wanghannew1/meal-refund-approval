"""
Excel parser for meal refund detail spreadsheets.

Supports both legacy .xls (via xlrd) and modern .xlsx (via openpyxl)
formats.  Locates the header row by matching column labels, extracts
data rows, and normalizes numeric values to clean strings.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path


@dataclass(frozen=True)
class MealRefundRow:
    """A single data row from the meal refund spreadsheet.

    Attributes:
        serial: Normalized serial number string ("1", "2", ...).
        name: Employee name.
        id_card: Identity card number (text, may end in X).
        amount: Plain decimal amount string ("1860" or "1860.50").
    """

    serial: str
    name: str
    id_card: str
    amount: str


@dataclass(frozen=True)
class MealRefundSheet:
    """Parsed result of a meal refund spreadsheet.

    Attributes:
        title: Sheet title from row 0 col 0.
        rows: List of data rows.
        total_amount: Normalized total amount string (e.g. "31620").
        total_chinese: Chinese uppercase amount text.
        maker: Name of the person who prepared the sheet.
        warnings: List of warning messages (Chinese text).
    """

    title: str
    rows: list[MealRefundRow]
    total_amount: str
    total_chinese: str
    maker: str
    warnings: list[str] = field(default_factory=list)


def _normalize_numeric(value: object) -> tuple[str, bool]:
    """Convert a cell value to a clean numeric string.

    Args:
        value: Raw cell value (float, int, or string).

    Returns:
        Tuple of (normalized_string, was_numeric_float).
        The boolean indicates whether the original value was a float
        that got converted via int(), which may lose precision for
        very large numbers.

    Raises:
        ValueError: If the value cannot be interpreted as a number.
    """
    if isinstance(value, float):
        # xlrd returns numeric cells as floats (1.0, 1860.0)
        if value == int(value):
            return str(int(value)), True
        # Has fractional part — format without trailing zeros
        d = Decimal(str(value))
        normalized = d.normalize()
        # Ensure at least integer form
        return str(normalized), True
    if isinstance(value, int):
        return str(value), False
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return "", False
        # Try parsing as number
        try:
            d = Decimal(stripped)
            if d == d.to_integral_value():
                return str(int(d)), False
            return str(d.normalize()), False
        except InvalidOperation:
            return stripped, False
    return str(value), False


def _normalize_id_card(value: object) -> tuple[str, bool]:
    """Convert an ID card cell value to string, preserving trailing X.

    Args:
        value: Raw cell value.

    Returns:
        Tuple of (id_card_string, was_numeric).
        If the cell was numeric (float), converting to int may lose
        precision for 18-digit IDs — caller should add a warning.
    """
    if isinstance(value, float):
        # Numeric cell — convert via int, but warn about precision
        return str(int(value)), True
    if isinstance(value, int):
        return str(value), False
    if isinstance(value, str):
        return value.strip(), False
    return str(value).strip(), False


def _extract_total_amount(raw: str) -> str:
    """Extract numeric amount from strings like '￥：31620' or '￥31,620.00'.

    Args:
        raw: Raw total amount cell text.

    Returns:
        Clean decimal string without currency symbols or commas.
    """
    # Remove currency symbols, commas, colons (full/half-width), whitespace
    cleaned = re.sub(r"[￥$：:,\s]", "", raw.strip())
    # Try to parse as number
    try:
        d = Decimal(cleaned)
        if d == d.to_integral_value():
            return str(int(d))
        return str(d.normalize())
    except InvalidOperation:
        return cleaned


def _extract_maker(text: str) -> str:
    """Extract maker name from text like '制表：孙世多'.

    Args:
        text: Raw cell text from the signature row.

    Returns:
        Maker name string, or empty string if not found.
    """
    match = re.search(r"制表[：:]\s*(.+)", text.strip())
    if match:
        return match.group(1).strip()
    return ""


def _is_blank_row(row_values: list[object]) -> bool:
    """Check if a row is entirely blank/empty."""
    for v in row_values:
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        if isinstance(v, (int, float)) and v == 0:
            continue
        return False
    return True


def _read_xls(path: Path) -> tuple[str, list[list[object]]]:
    """Read an .xls file and return (title, all_rows).

    Args:
        path: Path to the .xls file.

    Returns:
        Tuple of (title_string, list_of_row_lists).

    Raises:
        ValueError: If the file cannot be read.
    """
    import xlrd

    try:
        wb = xlrd.open_workbook(str(path))
    except Exception as exc:
        raise ValueError(f"无法读取 Excel 文件: {exc}") from exc

    sh = wb.sheet_by_index(0)
    title = str(sh.cell(0, 0).value) if sh.nrows > 0 else ""

    all_rows: list[list[object]] = []
    for r in range(sh.nrows):
        row: list[object] = []
        for c in range(sh.ncols):
            row.append(sh.cell(r, c).value)
        all_rows.append(row)

    return title, all_rows


def _read_xlsx(path: Path) -> tuple[str, list[list[object]]]:
    """Read an .xlsx file and return (title, all_rows).

    Args:
        path: Path to the .xlsx file.

    Returns:
        Tuple of (title_string, list_of_row_lists).

    Raises:
        ValueError: If the file cannot be read.
    """
    import openpyxl

    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError(f"无法读取 Excel 文件: {exc}") from exc

    ws = wb.active
    all_rows: list[list[object]] = []
    for row in ws.iter_rows(values_only=True):
        all_rows.append(list(row))

    wb.close()

    title = ""
    if all_rows and all_rows[0]:
        title = str(all_rows[0][0]) if all_rows[0][0] is not None else ""

    return title, all_rows


def parse_meal_refund_excel(path: str | Path) -> MealRefundSheet:
    """Parse a meal refund detail spreadsheet.

    Supports .xls (xlrd) and .xlsx (openpyxl) formats.  The header row
    is located by finding col0=="序号" AND col1=="姓名".  Data rows
    follow until a 总计 row (col0=="总计") or a blank row.

    Args:
        path: Path to the Excel file.

    Returns:
        MealRefundSheet with parsed data.

    Raises:
        ValueError: If the file is unreadable, header not found,
            or no data rows present.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == ".xls":
        title, all_rows = _read_xls(path)
    elif suffix == ".xlsx":
        title, all_rows = _read_xlsx(path)
    else:
        raise ValueError(f"不支持的文件格式: {suffix}，仅支持 .xls 和 .xlsx")

    # Locate header row
    header_row_idx: int | None = None
    for idx, row in enumerate(all_rows):
        if len(row) < 2:
            continue
        col0 = str(row[0]).strip() if row[0] is not None else ""
        col1 = str(row[1]).strip() if row[1] is not None else ""
        if col0 == "序号" and col1 == "姓名":
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError('未找到表头行（需要包含"序号"和"姓名"列）')

    # Parse data rows
    rows: list[MealRefundRow] = []
    warnings: list[str] = []
    total_amount = ""
    total_chinese = ""
    maker = ""

    data_start = header_row_idx + 1
    for idx in range(data_start, len(all_rows)):
        row = all_rows[idx]

        # Pad row to at least 4 columns
        while len(row) < 4:
            row.append(None)

        col0_raw = row[0]

        # Check for 总计 row
        col0_str = str(col0_raw).strip() if col0_raw is not None else ""
        if col0_str == "总计":
            # Extract total info
            total_chinese = str(row[1]).strip() if row[1] is not None else ""
            raw_total = str(row[3]).strip() if row[3] is not None else ""
            total_amount = _extract_total_amount(raw_total)

            # Look for maker in subsequent rows
            for subsequent_idx in range(idx + 1, min(idx + 5, len(all_rows))):
                sub_row = all_rows[subsequent_idx]
                for cell_idx in range(min(len(sub_row), 4)):
                    cell_text = str(sub_row[cell_idx]).strip() if sub_row[cell_idx] is not None else ""
                    if "制表" in cell_text:
                        maker = _extract_maker(cell_text)
                        break
                if maker:
                    break
            break

        # Check for blank row
        if _is_blank_row(row):
            continue

        # Parse data row
        serial_str, serial_was_float = _normalize_numeric(col0_raw)
        name = str(row[1]).strip() if row[1] is not None else ""
        id_card, id_was_numeric = _normalize_id_card(row[2])
        amount_str, _ = _normalize_numeric(row[3])

        if id_was_numeric:
            warnings.append(
                f"第{idx + 1}行身份证号码为数字格式，可能丢失精度: {id_card}"
            )

        if not name:
            continue

        rows.append(
            MealRefundRow(
                serial=serial_str,
                name=name,
                id_card=id_card,
                amount=amount_str,
            )
        )

    if not rows:
        raise ValueError("未找到数据行")

    # Validate sum
    if total_amount:
        try:
            row_sum = sum(Decimal(r.amount) for r in rows if r.amount)
            expected = Decimal(total_amount)
            if row_sum != expected:
                warnings.append(
                    f"明细金额合计 ({row_sum}) 与表格总计 ({expected}) 不一致"
                )
        except InvalidOperation:
            warnings.append("无法验证金额合计（存在非数字金额）")

    return MealRefundSheet(
        title=title,
        rows=rows,
        total_amount=total_amount,
        total_chinese=total_chinese,
        maker=maker,
        warnings=warnings,
    )
