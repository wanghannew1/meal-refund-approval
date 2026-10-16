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


def _normalize_money(value: object) -> str:
    """Normalize a money cell to a plain decimal string.

    Blank/None cells become "0" so amount widgets never receive the
    literal text "None" or an empty value.
    """
    if value is None:
        return "0"
    if isinstance(value, str) and not value.strip():
        return "0"
    text, _ = _normalize_numeric(value)
    return text or "0"


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


@dataclass(frozen=True)
class BonusRow:
    """A single data row from the bonus (科研业绩津贴) spreadsheet.

    Attributes:
        serial: Normalized serial number string (preview only, not uploaded).
        employee_id: Employee ID string (zero-padded text preserved).
        name: Employee name.
        id_card: Identity card number (text, may end in X).
        bonus: Plain decimal bonus amount string.
        tax: Plain decimal tax amount string.
        net_amount: Plain decimal net amount string.
        adjustment: 个人其他调整 amount string; "0" when the Excel
            column is absent or the cell is blank.
    """

    serial: str
    employee_id: str
    name: str
    id_card: str
    bonus: str
    tax: str
    net_amount: str
    adjustment: str = "0"


@dataclass(frozen=True)
class BonusSheet:
    """Parsed result of a bonus spreadsheet.

    Attributes:
        title: Sheet title from row 0 col 0.
        rows: List of data rows.
        bonus_total: Normalized bonus total string from 合计 row.
        tax_total: Normalized tax total string from 合计 row.
        net_total: Normalized net total string from 合计 row.
        maker: Name of the person who prepared the sheet.
        adjustment_total: 合计 row's 个人其他调整 cell; "0" when the
            column is absent.
        warnings: List of warning messages (Chinese text).
    """

    title: str
    rows: list[BonusRow]
    bonus_total: str
    tax_total: str
    net_total: str
    maker: str
    adjustment_total: str = "0"
    warnings: list[str] = field(default_factory=list)



def _resolve_bonus_header(cell_value: object) -> str | None:
    """Return the canonical field name for a bonus header cell, or None.

    Mapping (first match wins per field):
        serial:      stripped text == "序号"
        employee_id: "工号" in text
        name:        stripped text == "姓名"
        id_card:     text.startswith("身份证")
        bonus:       text.startswith("奖金")
        tax:         "个人所得税" in text  (matches 应纳个人所得税;
                     does NOT match 个人其他调整)
        net:         "实发" in text
        adjustment:  "调整" in text  (OPTIONAL)
    """
    if cell_value is None:
        return None
    text = str(cell_value).strip()
    if not text:
        return None
    if text == "序号":
        return "serial"
    if "工号" in text:
        return "employee_id"
    if text == "姓名":
        return "name"
    if text.startswith("身份证"):
        return "id_card"
    if text.startswith("奖金"):
        return "bonus"
    if "个人所得税" in text or text == "个税":
        return "tax"
    if "实发" in text:
        return "net"
    if "调整" in text:
        return "adjustment"
    return None


def parse_bonus_excel(path: str | Path) -> BonusSheet:
    """Parse a bonus (科研业绩津贴) detail spreadsheet.

    Supports .xls (xlrd) and .xlsx (openpyxl) formats.  The header row
    is located by matching column labels by name (not by fixed position),
    so an optional 个人其他调整 column may appear anywhere among the
    value columns without corrupting the other fields.  Data rows follow
    until a 合计 row (cell at the serial index == "合计").

    Args:
        path: Path to the Excel file.

    Returns:
        BonusSheet with parsed data.

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

    # Required field names (must all resolve); adjustment is optional.
    required_fields = {"serial", "employee_id", "name", "id_card", "bonus", "tax", "net"}

    header_row_idx: int | None = None
    col_map: dict[str, int] = {}
    for idx, row in enumerate(all_rows):
        candidate: dict[str, int] = {}
        for col_idx, cell in enumerate(row):
            field_name = _resolve_bonus_header(cell)
            if field_name is not None and field_name not in candidate:
                candidate[field_name] = col_idx
        if required_fields.issubset(candidate):
            header_row_idx = idx
            col_map = candidate
            break

    if header_row_idx is None:
        raise ValueError(
            '未找到表头行（需要包含"序号""工号""姓名""身份证号"'
            '"奖金""个人所得税""实发金额"列）'
        )

    idx_serial = col_map["serial"]
    idx_emp = col_map["employee_id"]
    idx_name = col_map["name"]
    idx_id = col_map["id_card"]
    idx_bonus = col_map["bonus"]
    idx_tax = col_map["tax"]
    idx_net = col_map["net"]
    idx_adj: int | None = col_map.get("adjustment")

    # Maximum column index we need to pad rows to.
    max_needed = max(idx_serial, idx_emp, idx_name, idx_id, idx_bonus, idx_tax, idx_net)
    if idx_adj is not None:
        max_needed = max(max_needed, idx_adj)

    rows: list[BonusRow] = []
    warnings: list[str] = []
    bonus_total = ""
    tax_total = ""
    net_total = ""
    adjustment_total = "0"
    maker = ""

    data_start = header_row_idx + 1
    totals_row_idx: int | None = None
    for idx in range(data_start, len(all_rows)):
        row = all_rows[idx]

        # Pad row to cover all needed indices.
        while len(row) <= max_needed:
            row.append(None)

        col0_raw = row[idx_serial]
        col0_str = str(col0_raw).strip() if col0_raw is not None else ""

        if col0_str == "合计":
            totals_row_idx = idx
            raw_bonus = str(row[idx_bonus]).strip() if row[idx_bonus] is not None else ""
            raw_tax = str(row[idx_tax]).strip() if row[idx_tax] is not None else ""
            raw_net = str(row[idx_net]).strip() if row[idx_net] is not None else ""
            bonus_total = _extract_total_amount(raw_bonus)
            tax_total = _extract_total_amount(raw_tax)
            net_total = _extract_total_amount(raw_net)
            if idx_adj is not None:
                raw_adj = str(row[idx_adj]).strip() if row[idx_adj] is not None else ""
                adjustment_total = _extract_total_amount(raw_adj) if raw_adj else "0"
            break

        if _is_blank_row(row):
            continue

        name = str(row[idx_name]).strip() if row[idx_name] is not None else ""
        if not name:
            continue

        serial_str, _ = _normalize_numeric(col0_raw if col0_raw is not None else "")

        # Preserve zero-padded text IDs; only normalize float/int.
        emp_raw = row[idx_emp]
        if isinstance(emp_raw, (int, float)):
            employee_id = str(int(emp_raw))
        elif emp_raw is None:
            employee_id = ""
        else:
            employee_id = str(emp_raw).strip()

        id_card, id_was_numeric = _normalize_id_card(row[idx_id] if row[idx_id] is not None else "")
        bonus_str = _normalize_money(row[idx_bonus])
        tax_str = _normalize_money(row[idx_tax])
        net_str = _normalize_money(row[idx_net])

        adj_str = _normalize_money(row[idx_adj]) if idx_adj is not None else "0"

        if id_was_numeric:
            warnings.append(
                f"第{idx + 1}行身份证号码为数字格式，可能丢失精度: {id_card}"
            )

        rows.append(
            BonusRow(
                serial=serial_str,
                employee_id=employee_id,
                name=name,
                id_card=id_card,
                bonus=bonus_str,
                tax=tax_str,
                net_amount=net_str,
                adjustment=adj_str,
            )
        )

    if not rows:
        raise ValueError("未找到数据行")

    if totals_row_idx is not None:
        for subsequent_idx in range(totals_row_idx + 1, min(totals_row_idx + 6, len(all_rows))):
            sub_row = all_rows[subsequent_idx]
            for cell_idx in range(len(sub_row)):
                cell_text = str(sub_row[cell_idx]).strip() if sub_row[cell_idx] is not None else ""
                if "制表" in cell_text:
                    match = re.search(r"制表人?\s*[：:]\s*(.+)", cell_text)
                    if match:
                        maker = match.group(1).strip()
                        break
            if maker:
                break

    # Sum-check warnings for bonus / tax / net (always) and adjustment (when present).
    if bonus_total:
        try:
            row_sum = sum(Decimal(r.bonus) for r in rows if r.bonus)
            expected = Decimal(bonus_total)
            if row_sum != expected:
                warnings.append(
                    f"奖金明细合计 ({row_sum}) 与表格合计 ({expected}) 不一致"
                )
        except InvalidOperation:
            pass

    if tax_total:
        try:
            row_sum = sum(Decimal(r.tax) for r in rows if r.tax)
            expected = Decimal(tax_total)
            if row_sum != expected:
                warnings.append(
                    f"个人所得税明细合计 ({row_sum}) 与表格合计 ({expected}) 不一致"
                )
        except InvalidOperation:
            pass

    if net_total:
        try:
            row_sum = sum(Decimal(r.net_amount) for r in rows if r.net_amount)
            expected = Decimal(net_total)
            if row_sum != expected:
                warnings.append(
                    f"实发价格明细合计 ({row_sum}) 与表格合计 ({expected}) 不一致"
                )
        except InvalidOperation:
            pass

    if idx_adj is not None and adjustment_total and adjustment_total != "0":
        try:
            row_sum = sum(Decimal(r.adjustment) for r in rows if r.adjustment)
            expected = Decimal(adjustment_total)
            if row_sum != expected:
                warnings.append(
                    f"个人其他调整明细合计 ({row_sum}) 与表格合计 ({expected}) 不一致"
                )
        except InvalidOperation:
            pass

    # Per-row net check: skip when adjustment is non-zero (sign convention unknown).
    for i, r in enumerate(rows, start=1):
        try:
            if r.bonus and r.tax and r.net_amount:
                if Decimal(r.adjustment) != 0:
                    continue
                b = Decimal(r.bonus)
                t = Decimal(r.tax)
                n = Decimal(r.net_amount)
                if b - t != n:
                    warnings.append(
                        f"第{i}行 实发金额 与 奖金-个人所得税 不一致: "
                        f"{r.bonus} - {r.tax} = {b - t}, 但实发为 {r.net_amount}"
                    )
        except InvalidOperation:
            pass

    return BonusSheet(
        title=title,
        rows=rows,
        bonus_total=bonus_total,
        tax_total=tax_total,
        net_total=net_total,
        maker=maker,
        adjustment_total=adjustment_total,
        warnings=warnings,
    )
