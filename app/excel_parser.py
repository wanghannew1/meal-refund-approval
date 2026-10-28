"""
Excel parser for meal refund detail spreadsheets.

Supports both legacy .xls (via xlrd) and modern .xlsx (via openpyxl)
formats.  Locates the header row by matching column labels, extracts
data rows, and normalizes numeric values to clean strings.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
        None / blank → ('', False).
    """
    if value is None:
        return "", False
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
    """Extract maker name from text like '制表：孙世多' or '制表人：张艳艳'.

    Args:
        text: Raw cell text from the signature row.

    Returns:
        Maker name string, or empty string if not found.
    """
    match = re.search(r"制表人?\s*[：:]\s*(.+)", text.strip())
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


def _resolve_meal_header(cell_value: object) -> str | None:
    """Return the canonical field name for a meal-refund header cell, or None.

    Mapping (first match wins per field):
        serial:   stripped text == "序号"
        name:     stripped text == "姓名"
        id_card:  text in {"身份证号码", "身份证号", "证件号"}
        amount:   stripped text == "金额"
    """
    if cell_value is None:
        return None
    text = str(cell_value).strip()
    if not text:
        return None
    if text == "序号":
        return "serial"
    if text == "姓名":
        return "name"
    if text in ("身份证号码", "身份证号", "证件号"):
        return "id_card"
    if text == "金额":
        return "amount"
    return None


@dataclass(frozen=True)
class SheetInfo:
    """Metadata about one sheet in a workbook.

    Attributes:
        name: Sheet name.
        has_valid_header: True if the sheet contains a recognized meal-refund
            header row (序号 + 姓名 + 金额 columns).
        is_default: True for the sheet the parser reads when no sheet is
            given: .xlsx -> the workbook's ACTIVE sheet; .xls -> the LAST
            sheet. Matches the read defaults in _read_xlsx/_read_xls.
    """

    name: str
    has_valid_header: bool
    is_default: bool = False


def _has_meal_header(rows: list[list[object]]) -> bool:
    """Return True if any row contains the required meal header fields."""
    required = {"serial", "name", "amount"}
    for row in rows:
        found: set[str] = set()
        for cell in row:
            field = _resolve_meal_header(cell)
            if field is not None:
                found.add(field)
        if required.issubset(found):
            return True
    return False


def _xls_saved_selected_index(path: str | Path) -> int | None:
    """Return the 0-based index of the sheet that was selected when the .xls
    file was saved (BIFF record WINDOW1, field itabCur), or None if it cannot
    be determined.

    The saved-selected tab is recoverable from the BIFF binary even though
    xlrd exposes no API for it. WINDOW1 record layout (type 0x003D, data
    length 0x12 = 18 bytes): 9 little-endian uint16 fields — xWn, yWn, dxWn,
    dyWn, grbit, itabCur, itabFirst, ctabSel, wTabRatio. itabCur is at data
    offset 10 (field index 5), 0-based sheet index.

    Preferred method: extract the Workbook stream via xlrd.compdoc.CompDoc
    (stream name b'Workbook', fall back to b'Book') and iterate BIFF records
    properly (2-byte type + 2-byte length + data, walking pos += 4+len from
    stream start).

    Fallback method: scan the raw file bytes for the exact 4-byte header
    b'\\x3d\\x00\\x12\\x00' and read itabCur from the following record data.

    Any exception → return None (never raise out of this helper).
    """
    import struct

    try:
        path = Path(path)
        raw_bytes = path.read_bytes()

        # Preferred method: extract Workbook stream via CompDoc
        try:
            import xlrd.compdoc

            compdoc = xlrd.compdoc.CompDoc(raw_bytes)
            # Try 'Workbook' first, then 'Book' (older BIFF versions)
            stream = None
            for stream_name in (b"Workbook", b"Book"):
                try:
                    stream = compdoc.get_named_stream(stream_name)
                    if stream is not None:
                        break
                except Exception:
                    continue

            if stream is not None:
                # Iterate BIFF records: 2-byte type + 2-byte length + data
                pos = 0
                while pos + 4 <= len(stream):
                    rec_type = struct.unpack_from("<H", stream, pos)[0]
                    rec_len = struct.unpack_from("<H", stream, pos + 2)[0]
                    if rec_type == 0x003D and rec_len >= 12:
                        # WINDOW1 record found; read itabCur at data offset 10
                        itab_cur = struct.unpack_from("<H", stream, pos + 4 + 10)[0]
                        return itab_cur
                    pos += 4 + rec_len
        except Exception:
            pass  # Fall through to raw-byte scan

        # Fallback method: scan raw bytes for WINDOW1 header
        # Header: type 0x003D (little-endian: b'\\x3d\\x00'), length 0x0012 (b'\\x12\\x00')
        header = b"\x3d\x00\x12\x00"
        idx = raw_bytes.find(header)
        if idx != -1 and idx + 4 + 12 <= len(raw_bytes):
            itab_cur = struct.unpack_from("<H", raw_bytes, idx + 4 + 10)[0]
            return itab_cur

        return None
    except Exception:
        return None


def list_sheets(path: str | Path) -> list[SheetInfo]:
    """Return sheet names and header-validity for a workbook.

    Supports .xls (xlrd) and .xlsx (openpyxl). Does not raise on unreadable
    files — returns an empty list with a best-effort error entry.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    results: list[SheetInfo] = []

    if suffix == ".xlsx":
        import openpyxl

        try:
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        except Exception:
            return results
        # Default sheet rule must match _read_xlsx: the workbook's ACTIVE sheet.
        active = wb.active
        default_name = active.title if active is not None else None
        for name in wb.sheetnames:
            ws = wb[name]
            rows: list[list[object]] = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows.append(list(row))
                if i >= 30:
                    break
            results.append(
                SheetInfo(
                    name=name,
                    has_valid_header=_has_meal_header(rows),
                    is_default=(name == default_name),
                )
            )
        wb.close()
    elif suffix == ".xls":
        import xlrd

        try:
            wb = xlrd.open_workbook(str(path))
        except Exception:
            return results
        # Default sheet rule: the saved-selected tab (WINDOW1.itabCur) if
        # recoverable, else the LAST sheet (fallback for older/corrupt files).
        saved_idx = _xls_saved_selected_index(path)
        nsheets = wb.nsheets
        if saved_idx is not None and 0 <= saved_idx < nsheets:
            default_index = saved_idx
        else:
            default_index = nsheets - 1
        for i, name in enumerate(wb.sheet_names()):
            sh = wb.sheet_by_name(name)
            rows = []
            for r in range(min(sh.nrows, 31)):
                rows.append([sh.cell(r, c).value for c in range(sh.ncols)])
            results.append(
                SheetInfo(
                    name=name,
                    has_valid_header=_has_meal_header(rows),
                    is_default=(i == default_index),
                )
            )
    return results


def _read_xls(path: Path, sheet_name: str | None = None, sheet_index: int | None = None) -> tuple[str, list[list[object]]]:
    """Read an .xls file and return (title, all_rows).

    Args:
        path: Path to the .xls file.
        sheet_name: Optional sheet name to read.
        sheet_index: Optional sheet index to read. If neither sheet_name nor
            sheet_index is given, the saved-selected tab (WINDOW1.itabCur) is
            used if recoverable, else the LAST sheet (fallback).

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

    if sheet_name is not None:
        try:
            sh = wb.sheet_by_name(sheet_name)
        except Exception:
            sh = wb.sheet_by_index(wb.nsheets - 1)
    elif sheet_index is not None:
        try:
            sh = wb.sheet_by_index(sheet_index)
        except Exception:
            sh = wb.sheet_by_index(wb.nsheets - 1)
    else:
        # Default sheet rule: saved-selected tab if recoverable, else last.
        saved_idx = _xls_saved_selected_index(path)
        if saved_idx is not None and 0 <= saved_idx < wb.nsheets:
            sh = wb.sheet_by_index(saved_idx)
        else:
            sh = wb.sheet_by_index(wb.nsheets - 1)

    title = str(sh.cell(0, 0).value) if sh.nrows > 0 else ""

    all_rows: list[list[object]] = []
    for r in range(sh.nrows):
        row: list[object] = []
        for c in range(sh.ncols):
            row.append(sh.cell(r, c).value)
        all_rows.append(row)

    return title, all_rows


def _read_xlsx(path: Path, sheet_name: str | None = None, sheet_index: int | None = None) -> tuple[str, list[list[object]]]:
    """Read an .xlsx file and return (title, all_rows).

    Args:
        path: Path to the .xlsx file.
        sheet_name: Optional sheet name to read.
        sheet_index: Optional sheet index to read. If neither is given, the
            workbook's ACTIVE sheet is used (openpyxl default).

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

    if sheet_name is not None and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    elif sheet_index is not None and 0 <= sheet_index < len(wb.sheetnames):
        ws = wb[wb.sheetnames[sheet_index]]
    else:
        ws = wb.active

    all_rows: list[list[object]] = []
    for row in ws.iter_rows(values_only=True):
        all_rows.append(list(row))

    wb.close()

    title = ""
    if all_rows and all_rows[0]:
        title = str(all_rows[0][0]) if all_rows[0][0] is not None else ""

    return title, all_rows


def parse_meal_refund_excel(
    path: str | Path,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
) -> MealRefundSheet:
    """Parse a meal refund detail spreadsheet.

    Supports .xls (xlrd) and .xlsx (openpyxl) formats.  The header row
    is located by matching column labels by name (not by fixed position):
      - Required: 序号, 姓名, 金额
      - Optional: 身份证号码 / 身份证号 / 证件号 (any alias)
    Extra columns are ignored; column order is flexible.

    Sheet selection:
      - sheet_name: pick by exact sheet name
      - sheet_index: pick by 0-based index
      - Neither given: .xlsx -> active sheet; .xls -> last sheet

    Two total-row labels are recognized:
      - 返餐费 (meal refund) format: col at serial index == "总计"
      - 工会会费 (union fee) format: col at serial index == "合计"

    Duplicate names (same name on multiple rows) are kept as independent
    rows; a non-blocking warning lists duplicated names and counts.

    Args:
        path: Path to the Excel file.
        sheet_name: Optional sheet name to parse.
        sheet_index: Optional 0-based sheet index to parse.

    Returns:
        MealRefundSheet with parsed data.

    Raises:
        ValueError: If the file is unreadable, header not found,
            or no data rows present.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == ".xls":
        title, all_rows = _read_xls(path, sheet_name=sheet_name, sheet_index=sheet_index)
    elif suffix == ".xlsx":
        title, all_rows = _read_xlsx(path, sheet_name=sheet_name, sheet_index=sheet_index)
    else:
        raise ValueError(f"不支持的文件格式: {suffix}，仅支持 .xls 和 .xlsx")

    # Locate header row by matching column labels
    required_fields = {"serial", "name", "amount"}
    header_row_idx: int | None = None
    col_map: dict[str, int] = {}
    for idx, row in enumerate(all_rows):
        candidate: dict[str, int] = {}
        for col_idx, cell in enumerate(row):
            field_name = _resolve_meal_header(cell)
            if field_name is not None and field_name not in candidate:
                candidate[field_name] = col_idx
        if required_fields.issubset(candidate):
            header_row_idx = idx
            col_map = candidate
            break

    if header_row_idx is None:
        raise ValueError('未找到表头行（需要包含"序号""姓名""金额"列）')

    idx_serial = col_map["serial"]
    idx_name = col_map["name"]
    idx_amount = col_map["amount"]
    idx_id: int | None = col_map.get("id_card")

    # Maximum column index we need to pad rows to.
    max_needed = max(idx_serial, idx_name, idx_amount)
    if idx_id is not None:
        max_needed = max(max_needed, idx_id)

    # Parse data rows
    rows: list[MealRefundRow] = []
    warnings: list[str] = []
    total_amount = ""
    total_chinese = ""
    maker = ""
    skipped_count = 0
    name_counts: dict[str, int] = {}

    data_start = header_row_idx + 1
    for idx in range(data_start, len(all_rows)):
        row = all_rows[idx]

        # Pad row to cover all needed indices.
        while len(row) <= max_needed:
            row.append(None)

        col0_raw = row[idx_serial]

        # Check for 总计 (返餐费) or 合计 (工会会费) total row
        col0_str = str(col0_raw).strip() if col0_raw is not None else ""
        if col0_str in ("总计", "合计"):
            # Extract total info.
            # total_chinese comes from the name column; may be empty for 工会会费.
            total_chinese = str(row[idx_name]).strip() if row[idx_name] is not None else ""
            raw_total = str(row[idx_amount]).strip() if row[idx_amount] is not None else ""
            total_amount = _extract_total_amount(raw_total)

            # Look for maker in subsequent rows. Scan ALL columns of each
            # row so "制表：..." in any column is still found.
            for subsequent_idx in range(idx + 1, min(idx + 8, len(all_rows))):
                sub_row = all_rows[subsequent_idx]
                for cell_idx in range(len(sub_row)):
                    cell_text = str(sub_row[cell_idx]).strip() if sub_row[cell_idx] is not None else ""
                    if "制表" in cell_text:
                        maker = _extract_maker(cell_text)
                        break
                if maker:
                    break
            break

        # Defensive end-of-data guard: signature rows like "总经理签字："
        # must never be parsed as data rows.
        if "签字" in col0_str:
            break

        # Check for blank row
        if _is_blank_row(row):
            skipped_count += 1
            continue

        # Parse data row
        serial_str, serial_was_float = _normalize_numeric(col0_raw)
        name = str(row[idx_name]).strip() if row[idx_name] is not None else ""
        if not name:
            skipped_count += 1
            continue

        # Track name counts for duplicate detection
        name_counts[name] = name_counts.get(name, 0) + 1

        # ID card: use mapped column if present, else empty string
        if idx_id is not None:
            id_card, id_was_numeric = _normalize_id_card(row[idx_id])
        else:
            id_card, id_was_numeric = "", False

        amount_str, _ = _normalize_numeric(row[idx_amount])

        if id_was_numeric:
            warnings.append(
                f"第{idx + 1}行身份证号码为数字格式，可能丢失精度: {id_card}"
            )

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

    # Duplicate name warning
    dup_names = {n: c for n, c in name_counts.items() if c > 1}
    if dup_names:
        dup_parts = [f"{n}({c}次)" for n, c in sorted(dup_names.items())]
        warnings.append(f"存在重复姓名: {', '.join(dup_parts)}（各行独立保留）")

    # Skipped row warning
    if skipped_count > 0:
        warnings.append(f"跳过 {skipped_count} 行（空行/缺姓名/签字行）")

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


# ── 五险一金数据模型 ──────────────────────────────────────────────────


@dataclass(frozen=True)
class InsuranceFeeItem:
    """单个费用对：费用名称 + 金额字符串。"""

    name: str
    amount: str


@dataclass(frozen=True)
class InsuranceUnitRow:
    """分单位表的一行数据：单位名称、参保人数、费用对列表。"""

    unit_name: str
    headcount: int
    fees: list[InsuranceFeeItem]


@dataclass(frozen=True)
class InsuranceSheet:
    """解析后的五险一金分单位表。"""

    title: str
    rows: list[InsuranceUnitRow]
    maker: str
    warnings: list[str] = field(default_factory=list)


_INSURANCE_EXCLUDED = frozenset({"序号"})
_UNIT_HEADER_LABELS = frozenset({"单位名称", "单位", "公司名称", "企业名称"})
_PERSON_LEVEL_MARKERS = frozenset({"所属单位", "姓名", "身份证号", "身份证号码"})


def _insurance_normalize_text(cell: object) -> str:
    """移除全部空白（含换行）后的表头文本。"""
    if cell is None:
        return ""
    return re.sub(r"\s+", "", str(cell)).strip()


def _is_unit_header_label(text: str) -> bool:
    return text in _UNIT_HEADER_LABELS


def _is_headcount_label(text: str) -> bool:
    return text.endswith("人数")


def _is_excluded_label(text: str) -> bool:
    return (
        text in _INSURANCE_EXCLUDED
        or _is_unit_header_label(text)
        or _is_headcount_label(text)
    )


def _find_unit_row(rows: list[list[object]]) -> int | None:
    for idx, row in enumerate(rows):
        for cell in row:
            if _is_unit_header_label(_insurance_normalize_text(cell)):
                return idx
    return None


def _group_label_for_col(header_rows: list[list[object]], name_row_idx: int, col: int, merged_groups: list[tuple[int, int, str]] | None = None) -> str:
    """金额列的组标签：优先命中合并单元格区间，其次同列组标签，最后向左查找。"""
    group_row_idx = name_row_idx - 1
    if merged_groups:
        for min_c, max_c, label in merged_groups:
            if min_c <= col <= max_c and label and not _is_excluded_label(label):
                return label
    if group_row_idx < 0:
        return ""
    above = header_rows[group_row_idx]
    if col < len(above):
        same_col = _insurance_normalize_text(above[col])
        if same_col and not _is_excluded_label(same_col):
            return same_col
    for c in range(col - 1, -1, -1):
        if c < len(above):
            text = _insurance_normalize_text(above[c])
            if text and not _is_excluded_label(text):
                return text
    return ""


def _is_insurance_monthly_header(rows: list[list[object]]) -> bool:
    """人员级表头含 所属单位/姓名/身份证号，一律判False；
    分单位月度表头需出现单位列与任意“人数”列，具体列名/组名不做写死。"""
    for row in rows[:40]:
        for cell in row:
            text = _insurance_normalize_text(cell)
            if text in _PERSON_LEVEL_MARKERS:
                return False
    has_unit = False
    has_count = False
    for row in rows[:40]:
        for cell in row:
            text = _insurance_normalize_text(cell)
            if _is_unit_header_label(text):
                has_unit = True
            if _is_headcount_label(text):
                has_count = True
    return has_unit and has_count


def _resolve_insurance_columns(header_rows: list[list[object]], merged_groups: list[tuple[int, int, str]] | None = None) -> tuple[int, int, list[tuple[str, int]]] | None:
    """解析分单位表头：(单位名称列, 参保人数列, [(组标签, 金额列), ...])。"""
    unit_row_idx = _find_unit_row(header_rows)
    if unit_row_idx is None:
        return None
    unit_row = header_rows[unit_row_idx]
    unit_col = None
    for col, cell in enumerate(unit_row):
        if _is_unit_header_label(_insurance_normalize_text(cell)):
            unit_col = col
            break
    if unit_col is None:
        return None

    name_row_idx = unit_row_idx + 1
    if name_row_idx >= len(header_rows):
        return None
    name_row = header_rows[name_row_idx]

    headcount_col = None
    for col, cell in enumerate(name_row):
        text = _insurance_normalize_text(cell)
        if _is_headcount_label(text):
            headcount_col = col
            break
    if headcount_col is None:
        return None

    money_cols: list[tuple[str, int]] = []
    for col, cell in enumerate(name_row):
        text = _insurance_normalize_text(cell)
        if not text or _is_excluded_label(text):
            continue
        group = _group_label_for_col(header_rows, name_row_idx, col, merged_groups)
        money_cols.append((group, col))
    if not money_cols:
        return None
    return (unit_col, headcount_col, money_cols)


def _sheet_merged_group_ranges(path: Path, suffix: str, sheet_name: str | None, sheet_index: int | None, group_row_idx: int) -> list[tuple[int, int, str]]:
    """返回指定行为合并单元格的组标签区间：(最小列0-based, 最大列0-based, 标签)。"""
    if group_row_idx < 0:
        return []
    if suffix == ".xlsx":
        import openpyxl

        try:
            wb = openpyxl.load_workbook(str(path), data_only=False)
        except Exception:
            return []
        try:
            if sheet_name is not None and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            elif sheet_index is not None and 0 <= sheet_index < len(wb.worksheets):
                ws = wb.worksheets[sheet_index]
            else:
                ws = wb.active
            ranges: list[tuple[int, int, str]] = []
            target_row = group_row_idx + 1
            for mr in ws.merged_cells.ranges:
                if mr.min_row <= target_row <= mr.max_row:
                    label = ws.cell(mr.min_row, mr.min_col).value
                    if label is not None and str(label).strip():
                        ranges.append((mr.min_col - 1, mr.max_col - 1, str(label)))
            return ranges
        finally:
            wb.close()
    elif suffix == ".xls":
        import xlrd

        try:
            wb = xlrd.open_workbook(str(path))
        except Exception:
            return []
        try:
            if sheet_name is not None:
                try:
                    sh = wb.sheet_by_name(sheet_name)
                except Exception:
                    sh = wb.sheet_by_index(0)
            elif sheet_index is not None:
                try:
                    sh = wb.sheet_by_index(sheet_index)
                except Exception:
                    sh = wb.sheet_by_index(0)
            else:
                sh = wb.sheet_by_index(0)
            ranges = []
            merged = getattr(sh, "merged_cells", None) or []
            target_row = group_row_idx
            for rng in merged:
                if len(rng) == 4:
                    r1, r2, c1, c2 = rng
                    if r1 <= target_row <= r2:
                        label = sh.cell_value(r1, c1)
                        if label is not None and str(label).strip():
                            ranges.append((c1, c2, str(label)))
            return ranges
        finally:
            pass
    return []


def _normalize_insurance_amount(value: object) -> str | None:
    """空单元格返回None（不提交）；非空金额统一2位小数字符串。"""
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    text, _ = _normalize_numeric(value)
    if not text:
        return None
    return str(Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def list_insurance_sheets(path: str | Path) -> list[SheetInfo]:
    """返回五险一金分单位表的 sheet 信息，仅标记通过表头识别的有效月度表。"""
    path = Path(path)
    suffix = path.suffix.lower()
    results: list[SheetInfo] = []

    if suffix == ".xlsx":
        import openpyxl

        try:
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        except Exception:
            return results
        active = wb.active
        default_name = active.title if active is not None else None
        for name in wb.sheetnames:
            ws = wb[name]
            rows: list[list[object]] = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows.append([str(c) if c is not None else "" for c in row])
                if i >= 40:
                    break
            results.append(
                SheetInfo(
                    name=name,
                    has_valid_header=_is_insurance_monthly_header(rows),
                    is_default=(name == default_name),
                )
            )
        wb.close()
    elif suffix == ".xls":
        import xlrd

        try:
            wb = xlrd.open_workbook(str(path))
        except Exception:
            return results
        saved_idx = _xls_saved_selected_index(path)
        nsheets = wb.nsheets
        if saved_idx is not None and 0 <= saved_idx < nsheets:
            default_index = saved_idx
        else:
            default_index = nsheets - 1
        for i, name in enumerate(wb.sheet_names()):
            sh = wb.sheet_by_name(name)
            rows = [
                [str(sh.cell_value(r, c)) for c in range(sh.ncols)]
                for r in range(min(sh.nrows, 41))
            ]
            results.append(
                SheetInfo(
                    name=name,
                    has_valid_header=_is_insurance_monthly_header(rows),
                    is_default=(i == default_index),
                )
            )
    return results


def parse_insurance_excel(
    path: str | Path,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
) -> InsuranceSheet:
    """解析五险一金分单位Excel，返回 InsuranceSheet。

    规则：
    - 按表头标签而非列位识别；仅解析第一张通过 _is_insurance_monthly_header 的 sheet，
      除非调用方指定 sheet_name / sheet_index。
    - 遇到首格为“合计”的行终止，合计行之后不再解析签批/转账块。
    - 空金额不生成费用对；非空金额统一2位小数；重复列名加组前缀。
    - 警告仅描述位置，不回显单元格数值/姓名/证件号。
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if sheet_name is not None:
        title, all_rows = _read_xls(path, sheet_name=sheet_name) if suffix == ".xls" else _read_xlsx(path, sheet_name=sheet_name)
    elif sheet_index is not None:
        title, all_rows = _read_xls(path, sheet_index=sheet_index) if suffix == ".xls" else _read_xlsx(path, sheet_index=sheet_index)
    else:
        infos = list_insurance_sheets(path)
        valid = [s for s in infos if s.has_valid_header]
        if not valid:
            raise ValueError("未找到有效分单位表")
        default = next((s for s in valid if s.is_default), valid[0])
        title, all_rows = (
            _read_xls(path, sheet_name=default.name)
            if suffix == ".xls"
            else _read_xlsx(path, sheet_name=default.name)
        )

    header_block = all_rows[:40]
    unit_row_idx = _find_unit_row(header_block)
    group_row_idx = unit_row_idx if unit_row_idx is not None else -1
    merged_groups = _sheet_merged_group_ranges(path, suffix, sheet_name, sheet_index, group_row_idx) if group_row_idx >= 0 else []
    resolved = _resolve_insurance_columns(header_block, merged_groups)
    if resolved is None:
        raise ValueError("未找到有效分单位表头")
    unit_col, headcount_col, money_cols = resolved
    if len(money_cols) > 17:
        raise ValueError("费用项超过17项，无法提交")
    unit_row_idx = _find_unit_row(header_block)
    name_row_idx = unit_row_idx + 1
    data_start = name_row_idx + 1

    max_col = max(unit_col, headcount_col, max(col for _, col in money_cols))

    # 费用命名：同 sheet 归一化后重名的列名全部加组前缀，否则裸名。
    normalized_names = [
        _insurance_normalize_text(header_block[name_row_idx][col])
        for _, col in money_cols
    ]
    # 多级表头：费用名 = 第一级组标签-第二级列名。
    # 没有第一级组标签的单层表保持裸列名。
    fee_names: list[str] = []
    for (group, col), name in zip(money_cols, normalized_names):
        if group:
            fee_names.append(f"{group}-{name}")
        else:
            fee_names.append(name)

    rows: list[InsuranceUnitRow] = []
    warnings: list[str] = []
    maker = ""
    total_row_idx: int | None = None

    col_sums = [Decimal(0) for _ in money_cols]

    for idx in range(data_start, len(all_rows)):
        row = all_rows[idx]
        while len(row) <= max_col:
            row.append(None)

        col0_raw = row[0] if len(row) > 0 else None
        col0_str = _insurance_normalize_text(col0_raw)
        if col0_str == "合计":
            total_row_idx = idx
            break

        if _is_blank_row(row):
            continue

        unit_name = ""
        if unit_col < len(row) and row[unit_col] is not None:
            unit_name = str(row[unit_col]).strip()

        headcount = 0
        raw_hc = row[headcount_col] if headcount_col < len(row) else None
        if raw_hc is not None:
            hc_text = str(raw_hc).strip()
            if hc_text:
                try:
                    headcount = int(float(hc_text))
                except ValueError:
                    headcount = 0
                    warnings.append(f"第{idx + 1}行参保人数非数字")
            else:
                warnings.append(f"第{idx + 1}行参保人数为空")

        fees: list[InsuranceFeeItem] = []
        for i, ((group, col), fee_name) in enumerate(zip(money_cols, fee_names)):
            raw = row[col] if col < len(row) else None
            amount = _normalize_insurance_amount(raw)
            if amount is not None:
                fees.append(InsuranceFeeItem(name=fee_name, amount=amount))
                col_sums[i] += Decimal(amount)

        rows.append(InsuranceUnitRow(unit_name=unit_name, headcount=headcount, fees=fees))

    if not rows:
        raise ValueError("未找到数据行")

    # 合计行逐列校验：不回显数值，仅描述位置。
    if total_row_idx is not None:
        total_row = all_rows[total_row_idx]
        for i, ((group, col), fee_name) in enumerate(zip(money_cols, fee_names)):
            raw_total = total_row[col] if col < len(total_row) else None
            total_text = str(raw_total).strip() if raw_total is not None else ""
            try:
                total_val = Decimal(total_text) if total_text else None
            except InvalidOperation:
                total_val = None
            if total_val is None:
                if col_sums[i] != 0:
                    warnings.append(f"第{col + 1}列合计为空")
            elif col_sums[i] != total_val:
                warnings.append(f"第{col + 1}列合计与明细不一致")

    # 签批行提取制表人，扫描合计行之后最多8行。
    if total_row_idx is not None:
        for subsequent_idx in range(total_row_idx + 1, min(total_row_idx + 9, len(all_rows))):
            sub_row = all_rows[subsequent_idx]
            for cell in sub_row:
                cell_text = str(cell).strip() if cell is not None else ""
                if "制表" in cell_text:
                    maker = _extract_maker(cell_text)
                    break
            if maker:
                break

    return InsuranceSheet(title=title, rows=rows, maker=maker, warnings=warnings)
