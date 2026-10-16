"""
Form builder for Feishu approval instances.

Constructs the widget list JSON required by the Feishu approval
v4 instances API, mapping parsed spreadsheet data to the verified
widget IDs of the meal refund approval form.
"""

from __future__ import annotations

from app.excel_parser import BonusSheet, MealRefundSheet

# ── Verified widget IDs (from live approval definition) ──────────────
WIDGET_TITLE = "widget17811628753290001"
WIDGET_DETAIL_LIST = "widget16496685286700001"
WIDGET_ROW_SERIAL = "widget17811629865160001"
WIDGET_ROW_NAME = "widget17920477234030001"
WIDGET_ROW_ID_CARD = "widget17920477572400001"
WIDGET_ROW_AMOUNT = "widget17822958110070001"
WIDGET_ATTACHMENT = "widget16496686388470001"
WIDGET_REMARK = "widget17811635300840001"

# ── Approval code ────────────────────────────────────────────────────
APPROVAL_CODE = "081B5624-B8B7-4334-AF42-D1075F8E650A"


def build_form_widgets(
    sheet: MealRefundSheet,
    file_code: str,
    title: str | None = None,
    remark: str = "",
) -> list[dict]:
    """Build the widget list for a Feishu approval instance.

    Constructs the exact JSON structure expected by the Feishu
    approval/v4/instances API, mapping spreadsheet data to the
    verified widget IDs.

    Args:
        sheet: Parsed meal refund spreadsheet data.
        file_code: File code returned by the approval file upload API.
        title: Optional override for the approval title.  Defaults
            to the sheet's title.
        remark: Optional remark text for the textarea widget.

    Returns:
        List of widget dicts ready for json.dumps().  The list
        contains exactly four widgets: title (input), detail list
        (fieldList), attachment (attachmentV2), and remark (textarea).
    """
    # Build detail list rows
    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_rows.append(
            [
                {
                    "id": WIDGET_ROW_SERIAL,
                    "type": "input",
                    "value": row.serial,
                },
                {
                    "id": WIDGET_ROW_NAME,
                    "type": "input",
                    "value": row.name,
                },
                {
                    "id": WIDGET_ROW_ID_CARD,
                    "type": "input",
                    "value": row.id_card,
                },
                {
                    "id": WIDGET_ROW_AMOUNT,
                    "type": "amount",
                    "value": row.amount,
                },
            ]
        )

    widgets: list[dict] = [
        {
            "id": WIDGET_TITLE,
            "type": "input",
            "value": title if title is not None else sheet.title,
        },
        {
            "id": WIDGET_DETAIL_LIST,
            "type": "fieldList",
            "value": detail_rows,
        },
        {
            "id": WIDGET_ATTACHMENT,
            "type": "attachmentV2",
            "value": [file_code],
        },
        {
            "id": WIDGET_REMARK,
            "type": "textarea",
            "value": remark,
        },
    ]

    return widgets


# ── Bonus approval widget IDs (from live approval definition) ────────
BONUS_APPROVAL_CODE = "01BA9019-4E75-4726-8F95-38C43282F543"
WIDGET_BONUS_TITLE = "widget17811628753290001"
WIDGET_BONUS_DETAIL_LIST = "widget16496685286700001"
WIDGET_BONUS_EMPLOYEE_ID = "widget17921134502450001"
WIDGET_BONUS_NAME = "widget17920477234030001"
WIDGET_BONUS_ID_CARD = "widget17920477572400001"
WIDGET_BONUS_AMOUNT = "widget17921134646590001"
WIDGET_BONUS_TAX = "widget17921134788590001"
WIDGET_BONUS_ADJUSTMENT = "widget17921148886730001"  # 个人其他调整 (amount)
WIDGET_BONUS_NET = "widget17822958110070001"
WIDGET_BONUS_ATTACHMENT = "widget16496686388470001"
WIDGET_BONUS_REMARK = "widget17811635300840001"


def build_bonus_form_widgets(
    sheet: BonusSheet,
    file_code: str,
    title: str | None = None,
    remark: str = "",
) -> list[dict]:
    """Build the widget list for a bonus Feishu approval instance.

    The serial (序号) field is intentionally excluded from the detail
    list — the Feishu bonus approval form has no serial widget.

    Each detail row contains exactly 7 widgets in the Feishu fieldList
    children order: 工号 (input), 姓名 (input), 身份证号码 (input),
    奖金 (amount), 个人所得税 (amount), 个人其他调整 (amount),
    实发金额 (amount).  The 个人其他调整 value defaults to "0" when
    the uploaded Excel lacks that column.

    Args:
        sheet: Parsed bonus spreadsheet data.
        file_code: File code returned by the approval file upload API.
        title: Optional override for the approval title.
        remark: Optional remark text for the textarea widget.

    Returns:
        List of widget dicts ready for json.dumps().  Contains exactly
        four widgets: title (input), detail list (fieldList),
        attachment (attachmentV2), and remark (textarea).
    """
    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_rows.append(
            [
                {
                    "id": WIDGET_BONUS_EMPLOYEE_ID,
                    "type": "input",
                    "value": row.employee_id,
                },
                {
                    "id": WIDGET_BONUS_NAME,
                    "type": "input",
                    "value": row.name,
                },
                {
                    "id": WIDGET_BONUS_ID_CARD,
                    "type": "input",
                    "value": row.id_card,
                },
                {
                    "id": WIDGET_BONUS_AMOUNT,
                    "type": "amount",
                    "value": row.bonus,
                },
                {
                    "id": WIDGET_BONUS_TAX,
                    "type": "amount",
                    "value": row.tax,
                },
                {
                    "id": WIDGET_BONUS_ADJUSTMENT,
                    "type": "amount",
                    "value": row.adjustment,
                },
                {
                    "id": WIDGET_BONUS_NET,
                    "type": "amount",
                    "value": row.net_amount,
                },
            ]
        )

    return [
        {
            "id": WIDGET_BONUS_TITLE,
            "type": "input",
            "value": title if title is not None else sheet.title,
        },
        {
            "id": WIDGET_BONUS_DETAIL_LIST,
            "type": "fieldList",
            "value": detail_rows,
        },
        {
            "id": WIDGET_BONUS_ATTACHMENT,
            "type": "attachmentV2",
            "value": [file_code],
        },
        {
            "id": WIDGET_BONUS_REMARK,
            "type": "textarea",
            "value": remark,
        },
    ]
