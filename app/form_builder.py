"""
Form builder for Feishu approval instances.

Constructs the widget list JSON required by the Feishu approval
v4 instances API, mapping parsed spreadsheet data to the verified
widget IDs of the meal refund approval form.
"""

from __future__ import annotations

from app.excel_parser import MealRefundSheet

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
