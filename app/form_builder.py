"""
Form builder for Feishu approval instances.

Constructs the widget list JSON required by the Feishu approval
v4 instances API, mapping parsed spreadsheet data to the verified
widget IDs of the meal refund approval form.
"""

from __future__ import annotations

from app.excel_parser import BonusSheet, InsuranceSheet, MealRefundSheet

# ── Verified widget IDs (from live approval definition) ──────────────
WIDGET_TITLE = "widget17811628753290001"
WIDGET_DETAIL_LIST = "widget16496685286700001"
WIDGET_ROW_SERIAL = "widget17811629865160001"
WIDGET_ROW_NAME = "widget17920477234030001"
WIDGET_ROW_ID_CARD = "widget17920477572400001"
WIDGET_ROW_AMOUNT = "widget17822958110070001"
WIDGET_ATTACHMENT = "widget16496686388470001"
WIDGET_REMARK = "widget17811635300840001"

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


# ── 五险一金审批明细 widget IDs（只读，不做任何提交动作） ──────────────
INS_UNITS_NAME = "widget17811629865160001"
INS_HEADCOUNT = "widget17930639143030001"
INS_FEE_NAME_1 = "widget17930638724660001"
INS_FEE_AMOUNT_1 = "widget17822958110070001"
INS_FEE_NAME_2 = "widget17930641142080001"
INS_FEE_AMOUNT_2 = "widget17822960394550001"
INS_FEE_NAME_3 = "widget17930641615980001"
INS_FEE_AMOUNT_3 = "widget16496685800140001"
INS_FEE_NAME_4 = "widget17930641998140001"
INS_FEE_AMOUNT_4 = "widget17851332232590001"
INS_FEE_NAME_5 = "widget17930642276560001"
INS_FEE_AMOUNT_5 = "widget17851332541430001"
INS_FEE_NAME_6 = "widget17930642815070001"
INS_FEE_AMOUNT_6 = "widget17822963272840001"
INS_FEE_NAME_7 = "widget17930643094440001"
INS_FEE_AMOUNT_7 = "widget17851332961860001"
INS_FEE_NAME_8 = "widget17930643375920001"
INS_FEE_AMOUNT_8 = "widget17851333370110001"
INS_FEE_NAME_9 = "widget17930643790500001"
INS_FEE_AMOUNT_9 = "widget17851333724920001"
INS_FEE_NAME_10 = "widget17930702589200001"
INS_FEE_AMOUNT_10 = "widget17811632175290001"
INS_FEE_NAME_11 = "widget17930702891640001"
INS_FEE_AMOUNT_11 = "widget17811632599910001"
INS_FEE_NAME_12 = "widget17930703239880001"
INS_FEE_AMOUNT_12 = "widget17811632838960001"
INS_FEE_NAME_13 = "widget17930703551120001"
INS_FEE_AMOUNT_13 = "widget17816837666160001"
INS_FEE_NAME_14 = "widget17930703888180001"
INS_FEE_AMOUNT_14 = "widget17811633075360001"
INS_FEE_NAME_15 = "widget17930704242670001"
INS_FEE_AMOUNT_15 = "widget17811633663510001"
INS_FEE_NAME_16 = "widget17930704609180001"
INS_FEE_AMOUNT_16 = "widget17811633849100001"
INS_FEE_NAME_17 = "widget17930704986480001"
INS_FEE_AMOUNT_17 = "widget17811634196160001"

INSURANCE_FEE_WIDGET_PAIRS: list[tuple[str, str]] = [
    (INS_FEE_NAME_1, INS_FEE_AMOUNT_1),
    (INS_FEE_NAME_2, INS_FEE_AMOUNT_2),
    (INS_FEE_NAME_3, INS_FEE_AMOUNT_3),
    (INS_FEE_NAME_4, INS_FEE_AMOUNT_4),
    (INS_FEE_NAME_5, INS_FEE_AMOUNT_5),
    (INS_FEE_NAME_6, INS_FEE_AMOUNT_6),
    (INS_FEE_NAME_7, INS_FEE_AMOUNT_7),
    (INS_FEE_NAME_8, INS_FEE_AMOUNT_8),
    (INS_FEE_NAME_9, INS_FEE_AMOUNT_9),
    (INS_FEE_NAME_10, INS_FEE_AMOUNT_10),
    (INS_FEE_NAME_11, INS_FEE_AMOUNT_11),
    (INS_FEE_NAME_12, INS_FEE_AMOUNT_12),
    (INS_FEE_NAME_13, INS_FEE_AMOUNT_13),
    (INS_FEE_NAME_14, INS_FEE_AMOUNT_14),
    (INS_FEE_NAME_15, INS_FEE_AMOUNT_15),
    (INS_FEE_NAME_16, INS_FEE_AMOUNT_16),
    (INS_FEE_NAME_17, INS_FEE_AMOUNT_17),
]


def build_insurance_form_widgets(
    sheet: InsuranceSheet,
    file_code: str,
    title: str | None = None,
    remark: str = "",
) -> list[dict]:
    """Build the widget list for a 五险一金 approval instance.

    The approval definition code is read from the environment
    (`FEISHU_INSURANCE_APPROVAL_CODE`) by the UI layer; it is not
    hard-coded here. Each unit row is mapped to:

    - 单位名称 (input)
    - 参保人数 (number, JSON int)
    - up to 17 pairs of 费用名称 (input) + 费用金额 (amount)

    Args:
        sheet: Parsed 五险一金 sheet.
        file_code: File code returned by the approval file upload API.
        title: Optional approval title override.
        remark: Optional remark text for the textarea widget.

    Returns:
        List of widget dicts ready for json.dumps(). Contains exactly
        four widgets: title, detail list, attachment, and remark.
    """
    detail_rows: list[list[dict]] = []

    for row in sheet.rows:
        if len(row.fees) > len(INSURANCE_FEE_WIDGET_PAIRS):
            raise ValueError("费用项超过17项，无法提交")

        detail_row: list[dict] = [
            {
                "id": INS_UNITS_NAME,
                "type": "input",
                "value": row.unit_name,
            },
            {
                "id": INS_HEADCOUNT,
                "type": "number",
                "value": int(row.headcount),
            },
        ]

        for fee, (name_widget_id, amount_widget_id) in zip(
            row.fees, INSURANCE_FEE_WIDGET_PAIRS, strict=False
        ):
            detail_row.extend(
                [
                    {
                        "id": name_widget_id,
                        "type": "input",
                        "value": fee.name,
                    },
                    {
                        "id": amount_widget_id,
                        "type": "amount",
                        "value": fee.amount,
                    },
                ]
            )

        detail_rows.append(detail_row)

    return [
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
