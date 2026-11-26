"""
Form builder for Feishu approval instances.

Constructs the widget list JSON required by the Feishu approval
v4 instances API, mapping parsed spreadsheet data to the verified
widget IDs of the meal refund approval form.
"""

from __future__ import annotations

from app.excel_parser import (
    BonusSheet,
    InsuranceSheet,
    MealRefundSheet,
    MealTwoAmountSheet,
    SealDetailSheet,
)

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
    # 仅当原始表有“个人其他调整”列，且至少有一行非零时，才包含该列。
    # 列不存在或全部为 0 时，不向飞书表单传该列数据。
    include_adjustment = sheet.has_adjustment_column and any(
        r.adjustment != "0" for r in sheet.rows
    )

    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_row: list[dict] = [
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
        ]
        if include_adjustment:
            detail_row.append(
                {
                    "id": WIDGET_BONUS_ADJUSTMENT,
                    "type": "amount",
                    "value": row.adjustment,
                }
            )
        detail_row.append(
            {
                "id": WIDGET_BONUS_NET,
                "type": "amount",
                "value": row.net_amount,
            }
        )
        detail_rows.append(detail_row)

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


# ── 彩虹公司五险一金审批（2026-11 表单改版）──────────────────────────
# 明细行结构：单位名称 + 参保人数 + 小计金额 + 明细1~16 / 金额1~16。
# 财务要求现阶段每行只把 Excel 最后一列的行合计（缴费合计 / 转款-合计）
# 填入 小计金额，明细1~16 留空，为后续升级预留。
#
# 控件 ID 与分公司表高度重合：单位名称/参保人数 完全相同；小计金额
# 沿用分公司 费用金额1 的 ID；明细1~16 与分公司表 费用名称/金额
# 第 2~17 对同 ID。改版初期表单里曾插入 说明1 静态文字控件
# （widget17951450334860001，值「以下为费用明细」），后已从表单删除，
# payload 随之不再携带。
CAIHONG_INS_SUBTOTAL = "widget17822958110070001"
# 预留：明细1~16/金额1~16，后续升级直接按此列表填写。
CAIHONG_DETAIL_WIDGET_PAIRS = INSURANCE_FEE_WIDGET_PAIRS[1:]


def build_caihong_insurance_form_widgets(
    sheet: InsuranceSheet,
    file_code: str,
    title: str | None = None,
    remark: str = "",
) -> list[dict]:
    """Build the widget list for the 彩虹公司五险一金 approval instance.

    The form was redesigned (2026-11): each detail row now holds
    单位名称 + 参保人数 + 小计金额 + 明细1~16/金额1~16 pairs.  Per
    finance's current requirement only the row total (the Excel's
    last money column, e.g. 缴费合计 / 转款-合计) is submitted, into
    小计金额; the 16 detail pairs stay empty and are reserved for a
    future upgrade (see ``CAIHONG_DETAIL_WIDGET_PAIRS``).

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
        if not row.fees:
            raise ValueError(
                f"{row.unit_name or '未命名单位'}：行内没有费用列，无法填写小计金额"
            )
        # 行合计 = 该行最后一列费用金额（缴费合计 / 转款-合计）。
        subtotal = row.fees[-1].amount

        detail_rows.append(
            [
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
                {
                    "id": CAIHONG_INS_SUBTOTAL,
                    "type": "amount",
                    "value": subtotal,
                },
            ]
        )

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


# ── 派遣员工简易费用明细表审批-两列金额 widget IDs ────────────────────
# （2026-11-25 从线上审批定义 63FBFB03-B321-4402-B85E-8404A18E6C2D
#   实时抓取核实；与在用的单金额流程完全独立，互不影响。）
MEAL2_TITLE = "widget17811628753290001"
MEAL2_COLUMN_MAPPING = "widget17955716825010001"      # 列名称映射 fieldList（必填）
MEAL2_MAPPING_FORM_COLUMN = "widget17955717008060001"  # 子控件：审批单中列名称
MEAL2_MAPPING_EXCEL_COLUMN = "widget17955717607840001"  # 子控件：excel表中的列名称
MEAL2_DETAIL_LIST = "widget16496685286700001"
MEAL2_ROW_NAME = "widget17920477234030001"
MEAL2_ROW_ID_CARD = "widget17920477572400001"
MEAL2_ROW_AMOUNT1 = "widget17822958110070001"          # 金额1 (amount CNY)
MEAL2_ROW_AMOUNT2 = "widget17955713713480001"          # 金额2 (amount CNY)
MEAL2_ATTACHMENT = "widget16496686388470001"
MEAL2_REMARK = "widget17811635300840001"


def build_meal2_form_widgets(
    sheet: MealTwoAmountSheet,
    file_code: str,
    title: str | None = None,
    remark: str = "",
) -> list[dict]:
    """Build the widget list for the two-amount-column meal approval.

    The approval definition contains an extra required top-level
    「列名称映射」 fieldList that maps the form's 金额1/金额2 columns to
    the original Excel column headers; its rows are filled from
    ``sheet.amount1_column`` / ``sheet.amount2_column``. All form
    children are optional (required: false), so when the source sheet
    lacks the id-card column or the second amount column, the
    corresponding child widgets and mapping row are omitted entirely.

    Detail rows carry 姓名/身份证号码/金额1/金额2 in the fieldList
    children order (missing columns omitted) — the form has no 序号
    widget.

    Args:
        sheet: Parsed two-amount-column spreadsheet data.
        file_code: File code returned by the approval file upload API.
        title: Optional override for the approval title.
        remark: Optional remark text for the textarea widget.

    Returns:
        List of widget dicts ready for json.dumps(): title (input),
        column mapping (fieldList), detail list (fieldList),
        attachment (attachmentV2), remark (textarea).
    """
    # 列名称映射：金额2 行仅在源表确有第二金额列时携带。
    mapping_rows: list[list[dict]] = [
        [
            {
                "id": MEAL2_MAPPING_FORM_COLUMN,
                "type": "input",
                "value": "金额1",
            },
            {
                "id": MEAL2_MAPPING_EXCEL_COLUMN,
                "type": "input",
                "value": sheet.amount1_column,
            },
        ]
    ]
    if sheet.amount2_column:
        mapping_rows.append(
            [
                {
                    "id": MEAL2_MAPPING_FORM_COLUMN,
                    "type": "input",
                    "value": "金额2",
                },
                {
                    "id": MEAL2_MAPPING_EXCEL_COLUMN,
                    "type": "input",
                    "value": sheet.amount2_column,
                },
            ]
        )

    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_row: list[dict] = [
            {
                "id": MEAL2_ROW_NAME,
                "type": "input",
                "value": row.name,
            }
        ]
        # 身份证号 / 金额2 控件在源表缺列时整列省略（均为非必填控件）。
        if sheet.has_id_column:
            detail_row.append(
                {
                    "id": MEAL2_ROW_ID_CARD,
                    "type": "input",
                    "value": row.id_card,
                }
            )
        detail_row.append(
            {
                "id": MEAL2_ROW_AMOUNT1,
                "type": "amount",
                "value": row.amount1,
            }
        )
        if sheet.amount2_column:
            detail_row.append(
                {
                    "id": MEAL2_ROW_AMOUNT2,
                    "type": "amount",
                    "value": row.amount2,
                }
            )
        detail_rows.append(detail_row)

    return [
        {
            "id": MEAL2_TITLE,
            "type": "input",
            "value": title if title is not None else sheet.title,
        },
        {
            "id": MEAL2_COLUMN_MAPPING,
            "type": "fieldList",
            "value": mapping_rows,
        },
        {
            "id": MEAL2_DETAIL_LIST,
            "type": "fieldList",
            "value": detail_rows,
        },
        {
            "id": MEAL2_ATTACHMENT,
            "type": "attachmentV2",
            "value": [file_code],
        },
        {
            "id": MEAL2_REMARK,
            "type": "textarea",
            "value": remark,
        },
    ]


# ═══════════════════════════════════════════════════════════════════════
# 用章/用印批量申请（3 个流程，2026-11-25 从线上审批定义实时抓取核实）
#
#   - 补签派遣员工合同用章申请   C17A8B67-4C63-4FD8-BACE-D398BEE3C2FE
#   - 解除（终止）劳动合同证明书用印申请 6B6D7EFF-C1A7-4021-B7D5-2963C01D509C
#   - 新签/续签派遣员工合同批量用章用印申请 FFF8A3B3-7917-44F1-8B06-3F4B0183A4F8
#
# 三表由同一模板复制而来，控件 ID 高度重合（明细列表/姓名/份数/单选/
# 事由/备注/附件完全同 ID）；差异控件：
#   - 补签：行内另有 合同起止 dateInterval + 补签原因 textarea；
#   - 新签/续签：行内另有 身份证复印件盖公章份数 input。
# 表单没有标题控件，实例名固定为审批定义名；顶层为
# 用章事由（widget0，必填）→ 明细列表 → 备注 → 附件。
# 2026-11-26 线上核实：用章事由控件显示名为 新签/续签「用章事由-仅
# 新签/续签可用」、补签「用章事由-仅补签可用」、解除「用章事由」；
# 提交按控件 ID 填值，不依赖显示名。
# 2026-11-26 线上核实（改名）：行级单选控件显示名统一为「用印主体」
# （原 补签/解除「印章类型」、新签/续签「签约主体」），控件 ID 未变；
# 选项文本 = 主体名+后缀：补签/新签续签「合同章」、解除「公章」，
# 选项 value 未变；解除流程另设默认值 彩虹公司公章（提交仍显式传值）。
# radioV2 的 value 必须传选项 value（非选项文本），映射见
# SEAL_ENTITY_OPTION_VALUES；dateInterval 的 value 为
# {start, end(RFC3339), interval(天)} 对象。
# ═══════════════════════════════════════════════════════════════════════

import re as _re

SEAL_REASON = "widget0"                            # 用章事由 textarea（必填）
SEAL_DETAIL_LIST = "widget17954194453400001"       # 明细 fieldList（3 表共用）
SEAL_ROW_NAME = "widget17954194530910001"          # 派遣员工姓名 input
SEAL_ROW_COPIES = "widget17954194884670001"        # 合同份数/证明份数 number
SEAL_ROW_OPTION = "widget17953432091600001"        # 用印主体 radioV2
SEAL_ROW_DATE_INTERVAL = "widget17955967811230001"  # 补签：合同起止 dateInterval
SEAL_ROW_REASON = "widget17955969605740001"        # 补签：补签原因 textarea
SEAL_ROW_ID_COPIES = "widget17954195767570001"     # 新签：身份证复印件盖公章份数 input
SEAL_REMARK = "widget17953417689960001"            # 备注 textarea（选填）
SEAL_ATTACHMENT = "widget15828104903330001"        # 附件 attachmentV2（选填）

# 行级单选控件显示名（3 表一致；2026-11-26 线上核实由 补签/解除的
# 「印章类型」与新签/续签的「签约主体」统一改名为「用印主体」）。
SEAL_OPTION_LABEL = "用印主体"

# 主体 → radioV2 选项 value（3 表的单选选项取值完全一致）。
# 选项文本 = 主体名 + 后缀（见 SEAL_KIND_OPTION_SUFFIXES）；value 只与
# 主体对应，与后缀无关（2026-11-26 线上核实）。
SEAL_ENTITY_OPTION_VALUES: dict[str, str] = {
    "彩虹公司": "mwrozq2w-dl9xxzxtd6-0",
    "长白山分公司": "mwrozq2w-5utstr4jpbi-0",
    "吉林分公司": "mwrozq2w-4sabhobhlo6-0",
    "白城分公司": "mwrozq41-zk8jw3vhlvl-1",
    "公主岭分公司": "mwrozq41-l2wrznbz6m-3",
}

# 各流程用印主体选项文本的后缀（kind → 后缀）。
SEAL_KIND_OPTION_SUFFIXES: dict[str, str] = {
    "seal_reseal": "合同章",
    "seal_terminate_cert": "公章",
    "seal_contract_batch": "合同章",
}


def seal_option_labels(suffix: str) -> list[str]:
    """按后缀返回用印主体的合法选项文本列表（错误提示用）。"""
    return [f"{entity}{suffix}" for entity in SEAL_ENTITY_OPTION_VALUES]


def resolve_seal_option_value(text: str, suffix: str = "公章") -> str:
    """把 Excel 中用印主体文本解析为 radioV2 选项 value。

    三个流程的选项仅后缀不同（公章/合同章），value 与主体一一对应，
    故解析只看主体名，后缀写法兼容（可省略如「彩虹公司」，写错后缀
    如解除流程里写「合同章」也可）。无法识别时抛 ValueError 并列出
    该流程的全部合法选项文本。
    """
    cleaned = _re.sub(r"\s+", "", str(text))
    core = cleaned
    for known in ("合同章", "公章"):
        if core.endswith(known):
            core = core[: -len(known)]
            break
    if core in SEAL_ENTITY_OPTION_VALUES:
        return SEAL_ENTITY_OPTION_VALUES[core]
    for entity, value in SEAL_ENTITY_OPTION_VALUES.items():
        if core and (cleaned.startswith(entity) or entity.startswith(core)):
            return value
    allowed = "、".join(seal_option_labels(suffix))
    raise ValueError(f"{SEAL_OPTION_LABEL}无法识别：{text}（可选：{allowed}）")


def build_seal_reseal_form_widgets(
    sheet: SealDetailSheet,
    file_code: str,
    reason: str | None = None,
    remark: str | None = None,
) -> list[dict]:
    """Build widgets for 补签派遣员工合同用章申请.

    Detail rows carry exactly 5 children in definition order:
    派遣员工姓名 (input) → 合同起止 (dateInterval) → 合同份数 (number)
    → 用印主体 (radioV2) → 补签原因 (textarea).  Top level:
    用章事由 (textarea, 必填) → 明细 → 备注 → 附件.

    Args:
        sheet: Parsed reseal sheet (``SealDetailSheet``).
        file_code: File code from the approval file upload API.
        reason: 用章事由 override; defaults to ``sheet.reason``.
        remark: 备注 override; defaults to ``sheet.remark``.

    Returns:
        List of 4 widget dicts: reason, detail fieldList, remark,
        attachment.  The form has no title widget.
    """
    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_rows.append(
            [
                {"id": SEAL_ROW_NAME, "type": "input", "value": row.name},
                {
                    "id": SEAL_ROW_DATE_INTERVAL,
                    "type": "dateInterval",
                    "value": {
                        "start": row.start_date,
                        "end": row.end_date,
                        "interval": float(row.interval_days),
                    },
                },
                {"id": SEAL_ROW_COPIES, "type": "number", "value": row.copies},
                {"id": SEAL_ROW_OPTION, "type": "radioV2", "value": row.option_value},
                {"id": SEAL_ROW_REASON, "type": "textarea", "value": row.reason},
            ]
        )

    return [
        {
            "id": SEAL_REASON,
            "type": "textarea",
            "value": reason if reason is not None else sheet.reason,
        },
        {"id": SEAL_DETAIL_LIST, "type": "fieldList", "value": detail_rows},
        {
            "id": SEAL_REMARK,
            "type": "textarea",
            "value": remark if remark is not None else sheet.remark,
        },
        {"id": SEAL_ATTACHMENT, "type": "attachmentV2", "value": [file_code]},
    ]


def build_seal_terminate_cert_form_widgets(
    sheet: SealDetailSheet,
    file_code: str,
    reason: str | None = None,
    remark: str | None = None,
) -> list[dict]:
    """Build widgets for 解除（终止）劳动合同证明书用印申请.

    Detail rows carry exactly 3 children in definition order:
    派遣员工姓名 (input) → 证明份数 (number) → 用印主体 (radioV2).

    Args:
        sheet: Parsed terminate-certificate sheet.
        file_code: File code from the approval file upload API.
        reason: 用章事由 override; defaults to ``sheet.reason``.
        remark: 备注 override; defaults to ``sheet.remark``.

    Returns:
        List of 4 widget dicts: reason, detail fieldList, remark,
        attachment.  The form has no title widget.
    """
    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_rows.append(
            [
                {"id": SEAL_ROW_NAME, "type": "input", "value": row.name},
                {"id": SEAL_ROW_COPIES, "type": "number", "value": row.copies},
                {"id": SEAL_ROW_OPTION, "type": "radioV2", "value": row.option_value},
            ]
        )

    return [
        {
            "id": SEAL_REASON,
            "type": "textarea",
            "value": reason if reason is not None else sheet.reason,
        },
        {"id": SEAL_DETAIL_LIST, "type": "fieldList", "value": detail_rows},
        {
            "id": SEAL_REMARK,
            "type": "textarea",
            "value": remark if remark is not None else sheet.remark,
        },
        {"id": SEAL_ATTACHMENT, "type": "attachmentV2", "value": [file_code]},
    ]


def build_seal_contract_batch_form_widgets(
    sheet: SealDetailSheet,
    file_code: str,
    reason: str | None = None,
    remark: str | None = None,
) -> list[dict]:
    """Build widgets for 新签/续签派遣员工合同批量用章用印申请.

    Detail rows carry exactly 4 children in definition order:
    派遣员工姓名 (input) → 合同份数 (number) → 用印主体 (radioV2)
    → 身份证复印件盖公章份数 (input; the form defines it as a text
    input, so the raw Excel text is passed through).

    Args:
        sheet: Parsed contract-batch sheet.
        file_code: File code from the approval file upload API.
        reason: 用章事由 override; defaults to ``sheet.reason``.
        remark: 备注 override; defaults to ``sheet.remark``.

    Returns:
        List of 4 widget dicts: reason, detail fieldList, remark,
        attachment.  The form has no title widget.
    """
    detail_rows: list[list[dict]] = []
    for row in sheet.rows:
        detail_rows.append(
            [
                {"id": SEAL_ROW_NAME, "type": "input", "value": row.name},
                {"id": SEAL_ROW_COPIES, "type": "number", "value": row.copies},
                {"id": SEAL_ROW_OPTION, "type": "radioV2", "value": row.option_value},
                {
                    "id": SEAL_ROW_ID_COPIES,
                    "type": "input",
                    "value": row.id_copy_copies,
                },
            ]
        )

    return [
        {
            "id": SEAL_REASON,
            "type": "textarea",
            "value": reason if reason is not None else sheet.reason,
        },
        {"id": SEAL_DETAIL_LIST, "type": "fieldList", "value": detail_rows},
        {
            "id": SEAL_REMARK,
            "type": "textarea",
            "value": remark if remark is not None else sheet.remark,
        },
        {"id": SEAL_ATTACHMENT, "type": "attachmentV2", "value": [file_code]},
    ]
