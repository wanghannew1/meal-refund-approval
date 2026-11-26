"""
Streamlit application for meal refund approval submission to Feishu.

Provides a web UI for uploading Excel files, parsing meal refund data,
looking up submitters by phone number, and submitting approval instances.
"""

import os
import sys
import tempfile
import traceback
import uuid
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

# Add project root to path for imports
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

# Load .env from project root
_env_path = _PROJECT_ROOT / ".env"
if _env_path.exists():
    load_dotenv(_env_path)

# Import core modules (implemented by parallel agent)
from app.excel_parser import (
    parse_meal_refund_excel,
    parse_meal_two_amount_excel,
    parse_bonus_excel,
    parse_insurance_excel,
    parse_seal_reseal_excel,
    parse_seal_terminate_cert_excel,
    parse_seal_contract_batch_excel,
    list_sheets,
    list_insurance_sheets,
    MealRefundSheet,
    BonusSheet,
    InsuranceSheet,
    SealDetailSheet,
)
from app.form_builder import (
    build_form_widgets,
    build_meal2_form_widgets,
    build_bonus_form_widgets,
    build_insurance_form_widgets,
    build_caihong_insurance_form_widgets,
    build_seal_reseal_form_widgets,
    build_seal_terminate_cert_form_widgets,
    build_seal_contract_batch_form_widgets,
    seal_option_labels,
)
from app.feishu_api import (
    get_tenant_token,
    lookup_user_by_mobile,
    get_user_name,
    upload_approval_file,
    get_approval_info,
    preview_instance,
    create_instance,
    get_instance_detail,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="审批提交工具",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state initialization
# ---------------------------------------------------------------------------
if "sheet" not in st.session_state:
    st.session_state.sheet = None
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "open_id" not in st.session_state:
    st.session_state.open_id = ""
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "title" not in st.session_state:
    st.session_state.title = ""
if "remark" not in st.session_state:
    st.session_state.remark = ""
if "last_instance" not in st.session_state:
    st.session_state.last_instance = None
if "submitted" not in st.session_state:
    st.session_state.submitted = False
if "tmp_file_path" not in st.session_state:
    st.session_state.tmp_file_path = None
if "approval_kind" not in st.session_state:
    # 默认新签/续签用章（审批类型选择器首项），与 visible_kind_options 首项一致。
    st.session_state.approval_kind = "seal_contract_batch"
if "approval_entity" not in st.session_state:
    # 初值写 DEFAULT_INSURANCE_ENTITY 对应的字面量（此处位于常量定义之前）。
    # 若该主体被隐藏（彩虹公司暂时停用），main() 会归位到第一个可见主体。
    st.session_state.approval_entity = "caihong"
if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = None
if "uploaded_file_sig" not in st.session_state:
    st.session_state.uploaded_file_sig = None
if "name_lookup_failed" not in st.session_state:
    st.session_state.name_lookup_failed = False
if "batch_items" not in st.session_state:
    st.session_state.batch_items = []
if "batch_submitting" not in st.session_state:
    st.session_state.batch_submitting = False
if "sheet_infos" not in st.session_state:
    st.session_state.sheet_infos = None
if "selected_sheet_name" not in st.session_state:
    st.session_state.selected_sheet_name = None

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _require_env_code(name: str, label: str) -> str:
    code = os.getenv(name, "").strip()
    if not code:
        raise RuntimeError(f"缺少环境变量 {name}，请在 .env 中配置{label}")
    return code


# 五险一金按公司主体路由到不同审批定义。Code 全部来自环境变量，
# 代码中不硬编码审批定义 UUID；新增主体只需在此登记一行。
# 各主体表单明细行控件结构不同，用 form_builder 指向各自的构建函数：
# 分公司表每行为 费用名称/金额 17 对；彩虹表 2026-11 改版后每行只有
# 小计金额（行合计）+ 明细1~16（预留未填）。
# 注册表顺序即「公司主体」选择器的选项顺序：彩虹公司放左侧。
# 彩虹公司五险一金暂时停用（2026-11-25）：解析/表单/提交代码全部保留，
# 仅隐藏「公司主体」选择器入口。恢复上线时改回 True。
CAIHONG_INSURANCE_ENABLED = False

INSURANCE_ENTITIES: dict[str, dict] = {
    "caihong": {
        "label": "彩虹公司",
        "env": "FEISHU_INSURANCE_CAIHONG_APPROVAL_CODE",
        "legacy_env": "",
        "form_builder": build_caihong_insurance_form_widgets,
        "enabled": CAIHONG_INSURANCE_ENABLED,
    },
    "branch": {
        "label": "分公司",
        "env": "FEISHU_INSURANCE_BRANCH_APPROVAL_CODE",
        # 旧部署只配了统一名称，继续兼容分公司主体
        "legacy_env": "FEISHU_INSURANCE_APPROVAL_CODE",
        "form_builder": build_insurance_form_widgets,
        "enabled": True,
    },
}

DEFAULT_INSURANCE_ENTITY = "caihong"


def visible_insurance_entities() -> dict[str, dict]:
    """Registry entries exposed in the 公司主体 selector (enabled only)."""
    return {k: m for k, m in INSURANCE_ENTITIES.items() if m.get("enabled", True)}


def get_default_insurance_entity() -> str:
    """First visible entity, so a hidden default never routes submissions."""
    if DEFAULT_INSURANCE_ENTITY in visible_insurance_entities():
        return DEFAULT_INSURANCE_ENTITY
    return next(iter(visible_insurance_entities()))


# 五险一金审批入口暂时整体下线（2026-11-25，含分公司与彩虹公司）：
# 上传/解析/表单构建/提交代码全部保留，仅从「审批类型」选择器隐藏。
# 恢复上线改回 True 即可；各主体的单独开关见 CAIHONG_INSURANCE_ENABLED。
INSURANCE_KIND_ENABLED = False

# 简易费用（单金额/两列金额）与奖金入口暂时整体下线（2026-11-25）：
# 三个工具的解析/表单构建/提交代码全部保留，仅从「审批类型」选择器
# 隐藏，恢复上线改回 True 即可。
MEAL_KIND_ENABLED = False
MEAL2_KIND_ENABLED = False
BONUS_KIND_ENABLED = False

# 用章/用印批量申请流程（2026-11-25 上线）。三个审批定义由同一模板
# 复制而来，控件 ID 高度重合；Code 全部来自环境变量，代码不硬编码。
# 各流程仅限对应类型文件使用（审批定义名称含「其他文件勿用」）。
# 注册表顺序即「审批类型」选择器与模板下载的展示顺序：
# 新签/续签 → 解除证明 → 补签。
SEAL_KINDS: dict[str, dict] = {
    "seal_contract_batch": {
        "label": "新签/续签派遣员工合同批量用章用印申请",
        "env": "FEISHU_SEAL_CONTRACT_BATCH_APPROVAL_CODE",
        "parser": parse_seal_contract_batch_excel,
        "form_builder": build_seal_contract_batch_form_widgets,
        "upload_label": "选择新签/续签派遣员工合同批量用章明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）",
        "option_label": "签约主体",
    },
    "seal_terminate_cert": {
        "label": "解除（终止）劳动合同证明书用印申请",
        "env": "FEISHU_SEAL_TERMINATE_CERT_APPROVAL_CODE",
        "parser": parse_seal_terminate_cert_excel,
        "form_builder": build_seal_terminate_cert_form_widgets,
        "upload_label": "选择解除（终止）劳动合同证明书用印明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）",
        "option_label": "印章类型",
    },
    "seal_reseal": {
        "label": "补签派遣员工合同用章申请",
        "env": "FEISHU_SEAL_RESEAL_APPROVAL_CODE",
        "parser": parse_seal_reseal_excel,
        "form_builder": build_seal_reseal_form_widgets,
        "upload_label": "选择补签派遣员工合同用章明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）",
        "option_label": "印章类型",
    },
}

# 各审批类型的上下架开关（False 时代码保留、仅从界面隐藏）。
KIND_ENABLED_FLAGS: dict[str, bool] = {
    "seal_contract_batch": True,
    "seal_terminate_cert": True,
    "seal_reseal": True,
    "meal2": MEAL2_KIND_ENABLED,
    "meal": MEAL_KIND_ENABLED,
    "bonus": BONUS_KIND_ENABLED,
    "insurance": INSURANCE_KIND_ENABLED,
}

# 模板文件 → 所属审批类型；类型隐藏时模板随之从界面隐藏（文件保留）。
# 字典顺序 = 下载区展示顺序（与审批类型选择器一致）。
TEMPLATE_FILES_BY_KIND: dict[str, str] = {
    "新签／续签派遣员工合同批量用章用印申请模板.xlsx": "seal_contract_batch",
    "解除（终止）劳动合同证明书用印申请模板.xlsx": "seal_terminate_cert",
    "补签派遣员工合同用章申请模板.xlsx": "seal_reseal",
    "两列金额明细表模板.xlsx": "meal2",
    "返餐费明细表模板.xlsx": "meal",
    "工会会费明细表模板.xlsx": "meal",
    "科研业绩津贴明细表模板.xlsx": "bonus",
}


def visible_kind_options() -> list[tuple[str, str]]:
    """(label, kind) pairs exposed in the 审批类型 selector, in display order."""
    options = []
    for kind, meta in SEAL_KINDS.items():
        if KIND_ENABLED_FLAGS.get(kind, True):
            options.append((meta["label"], kind))
    if MEAL2_KIND_ENABLED:
        options.append(("派遣员工简易费用明细表审批-两列金额", "meal2"))
    if MEAL_KIND_ENABLED:
        options.append(("派遣员工简易费用明细表审批", "meal"))
    if BONUS_KIND_ENABLED:
        options.append(("奖金发放审批", "bonus"))
    if INSURANCE_KIND_ENABLED:
        options.append(("五险一金审批", "insurance"))
    return options


def get_insurance_entity(entity: str) -> dict[str, str]:
    """Return the entity registry entry, falling back to 彩虹公司（默认主体）."""
    return INSURANCE_ENTITIES.get(entity, INSURANCE_ENTITIES[DEFAULT_INSURANCE_ENTITY])


def get_approval_code(kind: str, entity: str = DEFAULT_INSURANCE_ENTITY) -> str:
    """Return the effective approval code for the given kind (and entity).

    All approval codes come from environment variables; there are no
    built-in fallbacks. The legacy FEISHU_APPROVAL_CODE name is still
    accepted for meal approvals, and FEISHU_INSURANCE_APPROVAL_CODE is
    still accepted for the branch-company insurance approval, so existing
    deployments keep working.
    """
    if kind == "bonus":
        return _require_env_code("FEISHU_BONUS_APPROVAL_CODE", "奖金发放审批 Code")
    if kind in SEAL_KINDS:
        meta = SEAL_KINDS[kind]
        return _require_env_code(meta["env"], f"{meta['label']} Code")
    if kind == "meal2":
        # 两列金额派遣员工审批（2026-11 新增流程，独立于在用流程）
        return _require_env_code(
            "FEISHU_MEAL2_APPROVAL_CODE", "派遣员工简易费用明细表审批-两列金额 Code"
        )
    if kind == "insurance":
        meta = get_insurance_entity(entity)
        code = os.getenv(meta["env"], "").strip()
        if code:
            return code
        legacy_env = meta.get("legacy_env", "")
        if legacy_env:
            legacy_code = os.getenv(legacy_env, "").strip()
            if legacy_code:
                return legacy_code
        label = meta["label"]
        raise RuntimeError(
            f"缺少环境变量 {meta['env']}，请在 .env 中配置{label}五险一金审批 Code"
        )

    meal_code = os.getenv("FEISHU_MEAL_APPROVAL_CODE", "").strip()
    if meal_code:
        return meal_code

    legacy_code = os.getenv("FEISHU_APPROVAL_CODE", "").strip()
    if legacy_code:
        return legacy_code

    raise RuntimeError(
        "缺少环境变量 FEISHU_MEAL_APPROVAL_CODE，请在 .env 中配置派遣员工简易费用明细表审批 Code"
    )


def kind_meta(kind: str, entity: str = DEFAULT_INSURANCE_ENTITY) -> tuple[str, str]:
    """Return (label, approval_code) for a given kind (and entity).

    All approval codes come from environment variables; there are no
    built-in fallbacks. The insurance label carries the company entity,
    e.g. ``五险一金审批（分公司）``.
    """
    if kind == "bonus":
        return (
            "奖金发放审批",
            _require_env_code("FEISHU_BONUS_APPROVAL_CODE", "奖金发放审批 Code"),
        )
    if kind in SEAL_KINDS:
        meta = SEAL_KINDS[kind]
        return (meta["label"], get_approval_code(kind))
    if kind == "meal2":
        return (
            "派遣员工简易费用明细表审批-两列金额",
            get_approval_code("meal2"),
        )
    if kind == "insurance":
        meta = get_insurance_entity(entity)
        return (
            f"五险一金审批（{meta['label']}）",
            get_approval_code(kind, entity),
        )
    return "派遣员工简易费用明细表审批", get_approval_code(kind)


def _reset_preview_state():
    """Clear parsed sheet + editable title/remark widget state."""
    st.session_state.sheet = None
    st.session_state.title = ""
    st.session_state.remark = ""
    st.session_state.submitted = False
    # Widget keys must be removed too; otherwise Streamlit keeps the old
    # text_input/text_area display even though session values were reset.
    st.session_state.pop("title_input", None)
    st.session_state.pop("remark_input", None)


def _reset_sheet_state():
    """Reset uploaded-sheet state when switching approval kind."""
    _cleanup_tmp_file()
    _reset_preview_state()
    _clear_batch_items()
    st.session_state.uploaded_filename = None
    st.session_state.uploaded_file_sig = None
    st.session_state.sheet_infos = None
    st.session_state.selected_sheet_name = None


def _clear_batch_items():
    """清理所有批量上传状态及其临时文件。"""
    for item in st.session_state.batch_items:
        tmp_path = item.get("tmp_path")
        if tmp_path and Path(tmp_path).exists():
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
    st.session_state.batch_items = []
    # 批量编辑控件（审批标题/备注/用章事由）的 widget 状态一并清掉：
    # Streamlit 按 key 保留输入值，换文件或换审批类型后同序号的控件
    # 会沿用旧文件的编辑值，覆盖新生成的默认值。
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith(
            ("batch_title_", "batch_remark_", "batch_reason_")
        ):
            st.session_state.pop(key, None)


def _get_token() -> str:
    """Get tenant access token from environment variables."""
    app_id = os.getenv("FEISHU_APP_ID", "")
    app_secret = os.getenv("FEISHU_APP_SECRET", "")
    if not app_id or not app_secret:
        return ""
    try:
        return get_tenant_token(app_id, app_secret)
    except Exception:
        return ""


@st.cache_data(ttl=600)
def _get_approval_name_cached(approval_code: str) -> str:
    """Fetch approval name with caching."""
    token = _get_token()
    if not token:
        return ""
    try:
        info = get_approval_info(token, approval_code)
        return info.get("approval_name", "")
    except Exception:
        return ""


def _cleanup_tmp_file():
    """Clean up temporary file if it exists."""
    tmp_path = st.session_state.get("tmp_file_path")
    if tmp_path and Path(tmp_path).exists():
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    st.session_state.tmp_file_path = None


def _render_error_with_traceback(error_msg: str, error_obj: Exception = None):
    """Render error message with optional traceback in expander."""
    st.error(error_msg)
    if error_obj:
        with st.expander("技术详情"):
            st.code(traceback.format_exc())


# ---------------------------------------------------------------------------
# Login gate
# ---------------------------------------------------------------------------
def render_login_gate() -> bool:
    """Render the phone-number login gate.

    Returns True if the user is logged in (session has open_id).
    If not logged in, renders the login form and returns False.
    The caller should stop rendering the rest of the page.
    """
    if st.session_state.logged_in:
        # Logged-in view: welcome + logout
        col_welcome, col_logout = st.columns([4, 1])
        with col_welcome:
            st.subheader(f"欢迎，{st.session_state.user_name}")
            if st.session_state.name_lookup_failed:
                st.caption(
                    "提示：未能获取飞书姓名（可能缺少 contact:user.basic_profile:readonly 权限），"
                    "已显示 open_id"
                )
        with col_logout:
            if st.button("退出登录"):
                st.session_state.logged_in = False
                st.session_state.open_id = ""
                st.session_state.user_name = ""
                st.session_state.name_lookup_failed = False
                st.rerun()
        return True

    # Not logged in: show login form
    st.subheader("手机号登录")
    mobile = st.text_input("手机号", placeholder="请输入手机号")
    if st.button("登录"):
        if not mobile.strip():
            st.error("请输入手机号")
        else:
            token = _get_token()
            if not token:
                st.error("无法获取访问令牌，请检查飞书应用凭证配置")
            else:
                try:
                    user_info = lookup_user_by_mobile(token, mobile.strip())
                    open_id = user_info["open_id"]
                    # Best-effort name lookup
                    name = get_user_name(token, open_id)
                    display_name = name or open_id
                    st.session_state.logged_in = True
                    st.session_state.open_id = open_id
                    st.session_state.user_name = display_name
                    st.session_state.name_lookup_failed = name is None
                    st.success(f"登录成功：{display_name}")
                    st.rerun()
                except LookupError:
                    st.error("手机号未找到对应用户")
                except RuntimeError as e:
                    error_msg = str(e)
                    st.error(error_msg)
                    # Render grant URLs as clickable links
                    if "https://" in error_msg:
                        import re
                        urls = re.findall(r'https://[^\s\)）]+', error_msg)
                        if urls:
                            link_md = "  ".join(f"[开通链接]({u})" for u in urls)
                            st.markdown(link_md)
                    _render_error_with_traceback(error_msg, e)
    return False


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
def render_sidebar():
    """Render sidebar with tool info and environment status."""
    kind = st.session_state.approval_kind
    entity = st.session_state.approval_entity
    kind_label, _ = kind_meta(kind, entity)
    tool_title = f"{kind_label}提交工具"

    with st.sidebar:
        st.header(tool_title)
        st.divider()

        # Approval info
        try:
            approval_code = get_approval_code(kind, entity)
            routing_error = ""
        except RuntimeError as e:
            approval_code = ""
            routing_error = str(e)
        st.markdown(f"**审批类型：** {kind_label}")
        if kind == "insurance":
            st.markdown(f"**公司主体：** {get_insurance_entity(entity)['label']}")
        st.markdown(f"**审批定义：**")
        if routing_error:
            st.error(routing_error)
        else:
            st.code(approval_code, language=None)

            approval_name = _get_approval_name_cached(approval_code)
            if approval_name:
                st.success(f"审批名称：{approval_name}")
            else:
                app_id = os.getenv("FEISHU_APP_ID", "")
                app_secret = os.getenv("FEISHU_APP_SECRET", "")
                if not app_id or not app_secret:
                    st.warning("未配置飞书应用凭证")
                else:
                    st.warning("无法获取审批名称（请检查凭证或网络）")

        st.divider()

        # Environment check
        st.markdown("**环境检查：**")
        app_id = os.getenv("FEISHU_APP_ID", "")
        app_secret = os.getenv("FEISHU_APP_SECRET", "")

        if app_id and app_secret:
            st.success("✓ 飞书应用凭证已配置")
        else:
            st.error("✗ 缺少 FEISHU_APP_ID 或 FEISHU_APP_SECRET")

        st.divider()
        st.caption("制表人上传 Excel → 自动解析 → 提交飞书审批")


# ---------------------------------------------------------------------------
# Step 1: Upload Excel
# ---------------------------------------------------------------------------
def render_step1_upload():
    """Render Step 1: Upload and parse Excel file(s).

    Supports multi-file upload for all approval kinds. Each Excel file is
    parsed into one approval item stored in ``st.session_state.batch_items``.
    """
    st.subheader("步骤 1：上传表格")

    with st.expander("模板下载", expanded=False):
        st.markdown("**模板下载（可直接填写后上传）：**")
        templates_dir = _PROJECT_ROOT / "templates"
        # 模板随所属审批类型一起上下架：类型隐藏时模板文件保留在仓库，
        # 仅不在界面展示。顺序 = 字典定义顺序（与审批类型选择器一致）。
        template_paths = [
            templates_dir / name
            for name in TEMPLATE_FILES_BY_KIND
            if KIND_ENABLED_FLAGS.get(TEMPLATE_FILES_BY_KIND[name], False)
        ]
        if INSURANCE_KIND_ENABLED:
            template_paths += sorted((templates_dir / "insurance").glob("*.xlsx"))
        if template_paths:
            for chunk_start in range(0, len(template_paths), 4):
                chunk = template_paths[chunk_start : chunk_start + 4]
                cols = st.columns(4)
                for col, template_path in zip(cols, chunk):
                    with col:
                        st.download_button(
                            label=f"下载{template_path.stem}",
                            data=template_path.read_bytes(),
                            file_name=template_path.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
        else:
            st.warning("模板文件缺失")

        st.caption("模板已脱敏（不含真实个人信息）。")
        # 以下说明随所属审批类型一起显示，隐藏的工具不再展示旧文案。
        if MEAL_KIND_ENABLED:
            st.caption("返餐费/工会会费：请保留表头与合计行结构；返餐费末行写『总计』，工会会费末行写『合计』，工会会费多出的合计列可忽略。")
        if MEAL2_KIND_ENABLED:
            st.caption("两列金额：表头需含 姓名 和至少一个金额列；序号、身份证号、第二金额列可省略，单金额列表也可上传。末行合计行可选，有则自动校验。")
        if INSURANCE_KIND_ENABLED:
            st.caption("五险一金请使用对应的分单位模板（医保/社保/工伤/公积金），每月一个 sheet 或一个文件均可。")
        if st.session_state.approval_kind in SEAL_KINDS:
            meta = SEAL_KINDS[st.session_state.approval_kind]
            st.caption(
                f"用章/用印模板首行即列头，只需填明细列表；"
                f"{meta['option_label']}需填写：{'、'.join(seal_option_labels(meta['option_label']))}；"
                "合同/证明份数默认 4 份，可按需修改。"
            )
            st.caption("用章事由提交时按明细自动生成（人数、份数），步骤 2 中可修改；备注在步骤 3 填写。")
            if st.session_state.approval_kind == "seal_reseal":
                st.caption("补签合同：合同开始/结束时间填写日期（如 2026-01-01），期限天数自动计算；每行需填写补签原因。")
            if st.session_state.approval_kind == "seal_contract_batch":
                st.caption("新签/续签合同：身份证复印件盖公章份数按文本填写（如 1）。")

    kind = st.session_state.approval_kind
    if kind in SEAL_KINDS:
        label = SEAL_KINDS[kind]["upload_label"]
    elif kind == "bonus":
        label = "选择奖金明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"
    elif kind == "meal2":
        label = "选择两列金额明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"
    elif kind == "insurance":
        entity_label = get_insurance_entity(st.session_state.approval_entity)["label"]
        label = f"选择{entity_label}五险一金明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"
    else:
        label = "选择返餐费或工会会费明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"

    uploaded_files = st.file_uploader(label, type=["xls", "xlsx"], accept_multiple_files=True)

    st.caption("每个 Excel 文件会被解析为独立的审批单，提交时逐个生成飞书审批实例。")

    if not uploaded_files:
        _clear_batch_items()
        return

    # 计算当前上传的文件签名，用于检测文件集合是否变化
    current_sigs = [
        f"{f.name}:{getattr(f, 'size', None)}:{getattr(f, 'file_id', None)}"
        for f in uploaded_files
    ]
    existing_sigs = [item.get("file_sig") for item in st.session_state.batch_items]

    if current_sigs != existing_sigs:
        # 文件集合变化，先清理旧状态再重新解析全部
        _clear_batch_items()
        for uploaded_file in uploaded_files:
            file_sig = (
                f"{uploaded_file.name}:{getattr(uploaded_file, 'size', None)}:"
                f"{getattr(uploaded_file, 'file_id', None)}"
            )
            tmp_file = tempfile.NamedTemporaryFile(
                delete=False, suffix=Path(uploaded_file.name).suffix
            )
            tmp_file.write(uploaded_file.getvalue())
            tmp_file.close()

            try:
                if kind in SEAL_KINDS:
                    sheet = SEAL_KINDS[kind]["parser"](tmp_file.name)
                elif kind == "bonus":
                    sheet = parse_bonus_excel(tmp_file.name)
                elif kind == "meal2":
                    sheet = parse_meal_two_amount_excel(tmp_file.name)
                elif kind == "insurance":
                    sheet = parse_insurance_excel(tmp_file.name)
                else:
                    sheet = parse_meal_refund_excel(tmp_file.name)
                if kind in SEAL_KINDS:
                    # 用章流程：默认备注取表头信息区的「备注」单元格。
                    default_remark = sheet.remark
                else:
                    default_remark = f"制表：{sheet.maker}" if sheet.maker else ""
                st.session_state.batch_items.append(
                    {
                        "kind": kind,
                        "filename": uploaded_file.name,
                        "tmp_path": tmp_file.name,
                        "sheet": sheet,
                        "title": sheet.title,
                        "reason": sheet.reason if isinstance(sheet, SealDetailSheet) else "",
                        "remark": default_remark,
                        "file_code": None,
                        "instance_code": None,
                        "status": "pending",
                        "error": None,
                        "file_sig": file_sig,
                        "sheet_name": None,
                    }
                )
            except Exception as e:
                st.session_state.batch_items.append(
                    {
                        "kind": kind,
                        "filename": uploaded_file.name,
                        "tmp_path": tmp_file.name,
                        "sheet": None,
                        "title": "",
                        "reason": "",
                        "remark": "",
                        "file_code": None,
                        "instance_code": None,
                        "status": "error",
                        "error": str(e),
                        "file_sig": file_sig,
                        "sheet_name": None,
                    }
                )
    return


# ---------------------------------------------------------------------------
# Step 2: Preview data
# ---------------------------------------------------------------------------
def render_step2_preview():
    """Render Step 2: Preview parsed data and lookup submitter."""
    st.subheader("步骤 2：预览数据")

    kind = st.session_state.approval_kind
    items = st.session_state.batch_items

    # ── 批量预览 ──────────────────────────────────────────
    if items:
        error_count = sum(1 for it in items if it["status"] == "error")
        if error_count:
            st.error(f"{error_count} 个文件解析失败，请检查后重新上传")

        for idx, item in enumerate(items):
            sheet = item["sheet"]
            with st.expander(
                f"{idx + 1}. {item['filename']}"
                + ("" if item["status"] != "error" else " ⚠ 解析失败"),
                expanded=True,
            ):
                if item["status"] == "error":
                    st.error(f"解析失败：{item['error']}")
                    continue

                if kind in SEAL_KINDS:
                    # 用章表单没有标题控件，用「用章事由」替代标题编辑
                    # （解析时按明细自动生成默认值，此处可修改）。
                    reason = st.text_input(
                        "用章事由（自动生成，可修改）",
                        value=item.get("reason", ""),
                        key=f"batch_reason_{idx}",
                    )
                    item["reason"] = reason
                else:
                    # Editable title per file
                    title_key = f"batch_title_{idx}"
                    title = st.text_input(
                        "审批标题",
                        value=item["title"],
                        key=title_key,
                    )
                    item["title"] = title

                st.markdown("**明细数据：**")

                data = _build_preview_data(kind, sheet)
                if data:
                    st.dataframe(data, use_container_width=True, hide_index=True)

                _render_sheet_totals(kind, sheet)
                st.markdown(f"**数据行数：** {len(sheet.rows)}")

                if sheet.warnings:
                    for warning in sheet.warnings:
                        st.warning(warning)

        st.divider()
        return

    # ── 单文件预览（兼容旧状态）───────────────────────────
    sheet = st.session_state.sheet
    if sheet is None:
        st.info("请先上传表格")
        return

    # Editable title
    title = st.text_input("审批标题", value=st.session_state.title, key="title_input")
    st.session_state.title = title

    st.divider()

    # Data preview
    st.markdown("**明细数据：**")

    data = _build_preview_data(kind, sheet)
    if data:
        st.dataframe(data, use_container_width=True, hide_index=True)

    _render_sheet_totals(kind, sheet)
    st.markdown(f"**数据行数：** {len(sheet.rows)}")

    # Warnings
    if sheet.warnings:
        for warning in sheet.warnings:
            st.warning(warning)

    st.divider()


def _build_preview_data(kind, sheet):
    """Build preview dataframe rows for a parsed sheet."""
    data = []
    if kind in SEAL_KINDS:
        for i, row in enumerate(sheet.rows, start=1):
            item = {
                "序号": row.serial or str(i),
                "派遣员工姓名": row.name,
                sheet.copies_label: row.copies,
                sheet.option_label: row.option_text,
            }
            if kind == "seal_reseal":
                item["合同开始时间"] = row.start_date[:10] if row.start_date else ""
                item["合同结束时间"] = row.end_date[:10] if row.end_date else ""
                item["合同期限（天）"] = row.interval_days
                item["补签原因"] = row.reason
            if kind == "seal_contract_batch":
                item["身份证复印件盖公章份数"] = row.id_copy_copies
            data.append(item)
    elif kind == "meal2":
        # 列名跟随原始 Excel 表头，与「列名称映射」一致；
        # 源表缺列（序号/身份证号/第二金额列）时预览同步省略。
        for i, row in enumerate(sheet.rows, start=1):
            item = {
                "序号": row.serial if sheet.has_serial else str(i),
                "姓名": row.name,
            }
            if sheet.has_id_column:
                item["身份证号码"] = row.id_card
            item[sheet.amount1_column] = row.amount1
            if sheet.amount2_column:
                item[sheet.amount2_column] = row.amount2
            data.append(item)
    elif kind == "bonus":
        for row in sheet.rows:
            data.append({
                "序号": row.serial,
                "工号": row.employee_id,
                "姓名": row.name,
                "身份证号": row.id_card,
                "奖金": row.bonus,
                "应纳个人所得税": row.tax,
                "个人其他调整": row.adjustment,
                "实发金额": row.net_amount,
            })
    elif kind == "insurance":
        for row in sheet.rows:
            fee_pairs = "；".join(f"{f.name}={f.amount}" for f in row.fees)
            data.append({
                "单位名称": row.unit_name,
                "参保人数": row.headcount,
                "费用明细": fee_pairs,
            })
    else:
        for row in sheet.rows:
            data.append({
                "序号": row.serial,
                "姓名": row.name,
                "身份证号码": row.id_card,
                "金额": row.amount,
            })
    return data


def _render_sheet_totals(kind, sheet):
    """Render total lines for a parsed sheet."""
    if kind in SEAL_KINDS:
        st.markdown(f"**{sheet.copies_label}合计：** {sheet.total_copies}")
        return
    if kind == "meal2":
        st.markdown(f"**{sheet.amount1_column}合计：** {sheet.amount1_total or '—'}")
        if sheet.amount2_column:
            st.markdown(f"**{sheet.amount2_column}合计：** {sheet.amount2_total or '—'}")
    elif kind == "bonus":
        st.markdown(f"**奖金合计：** {sheet.bonus_total}")
        st.markdown(f"**个人所得税合计：** {sheet.tax_total}")
        st.markdown(f"**个人其他调整合计：** {sheet.adjustment_total}")
        st.markdown(f"**实发金额合计：** {sheet.net_total}")
    elif kind == "insurance":
        pass  # Insurance units carry per-unit fees; no single grand total line.
    else:
        if sheet.total_chinese:
            st.markdown(f"**合计金额：** {sheet.total_amount}（{sheet.total_chinese}）")
        else:
            st.markdown(f"**合计金额：** {sheet.total_amount}")


# ---------------------------------------------------------------------------
# Step 3: Remark
# ---------------------------------------------------------------------------
def render_step3_remark():
    """Render Step 3: Optional remark."""
    st.subheader("步骤 3：备注")

    kind = st.session_state.approval_kind
    items = st.session_state.batch_items

    # ── 批量备注 ──────────────────────────────────────────
    if items:
        for idx, item in enumerate(items):
            sheet = item["sheet"]
            if item["status"] == "error":
                continue
            if item["kind"] in SEAL_KINDS:
                # 用章流程的备注默认值已在解析时取自表头信息区。
                default_remark = ""
            else:
                default_remark = f"制表：{sheet.maker}" if sheet and sheet.maker else ""
            remark_key = f"batch_remark_{idx}"
            remark = st.text_area(
                f"{item['filename']} 备注（可选）",
                value=item["remark"] or default_remark,
                key=remark_key,
            )
            item["remark"] = remark
        st.divider()
        return

    # ── 单文件备注 ────────────────────────────────────────────────
    sheet = st.session_state.sheet
    default_remark = f"制表：{sheet.maker}" if sheet and sheet.maker else ""

    remark = st.text_area("备注（可选）", value=st.session_state.remark or default_remark, key="remark_input")
    st.session_state.remark = remark


# ---------------------------------------------------------------------------
# Step 4: Actions
# ---------------------------------------------------------------------------
def render_step4_actions():
    """Render Step 4: Preview and submit actions."""
    st.subheader("步骤 4：操作")

    kind = st.session_state.approval_kind
    items = st.session_state.batch_items

    # ── 批量提交 ──────────────────────────────────────────
    if items:
        pending = [it for it in items if it["status"] == "pending"]
        errors = [it for it in items if it["status"] == "error"]
        done = [it for it in items if it["status"] == "success"]

        if errors:
            st.error(f"{len(errors)} 个文件解析失败，无法提交，请检查后重新上传")

        # 用章流程：用章事由为表单必填，提交前逐单校验。
        missing_reason = [
            it for it in pending
            if kind in SEAL_KINDS and not (it.get("reason") or "").strip()
        ]
        if missing_reason:
            st.error(
                "以下审批单缺少用章事由（表单必填），请在步骤 2 中填写："
                + "、".join(it["filename"] for it in missing_reason)
            )

        if not pending:
            st.info("没有待提交的审批单")
        else:
            st.markdown(f"**待提交审批单：** {len(pending)} 个")
            for it in pending:
                st.caption(f"• {it['filename']}（{len(it['sheet'].rows)} 行数据）")

        col_submit, _ = st.columns(2)
        with col_submit:
            submit_all = st.button(
                "全部提交审批",
                type="primary",
                use_container_width=True,
                disabled=not pending
                or bool(missing_reason)
                or st.session_state.get("batch_submitting", False),
            )

        if submit_all:
            st.session_state.batch_submitting = True
            token = _get_token()
            if not token:
                st.error("无法获取访问令牌")
                st.session_state.batch_submitting = False
            else:
                try:
                    approval_code = get_approval_code(kind, st.session_state.approval_entity)
                except RuntimeError as e:
                    st.error(str(e))
                    st.session_state.batch_submitting = False
                    return
                if kind in SEAL_KINDS:
                    build_widgets = SEAL_KINDS[kind]["form_builder"]
                elif kind == "bonus":
                    build_widgets = build_bonus_form_widgets
                elif kind == "meal2":
                    build_widgets = build_meal2_form_widgets
                elif kind == "insurance":
                    build_widgets = get_insurance_entity(
                        st.session_state.approval_entity
                    )["form_builder"]
                else:
                    build_widgets = build_form_widgets

                results = []
                for idx, item in enumerate(items):
                    if item["status"] != "pending":
                        continue
                    sheet = item["sheet"]
                    try:
                        with st.spinner(f"正在提交 {idx + 1}/{len(pending)}：{item['filename']}..."):
                            # 上传附件
                            file_code = upload_approval_file(
                                token,
                                item["tmp_path"],
                                filename=item["filename"],
                            )
                            # 构建表单控件
                            if kind in SEAL_KINDS:
                                form_widgets = build_widgets(
                                    sheet,
                                    file_code=file_code,
                                    reason=item.get("reason", ""),
                                    remark=item.get("remark", ""),
                                )
                            else:
                                form_widgets = build_widgets(
                                    sheet,
                                    file_code=file_code,
                                    title=item["title"],
                                    remark=item["remark"],
                                )
                            # 创建审批实例
                            instance_code = create_instance(
                                token,
                                approval_code,
                                st.session_state.open_id,
                                form_widgets,
                                uuid=str(uuid.uuid4()),
                            )
                        item["file_code"] = file_code
                        item["instance_code"] = instance_code
                        item["status"] = "success"
                        results.append((item["filename"], instance_code, None))
                    except Exception as e:
                        item["status"] = "error"
                        item["error"] = str(e)
                        results.append((item["filename"], None, str(e)))

                st.session_state.batch_submitting = False

                # 显示结果
                success_count = sum(1 for _, _, err in results if err is None)
                fail_count = len(results) - success_count
                if success_count:
                    st.success(f"成功提交 {success_count} 个审批单")
                if fail_count:
                    st.error(f"提交失败 {fail_count} 个")

                for filename, instance_code, err in results:
                    if err is None:
                        st.markdown(f"✅ **{filename}** → 实例编号：`{instance_code}`")
                        try:
                            detail = get_instance_detail(token, instance_code)
                            instance_url = detail.get("instance_url", "")
                            if instance_url:
                                st.markdown(f"  [查看审批详情]({instance_url})")
                        except Exception:
                            pass
                    else:
                        st.markdown(f"❌ **{filename}** → 提交失败：{err}")

        st.divider()
        return

    # ── 单文件提交（返餐费 / 五险一金 / 兼容旧 bonus 状态）─────────
    sheet = st.session_state.sheet
    if sheet is None:
        st.warning("请先上传表格")
        return

    try:
        approval_code = get_approval_code(kind, st.session_state.approval_entity)
    except RuntimeError as e:
        st.error(str(e))
        return
    if kind in SEAL_KINDS:
        build_widgets = SEAL_KINDS[kind]["form_builder"]
    elif kind == "bonus":
        build_widgets = build_bonus_form_widgets
    elif kind == "meal2":
        build_widgets = build_meal2_form_widgets
    elif kind == "insurance":
        build_widgets = get_insurance_entity(st.session_state.approval_entity)[
            "form_builder"
        ]
    else:
        build_widgets = build_form_widgets

    def _build_current_form_widgets(file_code: str) -> list[dict]:
        """按当前类型构建表单控件（用章流程无标题控件，传用章事由）。"""
        if kind in SEAL_KINDS:
            return build_widgets(
                sheet,
                file_code=file_code,
                reason=sheet.reason,
                remark=st.session_state.remark,
            )
        return build_widgets(
            sheet,
            file_code=file_code,
            title=st.session_state.title,
            remark=st.session_state.remark,
        )

    col_preview, col_submit = st.columns(2)

    # Preview button
    with col_preview:
        if st.button("预览审批流程", use_container_width=True):
            token = _get_token()
            if not token:
                st.error("无法获取访问令牌")
            else:
                try:
                    # Build form with dummy file code for preview
                    form_widgets = _build_current_form_widgets("DRYRUN")
                    preview = preview_instance(
                        token,
                        approval_code,
                        st.session_state.open_id,
                        form_widgets,
                    )
                    nodes = preview.get("nodes", [])
                    if nodes:
                        st.markdown("**审批流程：**")
                        node_data = []
                        for node in nodes:
                            node_data.append({
                                "节点名称": node.get("node_name", ""),
                                "审批人": ", ".join(node.get("user_id_list", [])),
                            })
                        st.dataframe(node_data, use_container_width=True, hide_index=True)
                    else:
                        st.info("未获取到审批节点信息")
                except Exception as e:
                    _render_error_with_traceback(f"预览失败：{e}", e)

    # Submit button
    with col_submit:
        submit_clicked = st.button(
            "提交审批",
            type="primary",
            use_container_width=True,
            disabled=st.session_state.submitted,
        )

        if submit_clicked:
            if st.session_state.submitted:
                st.warning("已提交，请勿重复提交")
                return

            token = _get_token()
            if not token:
                st.error("无法获取访问令牌")
                return

            tmp_path = st.session_state.tmp_file_path
            if not tmp_path or not Path(tmp_path).exists():
                st.error("临时文件不存在，请重新上传")
                return

            try:
                with st.spinner("正在上传并提交..."):
                    original_filename = st.session_state.get("uploaded_filename")
                    file_code = upload_approval_file(
                        token, tmp_path, filename=original_filename
                    )

                    # Build form widgets
                    form_widgets = _build_current_form_widgets(file_code)

                    # Create instance
                    instance_code = create_instance(
                        token,
                        approval_code,
                        st.session_state.open_id,
                        form_widgets,
                        uuid=str(uuid.uuid4()),
                    )

                    st.session_state.submitted = True
                    st.session_state.last_instance = instance_code

                st.success("提交成功！")
                st.markdown(f"**实例编号：** `{instance_code}`")

                # Try to get instance URL
                try:
                    detail = get_instance_detail(token, instance_code)
                    instance_url = detail.get("instance_url", "")
                    if instance_url:
                        st.markdown(f"[查看审批详情]({instance_url})")
                except Exception:
                    # Failure to fetch URL is not critical
                    pass

            except RuntimeError as e:
                _render_error_with_traceback(f"提交失败：{e}", e)
            except Exception as e:
                _render_error_with_traceback(f"提交失败：{e}", e)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    """Main entry point."""
    kind_options = visible_kind_options()
    kind_keys = [k for _, k in kind_options]

    # 已隐藏的类型（五险一金暂时整体下线）不再路由：旧会话残留的隐藏
    # 类型归位到默认类型并清空已上传数据，避免标题/侧边栏闪现旧类型。
    if st.session_state.approval_kind not in kind_keys:
        st.session_state.approval_kind = kind_keys[0]
        _reset_sheet_state()

    # 已隐藏的主体（如彩虹公司）不再路由：旧会话残留的隐藏主体直接
    # 归位到第一个可见主体，避免标题/侧边栏闪现隐藏主体名称。
    if st.session_state.approval_entity not in visible_insurance_entities():
        st.session_state.approval_entity = get_default_insurance_entity()

    render_sidebar()

    kind = st.session_state.approval_kind
    entity_label = get_insurance_entity(st.session_state.approval_entity)["label"]
    if kind == "seal_reseal":
        st.title("补签派遣员工合同用章申请提交工具（批量）")
        st.markdown(
            "上传补签派遣员工合同用章明细表，每个文件生成一个独立审批单，"
            "表内每行一位员工（含合同起止时间、印章类型、补签原因、合同份数）。"
            "该流程仅限补签派遣员工合同使用，其他文件勿用。"
        )
    elif kind == "seal_terminate_cert":
        st.title("解除（终止）劳动合同证明书用印申请提交工具（批量）")
        st.markdown(
            "上传解除（终止）劳动合同证明书用印明细表，每个文件生成一个独立审批单，"
            "表内每行一位员工（含证明份数、印章类型）。"
            "该流程仅限解除（终止）劳动合同证明书使用，其他文件勿用。"
        )
    elif kind == "seal_contract_batch":
        st.title("新签/续签派遣员工合同批量用章用印申请提交工具（批量）")
        st.markdown(
            "上传新签/续签派遣员工合同批量用章明细表，每个文件生成一个独立审批单，"
            "表内每行一位员工（含签约主体、合同份数、身份证复印件盖公章份数）。"
            "该流程仅限新签/续签派遣员工合同使用，其他文件勿用。"
        )
    elif kind == "bonus":
        st.title("奖金发放审批提交工具（批量）")
        st.markdown("上传多个奖金明细表，每个文件生成一个独立审批单，逐个提交飞书审批流程。")
    elif kind == "insurance":
        st.title(f"五险一金审批提交工具（{entity_label}）")
        st.markdown("上传五险一金分单位明细表，自动解析并提交飞书审批流程。")
    elif kind == "meal2":
        st.title("派遣员工简易费用明细表审批提交工具（两列金额）")
        st.markdown("适用于一至两个金额列的明细表（如单位工会经费、工会会费，单金额列表亦可）。表头至少含 姓名 和一个金额列，序号、身份证号、第二金额列可省略；末行合计行可选，有则自动校验。审批表单的「列名称映射」按上传表的列名自动填写。")
    else:
        st.title("派遣员工简易费用明细表审批提交工具")
        st.markdown("适用于手工制作的退返餐费、工会会费等简易费用表。表中除序号、姓名、身份证号外，仅设一列金额，表格末尾为合计金额。审批顺序：夏璠→张宇佳→蒿敬恪。")

    st.divider()

    # Login gate — must pass before any other UI
    if not render_login_gate():
        return

    # Approval type selector：仅展示可见类型（五险一金暂时整体下线，代码保留）。
    kind_label = st.radio(
        "审批类型",
        [label for label, _ in kind_options],
        horizontal=True,
    )
    new_kind = dict(kind_options)[kind_label]
    if new_kind != st.session_state.approval_kind:
        st.session_state.approval_kind = new_kind
        _reset_sheet_state()
        st.rerun()

    # 五险一金主体选择器：分公司 / 彩虹公司路由到不同审批定义。
    # 解析结果与主体无关，切换主体不清空已上传数据，仅切换提交路由。
    # 仅展示 enabled 的主体（彩虹公司暂时停用，代码保留）。
    if new_kind == "insurance":
        visible_entities = visible_insurance_entities()
        entity_keys = list(visible_entities)
        entity_labels = [visible_entities[k]["label"] for k in entity_keys]
        current_entity = st.session_state.approval_entity
        default_idx = entity_keys.index(current_entity) if current_entity in entity_keys else 0
        entity_choice = st.radio(
            "公司主体",
            entity_labels,
            index=default_idx,
            horizontal=True,
        )
        new_entity = entity_keys[entity_labels.index(entity_choice)]
        if new_entity != st.session_state.approval_entity:
            st.session_state.approval_entity = new_entity
            st.rerun()

    st.divider()

    render_step1_upload()
    st.divider()

    render_step2_preview()
    st.divider()

    render_step3_remark()
    st.divider()

    render_step4_actions()


if __name__ == "__main__":
    main()
