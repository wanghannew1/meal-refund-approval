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
    parse_bonus_excel,
    MealRefundSheet,
    BonusSheet,
)
from app.form_builder import (
    build_form_widgets,
    build_bonus_form_widgets,
    APPROVAL_CODE,
    BONUS_APPROVAL_CODE,
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
    st.session_state.approval_kind = "meal"
if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = None
if "name_lookup_failed" not in st.session_state:
    st.session_state.name_lookup_failed = False

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def get_approval_code(kind: str) -> str:
    """Return the effective approval code for the given kind.

    Reads from environment variables first, falling back to the built-in
    constants. This ensures env overrides work consistently across sidebar,
    preview, and submit paths.
    """
    if kind == "bonus":
        return os.getenv("FEISHU_BONUS_APPROVAL_CODE", BONUS_APPROVAL_CODE)
    return os.getenv("FEISHU_APPROVAL_CODE", APPROVAL_CODE)


def _reset_sheet_state():
    """Reset uploaded-sheet state when switching approval kind."""
    _cleanup_tmp_file()
    st.session_state.sheet = None
    st.session_state.title = ""
    st.session_state.remark = ""
    st.session_state.submitted = False
    st.session_state.uploaded_filename = None


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
                    "提示：未能获取飞书姓名（可能缺少 contact:contact.base:readonly 权限），"
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
    kind_label = "返餐费审批" if kind == "meal" else "奖金发放审批"
    tool_title = f"{kind_label}提交工具"

    with st.sidebar:
        st.header(tool_title)
        st.divider()

        # Approval info
        approval_code = get_approval_code(kind)
        st.markdown(f"**审批类型：** {kind_label}")
        st.markdown(f"**审批定义：**")
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
    """Render Step 1: Upload and parse Excel file."""
    st.subheader("步骤 1：上传表格")

    kind = st.session_state.approval_kind
    if kind == "bonus":
        label = "选择奖金明细表（.xls / .xlsx）"
    else:
        label = "选择退返餐费明细表（.xls / .xlsx）"

    uploaded_file = st.file_uploader(label, type=["xls", "xlsx"])

    if uploaded_file is not None:
        # Clean up previous temp file
        _cleanup_tmp_file()

        # Save to temp file
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix)
        tmp_file.write(uploaded_file.getvalue())
        tmp_file.close()
        st.session_state.tmp_file_path = tmp_file.name
        st.session_state.uploaded_filename = uploaded_file.name

        # Parse the file
        try:
            if kind == "bonus":
                sheet = parse_bonus_excel(tmp_file.name)
            else:
                sheet = parse_meal_refund_excel(tmp_file.name)
            st.session_state.sheet = sheet
            st.session_state.title = sheet.title
            st.session_state.submitted = False  # Reset submitted flag
            st.success(f"解析成功：共 {len(sheet.rows)} 行数据")
        except ValueError as e:
            st.session_state.sheet = None
            _render_error_with_traceback(f"解析失败：{e}", e)
        except Exception as e:
            st.session_state.sheet = None
            _render_error_with_traceback(f"解析失败：{e}", e)


# ---------------------------------------------------------------------------
# Step 2: Preview data
# ---------------------------------------------------------------------------
def render_step2_preview():
    """Render Step 2: Preview parsed data and lookup submitter."""
    st.subheader("步骤 2：预览数据")

    sheet = st.session_state.sheet
    if sheet is None:
        st.info("请先上传表格")
        return

    kind = st.session_state.approval_kind

    # Editable title
    title = st.text_input("审批标题", value=st.session_state.title)
    st.session_state.title = title

    st.divider()

    # Data preview
    st.markdown("**明细数据：**")

    # Build dataframe
    data = []
    if kind == "bonus":
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
    else:
        for row in sheet.rows:
            data.append({
                "序号": row.serial,
                "姓名": row.name,
                "身份证号码": row.id_card,
                "金额": row.amount,
            })

    if data:
        st.dataframe(data, use_container_width=True, hide_index=True)

    # Total line
    if kind == "bonus":
        st.markdown(f"**奖金合计：** {sheet.bonus_total}")
        st.markdown(f"**个人所得税合计：** {sheet.tax_total}")
        st.markdown(f"**个人其他调整合计：** {sheet.adjustment_total}")
        st.markdown(f"**实发金额合计：** {sheet.net_total}")
    else:
        st.markdown(f"**合计金额：** {sheet.total_amount}（{sheet.total_chinese}）")
    st.markdown(f"**数据行数：** {len(sheet.rows)}")

    # Warnings
    if sheet.warnings:
        for warning in sheet.warnings:
            st.warning(warning)

    st.divider()


# ---------------------------------------------------------------------------
# Step 3: Remark
# ---------------------------------------------------------------------------
def render_step3_remark():
    """Render Step 3: Optional remark."""
    st.subheader("步骤 3：备注")

    sheet = st.session_state.sheet
    default_remark = f"制表：{sheet.maker}" if sheet and sheet.maker else ""

    remark = st.text_area("备注（可选）", value=st.session_state.remark or default_remark)
    st.session_state.remark = remark


# ---------------------------------------------------------------------------
# Step 4: Actions
# ---------------------------------------------------------------------------
def render_step4_actions():
    """Render Step 4: Preview and submit actions."""
    st.subheader("步骤 4：操作")

    sheet = st.session_state.sheet
    kind = st.session_state.approval_kind

    if sheet is None:
        st.warning("请先上传表格")
        return

    approval_code = get_approval_code(kind)
    build_widgets = build_bonus_form_widgets if kind == "bonus" else build_form_widgets

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
                    form_widgets = build_widgets(
                        sheet,
                        file_code="DRYRUN",
                        title=st.session_state.title,
                        remark=st.session_state.remark,
                    )
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
                    form_widgets = build_widgets(
                        sheet,
                        file_code=file_code,
                        title=st.session_state.title,
                        remark=st.session_state.remark,
                    )

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
    render_sidebar()

    kind = st.session_state.approval_kind
    if kind == "bonus":
        st.title("奖金发放审批提交工具")
        st.markdown("上传奖金明细表，自动解析并提交飞书审批流程。")
    else:
        st.title("返餐费审批提交工具")
        st.markdown("上传退返餐费明细表，自动解析并提交飞书审批流程。")

    st.divider()

    # Login gate — must pass before any other UI
    if not render_login_gate():
        return

    # Approval type selector
    kind_label = st.radio(
        "审批类型",
        ["返餐费审批", "奖金发放审批"],
        horizontal=True,
    )
    new_kind = "meal" if kind_label == "返餐费审批" else "bonus"
    if new_kind != st.session_state.approval_kind:
        st.session_state.approval_kind = new_kind
        _reset_sheet_state()
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
