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
    parse_insurance_excel,
    list_sheets,
    list_insurance_sheets,
    MealRefundSheet,
    BonusSheet,
    InsuranceSheet,
)
from app.form_builder import (
    build_form_widgets,
    build_bonus_form_widgets,
    build_insurance_form_widgets,
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


def get_approval_code(kind: str) -> str:
    """Return the effective approval code for the given kind.

    All approval codes come from environment variables; there are no
    built-in fallbacks. The legacy FEISHU_APPROVAL_CODE name is still
    accepted for meal approvals so existing deployments keep working.
    """
    if kind == "bonus":
        return _require_env_code("FEISHU_BONUS_APPROVAL_CODE", "奖金发放审批 Code")
    if kind == "insurance":
        return _require_env_code("FEISHU_INSURANCE_APPROVAL_CODE", "五险一金审批 Code")

    meal_code = os.getenv("FEISHU_MEAL_APPROVAL_CODE", "").strip()
    if meal_code:
        return meal_code

    legacy_code = os.getenv("FEISHU_APPROVAL_CODE", "").strip()
    if legacy_code:
        return legacy_code

    raise RuntimeError(
        "缺少环境变量 FEISHU_MEAL_APPROVAL_CODE，请在 .env 中配置派遣员工简易费用明细表审批 Code"
    )


def kind_meta(kind: str) -> tuple[str, str]:
    """Return (label, approval_code) for a given kind.

    All approval codes come from environment variables; there are no
    built-in fallbacks.
    """
    if kind == "bonus":
        return (
            "奖金发放审批",
            _require_env_code("FEISHU_BONUS_APPROVAL_CODE", "奖金发放审批 Code"),
        )
    if kind == "insurance":
        return (
            "五险一金审批",
            _require_env_code("FEISHU_INSURANCE_APPROVAL_CODE", "五险一金审批 Code"),
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
    kind_label, _ = kind_meta(kind)
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
    """Render Step 1: Upload and parse Excel file(s).

    Supports multi-file upload for all approval kinds. Each Excel file is
    parsed into one approval item stored in ``st.session_state.batch_items``.
    """
    st.subheader("步骤 1：上传表格")

    with st.expander("模板下载", expanded=False):
        st.markdown("**模板下载（可直接填写后上传）：**")
        templates_dir = _PROJECT_ROOT / "templates"
        template_paths = sorted(templates_dir.glob("*.xlsx")) + sorted(
            (templates_dir / "insurance").glob("*.xlsx")
        )
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

        st.caption("模板已脱敏（不含真实个人信息）；请保留表头与合计行结构后填写。")
        st.caption("返餐费末行写『总计』，工会会费末行写『合计』；工会会费多出的合计列可忽略。")
        st.caption("五险一金请使用对应的分单位模板（医保/社保/工伤/公积金），每月一个 sheet 或一个文件均可。")

    kind = st.session_state.approval_kind
    if kind == "bonus":
        label = "选择奖金明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"
    elif kind == "insurance":
        label = "选择五险一金明细表（.xls / .xlsx，可多选，每个文件生成一个审批单）"
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
                if kind == "bonus":
                    sheet = parse_bonus_excel(tmp_file.name)
                elif kind == "insurance":
                    sheet = parse_insurance_excel(tmp_file.name)
                else:
                    sheet = parse_meal_refund_excel(tmp_file.name)
                st.session_state.batch_items.append(
                    {
                        "kind": kind,
                        "filename": uploaded_file.name,
                        "tmp_path": tmp_file.name,
                        "sheet": sheet,
                        "title": sheet.title,
                        "remark": f"制表：{sheet.maker}" if sheet.maker else "",
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
    if kind == "bonus":
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
                disabled=not pending or st.session_state.get("batch_submitting", False),
            )

        if submit_all:
            st.session_state.batch_submitting = True
            token = _get_token()
            if not token:
                st.error("无法获取访问令牌")
                st.session_state.batch_submitting = False
            else:
                approval_code = get_approval_code(kind)
                if kind == "bonus":
                    build_widgets = build_bonus_form_widgets
                elif kind == "insurance":
                    build_widgets = build_insurance_form_widgets
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

    approval_code = get_approval_code(kind)
    if kind == "bonus":
        build_widgets = build_bonus_form_widgets
    elif kind == "insurance":
        build_widgets = build_insurance_form_widgets
    else:
        build_widgets = build_form_widgets

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
        st.title("奖金发放审批提交工具（批量）")
        st.markdown("上传多个奖金明细表，每个文件生成一个独立审批单，逐个提交飞书审批流程。")
    elif kind == "insurance":
        st.title("五险一金审批提交工具")
        st.markdown("上传五险一金分单位明细表，自动解析并提交飞书审批流程。")
    else:
        st.title("派遣员工简易费用明细表审批提交工具")
        st.markdown("适用于手工制作的退返餐费、工会会费等简易费用表。表中除序号、姓名、身份证号外，仅设一列金额，表格末尾为合计金额。审批顺序：夏璠→张宇佳→蒿敬恪。")

    st.divider()

    # Login gate — must pass before any other UI
    if not render_login_gate():
        return

    # Approval type selector
    kind_label = st.radio(
        "审批类型",
        ["派遣员工简易费用明细表审批", "奖金发放审批", "五险一金审批"],
        horizontal=True,
    )
    if kind_label == "派遣员工简易费用明细表审批":
        new_kind = "meal"
    elif kind_label == "奖金发放审批":
        new_kind = "bonus"
    else:
        new_kind = "insurance"
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
