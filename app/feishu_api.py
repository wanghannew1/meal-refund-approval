"""
Feishu API client for meal refund approval operations.

Provides functions to authenticate with Feishu Open Platform,
upload approval files, look up users, and create/preview approval
instances for the meal refund approval form.
"""

from __future__ import annotations

import json
import os
import re
import threading
import time
from pathlib import Path

import requests

# ── API endpoints ────────────────────────────────────────────────────
BASE_URL = "https://open.feishu.cn/open-apis"
TOKEN_URL = f"{BASE_URL}/auth/v3/tenant_access_token/internal/"
APPROVAL_FILE_UPLOAD_URL = "https://www.feishu.cn/approval/openapi/v2/file/upload"
APPROVAL_INFO_URL = f"{BASE_URL}/approval/v4/approvals/{{approval_code}}"
INSTANCES_URL = f"{BASE_URL}/approval/v4/instances"
INSTANCE_DETAIL_URL = f"{BASE_URL}/approval/v4/instances/{{instance_code}}"
PREVIEW_URL = f"{BASE_URL}/approval/v4/instances/preview"
CONTACT_LOOKUP_URL = f"{BASE_URL}/contact/v3/users/batch_get_id"

# ── Token cache ──────────────────────────────────────────────────────
CACHE_FILE = Path(".token_cache.json")
_CACHE_LOCK = threading.Lock()

# ── Scope error grant URL ────────────────────────────────────────────
_SCOPE_GRANT_URL = (
    "https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth"
    "?q=contact:user.id:readonly&op_from=openapi&token_type=tenant"
)


def _load_cached_token(cache_path: Path | None = None) -> dict | None:
    """Load cached token data if it exists and is not expired.

    Args:
        cache_path: Optional override for the cache file path.

    Returns:
        Cached token dict with tenant_access_token and expire_at,
        or None if not found or expired.
    """
    if cache_path is None:
        cache_path = CACHE_FILE
    if not cache_path.exists():
        return None
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
    if data.get("expire_at", 0) > time.time():
        return data
    return None


def _save_token(
    token: str, expire_in: int, cache_path: Path | None = None
) -> None:
    """Save token to cache with 5-minute early expiry.

    Uses atomic write (temp file + os.replace) to prevent
    corrupted cache files.

    Args:
        token: Tenant access token string.
        expire_in: Token lifetime in seconds.
        cache_path: Optional override for the cache file path.
    """
    if cache_path is None:
        cache_path = CACHE_FILE
    data = {
        "tenant_access_token": token,
        "expire_at": time.time() + expire_in - 300,
    }
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    # Atomic write to avoid corrupted cache files
    temp_path = cache_path.with_suffix(".tmp")
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(data, f)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temp_path, cache_path)


def get_tenant_token(app_id: str, app_secret: str) -> str:
    """Get tenant access token from Feishu API.

    Fetches a new token if none is cached or if the cached token has
    expired (with a 5-minute early expiry margin).  The cache is
    written atomically and protected by a thread lock.

    Args:
        app_id: Feishu application ID.
        app_secret: Feishu application secret.

    Returns:
        Tenant access token string.

    Raises:
        RuntimeError: If token acquisition fails.
    """
    with _CACHE_LOCK:
        cached = _load_cached_token()
        if cached:
            return cached["tenant_access_token"]

    payload = {"app_id": app_id, "app_secret": app_secret}
    resp = requests.post(TOKEN_URL, json=payload)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"获取 token 失败: code={data.get('code')} msg={data.get('msg')}"
        )

    token = data["tenant_access_token"]
    with _CACHE_LOCK:
        # Re-check in case another thread already saved a token
        cached = _load_cached_token()
        if cached:
            return cached["tenant_access_token"]
        _save_token(token, data.get("expire", 7200))
    return token


def get_auth_headers(token: str) -> dict:
    """Build authorization headers for API requests.

    Args:
        token: Valid tenant access token.

    Returns:
        Dictionary with Authorization header.
    """
    return {
        "Authorization": f"Bearer {token}",
    }


def _normalize_mobile(mobile: str) -> str:
    """Normalize a Chinese mobile number to +86 format.

    Args:
        mobile: Mobile number string (11 digits, or already with +86).

    Returns:
        Normalized mobile string in +86XXXXXXXXXXX format.
    """
    stripped = mobile.strip().replace(" ", "").replace("-", "")
    if stripped.startswith("+86"):
        return stripped
    if stripped.startswith("86") and len(stripped) == 13:
        return f"+{stripped}"
    if len(stripped) == 11 and stripped.isdigit():
        return f"+86{stripped}"
    return stripped


def lookup_user_by_mobile(
    token: str,
    mobile: str,
    user_id_type: str = "user_id",
) -> dict:
    """Look up a Feishu user by mobile number.

    Calls the contact batch_get_id API to resolve a mobile number
    to a user_id and open_id.

    Args:
        token: Valid tenant access token.
        mobile: Mobile number (11-digit CN or +86 prefixed).
        user_id_type: User ID type parameter (default "user_id").

    Returns:
        Dict with keys "user_id", "open_id", "mobile".

    Raises:
        RuntimeError: If the API returns a non-zero code.  For
            code 99991672 (missing scope), the error message
            includes the grant URL.
        LookupError: If no user is found for the mobile number.
    """
    normalized = _normalize_mobile(mobile)
    headers = get_auth_headers(token)
    headers["Content-Type"] = "application/json; charset=utf-8"

    params = {"user_id_type": user_id_type}
    payload = {"mobiles": [normalized]}

    resp = requests.post(
        CONTACT_LOOKUP_URL,
        headers=headers,
        params=params,
        json=payload,
    )
    data = resp.json()

    code = data.get("code", -1)
    msg = data.get("msg", "")

    if code == 99991672:
        # Extract grant URL from msg if present, else use hardcoded
        grant_url = _SCOPE_GRANT_URL
        url_match = re.search(r"https?://[^\s\"'<>]+", msg)
        if url_match:
            grant_url = url_match.group(0)
        raise RuntimeError(
            f"权限不足，需要 contact:user.id:readonly 权限。"
            f"请开通: {grant_url} "
            f"(code={code} msg={msg})"
        )

    if code != 0:
        raise RuntimeError(
            f"查询用户失败: code={code} msg={msg}"
        )

    user_list = data.get("data", {}).get("user_list", [])
    if not user_list:
        raise LookupError("未找到该手机号对应的员工")

    user = user_list[0]
    user_id = user.get("user_id", "")
    if not user_id:
        raise LookupError("未找到该手机号对应的员工")

    return {
        "user_id": user_id,
        "open_id": user.get("open_id", ""),
        "mobile": user.get("mobile", normalized),
    }


def upload_approval_file(token: str, file_path: str | Path) -> str:
    """Upload a file for use as an approval attachment.

    Uploads via the approval-specific file upload endpoint and
    returns the file code for use in the attachment widget.

    Args:
        token: Valid tenant access token.
        file_path: Path to the file to upload.

    Returns:
        File code string (data.code from the API response).

    Raises:
        RuntimeError: If the upload fails.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    headers = get_auth_headers(token)

    with open(file_path, "rb") as f:
        file_content = f.read()

    # Multipart form: name (filename), type, content
    form_data = {
        "name": (file_path.name, None),
        "type": (None, "attachment"),
    }
    files = {
        "content": (file_path.name, file_content),
    }

    resp = requests.post(
        APPROVAL_FILE_UPLOAD_URL,
        headers=headers,
        data={"name": file_path.name, "type": "attachment"},
        files=files,
    )
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"上传审批附件失败: code={data.get('code')} msg={data.get('msg')}"
        )

    file_code = data.get("data", {}).get("code", "")
    if not file_code:
        raise RuntimeError(
            f"上传审批附件失败: 未返回文件 code, response={data}"
        )

    return file_code


def get_approval_info(token: str, approval_code: str) -> dict:
    """Get approval form definition.

    Args:
        token: Valid tenant access token.
        approval_code: Approval definition code.

    Returns:
        Dict with "approval_name" (str) and "widgets" (list of
        widget dicts parsed from the form JSON string).

    Raises:
        RuntimeError: If the API returns a non-zero code.
    """
    headers = get_auth_headers(token)
    url = APPROVAL_INFO_URL.format(approval_code=approval_code)

    resp = requests.get(url, headers=headers)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"获取审批定义失败: code={data.get('code')} msg={data.get('msg')}"
        )

    result = data.get("data", {})
    approval_name = result.get("approval_name", "")
    form_str = result.get("form", "[]")

    try:
        widgets = json.loads(form_str)
        if not isinstance(widgets, list):
            widgets = []
    except (json.JSONDecodeError, TypeError):
        widgets = []

    return {
        "approval_name": approval_name,
        "widgets": widgets,
    }


def preview_instance(
    token: str,
    approval_code: str,
    user_id: str,
    form_widgets: list[dict],
) -> dict:
    """Preview an approval instance without creating it.

    Resolves the approver chain for the given form content.
    Note: preview does NOT validate form content, only resolves
    the approver chain.

    Args:
        token: Valid tenant access token.
        approval_code: Approval definition code.
        user_id: User ID of the instance initiator.
        form_widgets: List of form widget dicts.

    Returns:
        Dict with "nodes" key containing a list of approval node
        dicts, each with "node_name" and "user_id_list".

    Raises:
        RuntimeError: If the API returns a non-zero code.
    """
    headers = get_auth_headers(token)
    headers["Content-Type"] = "application/json; charset=utf-8"

    params = {"user_id_type": "user_id"}
    payload = {
        "approval_code": approval_code,
        "user_id": user_id,
        "form": json.dumps(form_widgets),
    }

    resp = requests.post(
        PREVIEW_URL,
        headers=headers,
        params=params,
        json=payload,
    )
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"预览审批流程失败: code={data.get('code')} msg={data.get('msg')}"
        )

    preview_nodes = data.get("data", {}).get("preview_nodes", [])
    nodes = []
    for node in preview_nodes:
        nodes.append(
            {
                "node_name": node.get("node_name", ""),
                "user_id_list": node.get("user_id_list", []),
            }
        )

    return {"nodes": nodes}


def create_instance(
    token: str,
    approval_code: str,
    user_id: str,
    form_widgets: list[dict],
    uuid: str | None = None,
) -> str:
    """Create an approval instance.

    Args:
        token: Valid tenant access token.
        approval_code: Approval definition code.
        user_id: User ID of the instance initiator.
        form_widgets: List of form widget dicts.
        uuid: Optional idempotency key.

    Returns:
        Instance code string.

    Raises:
        RuntimeError: If the API returns a non-zero code.
    """
    headers = get_auth_headers(token)
    headers["Content-Type"] = "application/json; charset=utf-8"

    params = {"user_id_type": "user_id"}
    payload: dict = {
        "approval_code": approval_code,
        "user_id": user_id,
        "form": json.dumps(form_widgets),
    }
    if uuid:
        payload["uuid"] = uuid

    resp = requests.post(
        INSTANCES_URL,
        headers=headers,
        params=params,
        json=payload,
    )
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"创建审批实例失败: code={data.get('code')} msg={data.get('msg')}"
        )

    instance_code = data.get("data", {}).get("instance_code", "")
    if not instance_code:
        raise RuntimeError(
            f"创建审批实例失败: 未返回 instance_code, response={data}"
        )

    return instance_code


def get_instance_detail(token: str, instance_code: str) -> dict:
    """Get detailed information about a specific approval instance.

    Args:
        token: Valid tenant access token.
        instance_code: Unique identifier for the instance.

    Returns:
        Instance detail data dictionary including instance_code,
        instance_url, status, user_id.

    Raises:
        RuntimeError: If the API returns a non-zero code.
    """
    headers = get_auth_headers(token)
    url = INSTANCE_DETAIL_URL.format(instance_code=instance_code)

    resp = requests.get(url, headers=headers)
    data = resp.json()

    if data.get("code") != 0:
        raise RuntimeError(
            f"获取实例详情失败: code={data.get('code')} msg={data.get('msg')}"
        )

    return data["data"]
