"""
Tests for Feishu API module.
"""

import json
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from app.feishu_api import (
    APPROVAL_FILE_UPLOAD_URL,
    BASE_URL,
    CACHE_FILE,
    CONTACT_LOOKUP_URL,
    INSTANCE_DETAIL_URL,
    INSTANCES_URL,
    PREVIEW_URL,
    TOKEN_URL,
    USER_INFO_URL,
    create_instance,
    get_approval_info,
    get_auth_headers,
    get_instance_detail,
    get_tenant_token,
    get_user_name,
    lookup_user_by_mobile,
    preview_instance,
    upload_approval_file,
)


class TestGetAuthHeaders:
    """Test suite for get_auth_headers."""

    def test_authorization_header(self, mock_token: str):
        """Test Authorization header format."""
        headers = get_auth_headers(mock_token)
        assert headers["Authorization"] == f"Bearer {mock_token}"

    def test_base_url_is_set(self):
        """Test that BASE_URL is properly configured."""
        assert BASE_URL.startswith("https://")

    def test_token_url_format(self):
        """Test that TOKEN_URL follows expected format."""
        assert "auth/v3/tenant_access_token" in TOKEN_URL


class TestGetTenantToken:
    """Test suite for get_tenant_token with caching."""

    def test_token_success(self, temp_cache_file):
        """Test fetching tenant token from API."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "tenant_access_token": "new_token_123",
            "expire": 7200,
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "new_token_123"
            mock_post.assert_called_once_with(
                TOKEN_URL, json={"app_id": "app_id", "app_secret": "app_secret"}
            )
            assert temp_cache_file.exists()

    def test_token_uses_cache(self, temp_cache_file):
        """Test that cached token is returned without HTTP request."""
        cache_data = {
            "tenant_access_token": "cached_token_456",
            "expire_at": time.time() + 3600,
        }
        temp_cache_file.write_text(json.dumps(cache_data))

        with patch("app.feishu_api.requests.post") as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "cached_token_456"
            mock_post.assert_not_called()

    def test_token_expired_refetches(self, temp_cache_file):
        """Test that expired cache triggers a new HTTP request."""
        cache_data = {
            "tenant_access_token": "old_token",
            "expire_at": time.time() - 100,
        }
        temp_cache_file.write_text(json.dumps(cache_data))

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "tenant_access_token": "new_token_789",
            "expire": 7200,
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            token = get_tenant_token("app_id", "app_secret")
            assert token == "new_token_789"
            mock_post.assert_called_once()

    def test_token_auth_failure(self, temp_cache_file):
        """Test that auth failure raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 10003,
            "msg": "Invalid app_id or app_secret",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="获取 token 失败"):
                get_tenant_token("bad_id", "bad_secret")


class TestLookupUserByMobile:
    """Test suite for lookup_user_by_mobile."""

    def test_lookup_found(self, mock_token: str):
        """Test successful user lookup returns open_id and mobile."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "user_list": [
                    {
                        "mobile": "+8613800138000",
                        "user_id": "ou_abc123",
                    }
                ]
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = lookup_user_by_mobile(mock_token, "13800138000")
            assert result["open_id"] == "ou_abc123"
            assert result["mobile"] == "+8613800138000"
            assert "user_id" not in result

            call_args = mock_post.call_args
            assert call_args.kwargs["json"] == {
                "mobiles": ["+8613800138000"],
                "include_resigned": False,
            }
            assert call_args.kwargs["params"] == {"user_id_type": "open_id"}

    def test_lookup_not_found(self, mock_token: str):
        """Test user not found raises LookupError after retry."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {"user_list": []},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(LookupError, match="未找到该手机号对应的员工"):
                lookup_user_by_mobile(mock_token, "13800138000")

    def test_lookup_empty_user_id(self, mock_token: str):
        """Test empty user_id raises LookupError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "user_list": [
                    {"mobile": "+8613800138000", "user_id": ""}
                ]
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(LookupError, match="未找到该手机号对应的员工"):
                lookup_user_by_mobile(mock_token, "13800138000")

    def test_lookup_scope_error_99991672(self, mock_token: str):
        """Test scope error 99991672 raises RuntimeError with grant URL."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 99991672,
            "msg": "Access denied. One of the following scopes is required: "
            "[contact:user.id:readonly]...点击链接申请并开通任一权限即可："
            "https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth"
            "?q=contact:user.id:readonly&op_from=openapi&token_type=tenant",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="权限不足"):
                lookup_user_by_mobile(mock_token, "13800138000")

    def test_lookup_scope_error_contains_url(self, mock_token: str):
        """Test that scope error message contains scope name and grant URL."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 99991672,
            "msg": "Access denied. One of the following scopes is required: "
            "[contact:user.employee_id:readonly]",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError) as exc_info:
                lookup_user_by_mobile(mock_token, "13800138000")
            error_str = str(exc_info.value)
            assert "权限不足" in error_str
            assert "contact:user.employee_id:readonly" in error_str
            assert "open.feishu.cn/app/" in error_str

    def test_lookup_api_error(self, mock_token: str):
        """Test generic API error raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "some error",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="查询用户失败"):
                lookup_user_by_mobile(mock_token, "13800138000")

    def test_lookup_mobile_normalization(self, mock_token: str):
        """Test that 11-digit mobile is normalized to +86 format."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {
                "user_list": [
                    {"mobile": "+8613800138000", "user_id": "ou_uid"}
                ]
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            lookup_user_by_mobile(mock_token, "13800138000")
            payload = mock_post.call_args.kwargs["json"]
            assert payload["mobiles"] == ["+8613800138000"]

    def test_lookup_retry_with_raw_mobile(self, mock_token: str):
        """Test that +86-normalized form retries with raw input if not found."""
        call_count = 0

        def side_effect(*args, **kwargs):
            nonlocal call_count
            call_count += 1
            mock_resp = MagicMock()
            if call_count == 1:
                # First call with +86 prefix: not found
                mock_resp.json.return_value = {
                    "code": 0,
                    "msg": "success",
                    "data": {"user_list": []},
                }
            else:
                # Second call with raw input: found
                mock_resp.json.return_value = {
                    "code": 0,
                    "msg": "success",
                    "data": {
                        "user_list": [
                            {"mobile": "13800138000", "user_id": "ou_raw"}
                        ]
                    },
                }
            return mock_resp

        with patch("app.feishu_api.requests.post", side_effect=side_effect):
            result = lookup_user_by_mobile(mock_token, "13800138000")
            assert result["open_id"] == "ou_raw"
            assert call_count == 2


class TestGetUserName:
    """Test suite for get_user_name."""

    def test_get_user_name_success(self, mock_token: str):
        """Test successful name lookup returns the name."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {"user": {"name": "张三"}},
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response) as mock_get:
            result = get_user_name(mock_token, "ou_abc123")
            assert result == "张三"
            mock_get.assert_called_once()
            call_args = mock_get.call_args
            assert "ou_abc123" in call_args.args[0]
            assert call_args.kwargs["params"] == {"user_id_type": "open_id"}

    def test_get_user_name_scope_error(self, mock_token: str):
        """Test scope error returns None (graceful fallback)."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 99991672,
            "msg": "Access denied. One of the following scopes is required: "
            "[contact:contact.base:readonly]",
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            result = get_user_name(mock_token, "ou_abc123")
            assert result is None

    def test_get_user_name_empty_name(self, mock_token: str):
        """Test empty name returns None."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "msg": "success",
            "data": {"user": {"name": ""}},
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            result = get_user_name(mock_token, "ou_abc123")
            assert result is None

    def test_get_user_name_api_error(self, mock_token: str):
        """Test generic API error returns None."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "not found",
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            result = get_user_name(mock_token, "ou_abc123")
            assert result is None

    def test_get_user_name_network_error(self, mock_token: str):
        """Test network error returns None."""
        with patch("app.feishu_api.requests.get", side_effect=Exception("network error")):
            result = get_user_name(mock_token, "ou_abc123")
            assert result is None


class TestUploadApprovalFile:
    """Test suite for upload_approval_file."""

    def test_upload_success(self, mock_token: str, tmp_path: Path):
        """Test successful file upload returns file code."""
        test_file = tmp_path / "test.xls"
        test_file.write_bytes(b"fake excel content")

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {
                "code": "file_code_xyz",
                "url": "https://example.com/file",
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = upload_approval_file(mock_token, test_file)
            assert result == "file_code_xyz"
            mock_post.assert_called_once()

    def test_upload_api_error(self, mock_token: str, tmp_path: Path):
        """Test upload API error raises RuntimeError."""
        test_file = tmp_path / "test.xls"
        test_file.write_bytes(b"fake content")

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "upload failed",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="上传审批附件失败"):
                upload_approval_file(mock_token, test_file)

    def test_upload_file_not_found(self, mock_token: str, tmp_path: Path):
        """Test missing file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            upload_approval_file(mock_token, tmp_path / "nonexistent.xls")

    def test_upload_no_code_in_response(self, mock_token: str, tmp_path: Path):
        """Test missing file code in response raises RuntimeError."""
        test_file = tmp_path / "test.xls"
        test_file.write_bytes(b"fake content")

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {"url": "https://example.com/file"},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="未返回文件 code"):
                upload_approval_file(mock_token, test_file)

    def test_upload_uses_override_filename(
        self, mock_token: str, tmp_path: Path
    ):
        """Test override filename is used in both data and files parts."""
        test_file = tmp_path / "tmpXXXX.xls"
        test_file.write_bytes(b"fake excel content")

        override_name = "科研部发放吉林大学2024年科研业绩津贴明细表.xls"
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {"code": "file_code_override"},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = upload_approval_file(mock_token, test_file, filename=override_name)
            assert result == "file_code_override"
            mock_post.assert_called_once()
            call_kwargs = mock_post.call_args.kwargs
            assert call_kwargs["data"] == {"name": override_name, "type": "attachment"}
            assert call_kwargs["files"]["content"][0] == override_name

    def test_upload_defaults_to_file_name(
        self, mock_token: str, tmp_path: Path
    ):
        """Test default filename falls back to file_path.name."""
        test_file = tmp_path / "test.xls"
        test_file.write_bytes(b"fake excel content")

        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {"code": "file_code_default"},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = upload_approval_file(mock_token, test_file)
            assert result == "file_code_default"
            call_kwargs = mock_post.call_args.kwargs
            assert call_kwargs["data"] == {"name": "test.xls", "type": "attachment"}
            assert call_kwargs["files"]["content"][0] == "test.xls"


class TestGetApprovalInfo:
    """Test suite for get_approval_info."""

    def test_get_info_success(self, mock_token: str):
        """Test successful approval info retrieval."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {
                "approval_name": "派遣员工向甲方单位退返餐费审批",
                "form": '[{"id":"w1","type":"input"}]',
            },
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response) as mock_get:
            result = get_approval_info(mock_token, "APPROVAL_CODE_123")
            assert result["approval_name"] == "派遣员工向甲方单位退返餐费审批"
            assert len(result["widgets"]) == 1
            mock_get.assert_called_once()

    def test_get_info_api_error(self, mock_token: str):
        """Test API error raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "not found",
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            with pytest.raises(RuntimeError, match="获取审批定义失败"):
                get_approval_info(mock_token, "BAD_CODE")

    def test_get_info_invalid_form_json(self, mock_token: str):
        """Test invalid form JSON returns empty widgets list."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {
                "approval_name": "Test",
                "form": "not valid json",
            },
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            result = get_approval_info(mock_token, "CODE")
            assert result["widgets"] == []


class TestPreviewInstance:
    """Test suite for preview_instance."""

    def test_preview_success(self, mock_token: str):
        """Test successful preview returns node list."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {
                "preview_nodes": [
                    {"node_name": "提交", "user_id_list": ["user1"]},
                    {"node_name": "业务审核", "user_id_list": ["user2"]},
                    {"node_name": "财务审核", "user_id_list": ["user3"]},
                ]
            },
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = preview_instance(
                mock_token, "APPROVAL_CODE", "ou_open123", [{"id": "w1"}]
            )
            assert len(result["nodes"]) == 3
            assert result["nodes"][0]["node_name"] == "提交"
            assert result["nodes"][1]["node_name"] == "业务审核"

            payload = mock_post.call_args.kwargs["json"]
            assert isinstance(payload["form"], str)
            assert json.loads(payload["form"]) == [{"id": "w1"}]
            assert payload["user_id"] == "ou_open123"
            assert "open_id" not in payload

            params = mock_post.call_args.kwargs["params"]
            assert params == {"user_id_type": "open_id"}

    def test_preview_api_error(self, mock_token: str):
        """Test API error raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "preview failed",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="预览审批流程失败"):
                preview_instance(mock_token, "CODE", "ou_user1", [])


class TestCreateInstance:
    """Test suite for create_instance."""

    def test_create_success(self, mock_token: str):
        """Test successful instance creation returns instance code."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {"instance_code": "INST_001"},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            result = create_instance(
                mock_token, "APPROVAL_CODE", "ou_user1", [{"id": "w1"}]
            )
            assert result == "INST_001"

            payload = mock_post.call_args.kwargs["json"]
            assert payload["approval_code"] == "APPROVAL_CODE"
            assert payload["open_id"] == "ou_user1"
            assert "user_id" not in payload
            assert isinstance(payload["form"], str)
            assert "uuid" in payload

    def test_create_with_uuid(self, mock_token: str):
        """Test idempotency key is included when provided."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {"instance_code": "INST_002"},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response) as mock_post:
            create_instance(
                mock_token, "CODE", "ou_user1", [], uuid="my-uuid-123"
            )
            payload = mock_post.call_args.kwargs["json"]
            assert payload["uuid"] == "my-uuid-123"

    def test_create_api_error(self, mock_token: str):
        """Test API error raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "create failed",
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="创建审批实例失败"):
                create_instance(mock_token, "CODE", "ou_user1", [])

    def test_create_no_instance_code(self, mock_token: str):
        """Test missing instance_code in response raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {},
        }
        with patch("app.feishu_api.requests.post", return_value=mock_response):
            with pytest.raises(RuntimeError, match="未返回 instance_code"):
                create_instance(mock_token, "CODE", "ou_user1", [])


class TestGetInstanceDetail:
    """Test suite for get_instance_detail."""

    def test_detail_success(self, mock_token: str):
        """Test successful detail retrieval."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 0,
            "data": {
                "instance_code": "INST_001",
                "instance_url": "https://example.com/inst",
                "status": "APPROVED",
                "user_id": "user1",
            },
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response) as mock_get:
            result = get_instance_detail(mock_token, "INST_001")
            assert result["instance_code"] == "INST_001"
            assert result["status"] == "APPROVED"
            mock_get.assert_called_once()

    def test_detail_api_error(self, mock_token: str):
        """Test API error raises RuntimeError."""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "code": 9999,
            "msg": "not found",
        }
        with patch("app.feishu_api.requests.get", return_value=mock_response):
            with pytest.raises(RuntimeError, match="获取实例详情失败"):
                get_instance_detail(mock_token, "BAD_CODE")
