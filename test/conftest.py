"""
Pytest configuration and shared fixtures for tests.
"""

from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def mock_token() -> str:
    """Provide a mock tenant access token."""
    return "test_tenant_access_token_12345"


@pytest.fixture
def mock_auth_headers(mock_token: str) -> dict:
    """Provide mock authorization headers."""
    return {
        "Authorization": f"Bearer {mock_token}",
    }


@pytest.fixture
def real_fixture_path() -> Path:
    """Provide path to the real template fixture file."""
    path = FIXTURES_DIR / "返餐费2026.xls"
    if not path.exists():
        pytest.skip("测试模板未提供（含个人信息，不随仓库分发），已跳过")
    return path


@pytest.fixture
def temp_cache_file(monkeypatch, tmp_path):
    """Provide a temporary cache file path for token tests."""
    cache = tmp_path / ".token_cache.json"
    monkeypatch.setattr("app.feishu_api.CACHE_FILE", cache)
    return cache
