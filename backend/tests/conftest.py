import os
from pathlib import Path

import pytest

TEST_DB = Path(__file__).resolve().parents[1] / "test_wordreplace.db"
os.environ.setdefault("DATABASE_URL", f"sqlite:///{TEST_DB}")
os.environ.setdefault("EXPORT_TOKEN_SECRET", "test-secret")
os.environ.setdefault("MAX_CONCURRENT_RUNS", "1")


@pytest.fixture(scope="session", autouse=True)
def cleanup_test_database():
    if TEST_DB.exists():
        TEST_DB.unlink()
    yield
    if TEST_DB.exists():
        TEST_DB.unlink()
