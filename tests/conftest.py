"""Shared pytest fixtures.

Hard rule (see CLAUDE.md): never write to data/Data.xlsx during tests. Every
test that writes goes through the test_excel fixture, which works on a throwaway
copy and a throwaway backup directory, both removed after the test.
"""

from __future__ import annotations

import os
import shutil
from typing import Iterator

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE_XLSX = os.path.join(REPO_ROOT, "data", "Data.xlsx")
TEST_COPY = os.path.join(REPO_ROOT, "data", "_test_copy.xlsx")
TEST_BACKUP_DIR = os.path.join(REPO_ROOT, "data", "_test_backup")


def _cleanup() -> None:
    if os.path.exists(TEST_COPY):
        os.remove(TEST_COPY)
    if os.path.exists(TEST_BACKUP_DIR):
        shutil.rmtree(TEST_BACKUP_DIR, ignore_errors=True)


@pytest.fixture
def test_excel() -> Iterator[str]:
    """Path to a disposable copy of data/Data.xlsx. Never the real file."""
    _cleanup()
    shutil.copy2(SOURCE_XLSX, TEST_COPY)
    try:
        yield TEST_COPY
    finally:
        _cleanup()


@pytest.fixture
def test_backup_dir() -> str:
    """Throwaway backup_dir for writer functions, matching the test_excel fixture."""
    return TEST_BACKUP_DIR
