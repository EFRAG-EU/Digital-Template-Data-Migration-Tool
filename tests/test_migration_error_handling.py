"""Tests for migrate_workbook behaviour with invalid / edge-case inputs."""

from __future__ import annotations

import pytest
from conftest import SAMPLES_DIR

from migration_tool.tool import migrate_workbook


class TestMissingOrWrongPath:
    """Paths that don't point to a valid .xlsx file."""

    def test_nonexistent_file_raises_file_not_found(self):
        fake = SAMPLES_DIR / "this-file-does-not-exist.xlsx"
        with pytest.raises(FileNotFoundError):
            migrate_workbook(fake)

    def test_none_path_raises_type_error(self):
        with pytest.raises((TypeError, AttributeError)):
            migrate_workbook(None)  # type: ignore[arg-type]

    def test_directory_path_raises(self, tmp_path):
        """Passing a directory instead of a file should not silently succeed."""
        with pytest.raises(Exception):
            migrate_workbook(tmp_path)


class TestInvalidFileContent:
    """Files that exist but aren't valid workbooks."""

    def test_non_xlsx_file_raises(self, tmp_path):
        bad_file = tmp_path / "not_a_workbook.txt"
        bad_file.write_text("hello world")
        with pytest.raises(Exception):
            migrate_workbook(bad_file)

    def test_corrupted_xlsx_raises(self, tmp_path):
        bad_xlsx = tmp_path / "corrupted.xlsx"
        bad_xlsx.write_bytes(b"\x00\x01\x02\x03")
        with pytest.raises(Exception):
            migrate_workbook(bad_xlsx)

    def test_empty_file_raises(self, tmp_path):
        empty = tmp_path / "empty.xlsx"
        empty.write_bytes(b"")
        with pytest.raises(Exception):
            migrate_workbook(empty)
