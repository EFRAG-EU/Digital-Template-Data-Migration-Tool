"""Tests that exercise migrate_workbook against every sample template."""

from __future__ import annotations

import pytest
from conftest import EXPECTED_ISSUES, SAMPLE_PATHS

if not SAMPLE_PATHS:
    pytest.fail(
        "No sample .xlsx files found in tests/templates_samples/ — "
        "cannot run sample-based tests.",
        pytrace=False,
    )


class TestMigrationContract:
    """Every sample must satisfy these invariants after migration."""

    def test_output_file_is_created(self, migration_result):
        assert migration_result.output_path.exists()

    def test_output_file_is_non_empty(self, migration_result):
        assert migration_result.output_path.stat().st_size > 0

    def test_elapsed_time_is_non_negative(self, migration_result):
        assert migration_result.elapsed >= 0.0

    def test_workbook_has_at_least_one_sheet(self, migration_result):
        assert len(migration_result.workbook.sheetnames) >= 1

    def test_no_blank_sheet_names(self, migration_result):
        for name in migration_result.workbook.sheetnames:
            assert name and name.strip(), (
                f"Blank sheet name found in migrated workbook for "
                f"{migration_result.sample_path.name}"
            )

    def test_issues_is_a_list_of_strings(self, migration_result):
        issues = migration_result.issues
        assert isinstance(issues, list)
        for issue in issues:
            assert isinstance(issue, str)

    def test_issues_contain_no_empty_strings(self, migration_result):
        for issue in migration_result.issues:
            assert issue.strip(), (
                f"Empty/blank issue string found for "
                f"{migration_result.sample_path.name}"
            )


class TestExpectedIssues:
    """Verify that reported issues match the documented baselines."""

    def test_issues_match_baseline(self, migration_result):
        name = migration_result.sample_path.name
        issues = migration_result.issues
        expected = EXPECTED_ISSUES.get(name, None)

        if expected is not None:
            assert issues == expected, (
                f"Issues mismatch for {name}\n"
                f"  Expected ({len(expected)}): {expected}\n"
                f"  Actual   ({len(issues)}):   {issues}\n\n"
                f"If the new issues are correct, update EXPECTED_ISSUES in conftest.py:\n"
                f'    "{name}": {issues!r},'
            )
        elif issues:
            pytest.fail(
                f"Unexpected issues found for {name}:\n"
                f"  {issues}\n\n"
                f"If these issues are expected, add this entry to EXPECTED_ISSUES "
                f"in conftest.py:\n"
                f'    "{name}": {issues!r},'
            )

    def test_no_duplicate_issues(self, migration_result):
        issues = migration_result.issues
        seen: set[str] = set()
        duplicates: list[str] = []
        for msg in issues:
            if msg in seen:
                duplicates.append(msg)
            seen.add(msg)
        assert not duplicates, (
            f"Duplicate issues for {migration_result.sample_path.name}: {duplicates}"
        )
