"""Determinism tests: repeated migrations must produce identical results.

These tests run the migration multiple times per sample so are inherently
slow.  They are marked with ``@pytest.mark.slow`` and can be skipped in
fast CI runs with::

    pytest -m "not slow"
"""

from __future__ import annotations

from pathlib import Path
from typing import NamedTuple

import pytest
from conftest import migrated
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula

from migration_tool.tool import migrate_workbook

pytestmark = pytest.mark.slow


class TwoRunResult(NamedTuple):
    """Outcome of running migration twice on the same sample."""

    sample_path: Path
    wb_a: Workbook
    wb_b: Workbook
    issues_a: list[str]
    issues_b: list[str]


@pytest.fixture(scope="module")
def two_runs(sample_path) -> TwoRunResult:
    """The shared migration of the sample, and a second, independent run."""
    first = migrated(sample_path)
    wb_b, _, issues_b = migrate_workbook(sample_path)
    return TwoRunResult(
        sample_path=sample_path,
        wb_a=first.workbook,
        wb_b=wb_b,
        issues_a=first.issues,
        issues_b=issues_b,
    )


class TestDeterminism:
    """Migration must be deterministic: same input → same output."""

    def test_issues_are_identical(self, two_runs: TwoRunResult):
        assert two_runs.issues_a == two_runs.issues_b, (
            f"Non-deterministic issues for {two_runs.sample_path.name}\n"
            f"  Run 1: {two_runs.issues_a}\n"
            f"  Run 2: {two_runs.issues_b}"
        )

    def test_cell_values_are_identical(self, two_runs: TwoRunResult):
        """Every cell in every sheet must have the same value across runs."""
        wb_a = two_runs.wb_a
        wb_b = two_runs.wb_b
        name = two_runs.sample_path.name

        assert wb_a.sheetnames == wb_b.sheetnames, (
            f"Sheet names differ for {name}: {wb_a.sheetnames} vs {wb_b.sheetnames}"
        )

        for sheet_name in wb_a.sheetnames:
            ws_a = wb_a[sheet_name]
            ws_b = wb_b[sheet_name]

            assert ws_a.max_row == ws_b.max_row, (
                f"Row count differs in sheet '{sheet_name}' for {name}"
            )
            assert ws_a.max_column == ws_b.max_column, (
                f"Column count differs in sheet '{sheet_name}' for {name}"
            )

            # Whole rows first: far cheaper than a ws.cell() lookup per cell.
            rows_a = ws_a.iter_rows(values_only=True)
            rows_b = ws_b.iter_rows(values_only=True)
            for row_idx, (row_a, row_b) in enumerate(zip(rows_a, rows_b), start=1):
                vals_a = [_comparable(v) for v in row_a]
                vals_b = [_comparable(v) for v in row_b]
                if vals_a == vals_b:
                    continue
                col_idx, val_a, val_b = next(
                    (i, a, b)
                    for i, (a, b) in enumerate(zip(vals_a, vals_b), start=1)
                    if a != b
                )
                raise AssertionError(
                    f"Cell ({row_idx}, {col_idx}) differs in sheet "
                    f"'{sheet_name}' for {name}: {val_a!r} vs {val_b!r}"
                )


def _comparable(value: object) -> object:
    """Normalise openpyxl rich objects to plain comparable representations.

    Some cell values (e.g. ``ArrayFormula``) are rich objects that compare
    by identity.  Convert them to a tuple of their meaningful attributes so
    equality checks work on *content* rather than memory address.
    """
    if isinstance(value, ArrayFormula):
        return ("ArrayFormula", value.text, value.ref)
    # Add further cases here as they surface, e.g.:
    #   from openpyxl.worksheet.merge import MergedCell
    #   if isinstance(value, SomeOtherType): ...
    return value
