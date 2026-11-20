"""Female and male employees at management level are migrated by label, not as one
block: VSME 1.x lists male first and VS female first, and the block put female
over the ratio formula below them."""

from __future__ import annotations

from importlib.resources import as_file, files
from pathlib import Path

import pytest
from conftest import SAMPLE_PARAMS, migrated
from openpyxl import Workbook

from migration_tool.data.validations import VersionCollection
from migration_tool.outils import get_version, resolveAnchors
from migration_tool.tool import (
    NEW_TEMPLATE_NAME,
    IssuesCollector,
    _missing_NR,
    load_workbook_quietly,
)

COUNTS = ("FemaleEmployeesAtManagementLevel", "MaleEmployeesAtManagementLevel")
RATIO = "FemaleToMaleRatioAtManagementLevelForTheReportingPeriod"


def counts(wb: Workbook) -> dict[str, object]:
    """The management-level counts in wb, found as its own version describes them."""
    version = get_version(wb, IssuesCollector())
    refs = VersionCollection.from_dict(_missing_NR())
    resolveAnchors(refs, version, wb)
    return {
        ref.label: wb[ref.sheet][ref.range.split(":")[0].replace("$", "")].value
        for ref in refs.getRefs(version)
        if ref.label in COUNTS
    }


def ratioFormula(wb: Workbook) -> object:
    [(sheet, cell)] = wb.defined_names[RATIO].destinations
    return wb[sheet][cell.replace("$", "")].value


@pytest.mark.parametrize("sample_path", SAMPLE_PARAMS)
def test_counts_land_by_gender_and_keep_the_ratio_formula(sample_path: Path):
    old = load_workbook_quietly(sample_path)
    new = migrated(sample_path).workbook
    with as_file(files("migration_tool.data").joinpath(NEW_TEMPLATE_NAME)) as path:
        template = load_workbook_quietly(path)

    assert counts(new) == counts(old)
    assert ratioFormula(new) == ratioFormula(template)
