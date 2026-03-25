"""Shared fixtures and configuration for migration tests."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import NamedTuple

import pytest

# sys.path must be adjusted *before* importing project code.
REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from openpyxl import Workbook  # noqa: E402

from migration_tool.tool import migrate_workbook  # noqa: E402

SAMPLES_DIR = REPO_ROOT / "tests" / "templates_samples"
RESULTS_DIR = REPO_ROOT / "tests" / "results"
SAMPLE_PATHS = sorted(SAMPLES_DIR.glob("*.xlsx"))

# ---------------------------------------------------------------------------
# Map filename → reason for samples that are *expected* to fail today.
# Remove an entry once the corresponding fix lands.
# ---------------------------------------------------------------------------
XFAIL_SAMPLES: dict[str, str] = {
    # "VSME-Digital-Template-Sample-broken.xlsx": "Issue #42 – crash on missing sheet",
}

# ---------------------------------------------------------------------------
# Baseline of known migration issues per sample.
# If a sample is not listed here it is expected to produce zero issues.
# ---------------------------------------------------------------------------
EXPECTED_ISSUES: dict[str, list[str]] = {
    "VSME-Digital-Template-Sample-1.0.0.xlsx": [
        "Waste category --        101102 Non-Hazardous Waste - Waste glass-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
        "Waste category --        040201 Non-Hazardous Waste - Waste from unprocessed textile fibres and other natural fibrous substances mainly of vegetable origin-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
    ],
    "VSME-Digital-Template-Sample-1.0.1.xlsx": [
        "Waste category --        101102 Non-Hazardous Waste - Waste glass-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
        "Waste category --        040201 Non-Hazardous Waste - Waste from unprocessed textile fibres and other natural fibrous substances mainly of vegetable origin-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
    ],
    "VSME-Digital-Template-Sample-1.0.1_notcomplete.xlsx": [
        "The old workbook is incomplete. Migration happened only for the filled-out cells, but some data might be missing.",
        "Waste category --        101102 Non-Hazardous Waste - Waste glass-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
        "Waste category --        040201 Non-Hazardous Waste - Waste from unprocessed textile fibres and other natural fibrous substances mainly of vegetable origin-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955",
    ],
    "VSME-Digital-Template-Sample-1.1.0.xlsx": [
        "Name range 'MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives' refers to sheet '[1]General Information' which is not present in the old workbook. This name range has been ignored in the migration.",
    ],
    "VSME-Digital-Template-Sample-1.1.1.xlsx": [
        "The old workbook is incomplete. Migration happened only for the filled-out cells, but some data might be missing.",
    ],
}


class MigrationResult(NamedTuple):
    """Immutable container for the outcome of a single migration run."""

    workbook: Workbook
    elapsed: float
    issues: list[str]
    output_path: Path
    sample_path: Path


def _make_sample_param(path: Path):
    """Wrap a sample path in ``pytest.param``, applying xfail when configured."""
    name = path.name
    if name in XFAIL_SAMPLES:
        return pytest.param(
            path,
            id=name,
            marks=pytest.mark.xfail(reason=XFAIL_SAMPLES[name], strict=True),
        )
    return pytest.param(path, id=name)


SAMPLE_PARAMS = [_make_sample_param(p) for p in SAMPLE_PATHS]


@pytest.fixture(scope="session")
def results_dir() -> Path:
    """Provide (and lazily create) the directory for migration output files."""
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    return RESULTS_DIR


@pytest.fixture(params=SAMPLE_PARAMS, scope="module")
def sample_path(request) -> Path:
    """Parametrised fixture yielding each sample workbook path."""
    return request.param


@pytest.fixture(scope="module")
def migration_result(sample_path, results_dir) -> MigrationResult:
    """Run the migration for a single sample and persist the output workbook.

    Scoped to ``module`` so the (expensive) migration runs once per sample
    per test file rather than once per test method.
    """
    new_wb, elapsed, issues = migrate_workbook(sample_path)
    output_path = results_dir / f"{sample_path.stem}-migrated.xlsx"
    new_wb.save(output_path)
    return MigrationResult(
        workbook=new_wb,
        elapsed=elapsed,
        issues=issues,
        output_path=output_path,
        sample_path=sample_path,
    )


# ---------------------------------------------------------------------------
# Test file collection order
# ---------------------------------------------------------------------------

_FILE_ORDER = [
    "test_migration_error_handling.py",
    "test_migration_samples.py",
    "test_migration_determinism.py",
]


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    """Sort collected tests so error-handling runs first, then samples, then determinism."""
    original_order = {id(item): idx for idx, item in enumerate(items)}

    def _sort_key(item: pytest.Item) -> tuple[int, int]:
        filename = Path(item.fspath).name
        try:
            file_idx = _FILE_ORDER.index(filename)
        except ValueError:
            file_idx = len(_FILE_ORDER)  # unknown files run last
        return (file_idx, original_order[id(item)])

    items.sort(key=_sort_key)
