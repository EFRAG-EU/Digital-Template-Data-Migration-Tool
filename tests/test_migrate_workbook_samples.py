import sys
from pathlib import Path

import pytest

from migration_tool.tool import migrate_workbook

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))


# Expected issues for each sample file
# Update this dict as you discover expected issues for each template
EXPECTED_ISSUES: dict[str, list[str]] = {
    # "sample-file-name.xlsx": ["Expected issue 1", "Expected issue 2"],
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
        "Name range 'MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives' refers to sheet '[1]General Information' which is not present in the old workbook. This name range has been ignored in the migration."
    ],
    "VSME-Digital-Template-Sample-1.1.1.xlsx": [
        "The old workbook is incomplete. Migration happened only for the filled-out cells, but some data might be missing."
    ],
}


@pytest.fixture
def results_dir():
    """Fixture providing the directory for migration results."""
    results_dir = REPO_ROOT / "tests" / "results"
    results_dir.mkdir(parents=True, exist_ok=True)
    return results_dir


@pytest.mark.parametrize(
    "sample_path",
    sorted((REPO_ROOT / "tests" / "templates_samples").glob("*.xlsx")),
    ids=lambda p: p.name,
)
def test_migrate_sample_workbook(sample_path, results_dir):
    """Test migration of a single sample workbook."""
    new_wb, elapsed, issues = migrate_workbook(sample_path)
    output_path = results_dir / f"{sample_path.stem}-migrated.xlsx"
    new_wb.save(output_path)

    # Basic assertions
    assert output_path.exists(), f"Output file was not created: {output_path}"
    assert elapsed >= 0.0, f"Elapsed time should be non-negative, got: {elapsed}"
    assert isinstance(issues, list), f"Issues should be a list, got: {type(issues)}"

    # Validate issues against expected baseline
    if sample_path.name in EXPECTED_ISSUES:
        # File has documented expected issues - verify they match exactly
        expected = EXPECTED_ISSUES[sample_path.name]
        assert issues == expected, (
            f"Issues mismatch for {sample_path.name}\n"
            f"Expected: {expected}\n"
            f"Actual:   {issues}"
        )
    elif issues:
        # File has issues but they're not documented - this should fail
        pytest.fail(
            f"Unexpected issues found in {sample_path.name}:\n"
            f"  {issues}\n\n"
            f"If these issues are expected, add this entry to EXPECTED_ISSUES:\n"
            f"  '{sample_path.name}': {issues!r}"
        )
