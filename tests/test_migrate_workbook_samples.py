from pathlib import Path
import sys

import pytest

from migration_tool.tool import migrate_workbook


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))


@pytest.fixture
def samples_dir():
    """Fixture providing the directory containing sample templates."""
    return REPO_ROOT / "tests" / "templates_samples"


@pytest.fixture
def results_dir():
    """Fixture providing the directory for migration results."""
    results_dir = REPO_ROOT / "tests" / "results"
    results_dir.mkdir(parents=True, exist_ok=True)
    return results_dir


@pytest.fixture
def sample_files(samples_dir):
    """Fixture providing sorted list of sample Excel files."""
    samples = sorted(samples_dir.glob("*.xlsx"))
    assert samples, f"No sample files found in {samples_dir}"
    return samples


@pytest.mark.parametrize("sample_path", sorted((REPO_ROOT / "tests" / "templates_samples").glob("*.xlsx")), ids=lambda p: p.name)
def test_migrate_sample_workbook(sample_path, results_dir):
    """Test migration of a single sample workbook."""
    new_wb, elapsed, issues = migrate_workbook(sample_path)
    output_path = results_dir / f"{sample_path.stem}-migrated.xlsx"
    new_wb.save(output_path)

    assert output_path.exists(), f"Output file was not created: {output_path}"
    assert elapsed >= 0.0, f"Elapsed time should be non-negative, got: {elapsed}"
    assert isinstance(issues, list), f"Issues should be a list, got: {type(issues)}"
