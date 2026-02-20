from pathlib import Path
import sys
import unittest

from migration_tool.tool import migrate_workbook


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))


class TestMigrateWorkbookSamples(unittest.TestCase):
    def test_migrate_all_samples_and_write_results(self):
        samples_dir = REPO_ROOT / "tests" / "templates_samples"
        results_dir = REPO_ROOT / "tests" / "results"
        results_dir.mkdir(parents=True, exist_ok=True)

        samples = sorted(samples_dir.glob("*.xlsx"))
        self.assertTrue(samples, f"No sample files found in {samples_dir}")

        for sample_path in samples:
            with self.subTest(sample=sample_path.name):
                new_wb, elapsed, issues = migrate_workbook(sample_path)
                output_path = results_dir / f"{sample_path.stem}-migrated.xlsx"
                new_wb.save(output_path)

                self.assertTrue(output_path.exists())
                self.assertGreaterEqual(elapsed, 0.0)
                self.assertIsInstance(issues, list)


if __name__ == "__main__":
    unittest.main(verbosity=2)
