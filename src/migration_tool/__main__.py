"""Entry point for running migration_tool as a module.
Command line usage: python -m migration_tool <path_to_excel_file>
If no path is provided, it defaults to a sample template in the tests directory (see line 27)."""

import sys
from pathlib import Path

from .tool import migrate_workbook


def main():
    # Determine repo root (parent of src/ when running from src/)
    cwd = Path.cwd()
    if cwd.name == "src":
        repo_root = cwd.parent
    else:
        repo_root = cwd

    # Accept command line argument or use default
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = str(
            repo_root
            / "tests"
            / "templates_samples"
            / "VSME-Digital-Template-Sample-1.0.0.xlsx"
        )

    workbook, elapsed, issues = migrate_workbook(file_path)

    # Save the migrated workbook
    output_path = Path(file_path).stem + "_migrated.xlsx"
    workbook.save(output_path)

    print(f"Migration completed in {elapsed:.2f} seconds")
    print(f"Migrated workbook saved to: {output_path}")
    if issues:
        print(f"Migration issues: {issues}")
    else:
        print("No migration issues found")


if __name__ == "__main__":
    main()
