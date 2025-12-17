"""Entry point for running migration_tool as a module."""

import sys
from pathlib import Path
from .tool import tool


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
            repo_root / "Template" / "VSME-Digital-Template-Sample-1.0.0.xlsx"
        )

    workbook, elapsed, issues = tool(file_path)

    print(f"Migration completed in {elapsed:.2f} seconds")
    if issues:
        print(f"Migration issues: {issues}")
    else:
        print("No migration issues found")


if __name__ == "__main__":
    main()
