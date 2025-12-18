"""
Lightweight runner to execute `migration_tool.tool()` across all sample templates.

Usage examples (from repo root):

  - Run on all samples and save outputs:
      python -m migration_tool.run_samples

  - Custom sample/output directories:
      python -m migration_tool.run_samples --samples-dir tests/templates_samples --output-dir tests/results

  - Filter by pattern and fail fast on first error:
      python -m migration_tool.run_samples -p "*1.0.0*.xlsx" --fail-fast

This script avoids test frameworks and provides timing + issue summaries.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import List

from .tool import tool


@dataclass
class SampleResult:
    sample: Path
    output: Path | None
    elapsed: float | None
    issues: List[str] | None
    error: str | None


def resolve_repo_root() -> Path:
    # src/migration_tool/run_samples.py -> repo_root is parent of src
    return Path(__file__).resolve().parents[2]


def run_sample(sample_path: Path, output_dir: Path | None) -> SampleResult:
    try:
        new_wb, elapsed, issues = tool(str(sample_path))
        output_path = None
        if output_dir is not None:
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f"{sample_path.stem}-migrated.xlsx"
            new_wb.save(output_path)
        return SampleResult(
            sample=sample_path,
            output=output_path,
            elapsed=elapsed,
            issues=issues,
            error=None,
        )
    except Exception as exc:  # keep it simple; report error and continue
        return SampleResult(
            sample=sample_path, output=None, elapsed=None, issues=None, error=str(exc)
        )


def find_samples(samples_dir: Path, pattern: str) -> List[Path]:
    return sorted(samples_dir.glob(pattern))


def main(argv: List[str] | None = None) -> int:
    repo_root = resolve_repo_root()

    parser = argparse.ArgumentParser(
        description="Run migration_tool over sample templates"
    )
    parser.add_argument(
        "--samples-dir",
        type=Path,
        default=repo_root / "tests" / "templates_samples",
        help="Directory containing input sample .xlsx files",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=repo_root / "tests" / "results",
        help="Directory to write migrated workbooks (.xlsx)",
    )
    parser.add_argument(
        "-p",
        "--pattern",
        default="*.xlsx",
        help="Glob pattern to select samples (default: *.xlsx)",
    )
    parser.add_argument(
        "--fail-fast",
        action="store_true",
        help="Stop on first failure/error",
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Do not write migrated workbooks to disk",
    )

    args = parser.parse_args(argv)

    samples_dir: Path = args.samples_dir
    output_dir: Path | None = None if args.no_save else args.output_dir
    pattern: str = args.pattern

    if not samples_dir.exists():
        print(f"Samples directory not found: {samples_dir}")
        return 2

    samples = find_samples(samples_dir, pattern)
    if not samples:
        print(f"No samples match pattern '{pattern}' in {samples_dir}")
        return 0

    print(f"Running migration on {len(samples)} sample(s) from {samples_dir}...")

    results: List[SampleResult] = []
    for sample in samples:
        print(f"- {sample.name} ...", end=" ")
        res = run_sample(sample, output_dir)
        results.append(res)
        if res.error:
            print(f"ERROR: {res.error}")
            if args.fail_fast:
                break
        else:
            issue_count = len(res.issues or [])
            print(f"OK in {res.elapsed:.2f}s | issues: {issue_count}")

    # Summary
    ok = [r for r in results if not r.error]
    failed = [r for r in results if r.error]
    total_issues = sum(len(r.issues or []) for r in ok)

    print("\nSummary:")
    print(f"  Succeeded: {len(ok)}")
    print(f"  Failed:   {len(failed)}")
    print(f"  Issues:   {total_issues}")

    if failed:
        print("\nFailures:")
        for r in failed:
            print(f"  - {r.sample.name}: {r.error}")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
