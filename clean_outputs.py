#!/usr/bin/env python3
"""Clean all generated pipeline output files."""

import shutil
from pathlib import Path


def clean_outputs() -> None:
    """Remove all generated Excel, CSV, JSON, VBA, and other output files."""
    workspace = Path.cwd()
    
    output_dir = workspace / "output"

    # Explicit generated filenames at the root plus legacy compatibility patterns.
    patterns = [
        "coverage_matrix_*.xlsx",
        "matrix_*.csv",
        "charts.json",
        "charts.xlsx",
        "manifest.json",
        "statistics_summary.json",
        "statistics_summary_*.csv",
        "StatisticsSummaryMacro.bas",
        "matrius_cobertura_excel_*.xlsx",
        "matriu_*.csv",
    ]

    statistics_dir = output_dir / "statistics"
    
    deleted_files = []
    deleted_dirs = []
    
    # Delete files matching patterns in output/ (and root for backwards compatibility)
    for pattern in patterns:
        for file_path in output_dir.glob(pattern):
            if file_path.is_file():
                file_path.unlink()
                deleted_files.append(str(file_path.relative_to(workspace)))
        for file_path in workspace.glob(pattern):
            if file_path.is_file():
                file_path.unlink()
                deleted_files.append(file_path.name)

    if statistics_dir.exists() and statistics_dir.is_dir():
        for file_path in statistics_dir.glob("statistics_summary_*.csv"):
            if file_path.is_file():
                file_path.unlink()
                deleted_files.append(str(file_path.relative_to(workspace)))
    
    # Delete the entire output directory (all pipeline steps)
    if output_dir.exists() and output_dir.is_dir():
        shutil.rmtree(output_dir)
        deleted_dirs.append(output_dir.name)

    # Legacy output directory from previous versions
    sortida_dir = workspace / "sortida_instancies_completa"
    if sortida_dir.exists() and sortida_dir.is_dir():
        shutil.rmtree(sortida_dir)
        deleted_dirs.append(sortida_dir.name)
    
    # Print summary
    if deleted_files or deleted_dirs:
        total_count = len(deleted_files) + len(deleted_dirs)
        print(f"Deleted {total_count} item(s):")
        for filename in sorted(deleted_files):
            print(f"  - {filename}")
        for dirname in sorted(deleted_dirs):
            print(f"  - {dirname}/ (directory)")
    else:
        print("No output files found to delete.")


if __name__ == "__main__":
    clean_outputs()
