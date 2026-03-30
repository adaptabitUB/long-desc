#!/usr/bin/env python3
"""Clean all generated output files (Excel and CSV)."""

import shutil
from pathlib import Path


def clean_outputs() -> None:
    """Remove all generated Excel, CSV, JSON, and other output files."""
    workspace = Path.cwd()
    
    # File patterns for generated files (Step 1: genera_matriu_cobertura)
    patterns = [
        "matrius_cobertura_excel_*.xlsx",
        "matriu_*.csv",
    ]
    
    deleted_files = []
    deleted_dirs = []
    
    # Delete files matching patterns
    for pattern in patterns:
        for file_path in workspace.glob(pattern):
            if file_path.is_file():
                file_path.unlink()
                deleted_files.append(file_path.name)
    
    # Delete the entire output directory (Steps 2-4: genera_instancies_canoniques, 
    # genera_resum_estadistic, genera_macro_vba_resum_estadistic)
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
