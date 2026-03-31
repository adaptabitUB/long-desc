"""Main orchestration script to run all generators in sequence."""

import time
from pathlib import Path

from . import (
    generate_coverage_matrix,
    generate_charts,
    generate_statistics_summary,
    generate_macro_vba_statistics_summary,
)


def main() -> None:
    """Run all generation scripts in sequence."""
    print("=" * 80)
    print("Starting Long Descriptions Generation Pipeline")
    print("=" * 80)
    print()
    
    start_time = time.time()
    
    # Step 1: Generate coverage matrix
    print("[1/5] Running generate_coverage_matrix...")
    step_start = time.time()
    generate_coverage_matrix.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 2: Generate canonical instances
    print("[2/5] Running generate_charts...")
    step_start = time.time()
    generate_charts.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 3: Generate statistical summary
    print("[3/5] Running generate_statistics_summary...")
    step_start = time.time()
    generate_statistics_summary.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 4: Generate VBA macro for statistical summary
    print("[4/5] Running generate_macro_vba_statistics_summary...")
    step_start = time.time()
    generate_macro_vba_statistics_summary.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    total_time = time.time() - start_time
    print("=" * 80)
    print(f"Pipeline completed successfully in {total_time:.2f}s ({total_time/60:.2f}m)")
    print("=" * 80)


if __name__ == "__main__":
    main()
