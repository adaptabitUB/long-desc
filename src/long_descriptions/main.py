"""Main orchestration script to run all generators in sequence."""

import time
from pathlib import Path

from . import (
    genera_matriu_cobertura,
    genera_instancies_canoniques,
    genera_resum_estadistic,
    genera_macro_vba_resum_estadistic,
    genera_metriques,
)


def main() -> None:
    """Run all generation scripts in sequence."""
    print("=" * 80)
    print("Starting Long Descriptions Generation Pipeline")
    print("=" * 80)
    print()
    
    start_time = time.time()
    
    # Step 1: Generate coverage matrix
    print("[1/5] Running genera_matriu_cobertura...")
    step_start = time.time()
    genera_matriu_cobertura.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 2: Generate canonical instances
    print("[2/5] Running genera_instancies_canoniques...")
    step_start = time.time()
    genera_instancies_canoniques.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 3: Generate statistical summary
    print("[3/5] Running genera_resum_estadistic...")
    step_start = time.time()
    genera_resum_estadistic.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Step 4: Generate VBA macro for statistical summary
    print("[4/5] Running genera_macro_vba_resum_estadistic...")
    step_start = time.time()
    genera_macro_vba_resum_estadistic.main()
    print(f"✓ Completed in {time.time() - step_start:.2f}s")
    print()
    
    # Feature Flag: Step 5 disabled - requires external data files
    # TODO: Enable when provider data files are available
    # Requires: long-descriptions/claude/claude_500_casos.xlsx (and other provider files)
    ENABLE_METRICS_GENERATION = False
    
    if ENABLE_METRICS_GENERATION:
        # Step 5: Generate metrics
        print("[5/5] Running genera_metriques...")
        step_start = time.time()
        genera_metriques.main()
        print(f"✓ Completed in {time.time() - step_start:.2f}s")
        print()
    else:
        print("[5/5] Skipping genera_metriques (feature disabled - missing provider data files)")
        print()
    
    total_time = time.time() - start_time
    print("=" * 80)
    print(f"Pipeline completed successfully in {total_time:.2f}s ({total_time/60:.2f}m)")
    print("=" * 80)


if __name__ == "__main__":
    main()
