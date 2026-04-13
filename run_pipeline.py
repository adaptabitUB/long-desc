#!/usr/bin/env python3
"""
Pipeline runner for Long Descriptions generation.

This script runs the complete pipeline:
1. generate_coverage_matrix - Generate coverage matrix with chart types
2. generate_charts - Generate canonical instances
3. generate_statistics_summary - Generate statistical summary
4. generate_macro_vba_statistics_summary - Generate VBA macro
5. generate_alt_text_openai - Generate long alt texts (requires API key)

With experiment versioning enabled (default), all artifacts are saved
to experiments/runs/exp_<timestamp>_<model>_<version>/ with full lineage tracking.

Usage:
    python run_pipeline.py
    uv run python run_pipeline.py
"""

import sys
from pathlib import Path

# Add src to path so we can import the module
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from long_descriptions.main import main

if __name__ == "__main__":
    # Run the pipeline with default settings
    # Set generate_descriptions=True to include alt-text generation (Step 5)
    # Set use_mock_api=False to use real OpenAI API (incurs costs!)
    main(
        enable_versioning=True,       # Use experiment versioning
        generate_descriptions=True,   # Include alt-text descriptions (Step 5)
        use_mock_api=True,            # Use mock responses (no API costs)
        model_name="gpt-5.4",
        prompt_version="v1",
        random_seed=20260311,
        description_start_case=1,     # Start from first available case
        description_end_case=100,     # Generate for cases up to 100
        description_batch_size=10     # 10 cases per batch file
    )
