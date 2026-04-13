#!/usr/bin/env python3
"""Test pipeline without descriptions."""

import sys
from pathlib import Path

# Add src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from long_descriptions.main import main

if __name__ == "__main__":
    # Test without descriptions (4 steps only)
    print("=" * 80)
    print("Testing pipeline WITHOUT descriptions (steps 1-4 only)")
    print("=" * 80)
    print()
    
    main(
        enable_versioning=True,
        generate_descriptions=False,  # Skip descriptions
        model_name="gpt-5.4",
        prompt_version="v1",
        random_seed=20260311
    )
