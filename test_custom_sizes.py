#!/usr/bin/env python3
"""Test script to demonstrate custom case sizes."""

import sys
sys.path.insert(0, 'src')

from long_descriptions.generate_coverage_matrix import main

# Test with 5000 full cases and 500 sample cases
print("Generating 5000/500 dataset...")
main(target_full_cases=5000, target_sample_cases=500)

print("\nGenerating 1000/100 dataset...")
main(target_full_cases=1000, target_sample_cases=100)
