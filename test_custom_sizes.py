#!/usr/bin/env python3
"""
Test script to generate coverage matrix with custom sizes.

This demonstrates how to call generate_coverage_matrix with custom parameters
for target_full_cases and target_sample_cases.

Usage:
    python test_custom_sizes.py
    uv run python test_custom_sizes.py
"""

import sys
from pathlib import Path

# Add src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from long_descriptions.generate_coverage_matrix import main as generate_matrix


def test_small_sizes():
    """Test with small matrix sizes (for quick testing)."""
    print("=" * 80)
    print("Testing with small sizes (100 full, 50 sample)")
    print("=" * 80)
    print()
    
    generate_matrix(
        target_full_cases=100,
        target_sample_cases=50,
        output_dir="output"
    )
    
    print()
    print("✅ Small test complete!")


def test_medium_sizes():
    """Test with medium matrix sizes."""
    print("=" * 80)
    print("Testing with medium sizes (1000 full, 200 sample)")
    print("=" * 80)
    print()
    
    generate_matrix(
        target_full_cases=1000,
        target_sample_cases=200,
        output_dir="output"
    )
    
    print()
    print("✅ Medium test complete!")


def test_custom_sizes():
    """Test with user-specified sizes."""
    print("=" * 80)
    print("Custom Size Test")
    print("=" * 80)
    print()
    
    try:
        full = int(input("Enter target_full_cases (e.g., 5000): "))
        sample = int(input("Enter target_sample_cases (e.g., 500): "))
        
        print()
        print(f"Generating matrix with {full} full cases and {sample} sample cases...")
        print()
        
        generate_matrix(
            target_full_cases=full,
            target_sample_cases=sample,
            output_dir="output"
        )
        
        print()
        print("✅ Custom test complete!")
        
    except ValueError:
        print("❌ Invalid input. Please enter integers.")
        sys.exit(1)


if __name__ == "__main__":
    print()
    print("Coverage Matrix Size Testing")
    print()
    print("Options:")
    print("  1. Small (100 full, 50 sample) - Quick test")
    print("  2. Medium (1000 full, 200 sample)")
    print("  3. Custom sizes")
    print()
    
    choice = input("Select option (1-3): ").strip()
    print()
    
    if choice == "1":
        test_small_sizes()
    elif choice == "2":
        test_medium_sizes()
    elif choice == "3":
        test_custom_sizes()
    else:
        print("❌ Invalid option")
        sys.exit(1)
