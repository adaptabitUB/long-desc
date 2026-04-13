"""
Compare multiple experiments to identify differences in configuration and results.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List


def load_manifest(experiment_id: str) -> Dict[str, Any]:
    """Load experiment manifest."""
    project_root = Path.cwd()
    manifest_path = project_root / "experiments" / "runs" / experiment_id / "manifest.json"
    
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found for experiment: {experiment_id}")
    
    with manifest_path.open("r") as f:
        return json.load(f)


def compare_field(field_name: str, values: List[Any]) -> str:
    """Compare values for a specific field."""
    unique_values = set(str(v) for v in values if v is not None)
    
    if len(unique_values) == 1:
        return f"  {field_name}: {values[0]} (same in all)"
    else:
        return f"  {field_name}:\n" + "\n".join(f"    - {exp_id}: {val}" for exp_id, val in zip(range(len(values)), values))


def compare_experiments(experiment_ids: List[str]) -> None:
    """
    Compare multiple experiments.
    
    Args:
        experiment_ids: List of experiment IDs to compare
    """
    print("=" * 80)
    print(f"Comparing {len(experiment_ids)} experiments")
    print("=" * 80)
    print()
    
    # Load all manifests
    manifests = {}
    for exp_id in experiment_ids:
        try:
            manifests[exp_id] = load_manifest(exp_id)
            print(f"✓ Loaded manifest for {exp_id}")
        except FileNotFoundError as e:
            print(f"✗ {e}")
            return
    
    print()
    
    # Compare Git information
    print("## Git Information")
    print("-" * 80)
    for exp_id in experiment_ids:
        manifest = manifests[exp_id]
        print(f"\n{exp_id}:")
        print(f"  Commit: {manifest.get('git_commit', 'N/A')[:12]}")
        print(f"  Branch: {manifest.get('git_branch', 'N/A')}")
        print(f"  Dirty: {manifest.get('git_dirty', 'N/A')}")
    print()
    
    # Compare Environment
    print("## Environment")
    print("-" * 80)
    py_versions = [manifests[exp_id].get("environment", {}).get("python_version") for exp_id in experiment_ids]
    print(compare_field("Python Version", py_versions))
    
    # Compare key packages
    all_packages = set()
    for exp_id in experiment_ids:
        packages = manifests[exp_id].get("environment", {}).get("packages", {})
        all_packages.update(packages.keys())
    
    for package in sorted(all_packages):
        versions = [manifests[exp_id].get("environment", {}).get("packages", {}).get(package, "N/A") for exp_id in experiment_ids]
        print(compare_field(f"  {package}", versions))
    print()
    
    # Compare Model Configuration
    print("## Model Configuration")
    print("-" * 80)
    for exp_id in experiment_ids:
        manifest = manifests[exp_id]
        model_config = manifest.get("model_config", {})
        print(f"\n{exp_id}:")
        print(f"  Provider: {model_config.get('provider', 'N/A')}")
        print(f"  Model: {model_config.get('model_name', 'N/A')}")
        print(f"  Inference params: {json.dumps(model_config.get('inference_params', {}), indent=4)}")
    print()
    
    # Compare Prompt
    print("## Prompt")
    print("-" * 80)
    prompt_versions = [manifests[exp_id].get("prompt", {}).get("version") for exp_id in experiment_ids]
    print(compare_field("Version", prompt_versions))
    
    prompt_hashes = [manifests[exp_id].get("prompt", {}).get("hash") for exp_id in experiment_ids]
    print(compare_field("Hash", prompt_hashes))
    print()
    
    # Compare Data
    print("## Data")
    print("-" * 80)
    random_seeds = [manifests[exp_id].get("data", {}).get("random_seed") for exp_id in experiment_ids]
    print(compare_field("Random Seed", random_seeds))
    
    num_cases = [manifests[exp_id].get("data", {}).get("num_cases") for exp_id in experiment_ids]
    print(compare_field("Number of Cases", num_cases))
    print()
    
    # Compare Results
    print("## Results")
    print("-" * 80)
    for exp_id in experiment_ids:
        manifest = manifests[exp_id]
        results = manifest.get("results", {})
        print(f"\n{exp_id}:")
        print(f"  Pipeline completed: {results.get('pipeline_completed', 'N/A')}")
        print(f"  Total time: {results.get('total_pipeline_time', 'N/A'):.2f}s" if results.get('total_pipeline_time') else "  Total time: N/A")
        
        # Show metrics if available
        if "charts_generated" in results:
            print(f"  Charts generated: {results.get('charts_generated')}")
        if "descriptions_generated" in results:
            print(f"  Descriptions generated: {results.get('descriptions_generated')}")
    print()
    
    # Summary of differences
    print("## Summary of Key Differences")
    print("-" * 80)
    differences = []
    
    if len(set(str(v) for v in prompt_versions)) > 1:
        differences.append("- Prompt versions differ")
    
    if len(set(str(v) for v in py_versions)) > 1:
        differences.append("- Python versions differ")
    
    git_commits = [manifests[exp_id].get("git_commit") for exp_id in experiment_ids]
    if len(set(str(v) for v in git_commits)) > 1:
        differences.append("- Git commits differ")
    
    if len(set(str(v) for v in random_seeds)) > 1:
        differences.append("- Random seeds differ")
    
    if differences:
        for diff in differences:
            print(diff)
    else:
        print("No significant differences found in configuration.")
    print()
    
    print("=" * 80)


def main():
    parser = argparse.ArgumentParser(description="Compare multiple experiments")
    parser.add_argument("experiment_ids", nargs="+", help="Experiment IDs to compare")
    
    args = parser.parse_args()
    
    if len(args.experiment_ids) < 2:
        print("Error: Need at least 2 experiments to compare")
        return
    
    compare_experiments(args.experiment_ids)


if __name__ == "__main__":
    main()
