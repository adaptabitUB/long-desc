"""
List all experiments with summary information.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List


def list_experiments(limit: int | None = None, sort_by: str = "timestamp") -> List[Dict[str, Any]]:
    """
    List all experiments.
    
    Args:
        limit: Maximum number of experiments to show (None for all)
        sort_by: Field to sort by (timestamp, experiment_id)
    
    Returns:
        List of experiment summaries
    """
    project_root = Path.cwd()
    experiments_dir = project_root / "experiments" / "runs"
    
    if not experiments_dir.exists():
        print(f"No experiments directory found at {experiments_dir}")
        return []
    
    experiment_dirs = [d for d in experiments_dir.iterdir() if d.is_dir()]
    
    if not experiment_dirs:
        print("No experiments found")
        return []
    
    experiments = []
    
    for exp_dir in experiment_dirs:
        manifest_path = exp_dir / "manifest.json"
        
        if not manifest_path.exists():
            continue
        
        try:
            with manifest_path.open("r") as f:
                manifest = json.load(f)
            
            # Extract key information
            summary = {
                "experiment_id": manifest.get("experiment_id", exp_dir.name),
                "timestamp": manifest.get("timestamp"),
                "git_commit": manifest.get("git_commit", "N/A")[:12] if manifest.get("git_commit") else "N/A",
                "git_branch": manifest.get("git_branch", "N/A"),
                "git_dirty": manifest.get("git_dirty", False),
                "model": manifest.get("model_config", {}).get("model_name", "N/A"),
                "prompt_version": manifest.get("prompt", {}).get("version", "N/A"),
                "python_version": manifest.get("environment", {}).get("python_version", "N/A"),
                "random_seed": manifest.get("data", {}).get("random_seed", "N/A"),
                "pipeline_completed": manifest.get("results", {}).get("pipeline_completed", False),
                "total_time": manifest.get("results", {}).get("total_pipeline_time"),
            }
            
            experiments.append(summary)
        
        except (json.JSONDecodeError, KeyError) as e:
            print(f"Warning: Could not parse manifest for {exp_dir.name}: {e}")
            continue
    
    # Sort experiments
    if sort_by == "timestamp":
        experiments.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
    elif sort_by == "experiment_id":
        experiments.sort(key=lambda x: x.get("experiment_id", ""))
    
    # Limit if requested
    if limit:
        experiments = experiments[:limit]
    
    return experiments


def print_experiments_table(experiments: List[Dict[str, Any]]) -> None:
    """Print experiments in a formatted table."""
    if not experiments:
        print("No experiments to display")
        return
    
    print("\n" + "=" * 150)
    print(f"{'Experiment ID':<40} {'Timestamp':<20} {'Model':<12} {'Prompt':<8} {'Commit':<14} {'Status':<12} {'Time':<10}")
    print("=" * 150)
    
    for exp in experiments:
        exp_id = exp["experiment_id"][:38]
        timestamp = exp.get("timestamp", "N/A")[:19] if exp.get("timestamp") else "N/A"
        model = exp.get("model", "N/A")[:10]
        prompt = exp.get("prompt_version", "N/A")[:6]
        commit = exp.get("git_commit", "N/A")[:12]
        
        if exp["git_dirty"]:
            commit += "*"
        
        status = "✓ Complete" if exp.get("pipeline_completed") else "✗ Incomplete"
        
        total_time = exp.get("total_time")
        time_str = f"{total_time:.1f}s" if total_time else "N/A"
        
        print(f"{exp_id:<40} {timestamp:<20} {model:<12} {prompt:<8} {commit:<14} {status:<12} {time_str:<10}")
    
    print("=" * 150)
    print(f"\nTotal experiments: {len(experiments)}")
    print()


def print_experiment_details(experiment_id: str) -> None:
    """Print detailed information about a specific experiment."""
    project_root = Path.cwd()
    manifest_path = project_root / "experiments" / "runs" / experiment_id / "manifest.json"
    
    if not manifest_path.exists():
        print(f"Experiment not found: {experiment_id}")
        return
    
    with manifest_path.open("r") as f:
        manifest = json.load(f)
    
    print("\n" + "=" * 80)
    print(f"Experiment Details: {experiment_id}")
    print("=" * 80)
    
    print(f"\nTimestamp: {manifest.get('timestamp')}")
    print(f"Git Commit: {manifest.get('git_commit')}")
    print(f"Git Branch: {manifest.get('git_branch')}")
    print(f"Git Dirty: {manifest.get('git_dirty')}")
    
    print("\n## Environment")
    env = manifest.get("environment", {})
    print(f"Python: {env.get('python_version')}")
    print(f"Platform: {env.get('platform')}")
    print("Packages:")
    for pkg, version in env.get("packages", {}).items():
        print(f"  - {pkg}: {version}")
    
    print("\n## Model Configuration")
    model = manifest.get("model_config", {})
    print(f"Provider: {model.get('provider')}")
    print(f"Model: {model.get('model_name')}")
    print(f"Parameters: {json.dumps(model.get('inference_params', {}), indent=2)}")
    
    print("\n## Prompt")
    prompt = manifest.get("prompt", {})
    print(f"Version: {prompt.get('version')}")
    print(f"File: {prompt.get('file')}")
    print(f"Hash: {prompt.get('hash')}")
    
    print("\n## Data")
    data = manifest.get("data", {})
    print(f"Random Seed: {data.get('random_seed')}")
    print(f"Number of Cases: {data.get('num_cases')}")
    
    print("\n## Results")
    results = manifest.get("results", {})
    for key, value in results.items():
        print(f"{key}: {value}")
    
    print("\n" + "=" * 80)


def main():
    parser = argparse.ArgumentParser(description="List experiments")
    parser.add_argument("--limit", type=int, help="Maximum number of experiments to show")
    parser.add_argument("--sort", choices=["timestamp", "experiment_id"], default="timestamp", help="Sort by field")
    parser.add_argument("--details", help="Show detailed information for a specific experiment ID")
    
    args = parser.parse_args()
    
    if args.details:
        print_experiment_details(args.details)
    else:
        experiments = list_experiments(limit=args.limit, sort_by=args.sort)
        print_experiments_table(experiments)


if __name__ == "__main__":
    main()
