"""Main orchestration script to run all generators in sequence."""

import os
import time
from datetime import datetime
from pathlib import Path

from . import (
    generate_coverage_matrix,
    generate_charts,
    generate_statistics_summary,
    generate_macro_vba_statistics_summary,
)
from .experiment_tracker import ExperimentTracker


def main(
    enable_versioning: bool = True,
    model_name: str = "gpt-5.4",
    prompt_version: str = "v1",
    random_seed: int = 20260311
) -> None:
    """
    Run all generation scripts in sequence.
    
    Args:
        enable_versioning: If True, use experiment versioning system
        model_name: Name of the model to use (for experiment ID)
        prompt_version: Version of the prompt template
        random_seed: Random seed for reproducibility
    """
    print("=" * 80)
    print("Starting Long Descriptions Generation Pipeline")
    print("=" * 80)
    print()
    
    # Setup experiment tracking
    project_root = Path(__file__).resolve().parent.parent.parent
    
    if enable_versioning:
        # Generate experiment ID and create tracker
        timestamp = datetime.now()
        tracker = ExperimentTracker.__new__(ExperimentTracker)
        experiment_id = tracker.generate_experiment_id(model_name, prompt_version, timestamp)
        
        experiment_dir = project_root / "experiments" / "runs" / experiment_id
        tracker.__init__(experiment_dir, project_root)
        
        print(f"� Experiment ID: {experiment_id}")
        print(f"📁 Experiment directory: {experiment_dir}")
        print()
        
        # Initialize manifest
        prompt_file = project_root / "experiments" / "prompts" / "versions" / f"{prompt_version}_prompt.txt"
        prompt_hash = tracker.calculate_file_hash(prompt_file) if prompt_file.exists() else None
        
        model_config = {
            "provider": "openai",
            "model_name": model_name,
            "model_version": f"{model_name}-{timestamp.strftime('%Y-%m-%d')}",
            "inference_params": {
                "temperature": 0.7,
                "max_tokens": 1500,
                "top_p": 1.0,
                "frequency_penalty": 0,
                "presence_penalty": 0
            }
        }
        
        prompt_info = {
            "version": prompt_version,
            "file": str(prompt_file.relative_to(project_root)),
            "hash": f"sha256:{prompt_hash[:16]}..." if prompt_hash else None,
            "commit": tracker.capture_git_info()["git_commit"]
        }
        
        data_info = {
            "matrix_file": f"data_snapshots/matrix_seed{random_seed}_v1.csv",
            "charts_snapshot": f"data_snapshots/charts_seed{random_seed}_v1.json",
            "random_seed": random_seed,
            "num_cases": 500
        }
        
        pipeline_params = {
            "coverage_matrix": {
                "target_full_cases": 5000,
                "target_sample_cases": 500
            },
            "chart_generation": {
                "random_seed": random_seed
            }
        }
        
        tracker.initialize_manifest(
            experiment_id,
            model_config,
            prompt_info,
            data_info,
            pipeline_params,
            timestamp
        )
        
        print("✓ Experiment manifest initialized")
        print()
    else:
        tracker = None
        experiment_dir = None
    
    start_time = time.time()
    
    # Step 1: Generate coverage matrix
    print("[1/4] Running generate_coverage_matrix...")
    step_start = time.time()
    generate_coverage_matrix.main(
        target_full_cases=5000,
        target_sample_cases=500,
        experiment_dir=experiment_dir
    )
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    if tracker:
        tracker.update_results({"coverage_matrix_time": step_time})
    print()
    
    # Step 2: Generate canonical instances
    print("[2/4] Running generate_charts...")
    step_start = time.time()
    generate_charts.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    if tracker:
        tracker.update_results({"charts_generation_time": step_time})
    print()
    
    # Step 3: Generate statistical summary
    print("[3/4] Running generate_statistics_summary...")
    step_start = time.time()
    generate_statistics_summary.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    if tracker:
        tracker.update_results({"statistics_summary_time": step_time})
    print()
    
    # Step 4: Generate VBA macro for statistical summary
    print("[4/4] Running generate_macro_vba_statistics_summary...")
    step_start = time.time()
    generate_macro_vba_statistics_summary.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    if tracker:
        tracker.update_results({"vba_macro_time": step_time})
    print()
    
    total_time = time.time() - start_time
    
    if tracker:
        tracker.update_results({
            "total_pipeline_time": total_time,
            "pipeline_completed": True,
            "completion_timestamp": datetime.now().isoformat()
        })
    
    print("=" * 80)
    print(f"Pipeline completed successfully in {total_time:.2f}s ({total_time/60:.2f}m)")
    if enable_versioning:
        print(f"📊 Results saved to: {experiment_dir}")
        print(f"📄 Manifest: {experiment_dir / 'manifest.json'}")
    print("=" * 80)


if __name__ == "__main__":
    main()
