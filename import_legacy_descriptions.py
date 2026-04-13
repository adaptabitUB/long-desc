#!/usr/bin/env python3
"""
Import legacy OpenAI descriptions into a versioned experiment.
This script:
1. Runs pipeline steps 1-4 (without generating new descriptions)
2. Copies legacy description files from longdescriptions_byprompt/openai
3. Runs step 6 (programmatic metrics) on the legacy descriptions
"""

import shutil
import time
from datetime import datetime
from pathlib import Path

from src.long_descriptions import (
    generate_coverage_matrix,
    generate_charts,
    generate_statistics_summary,
    generate_macro_vba_statistics_summary,
    metriques_programatiques_openai,
)
from src.long_descriptions.experiment_tracker import ExperimentTracker


def main() -> None:
    """Import legacy descriptions into versioned experiment."""
    print("=" * 80)
    print("Importing Legacy OpenAI Descriptions into Versioned Experiment")
    print("=" * 80)
    print()
    
    project_root = Path(__file__).resolve().parent
    legacy_descriptions_dir = project_root / "longdescriptions_byprompt" / "openai"
    
    # Check legacy descriptions exist
    if not legacy_descriptions_dir.exists():
        print(f"❌ Legacy descriptions directory not found: {legacy_descriptions_dir}")
        return
    
    legacy_files = list(legacy_descriptions_dir.glob("alt_text_descriptions_*.md"))
    if not legacy_files:
        print(f"❌ No legacy description files found in {legacy_descriptions_dir}")
        return
    
    print(f"Found {len(legacy_files)} legacy description files to import")
    print()
    
    # Setup experiment tracking for legacy import
    timestamp = datetime.now()
    tracker = ExperimentTracker.__new__(ExperimentTracker)
    experiment_id = tracker.generate_experiment_id("gpt-4o", "legacy", timestamp)
    
    experiment_dir = project_root / "experiments" / "runs" / experiment_id
    tracker.__init__(experiment_dir, project_root)
    
    print(f"🔬 Experiment ID: {experiment_id}")
    print(f"📁 Experiment directory: {experiment_dir}")
    print()
    
    # Initialize manifest for legacy import
    model_config = {
        "provider": "openai",
        "model_name": "gpt-4o",
        "model_version": "gpt-4o-2024-08-06",
        "inference_params": {
            "temperature": 0.7,
            "max_tokens": 1500,
            "top_p": 1.0,
            "frequency_penalty": 0,
            "presence_penalty": 0
        }
    }
    
    prompt_info = {
        "version": "legacy",
        "file": "longdescriptions_byprompt/old/prompt.txt",
        "hash": None,
        "commit": tracker.capture_git_info()["git_commit"],
        "note": "Legacy descriptions imported from longdescriptions_byprompt/openai"
    }
    
    data_info = {
        "matrix_file": "data_snapshots/matrix_seed20260311_v1.csv",
        "charts_snapshot": "data_snapshots/charts_seed20260311_v1.json",
        "random_seed": 20260311,
        "num_cases": 500
    }
    
    pipeline_params = {
        "coverage_matrix": {
            "target_full_cases": 5000,
            "target_sample_cases": 500
        },
        "chart_generation": {
            "random_seed": 20260311
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
    
    start_time = time.time()
    
    # Step 1-4: Generate baseline artifacts
    print("[1/4] Running generate_coverage_matrix...")
    step_start = time.time()
    generate_coverage_matrix.main(
        target_full_cases=5000,
        target_sample_cases=500,
        experiment_dir=experiment_dir
    )
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    tracker.update_results({"coverage_matrix_time": step_time})
    print()
    
    print("[2/4] Running generate_charts...")
    step_start = time.time()
    generate_charts.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    tracker.update_results({"charts_generation_time": step_time})
    print()
    
    print("[3/4] Running generate_statistics_summary...")
    step_start = time.time()
    generate_statistics_summary.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    tracker.update_results({"statistics_summary_time": step_time})
    print()
    
    print("[4/4] Running generate_macro_vba_statistics_summary...")
    step_start = time.time()
    generate_macro_vba_statistics_summary.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    tracker.update_results({"vba_macro_time": step_time})
    print()
    
    # Copy legacy descriptions
    print("=" * 80)
    print("Copying legacy descriptions...")
    print("=" * 80)
    descriptions_dir = experiment_dir / "artifacts" / "descriptions"
    descriptions_dir.mkdir(parents=True, exist_ok=True)
    
    copied_files = []
    for legacy_file in sorted(legacy_files):
        dest_file = descriptions_dir / legacy_file.name
        shutil.copy2(legacy_file, dest_file)
        print(f"  ✓ Copied {legacy_file.name}")
        copied_files.append(legacy_file.name)
    
    print(f"\n✓ Copied {len(copied_files)} description files")
    print()
    
    # Update manifest with description info
    tracker.update_results({
        "descriptions_generated": True,
        "descriptions_mock": False,
        "descriptions_source": "legacy_import",
        "description_files": copied_files
    })
    
    # Step 5: Generate metrics for legacy descriptions
    print("=" * 80)
    print("Running programmatic metrics on legacy descriptions...")
    print("=" * 80)
    step_start = time.time()
    metriques_programatiques_openai.main(experiment_dir=experiment_dir)
    step_time = time.time() - step_start
    print(f"✓ Completed in {step_time:.2f}s")
    tracker.update_results({"metrics_time": step_time})
    print()
    
    total_time = time.time() - start_time
    tracker.update_results({
        "total_pipeline_time": total_time,
        "pipeline_completed": True,
        "completion_timestamp": datetime.now().isoformat()
    })
    
    print("=" * 80)
    print(f"Legacy import completed successfully in {total_time:.2f}s ({total_time/60:.2f}m)")
    print(f"📊 Results saved to: {experiment_dir}")
    print(f"📄 Manifest: {experiment_dir / 'manifest.json'}")
    print("=" * 80)


if __name__ == "__main__":
    main()
