"""
Experiment tracking module for versioning AI experiments.

This module provides functionality to track experiment lineage including:
- Git commit information
- Python environment and dependencies
- Model configuration
- Prompt versions
- Data snapshots
- Results and metrics
"""

from __future__ import annotations

import hashlib
import json
import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

import importlib.metadata


class ExperimentTracker:
    """Tracks experiment metadata and generates manifest files."""
    
    def __init__(self, experiment_dir: Path, project_root: Path):
        """
        Initialize experiment tracker.
        
        Args:
            experiment_dir: Directory where experiment artifacts will be stored
            project_root: Root directory of the project (for Git operations)
        """
        self.experiment_dir = Path(experiment_dir)
        self.project_root = Path(project_root)
        self.manifest_path = self.experiment_dir / "manifest.json"
        self.manifest: Dict[str, Any] = {}
        
        # Create experiment directory structure
        self._create_directory_structure()
    
    def _create_directory_structure(self) -> None:
        """Create standard experiment directory structure."""
        (self.experiment_dir / "artifacts").mkdir(parents=True, exist_ok=True)
        (self.experiment_dir / "artifacts" / "statistics").mkdir(exist_ok=True)
        (self.experiment_dir / "artifacts" / "descriptions").mkdir(exist_ok=True)
        (self.experiment_dir / "artifacts" / "metrics").mkdir(exist_ok=True)
        (self.experiment_dir / "configs").mkdir(exist_ok=True)
        (self.experiment_dir / "logs").mkdir(exist_ok=True)
    
    def capture_git_info(self) -> Dict[str, Any]:
        """
        Capture current Git repository information.
        
        Returns:
            Dictionary with git commit, branch, and dirty status
        """
        git_info = {
            "git_commit": None,
            "git_branch": None,
            "git_dirty": False
        }
        
        try:
            # Get current commit SHA
            result = subprocess.run(
                ["git", "rev-parse", "HEAD"],
                cwd=self.project_root,
                capture_output=True,
                text=True,
                check=True
            )
            git_info["git_commit"] = result.stdout.strip()
            
            # Get current branch
            result = subprocess.run(
                ["git", "rev-parse", "--abbrev-ref", "HEAD"],
                cwd=self.project_root,
                capture_output=True,
                text=True,
                check=True
            )
            git_info["git_branch"] = result.stdout.strip()
            
            # Check if there are uncommitted changes
            result = subprocess.run(
                ["git", "status", "--porcelain"],
                cwd=self.project_root,
                capture_output=True,
                text=True,
                check=True
            )
            git_info["git_dirty"] = bool(result.stdout.strip())
            
        except (subprocess.CalledProcessError, FileNotFoundError):
            # Git not available or not a git repository
            pass
        
        return git_info
    
    def capture_environment(self, package_names: Optional[list[str]] = None) -> Dict[str, Any]:
        """
        Capture Python environment information.
        
        Args:
            package_names: List of package names to track versions for.
                          If None, uses common packages: openai, pandas, numpy
        
        Returns:
            Dictionary with Python version, packages, and platform info
        """
        if package_names is None:
            package_names = ["openai", "pandas", "numpy", "openpyxl"]
        
        packages = {}
        for package_name in package_names:
            try:
                version = importlib.metadata.version(package_name)
                packages[package_name] = version
            except importlib.metadata.PackageNotFoundError:
                packages[package_name] = "not installed"
        
        return {
            "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            "packages": packages,
            "platform": platform.platform()
        }
    
    def calculate_file_hash(self, file_path: Path) -> str:
        """
        Calculate SHA256 hash of a file.
        
        Args:
            file_path: Path to the file
        
        Returns:
            SHA256 hash as hex string
        """
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def generate_experiment_id(
        self,
        model_name: str,
        prompt_version: str,
        timestamp: Optional[datetime] = None
    ) -> str:
        """
        Generate unique experiment ID.
        
        Args:
            model_name: Name of the model used (e.g., "gpt-5.4")
            prompt_version: Version of the prompt (e.g., "v1")
            timestamp: Timestamp for the experiment (defaults to now)
        
        Returns:
            Experiment ID string like "exp_2026-04-13_14-30_gpt5.4_v1"
        """
        if timestamp is None:
            timestamp = datetime.now()
        
        # Format: exp_YYYY-MM-DD_HH-MM_model_promptversion
        model_safe = model_name.replace("-", "").replace(".", "")
        time_str = timestamp.strftime("%Y-%m-%d_%H-%M")
        
        return f"exp_{time_str}_{model_safe}_{prompt_version}"
    
    def initialize_manifest(
        self,
        experiment_id: str,
        model_config: Dict[str, Any],
        prompt_info: Dict[str, Any],
        data_info: Dict[str, Any],
        pipeline_params: Dict[str, Any],
        timestamp: Optional[datetime] = None
    ) -> None:
        """
        Initialize experiment manifest with all configuration.
        
        Args:
            experiment_id: Unique identifier for the experiment
            model_config: Model configuration (provider, name, params)
            prompt_info: Prompt information (version, file, hash)
            data_info: Data snapshot information
            pipeline_params: Pipeline execution parameters
            timestamp: Experiment timestamp (defaults to now)
        """
        if timestamp is None:
            timestamp = datetime.now()
        
        git_info = self.capture_git_info()
        env_info = self.capture_environment()
        
        self.manifest = {
            "experiment_id": experiment_id,
            "timestamp": timestamp.isoformat(),
            **git_info,
            "environment": env_info,
            "model_config": model_config,
            "prompt": prompt_info,
            "data": data_info,
            "pipeline_params": pipeline_params,
            "results": {},
            "evaluation": {}
        }
        
        self.save_manifest()
    
    def update_results(self, results: Dict[str, Any]) -> None:
        """
        Update results section of manifest.
        
        Args:
            results: Results dictionary to merge into manifest
        """
        self.manifest.setdefault("results", {})
        self.manifest["results"].update(results)
        self.save_manifest()
    
    def update_evaluation(self, evaluation: Dict[str, Any]) -> None:
        """
        Update evaluation section of manifest.
        
        Args:
            evaluation: Evaluation dictionary to merge into manifest
        """
        self.manifest.setdefault("evaluation", {})
        self.manifest["evaluation"].update(evaluation)
        self.save_manifest()
    
    def save_manifest(self) -> None:
        """Save manifest to disk."""
        with open(self.manifest_path, "w", encoding="utf-8") as f:
            json.dump(self.manifest, f, indent=2, ensure_ascii=False)
    
    def load_manifest(self) -> Dict[str, Any]:
        """
        Load manifest from disk.
        
        Returns:
            Manifest dictionary
        """
        if self.manifest_path.exists():
            with open(self.manifest_path, "r", encoding="utf-8") as f:
                self.manifest = json.load(f)
        return self.manifest
    
    def get_artifact_path(self, artifact_type: str) -> Path:
        """
        Get path for a specific artifact type.
        
        Args:
            artifact_type: Type of artifact (e.g., "charts", "statistics", "descriptions", "metrics")
        
        Returns:
            Path to artifact directory
        """
        if artifact_type in ["statistics", "descriptions", "metrics"]:
            return self.experiment_dir / "artifacts" / artifact_type
        else:
            return self.experiment_dir / "artifacts"
    
    def get_config_path(self) -> Path:
        """Get path to configs directory."""
        return self.experiment_dir / "configs"
    
    def get_logs_path(self) -> Path:
        """Get path to logs directory."""
        return self.experiment_dir / "logs"
