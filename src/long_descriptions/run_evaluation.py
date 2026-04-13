"""
Evaluation script for experiment results.

This script evaluates experiment outputs against predefined quality metrics
and ground truth data.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List

# Simple readability metrics (simplified versions)


def calculate_flesch_reading_ease(text: str) -> float:
    """Calculate Flesch Reading Ease score (simplified)."""
    words = text.split()
    sentences = text.count('.') + text.count('!') + text.count('?')
    if not words or not sentences:
        return 0.0
    
    syllables = sum(count_syllables(word) for word in words)
    words_per_sentence = len(words) / max(sentences, 1)
    syllables_per_word = syllables / max(len(words), 1)
    
    # Flesch formula: 206.835 - 1.015 * (words/sentences) - 84.6 * (syllables/words)
    score = 206.835 - (1.015 * words_per_sentence) - (84.6 * syllables_per_word)
    return max(0, min(100, score))


def calculate_gunning_fog(text: str) -> float:
    """Calculate Gunning Fog index (simplified)."""
    words = text.split()
    sentences = text.count('.') + text.count('!') + text.count('?')
    if not words or not sentences:
        return 0.0
    
    complex_words = sum(1 for word in words if count_syllables(word) > 2)
    words_per_sentence = len(words) / max(sentences, 1)
    complex_word_ratio = complex_words / max(len(words), 1)
    
    # Gunning Fog: 0.4 * (words/sentences + 100 * complex_words/words)
    score = 0.4 * (words_per_sentence + 100 * complex_word_ratio)
    return max(0, score)


def count_syllables(word: str) -> int:
    """Count syllables in a word (simplified)."""
    word = word.lower().strip(".,!?;:")
    vowels = "aeiouy"
    syllable_count = 0
    previous_was_vowel = False
    
    for char in word:
        is_vowel = char in vowels
        if is_vowel and not previous_was_vowel:
            syllable_count += 1
        previous_was_vowel = is_vowel
    
    # Adjust for silent 'e'
    if word.endswith('e'):
        syllable_count -= 1
    
    return max(1, syllable_count)


def check_structure_completeness(description: str) -> Dict[str, Any]:
    """Check if all mandatory sections are present."""
    required_sections = [
        "Overview and main message",
        "Chart structure",
        "Relevant patterns, trends, and comparisons",
        "Essential key details"
    ]
    
    found_sections = []
    for section in required_sections:
        # Check for section heading (with ###)
        pattern = rf"###\s*{re.escape(section)}"
        if re.search(pattern, description, re.IGNORECASE):
            found_sections.append(section)
    
    completeness = len(found_sections) / len(required_sections)
    
    return {
        "score": completeness * 100,
        "found_sections": len(found_sections),
        "total_sections": len(required_sections),
        "missing": [s for s in required_sections if s not in found_sections]
    }


def check_visual_references(description: str, forbidden_phrases: List[str]) -> Dict[str, Any]:
    """Check for prohibited visual references."""
    found_violations = []
    lower_desc = description.lower()
    
    for phrase in forbidden_phrases:
        if phrase.lower() in lower_desc:
            found_violations.append(phrase)
    
    score = 0 if found_violations else 100
    
    return {
        "score": score,
        "violations": found_violations,
        "count": len(found_violations)
    }


def check_length_appropriateness(description: str, min_words: int = 100, max_words: int = 500) -> Dict[str, Any]:
    """Check if word count is within acceptable range."""
    words = description.split()
    word_count = len(words)
    
    if min_words <= word_count <= max_words:
        score = 100
    elif word_count < min_words:
        score = (word_count / min_words) * 100
    else:  # word_count > max_words
        excess = word_count - max_words
        penalty = min(50, excess / 2)  # Penalty for being too long
        score = max(50, 100 - penalty)
    
    return {
        "score": score,
        "word_count": word_count,
        "min": min_words,
        "max": max_words,
        "appropriate": min_words <= word_count <= max_words
    }


def evaluate_description(description: str, case_id: str, eval_config: Dict[str, Any]) -> Dict[str, Any]:
    """Evaluate a single description against all metrics."""
    metrics_results = {}
    
    # Structure completeness
    structure = check_structure_completeness(description)
    metrics_results["structure_completeness"] = structure
    
    # Readability
    flesch = calculate_flesch_reading_ease(description)
    gunning = calculate_gunning_fog(description)
    readability_score = (flesch + (100 - gunning * 5)) / 2  # Normalize
    metrics_results["readability_score"] = {
        "score": max(0, min(100, readability_score)),
        "flesch_reading_ease": flesch,
        "gunning_fog": gunning
    }
    
    # Visual references
    forbidden = eval_config["metrics"][4].get("forbidden_phrases", [])
    visual_refs = check_visual_references(description, forbidden)
    metrics_results["no_visual_references"] = visual_refs
    
    # Length
    length = check_length_appropriateness(description)
    metrics_results["length_appropriateness"] = length
    
    # Calculate weighted score
    weights = {m["name"]: m["weight"] for m in eval_config["metrics"]}
    total_score = 0
    for metric_name, result in metrics_results.items():
        weight = weights.get(metric_name, 0)
        metric_score = result.get("score", 0)
        total_score += metric_score * weight
    
    return {
        "case_id": case_id,
        "total_score": round(total_score, 2),
        "metrics": metrics_results
    }


def run_evaluation(experiment_id: str, eval_suite: str) -> Dict[str, Any]:
    """
    Run evaluation on an experiment.
    
    Args:
        experiment_id: ID of the experiment to evaluate
        eval_suite: Name of the evaluation suite to use (e.g., "v1")
    
    Returns:
        Evaluation results dictionary
    """
    project_root = Path.cwd()
    experiment_dir = project_root / "experiments" / "runs" / experiment_id
    eval_suite_dir = project_root / "experiments" / "evaluations" / f"eval_suite_{eval_suite}"
    
    # Load experiment manifest
    manifest_path = experiment_dir / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Experiment manifest not found: {manifest_path}")
    
    with manifest_path.open("r") as f:
        manifest = json.load(f)
    
    # Load evaluation config
    eval_config_path = eval_suite_dir / "eval_config.json"
    if not eval_config_path.exists():
        raise FileNotFoundError(f"Evaluation config not found: {eval_config_path}")
    
    with eval_config_path.open("r") as f:
        eval_config = json.load(f)
    
    # Load descriptions from experiment
    descriptions_dir = experiment_dir / "artifacts" / "descriptions"
    if not descriptions_dir.exists():
        print(f"Warning: No descriptions directory found at {descriptions_dir}")
        print("Evaluation will be limited to available artifacts.")
        return {
            "experiment_id": experiment_id,
            "eval_suite": eval_suite,
            "status": "incomplete",
            "message": "No descriptions found"
        }
    
    # Process descriptions
    description_files = sorted(descriptions_dir.glob("*.md"))
    all_results = []
    
    for desc_file in description_files:
        content = desc_file.read_text(encoding="utf-8")
        # Extract individual chart descriptions (assuming ## Chart {id}: format)
        chart_pattern = re.compile(r"##\s+Chart\s+(\d+)\s*:.*?\n(.*?)(?=##\s+Chart\s+\d+\s*:|$)", re.DOTALL)
        matches = chart_pattern.findall(content)
        
        for case_id, description in matches:
            result = evaluate_description(description, case_id, eval_config)
            all_results.append(result)
    
    # Aggregate results
    if all_results:
        mean_score = sum(r["total_score"] for r in all_results) / len(all_results)
        failed_cases = [r for r in all_results if r["total_score"] < eval_config["thresholds"]["mean_score"]]
        
        evaluation_results = {
            "experiment_id": experiment_id,
            "eval_suite": eval_suite,
            "timestamp": manifest.get("timestamp"),
            "model": manifest.get("model_config", {}).get("model_name"),
            "prompt_version": manifest.get("prompt", {}).get("version"),
            "summary": {
                "total_cases": len(all_results),
                "mean_score": round(mean_score, 2),
                "min_score": round(min(r["total_score"] for r in all_results), 2),
                "max_score": round(max(r["total_score"] for r in all_results), 2),
                "failures": len(failed_cases),
                "pass_rate": round((len(all_results) - len(failed_cases)) / len(all_results) * 100, 2)
            },
            "thresholds": eval_config["thresholds"],
            "passed": mean_score >= eval_config["thresholds"]["mean_score"] and len(failed_cases) <= eval_config["thresholds"]["max_failures"],
            "individual_results": all_results
        }
    else:
        evaluation_results = {
            "experiment_id": experiment_id,
            "eval_suite": eval_suite,
            "status": "no_results",
            "message": "No descriptions found to evaluate"
        }
    
    # Save results
    results_dir = eval_suite_dir / "results"
    results_dir.mkdir(parents=True, exist_ok=True)
    results_path = results_dir / f"{experiment_id}_eval_results.json"
    
    with results_path.open("w", encoding="utf-8") as f:
        json.dump(evaluation_results, f, indent=2, ensure_ascii=False)
    
    print(f"✓ Evaluation results saved to: {results_path}")
    
    return evaluation_results


def main():
    parser = argparse.ArgumentParser(description="Evaluate experiment results")
    parser.add_argument("--experiment-id", required=True, help="Experiment ID to evaluate")
    parser.add_argument("--eval-suite", default="v1", help="Evaluation suite version")
    
    args = parser.parse_args()
    
    results = run_evaluation(args.experiment_id, args.eval_suite)
    
    if "summary" in results:
        print("\n" + "=" * 80)
        print(f"Evaluation Results for {args.experiment_id}")
        print("=" * 80)
        print(f"Evaluation Suite: {args.eval_suite}")
        print(f"Total Cases: {results['summary']['total_cases']}")
        print(f"Mean Score: {results['summary']['mean_score']:.2f}")
        print(f"Min Score: {results['summary']['min_score']:.2f}")
        print(f"Max Score: {results['summary']['max_score']:.2f}")
        print(f"Failures: {results['summary']['failures']}")
        print(f"Pass Rate: {results['summary']['pass_rate']:.2f}%")
        print(f"\n{'PASSED' if results['passed'] else 'FAILED'}")
        print("=" * 80)


if __name__ == "__main__":
    main()
