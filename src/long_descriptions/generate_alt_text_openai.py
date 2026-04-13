from __future__ import annotations

import argparse
import json
import os
import time
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv
from openai import OpenAI


# Legacy mode constants (used only when experiment_dir is None)
BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent.parent
DEFAULT_INPUT_JSON = PROJECT_ROOT / "output" / "charts.json"
DEFAULT_PROMPT_FILE = PROJECT_ROOT / "experiments" / "prompts" / "active_prompt.txt"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "longdescriptions_byprompt" / "openai"
DEFAULT_ENV_FILE = PROJECT_ROOT / ".env"
DEFAULT_BATCH_SIZE = 50
DEFAULT_MODEL = "gpt-5.4"


def main(
    experiment_dir: Path | None = None,
    start_case: int = 1,
    end_case: int | None = None,
    batch_size: int = DEFAULT_BATCH_SIZE,
    model: str | None = None,
    sleep_seconds: float = 0.0,
    use_mock: bool = False
) -> None:
    """
    Generate long alt-text descriptions using OpenAI API.
    
    Args:
        experiment_dir: Optional experiment directory for versioned output
        start_case: First case number to process
        end_case: Last case number to process (None = all)
        batch_size: Number of charts per batch file
        model: OpenAI model to use (None = use env var or default)
        sleep_seconds: Seconds to sleep between API calls
        use_mock: If True, use mock responses instead of real API calls
    """
    # Determine paths based on mode (versioned vs legacy)
    if experiment_dir is not None:
        artifacts_dir = Path(experiment_dir) / "artifacts"
        input_json = artifacts_dir / "charts.json"
        output_dir = artifacts_dir / "descriptions"
        # Use active prompt from experiments/prompts/
        prompt_file = PROJECT_ROOT / "experiments" / "prompts" / "active_prompt.txt"
        if not prompt_file.exists():
            # Fallback to v1 if active_prompt doesn't exist
            prompt_file = PROJECT_ROOT / "experiments" / "prompts" / "versions" / "v1_prompt.txt"
    else:
        # Legacy mode: use output/ directory
        input_json = DEFAULT_INPUT_JSON
        output_dir = DEFAULT_OUTPUT_DIR
        prompt_file = DEFAULT_PROMPT_FILE
    
    # Load instances and prompt
    instances = load_instances(input_json)
    prompt_template = load_prompt_template(prompt_file)
    selected_instances = select_instances(instances, start_case, end_case)

    if not selected_instances:
        raise RuntimeError("No chart instances matched the requested range")

    # Determine model to use
    if model is None:
        model = os.environ.get("OPENAI_MODEL", DEFAULT_MODEL)
    
    # Initialize client (or mock)
    if use_mock:
        print("⚠️  Using MOCK mode - no real API calls will be made")
        client = None  # Will be handled in call_model
    else:
        api_key = load_api_key()
        client = OpenAI(api_key=api_key)
    
    # Process in batches
    batch_size = max(1, batch_size)
    created_files: List[Path] = []
    total = len(selected_instances)
    print(f"Starting generation for {total} charts with batch size {batch_size}...", flush=True)

    for batch_start in range(0, len(selected_instances), batch_size):
        batch_instances = selected_instances[batch_start : batch_start + batch_size]
        batch_outputs: List[str] = []
        batch_first = normalize_case_number(str(batch_instances[0]["id"]))
        batch_last = normalize_case_number(str(batch_instances[-1]["id"]))
        print(f"Batch {batch_first}-{batch_last}: started", flush=True)

        for local_index, instance in enumerate(batch_instances, start=1):
            prompt = render_prompt(prompt_template, instance)
            case_number = normalize_case_number(str(instance["id"]))
            global_index = batch_start + local_index
            print(f"  [{global_index}/{total}] case {case_number}: requesting model...", flush=True)
            
            if use_mock:
                response_text = generate_mock_response(instance)
            else:
                response_text = call_model(client, model, prompt)
            
            batch_outputs.append(response_text)
            print(f"  [{global_index}/{total}] case {case_number}: done", flush=True)
            
            if sleep_seconds > 0:
                time.sleep(sleep_seconds)

        written_file = write_batch_file(output_dir, batch_instances, batch_outputs)
        created_files.append(written_file)
        print(f"Batch {batch_first}-{batch_last}: written {written_file}", flush=True)

    print(f"Model used: {model}")
    print(f"Output dir: {output_dir}")
    print(f"Charts processed: {len(selected_instances)}")
    for path in created_files:
        print(f" - {path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate long alt-text descriptions for chart instances using OpenAI."
    )
    parser.add_argument("--input-json", type=Path, default=DEFAULT_INPUT_JSON)
    parser.add_argument("--prompt-file", type=Path, default=DEFAULT_PROMPT_FILE)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE)
    parser.add_argument("--start-case", type=int, default=1)
    parser.add_argument("--end-case", type=int)
    parser.add_argument("--model", default=os.environ.get("OPENAI_MODEL", DEFAULT_MODEL))
    parser.add_argument("--sleep-seconds", type=float, default=0.0)
    return parser.parse_args()


def load_api_key() -> str:
    # Always resolve .env from project root so execution directory does not matter.
    load_dotenv(dotenv_path=DEFAULT_ENV_FILE)
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY no definida. Configura-la com a variable d'entorn o en un fitxer .env a l'arrel del projecte: "
            f"{DEFAULT_ENV_FILE}"
        )
    return api_key


def load_instances(input_json: Path) -> List[Dict[str, Any]]:
    if not input_json.exists():
        raise FileNotFoundError(f"Input JSON not found: {input_json}")
    instances = json.loads(input_json.read_text(encoding="utf-8"))
    if not isinstance(instances, list):
        raise ValueError(f"Expected a list of chart instances in {input_json}")
    return instances


def load_prompt_template(prompt_file: Path) -> str:
    if not prompt_file.exists():
        raise FileNotFoundError(f"Prompt file not found: {prompt_file}")
    return prompt_file.read_text(encoding="utf-8")


def normalize_case_number(case_id: str) -> int:
    if case_id.startswith("CASE_"):
        return int(case_id.removeprefix("CASE_"))
    return int(case_id)


def select_instances(
    instances: List[Dict[str, Any]], start_case: int, end_case: int | None
) -> List[Dict[str, Any]]:
    selected: List[Dict[str, Any]] = []
    for instance in instances:
        case_id = str(instance.get("id") or "")
        if not case_id:
            continue
        case_number = normalize_case_number(case_id)
        if case_number < start_case:
            continue
        if end_case is not None and case_number > end_case:
            continue
        selected.append(instance)
    return sorted(selected, key=lambda item: normalize_case_number(str(item["id"])))


def render_prompt(template: str, instance: Dict[str, Any]) -> str:
    case_id = str(instance.get("id") or "")
    chart_title = str(instance.get("title") or "")
    chart_json = json.dumps(instance, ensure_ascii=False, indent=2)
    return template.format(case_id=case_id, chart_title=chart_title, chart_json=chart_json)


def extract_text(response: Any) -> str:
    output_text = getattr(response, "output_text", None)
    if output_text:
        return output_text.strip()

    parts: List[str] = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            text = getattr(content, "text", None)
            if text:
                parts.append(text)
    return "\n".join(parts).strip()


def call_model(client: OpenAI, model: str, prompt: str) -> str:
    response = client.responses.create(model=model, input=prompt)
    text = extract_text(response)
    if not text:
        raise RuntimeError("The model response was empty")
    return text


def generate_mock_response(instance: Dict[str, Any]) -> str:
    """Generate a mock response for testing without API calls."""
    chart_type = instance.get("chart", {}).get("excel_family", "Unknown")
    title = instance.get("title", "Untitled chart")
    case_id = instance.get("id", "CASE_00000")
    
    # Generate mock content with expected sections
    mock_sections = [
        "### Overview and main message",
        f"This {chart_type} chart titled '{title}' presents data visualization for {case_id}. "
        "The chart effectively communicates key data relationships and patterns.",
        "",
        "### Chart structure",
        f"The chart is a {chart_type} type visualization with multiple data series. "
        "Each element is clearly labeled and follows a consistent visual hierarchy.",
        "",
        "### Relevant patterns, trends, and comparisons",
        "The data shows several notable patterns across the visualization. "
        "Key trends indicate variations in the measured values, with some categories "
        "showing higher values than others. Comparisons between different data series "
        "reveal interesting relationships.",
        "",
        "### Essential key details",
        "The visualization includes specific numeric values and categorical breakdowns. "
        "Data points are distributed according to their respective categories, "
        "providing a comprehensive view of the underlying dataset.",
    ]
    
    return "\n".join(mock_sections)


def write_batch_file(
    output_dir: Path,
    batch_instances: List[Dict[str, Any]],
    batch_outputs: List[str],
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    start_number = normalize_case_number(str(batch_instances[0]["id"]))
    end_number = normalize_case_number(str(batch_instances[-1]["id"]))
    batch_file = output_dir / f"alt_text_descriptions_{start_number}-{end_number}.md"

    sections: List[str] = []
    for instance, text in zip(batch_instances, batch_outputs):
        chart_number = normalize_case_number(str(instance["id"]))
        sections.append(f"## Chart {chart_number}: {instance.get('title', '')}")
        sections.append("")
        sections.append(text.strip())
        sections.append("")

    batch_file.write_text("\n".join(sections).strip() + "\n", encoding="utf-8")
    return batch_file


def cli_main() -> None:
    """CLI entry point for backwards compatibility."""
    args = parse_args()
    
    # Call the main function with parsed arguments
    main(
        experiment_dir=None,  # CLI always uses legacy mode
        start_case=args.start_case,
        end_case=args.end_case,
        batch_size=args.batch_size,
        model=args.model,
        sleep_seconds=args.sleep_seconds,
        use_mock=False  # CLI uses real API
    )


if __name__ == "__main__":
    cli_main()