from __future__ import annotations

import csv
import json
import math
import re
from pathlib import Path
from functools import lru_cache
from typing import Any, Dict, Iterable, List

import pronouncing
import pyphen


BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent.parent
LONG_DESCRIPTIONS_DIR = PROJECT_ROOT / "longdescriptions_byprompt"
OUTPUT_DIR = PROJECT_ROOT / "output"
METRICS_DIR = OUTPUT_DIR / "metriques"

# Configuració de proveedors
PROVIDERS = {
	"claude": LONG_DESCRIPTIONS_DIR / "claude",
	"gemini": LONG_DESCRIPTIONS_DIR / "gemini",
	"chatgpt": LONG_DESCRIPTIONS_DIR / "chatgpt",
}

OUTPUT_CSV_UNIFIED = METRICS_DIR / "metriques_unificat.csv"
OUTPUT_JSON_UNIFIED = METRICS_DIR / "metriques_unificat.json"

WORD_PATTERN = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
NUMERIC_PATTERN = re.compile(r"(?<![A-Za-z])[+-]?(?:\d+[\.,]?\d*|\d*[\.,]\d+)(?![A-Za-z])")
CASE_ID_PATTERN = re.compile(r"CASE_\d+")
CHART_HEADER_PATTERN = re.compile(r"^##\s+Chart\s+\d+:\s+(.+?)\s*\((\d+)\)\s*$")
# Split sentences on '.', '!', '?' but keep decimal numbers intact (e.g., 38.6)
SENTENCE_SPLIT_PATTERN = re.compile(r"(?:\.(?!\d)|[!?])+")
HYPHENATOR = pyphen.Pyphen(lang="en_US")
SECTION_KEYS = {
	"overview and main message": "overview_and_main_message",
	"chart structure": "chart_structure",
	"relevant patterns, trends, and comparisons": "relevant_patterns_trends_and_comparisons",
	"essential key details": "essential_key_details",
}

CARDINAL_UNITS = {
	"zero",
	"one",
	"two",
	"three",
	"four",
	"five",
	"six",
	"seven",
	"eight",
	"nine",
	"ten",
	"eleven",
	"twelve",
	"thirteen",
	"fourteen",
	"fifteen",
	"sixteen",
	"seventeen",
	"eighteen",
	"nineteen",
}
CARDINAL_TENS = {
	"twenty",
	"thirty",
	"forty",
	"fifty",
	"sixty",
	"seventy",
	"eighty",
	"ninety",
}
CARDINAL_SCALES = {"hundred", "thousand", "million", "billion", "trillion"}
ORDINAL_WORDS = {
	"first",
	"second",
	"third",
	"fourth",
	"fifth",
	"sixth",
	"seventh",
	"eighth",
	"ninth",
	"tenth",
	"eleventh",
	"twelfth",
	"thirteenth",
	"fourteenth",
	"fifteenth",
	"sixteenth",
	"seventeenth",
	"eighteenth",
	"nineteenth",
	"twentieth",
	"thirtieth",
	"fortieth",
	"fiftieth",
	"sixtieth",
	"seventieth",
	"eightieth",
	"ninetieth",
	"hundredth",
	"thousandth",
	"millionth",
	"billionth",
	"trillionth",
}
CARDINAL_NUMBER_WORDS = CARDINAL_UNITS | CARDINAL_TENS | CARDINAL_SCALES



def extract_case_id(chart_label: str) -> str:
	match = CASE_ID_PATTERN.search(chart_label or "")
	if match:
		return match.group(0)
	return chart_label.strip()


def normalize_spaces(text: str) -> str:
	return re.sub(r"\s+", " ", text).strip()


def tokenize_words(text: str) -> List[str]:
	return WORD_PATTERN.findall(text)


def count_numeric_tokens(text: str) -> int:
	return len(NUMERIC_PATTERN.findall(text))


def count_sentences(text: str) -> int:
	parts = [part.strip() for part in SENTENCE_SPLIT_PATTERN.split(text) if part.strip()]
	return len(parts)


@lru_cache(maxsize=50000)
def estimate_syllables_in_word(word: str) -> int:
	clean_word = re.sub(r"[^a-z]", "", word.lower())
	if not clean_word:
		return 0

	phones = pronouncing.phones_for_word(clean_word)
	if phones:
		return max(1, pronouncing.syllable_count(phones[0]))

	hyphenated = HYPHENATOR.inserted(clean_word)
	if hyphenated:
		parts = [part for part in hyphenated.split("-") if part]
		if parts:
			return max(1, len(parts))

	return 1


@lru_cache(maxsize=50000)
def is_word_in_dictionary(word: str) -> bool:
	clean_word = re.sub(r"[^a-z]", "", word.lower())
	if not clean_word:
		return False
	return bool(pronouncing.phones_for_word(clean_word))


def count_dictionary_confidence(text: str) -> float | None:
	words = tokenize_words(text)
	if not words:
		return None
	in_dict = sum(1 for word in words if is_word_in_dictionary(word))
	return round((in_dict / len(words)) * 100, 6)


def count_total_syllables(text: str) -> int:
	total = 0
	for token in tokenize_words(text):
		parts = [part for part in re.split(r"[-']", token) if part]
		for part in parts:
			total += estimate_syllables_in_word(part)
	return total


def count_characters(text: str) -> int:
	return len(re.sub(r"\s", "", text))


def gunning_fog_index(word_count: int, sentence_count: int, polysyllabic_count: int) -> float | None:
	if word_count == 0 or sentence_count == 0:
		return None
	if word_count < 100:
		return None
	words_per_sentence = word_count / sentence_count
	polysyllabic_ratio = (polysyllabic_count / word_count) * 100
	return round(0.4 * (words_per_sentence + 100 * (polysyllabic_ratio / 100)), 6)


def flesch_kincaid_grade(word_count: int, sentence_count: int, syllable_count: int) -> float | None:
	if word_count == 0 or sentence_count == 0:
		return None
	syllables_per_word = syllable_count / word_count
	words_per_sentence = word_count / sentence_count
	return round(0.39 * words_per_sentence + 11.8 * syllables_per_word - 15.59, 6)


def coleman_liau_index(char_count: int, sentence_count: int, word_count: int) -> float | None:
	if word_count == 0:
		return None
	# L = characters per 100 words, S = sentences per 100 words
	L = (char_count / word_count) * 100
	S = (sentence_count / word_count) * 100
	return round(0.0588 * L - 0.296 * S - 15.8, 6)


def count_polysyllabic_words(text: str) -> int:
	count = 0
	for token in tokenize_words(text):
		parts = [part for part in re.split(r"[-']", token) if part]
		syllables = sum(estimate_syllables_in_word(part) for part in parts)
		if syllables >= 3:
			count += 1
	return count


def smog_index(polysyllabic_word_count: int, sentence_count: int) -> float | None:
	if polysyllabic_word_count <= 0:
		return 0.0
	if sentence_count < 0:
		return None
	return round(1.0430 * math.sqrt(polysyllabic_word_count * (30 / (sentence_count + 3.1291))), 6)


def count_number_words(text: str) -> int:
	normalized = text.lower()
	normalized = normalized.replace("—", " ").replace("–", " ")
	normalized = normalized.replace("-", " ")
	tokens = re.findall(r"[a-z]+", normalized)

	count = 0
	index = 0

	while index < len(tokens):
		token = tokens[index]

		if token in ORDINAL_WORDS:
			count += 1
			index += 1
			continue

		if token not in CARDINAL_NUMBER_WORDS:
			index += 1
			continue

		index += 1
		while index < len(tokens):
			current = tokens[index]
			if current in CARDINAL_NUMBER_WORDS:
				index += 1
				continue

			if current == "and" and index + 1 < len(tokens) and tokens[index + 1] in CARDINAL_NUMBER_WORDS:
				index += 1
				continue

			break

		count += 1

	return count


def words_per_number(word_count: int, number_count: int) -> float | None:
	if number_count == 0:
		return None
	return round(word_count / number_count, 6)


def build_description(parts: Iterable[Any]) -> str:
	clean_parts = [normalize_spaces(str(part)) for part in parts if part not in (None, "")]
	return " ".join(part for part in clean_parts if part)


def extract_case_blocks_from_markdown(markdown_text: str) -> List[Dict[str, Any]]:
	cases: List[Dict[str, Any]] = []
	current_case: Dict[str, Any] | None = None
	current_section_key: str | None = None

	for raw_line in markdown_text.splitlines():
		line = raw_line.rstrip()
		header_match = CHART_HEADER_PATTERN.match(line.strip())
		if header_match:
			if current_case is not None:
				for section_name in SECTION_KEYS.values():
					current_case["sections"][section_name] = normalize_spaces(current_case["sections"].get(section_name, ""))
				cases.append(current_case)
			chart_title = header_match.group(1).strip()
			case_id = header_match.group(2).strip()
			current_case = {
				"id": case_id,
				"chart": chart_title,
				"sections": {section_name: "" for section_name in SECTION_KEYS.values()},
			}
			current_section_key = None
			continue

		if current_case is None:
			continue

		if line.strip().startswith("###"):
			section_title = line.strip().removeprefix("###").strip().lower()
			current_section_key = SECTION_KEYS.get(section_title)
			continue

		if current_section_key is None:
			continue

		if line.strip():
			current_case["sections"][current_section_key] += (line.strip() + " ")

	if current_case is not None:
		for section_name in SECTION_KEYS.values():
			current_case["sections"][section_name] = normalize_spaces(current_case["sections"].get(section_name, ""))
		cases.append(current_case)

	return cases


def analyse_row(
	case_entry: Dict[str, Any],
) -> Dict[str, Any]:
	chart_label = str(case_entry.get("chart") or "")
	sections = case_entry.get("sections") or {}
	description_text = build_description([
		sections.get("overview_and_main_message"),
		sections.get("chart_structure"),
		sections.get("relevant_patterns_trends_and_comparisons"),
		sections.get("essential_key_details"),
	])

	word_count = len(tokenize_words(description_text))
	sentence_count = count_sentences(description_text)
	polysyllabic_word_count = count_polysyllabic_words(description_text)
	total_syllables = count_total_syllables(description_text)
	total_chars = count_characters(description_text)
	numeric_token_count = count_numeric_tokens(description_text)
	number_word_count = count_number_words(description_text)
	total_number_mentions = numeric_token_count + number_word_count
	smog = smog_index(polysyllabic_word_count, sentence_count)
	gunning_fog = gunning_fog_index(word_count, sentence_count, polysyllabic_word_count)
	fk_grade = flesch_kincaid_grade(word_count, sentence_count, total_syllables)
	cl_index = coleman_liau_index(total_chars, sentence_count, word_count)
	dict_confidence = count_dictionary_confidence(description_text)
	
	case_id = str(case_entry.get("id") or extract_case_id(chart_label)).strip()

	return {
		"chart": chart_label,
		"id": case_id,
		"overview_and_main_message": sections.get("overview_and_main_message"),
		"chart_structure": sections.get("chart_structure"),
		"relevant_patterns_trends_and_comparisons": sections.get("relevant_patterns_trends_and_comparisons"),
		"essential_key_details": sections.get("essential_key_details"),
		"nd_score": round((total_number_mentions / word_count) * 100, 6) if word_count else None,
		"smog_index": smog,
		"gunning_fog_index": gunning_fog,
		"flesch_kincaid_grade": fk_grade,
		"coleman_liau_index": cl_index,
		"dictionary_confidence": dict_confidence,
		"word_count": word_count,
		"sentence_count": sentence_count,
		"polysyllabic_word_count": polysyllabic_word_count,
		"total_syllables": total_syllables,
		"total_characters": total_chars,
		"numeric_token_count": numeric_token_count,
		"number_word_count": number_word_count,
		"total_number_mentions": total_number_mentions,
		"words_per_number": words_per_number(word_count, total_number_mentions),
		"description_text": description_text,
	}

def load_metrics() -> Dict[str, List[Dict[str, Any]]]:
	"""Load metrics for all configured providers from markdown prompt files."""
	results = {}
	
	for provider_name, provider_dir in PROVIDERS.items():
		provider_results: List[Dict[str, Any]] = []
		markdown_files = sorted(provider_dir.glob("*.md"))
		if not markdown_files:
			raise FileNotFoundError(
				f"No s'ha trobat cap fitxer markdown d'entrada per a {provider_name} a {provider_dir}"
			)

		for markdown_path in markdown_files:
			content = markdown_path.read_text(encoding="utf-8")
			for case_entry in extract_case_blocks_from_markdown(content):
				provider_results.append(analyse_row(case_entry))

		results[provider_name] = provider_results
	
	return results

def build_unified_cases(all_metrics: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
	"""Combine metrics from all providers into unified case records."""
	# Group by case ID
	cases_by_id = {}
	
	for provider_name, cases in all_metrics.items():
		for case in cases:
			case_id = case["id"]
			if case_id not in cases_by_id:
				cases_by_id[case_id] = {"id": case_id, "chart": case["chart"]}
			
			# Add provider-prefixed metrics
			for key, value in case.items():
				if key not in ("id", "chart", "description_text"):
					prefixed_key = f"{provider_name}_{key}"
					cases_by_id[case_id][prefixed_key] = value
	
	return sorted(cases_by_id.values(), key=lambda x: x["id"])


def build_cases_by_id(unified_cases: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
	"""Return a mapping keyed by the numeric case id found in chart titles."""
	return {str(case["id"]): case for case in unified_cases}


def write_csv_unified(unified_cases: List[Dict[str, Any]]) -> None:
	"""Write unified CSV with metrics from all providers."""
	OUTPUT_CSV_UNIFIED.parent.mkdir(parents=True, exist_ok=True)
	
	# Determine all possible headers
	all_keys = set()
	for case in unified_cases:
		all_keys.update(case.keys())
	
	# Score metrics first (in desired order), then remaining data fields alphabetically
	_SCORE_ORDER = [
		"nd_score",
		"smog_index", "gunning_fog_index", "flesch_kincaid_grade", "coleman_liau_index",
		"dictionary_confidence",
	]
	headers = ["id", "chart"]
	remaining = set(k for k in all_keys if k not in ("id", "chart"))
	for metric in _SCORE_ORDER:
		for p in PROVIDERS:
			key = f"{p}_{metric}"
			if key in remaining:
				headers.append(key)
				remaining.discard(key)
	headers.extend(sorted(remaining))
	
	with OUTPUT_CSV_UNIFIED.open("w", encoding="utf-8", newline="") as handle:
		writer = csv.DictWriter(handle, fieldnames=headers)
		writer.writeheader()
		for case in unified_cases:
			writer.writerow({header: case.get(header) for header in headers})


def write_json_unified(unified_cases: List[Dict[str, Any]]) -> None:
	"""Write unified JSON with metrics from all providers."""
	OUTPUT_JSON_UNIFIED.parent.mkdir(parents=True, exist_ok=True)
	payload = {
		"providers": list(PROVIDERS.keys()),
		"cases": build_cases_by_id(unified_cases),
	}
	with OUTPUT_JSON_UNIFIED.open("w", encoding="utf-8") as handle:
		json.dump(payload, handle, ensure_ascii=False, indent=2)

def main() -> None:
	# Load metrics for all providers
	all_metrics = load_metrics()
	
	# Generate unified case data
	unified_cases = build_unified_cases(all_metrics)
	
	# Write unified outputs
	write_csv_unified(unified_cases)
	write_json_unified(unified_cases)
	
	print("\n" + "=" * 80)
	print("FITXERS GENERATS:")
	print("=" * 80)
	print(f"CSV unificat: {OUTPUT_CSV_UNIFIED}")
	print(f"JSON unificat: {OUTPUT_JSON_UNIFIED}")


if __name__ == "__main__":
	main()
