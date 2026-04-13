from __future__ import annotations

import csv
import json
import math
import re
import statistics
from pathlib import Path
from functools import lru_cache
from typing import Any, Dict, Iterable, List

import pronouncing
import pyphen


BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent.parent
LONG_DESCRIPTIONS_DIR = PROJECT_ROOT / "longdescriptions_byprompt"
OUTPUT_DIR = PROJECT_ROOT / "output"
METRICS_DIR = BASE_DIR

# Configuració de proveedors
PROVIDERS = {
	"claude": LONG_DESCRIPTIONS_DIR / "claude",
	"gemini": LONG_DESCRIPTIONS_DIR / "gemini",
	"chatgpt": LONG_DESCRIPTIONS_DIR / "chatgpt",
}

OUTPUT_CSV_UNIFIED = METRICS_DIR / "metriques_unificat.csv"
OUTPUT_JSON_UNIFIED = METRICS_DIR / "metriques_unificat.json"
OUTPUT_CSV_FACTUAL_CLAIMS = METRICS_DIR / "afirmacions_factuals.csv"
OUTPUT_JSON_FACTUAL_CLAIMS = METRICS_DIR / "afirmacions_factuals.json"
OUTPUT_CSV_FACTUAL_CHECK = METRICS_DIR / "verificacio_afirmacions.csv"
OUTPUT_JSON_FACTUAL_CHECK = METRICS_DIR / "verificacio_afirmacions.json"

GROUNDTRUTH_CANONICAL_TABLE = BASE_DIR / "chartsold.json"
LONG_DESCRIPTIONS_COMBINED = BASE_DIR / "long_descriptions_combined.json"

WORD_PATTERN = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
NUMERIC_PATTERN = re.compile(r"(?<![A-Za-z])[+-]?(?:\d+[\.,]?\d*|\d*[\.,]\d+)(?![A-Za-z])")
CASE_ID_PATTERN = re.compile(r"CASE_\d+")
CHART_HEADER_PATTERN = re.compile(r"^##\s+Chart\s+\d+:\s+(.+?)\s*\((\d+)\)\s*$")
# Split sentences on '.', '!', '?' but keep decimal numbers intact (e.g., 38.6)
SENTENCE_SPLIT_PATTERN = re.compile(r"(?:\.(?!\d)|[!?])+")
YEAR_PATTERN = re.compile(r"\b(?:19|20)\d{2}\b")
HYPHENATOR = pyphen.Pyphen(lang="en_US")
SECTION_KEYS = {
	"overview and main message": "overview_and_main_message",
	"chart structure": "chart_structure",
	"relevant patterns, trends, and comparisons": "relevant_patterns_trends_and_comparisons",
	"essential key details": "essential_key_details",
}

PROVIDER_RELEVANT_FIELDS_ORDER = [
	"nd_score",
	"smog_index",
	"gunning_fog_index",
	"flesch_kincaid_grade",
	"coleman_liau_index",
	"dictionary_confidence",
	"verification_supported_rate",
	"verification_supported_claims",
	"verification_contradicted_claims",
	"verification_insufficient_claims",
	"verification_total_claims",
]

PROVIDER_AUX_FIELDS_ORDER = [
	"word_count",
	"sentence_count",
	"words_per_number",
	"total_number_mentions",
	"numeric_token_count",
	"number_word_count",
	"polysyllabic_word_count",
	"total_syllables",
	"total_characters",
	"overview_and_main_message",
	"chart_structure",
	"relevant_patterns_trends_and_comparisons",
	"essential_key_details",
]


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
FACTUAL_HINT_WORDS = {
	"higher",
	"lower",
	"more",
	"less",
	"increase",
	"increased",
	"decrease",
	"decreased",
	"grew",
	"declined",
	"rose",
	"fell",
	"compared",
	"versus",
	"between",
	"from",
	"to",
	"highest",
	"lowest",
	"maximum",
	"minimum",
	"peak",
	"top",
	"rank",
	"share",
	"percent",
	"percentage",
	"total",
	"average",
}
DEFAULT_TOLERANCE = 0.2



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


def split_sentences(text: str) -> List[str]:
	return [part.strip() for part in SENTENCE_SPLIT_PATTERN.split(text) if part.strip()]


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


def detect_claim_type(sentence: str, has_numeric_value: bool) -> str:
	lower = sentence.lower()
	if any(token in lower for token in ("highest", "lowest", "maximum", "minimum", "peak", "top")):
		return "extreme"
	if any(token in lower for token in ("rank", "first", "second", "third")):
		return "ranking"
	if any(token in lower for token in ("higher than", "lower than", "more than", "less than", "compared", "versus")):
		return "comparison"
	if any(token in lower for token in ("increase", "increased", "decrease", "decreased", "grew", "declined", "rose", "fell")):
		return "trend"
	if any(token in lower for token in ("share", "percent", "percentage", "%")):
		return "proportion"
	if has_numeric_value:
		return "value"
	return "other"


def detect_operator(sentence: str) -> str | None:
	lower = sentence.lower()
	if any(token in lower for token in ("higher than", "more than", "greater than", "above", "exceeds")):
		return ">"
	if any(token in lower for token in ("lower than", "less than", "below")):
		return "<"
	if any(token in lower for token in ("equal to", "same as", "equals")):
		return "="
	if any(token in lower for token in ("increase", "increased", "grew", "rose")):
		return "increase"
	if any(token in lower for token in ("decrease", "decreased", "declined", "fell")):
		return "decrease"
	if any(token in lower for token in ("highest", "maximum", "peak", "top")):
		return "max"
	if any(token in lower for token in ("lowest", "minimum", "bottom")):
		return "min"
	return None


def is_factual_candidate(sentence: str, has_numeric_value: bool) -> bool:
	lower = sentence.lower()
	if has_numeric_value:
		return True
	return any(hint in lower for hint in FACTUAL_HINT_WORDS)


def assess_verifiability(claim_type: str, numeric_values: List[str], periods: List[str]) -> str:
	if numeric_values and (periods or claim_type in {"comparison", "extreme", "value", "proportion"}):
		return "high"
	if claim_type in {"comparison", "trend", "ranking", "extreme", "proportion"}:
		return "medium"
	return "low"


def extract_factual_claims_for_case(provider: str, case: Dict[str, Any]) -> List[Dict[str, Any]]:
	claims: List[Dict[str, Any]] = []
	case_id = str(case.get("id", "")).strip()
	chart = str(case.get("chart", "")).strip()
	section_order = [
		"overview_and_main_message",
		"chart_structure",
		"relevant_patterns_trends_and_comparisons",
		"essential_key_details",
	]

	for section_name in section_order:
		section_text = normalize_spaces(str(case.get(section_name) or ""))
		if not section_text:
			continue

		for sentence in split_sentences(section_text):
			sentence = normalize_spaces(sentence)
			if not sentence:
				continue

			numeric_values = [value.replace(",", ".") for value in NUMERIC_PATTERN.findall(sentence)]
			periods = YEAR_PATTERN.findall(sentence)
			has_numeric_value = len(numeric_values) > 0
			if not is_factual_candidate(sentence, has_numeric_value):
				continue

			claim_type = detect_claim_type(sentence, has_numeric_value)
			operator = detect_operator(sentence)
			verifiability = assess_verifiability(claim_type, numeric_values, periods)
			claim_id = f"{provider}_{case_id}_{len(claims) + 1:03d}"

			claims.append(
				{
					"claim_id": claim_id,
					"provider": provider,
					"case_id": case_id,
					"chart": chart,
					"section": section_name,
					"claim_text": sentence,
					"claim_type": claim_type,
					"operator": operator,
					"value_expected": numeric_values,
					"period": periods,
					"evidence_span": sentence,
					"verifiability": verifiability,
					"needs_human_review": verifiability != "high",
				}
			)

	return claims


def safe_float(value: Any) -> float | None:
	if isinstance(value, (int, float)):
		return float(value)
	if isinstance(value, str):
		candidate = value.strip().replace(",", ".")
		try:
			return float(candidate)
		except ValueError:
			return None
	return None


def normalize_case_numeric_id(case_id: str) -> str:
	case_id = (case_id or "").strip()
	match = CASE_ID_PATTERN.search(case_id)
	if match:
		return match.group(0).removeprefix("CASE_")
	if case_id.isdigit():
		return case_id.zfill(5)
	return case_id


def collect_chart_values(chart_entry: Dict[str, Any]) -> List[float]:
	values: List[float] = []
	rows = ((chart_entry.get("data") or {}).get("source") or {}).get("values") or []
	for row in rows:
		if not isinstance(row, dict):
			continue
		for key in ("valor", "value"):
			numeric = safe_float(row.get(key))
			if numeric is not None:
				values.append(numeric)
				break
	return values


def collect_unique_field_count(chart_entry: Dict[str, Any], field_candidates: List[str]) -> int:
	rows = ((chart_entry.get("data") or {}).get("source") or {}).get("values") or []
	values = set()
	for row in rows:
		if not isinstance(row, dict):
			continue
		for field in field_candidates:
			value = row.get(field)
			if value not in (None, ""):
				values.add(str(value).strip().lower())
				break
	return len(values)


def build_groundtruth_stats(values: List[float]) -> Dict[str, float | None]:
	if not values:
		return {
			"count": 0,
			"min": None,
			"max": None,
			"mean": None,
			"median": None,
			"std_dev": None,
			"range": None,
		}
	std_dev = None
	if len(values) >= 2:
		std_dev = statistics.pstdev(values)
	return {
		"count": float(len(values)),
		"min": min(values),
		"max": max(values),
		"mean": statistics.fmean(values),
		"median": statistics.median(values),
		"std_dev": std_dev,
		"range": max(values) - min(values),
	}


def extract_stats_from_summary_entry(summary_entry: Dict[str, Any]) -> Dict[str, float | None] | None:
	numeric_summary = summary_entry.get("resum_numeric") or {}
	global_stats = numeric_summary.get("estadistica_global") or {}
	if not isinstance(global_stats, dict) or not global_stats:
		return None

	return {
		"count": safe_float(global_stats.get("n")),
		"min": safe_float(global_stats.get("min")),
		"max": safe_float(global_stats.get("max")),
		"mean": safe_float(global_stats.get("mitjana")),
		"median": safe_float(global_stats.get("mediana")),
		"std_dev": safe_float(global_stats.get("desviacio_estandard")),
		"range": safe_float(global_stats.get("rang")),
	}



def load_groundtruth_cases() -> Dict[str, Dict[str, Any]]:
	if not GROUNDTRUTH_CANONICAL_TABLE.exists():
		raise FileNotFoundError(
			f"No s'ha trobat el fitxer de ground truth: {GROUNDTRUTH_CANONICAL_TABLE}"
		)

	entries = json.loads(GROUNDTRUTH_CANONICAL_TABLE.read_text(encoding="utf-8-sig"))
	if not isinstance(entries, list):
		raise ValueError(f"Format de ground truth invàlid a {GROUNDTRUTH_CANONICAL_TABLE}: s'esperava una llista")

	groundtruth_by_case: Dict[str, Dict[str, Any]] = {}
	for entry in entries:
		if not isinstance(entry, dict):
			continue
		raw_case_id = str(entry.get("id") or (entry.get("source_case") or {}).get("case_id") or "")
		case_id = normalize_case_numeric_id(raw_case_id)
		if not case_id:
			continue

		chart_values = collect_chart_values(entry)
		stats = build_groundtruth_stats(chart_values)
		groundtruth_by_case[case_id] = {
			"case_id": case_id,
			"source": {
				"table": GROUNDTRUTH_CANONICAL_TABLE.name,
				"stats": None,
			},
			"values": chart_values,
			"category_count": collect_unique_field_count(entry, ["categoria", "category", "x"]),
			"series_count": collect_unique_field_count(entry, ["serie", "series", "color"]),
			"stats": stats,
			"title": str(entry.get("title") or ""),
		}

	return groundtruth_by_case


def approx_equal(a: float, b: float, tolerance: float = DEFAULT_TOLERANCE) -> bool:
	return abs(a - b) <= tolerance


def verify_claim_against_groundtruth(
	claim: Dict[str, Any],
	groundtruth_case: Dict[str, Any] | None,
) -> Dict[str, Any]:
	result = dict(claim)

	if groundtruth_case is None:
		result.update(
			{
				"verification_status": "insufficient_data",
				"verification_reason": "case_not_found_in_groundtruth",
				"gt_source": None,
				"gt_source_table": None,
				"gt_source_stats": None,
				"gt_values_count": None,
				"gt_min": None,
				"gt_max": None,
				"gt_mean": None,
				"gt_median": None,
				"gt_std_dev": None,
				"matched_values": [],
			}
		)
		return result

	values = [safe_float(item) for item in (claim.get("value_expected") or [])]
	values = [value for value in values if value is not None]
	stats = groundtruth_case["stats"]
	gt_values = groundtruth_case["values"]

	candidate_numbers = []
	candidate_numbers.extend(gt_values)
	for key in ("min", "max", "mean", "median", "std_dev", "range", "count"):
		numeric = safe_float(stats.get(key))
		if numeric is not None:
			candidate_numbers.append(numeric)
	category_count = int(groundtruth_case.get("category_count") or 0)
	series_count = int(groundtruth_case.get("series_count") or 0)
	if category_count:
		candidate_numbers.append(float(category_count))
	if series_count:
		candidate_numbers.append(float(series_count))

	matched_values: List[float] = []
	for expected in values:
		if any(approx_equal(expected, candidate) for candidate in candidate_numbers):
			matched_values.append(expected)

	operator = str(claim.get("operator") or "")
	status = "insufficient_data"
	reason = "no_numeric_values_in_claim"

	if values:
		if operator == "max" and safe_float(stats.get("max")) is not None:
			gt_max = float(stats["max"])
			if any(approx_equal(value, gt_max) for value in values):
				status = "supported"
				reason = "checked_against_groundtruth_max"
			elif matched_values:
				status = "insufficient_data"
				reason = "composite_extreme_not_fully_checkable"
			else:
				status = "contradicted"
				reason = "checked_against_groundtruth_max"
		elif operator == "min" and safe_float(stats.get("min")) is not None:
			gt_min = float(stats["min"])
			if any(approx_equal(value, gt_min) for value in values):
				status = "supported"
				reason = "checked_against_groundtruth_min"
			elif matched_values:
				status = "insufficient_data"
				reason = "composite_extreme_not_fully_checkable"
			else:
				status = "contradicted"
				reason = "checked_against_groundtruth_min"
		else:
			status = "supported" if len(matched_values) == len(values) else "contradicted"
			reason = "all_numeric_values_matched" if status == "supported" else "some_numeric_values_not_found"

	result.update(
		{
			"verification_status": status,
			"verification_reason": reason,
			"gt_source": groundtruth_case.get("source"),
			"gt_source_table": (groundtruth_case.get("source") or {}).get("table"),
			"gt_source_stats": (groundtruth_case.get("source") or {}).get("stats"),
			"gt_values_count": int(stats.get("count") or 0),
			"gt_category_count": category_count,
			"gt_series_count": series_count,
			"gt_min": stats.get("min"),
			"gt_max": stats.get("max"),
			"gt_mean": stats.get("mean"),
			"gt_median": stats.get("median"),
			"gt_std_dev": stats.get("std_dev"),
			"matched_values": matched_values,
		}
	)
	return result


def verify_factual_claims(claims: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
	groundtruth_by_case = load_groundtruth_cases()
	verified: List[Dict[str, Any]] = []

	for claim in claims:
		case_id = normalize_case_numeric_id(str(claim.get("case_id") or ""))
		gt_case = groundtruth_by_case.get(case_id)
		verified.append(verify_claim_against_groundtruth(claim, gt_case))

	return verified


def build_factual_claims(all_metrics: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
	claims: List[Dict[str, Any]] = []
	for provider, provider_cases in all_metrics.items():
		for case in provider_cases:
			claims.extend(extract_factual_claims_for_case(provider, case))

	return sorted(claims, key=lambda item: (str(item["case_id"]), item["provider"], item["claim_id"]))


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
	"""Load metrics for all configured providers from long_descriptions_combined.json."""
	if not LONG_DESCRIPTIONS_COMBINED.exists():
		raise FileNotFoundError(f"No s'ha trobat el fitxer: {LONG_DESCRIPTIONS_COMBINED}")
	if not GROUNDTRUTH_CANONICAL_TABLE.exists():
		raise FileNotFoundError(f"No s'ha trobat el fitxer: {GROUNDTRUTH_CANONICAL_TABLE}")

	combined = json.loads(LONG_DESCRIPTIONS_COMBINED.read_text(encoding="utf-8-sig"))
	chartsold = json.loads(GROUNDTRUTH_CANONICAL_TABLE.read_text(encoding="utf-8-sig"))

	titles_by_id: Dict[str, str] = {}
	for entry in chartsold:
		raw_id = str(entry.get("id") or "")
		numeric_id = raw_id.removeprefix("CASE_")
		titles_by_id[numeric_id] = str(entry.get("title") or "")

	results: Dict[str, List[Dict[str, Any]]] = {provider: [] for provider in PROVIDERS}

	for entry in combined:
		case_id = str(entry.get("id") or "").strip()
		chart_title = titles_by_id.get(case_id, "")

		for provider in PROVIDERS:
			provider_data = (entry.get(provider) or {}) if isinstance(entry, dict) else {}
			case_entry = {
				"id": case_id,
				"chart": chart_title,
				"sections": {
					"overview_and_main_message": str(provider_data.get("overview_and_main_message") or ""),
					"chart_structure": str(provider_data.get("chart_structure") or ""),
					"relevant_patterns_trends_and_comparisons": str(provider_data.get("relevant_patterns_trends_and_comparisons") or ""),
					"essential_key_details": str(provider_data.get("essential_key_details") or ""),
				},
			}
			results[provider].append(analyse_row(case_entry))

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


def add_verification_summary_to_unified_cases(
	unified_cases: List[Dict[str, Any]],
	verified_claims: List[Dict[str, Any]],
) -> None:
	"""Enrich unified case records with compact verification summaries."""
	status_keys = ["supported", "contradicted", "insufficient_data"]
	legacy_global_keys = [
		"verification_supported_rate",
		"verification_supported_claims",
		"verification_contradicted_claims",
		"verification_insufficient_claims",
		"verification_total_claims",
	]

	by_case_provider: Dict[str, Dict[str, Dict[str, int]]] = {}
	for claim in verified_claims:
		case_id = str(claim.get("case_id") or "").strip()
		provider = str(claim.get("provider") or "").strip()
		status = str(claim.get("verification_status") or "insufficient_data").strip()
		if not case_id or not provider:
			continue
		if status not in status_keys:
			status = "insufficient_data"

		provider_bucket = by_case_provider.setdefault(case_id, {}).setdefault(
			provider,
			{"total": 0, "supported": 0, "contradicted": 0, "insufficient_data": 0},
		)
		provider_bucket["total"] += 1
		provider_bucket[status] += 1

	for case in unified_cases:
		case_id = str(case.get("id") or "").strip()
		for legacy_key in legacy_global_keys:
			case.pop(legacy_key, None)

		provider_stats = by_case_provider.get(case_id, {})

		overall_total = 0
		overall_supported = 0
		overall_contradicted = 0
		overall_insufficient = 0

		for provider in PROVIDERS:
			stats = provider_stats.get(provider) or {
				"total": 0,
				"supported": 0,
				"contradicted": 0,
				"insufficient_data": 0,
			}
			total = stats["total"]
			supported = stats["supported"]
			contradicted = stats["contradicted"]
			insufficient = stats["insufficient_data"]

			case[f"{provider}_verification_total_claims"] = total
			case[f"{provider}_verification_supported_claims"] = supported
			case[f"{provider}_verification_contradicted_claims"] = contradicted
			case[f"{provider}_verification_insufficient_claims"] = insufficient
			case[f"{provider}_verification_supported_rate"] = (
				round((supported / total) * 100, 6) if total > 0 else None
			)

			overall_total += total
			overall_supported += supported
			overall_contradicted += contradicted
			overall_insufficient += insufficient



def build_cases_by_id(unified_cases: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
	"""Return a mapping keyed by the numeric case id found in chart titles."""
	return {str(case["id"]): case for case in unified_cases}


def reorder_unified_case(case: Dict[str, Any]) -> Dict[str, Any]:
	"""Group fields by provider, showing relevant metrics before auxiliary fields."""
	ordered: Dict[str, Any] = {}
	ordered["id"] = case.get("id")
	ordered["chart"] = case.get("chart")

	for provider in PROVIDERS:
		for field in PROVIDER_RELEVANT_FIELDS_ORDER:
			key = f"{provider}_{field}"
			if key in case:
				ordered[key] = case.get(key)

		for field in PROVIDER_AUX_FIELDS_ORDER:
			key = f"{provider}_{field}"
			if key in case:
				ordered[key] = case.get(key)

		provider_remaining = sorted(
			key
			for key in case
			if key.startswith(f"{provider}_") and key not in ordered
		)
		for key in provider_remaining:
			ordered[key] = case.get(key)

	for key in sorted(key for key in case if key not in ordered):
		ordered[key] = case.get(key)

	return ordered


def reorder_unified_cases(unified_cases: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
	return [reorder_unified_case(case) for case in unified_cases]


def write_csv_unified(unified_cases: List[Dict[str, Any]]) -> None:
	"""Write unified CSV with metrics from all providers."""
	OUTPUT_CSV_UNIFIED.parent.mkdir(parents=True, exist_ok=True)

	ordered_cases = reorder_unified_cases(unified_cases)
	headers: List[str] = []
	for case in ordered_cases:
		for key in case.keys():
			if key not in headers:
				headers.append(key)
	
	with OUTPUT_CSV_UNIFIED.open("w", encoding="utf-8", newline="") as handle:
		writer = csv.DictWriter(handle, fieldnames=headers)
		writer.writeheader()
		for case in ordered_cases:
			writer.writerow({header: case.get(header) for header in headers})


def write_json_unified(unified_cases: List[Dict[str, Any]]) -> None:
	"""Write unified JSON with metrics from all providers."""
	OUTPUT_JSON_UNIFIED.parent.mkdir(parents=True, exist_ok=True)
	ordered_cases = reorder_unified_cases(unified_cases)
	payload = {
		"providers": list(PROVIDERS.keys()),
		"cases": build_cases_by_id(ordered_cases),
	}
	with OUTPUT_JSON_UNIFIED.open("w", encoding="utf-8") as handle:
		json.dump(payload, handle, ensure_ascii=False, indent=2)


def write_csv_factual_claims(claims: List[Dict[str, Any]]) -> None:
	OUTPUT_CSV_FACTUAL_CLAIMS.parent.mkdir(parents=True, exist_ok=True)
	headers = [
		"claim_id",
		"provider",
		"case_id",
		"chart",
		"section",
		"claim_text",
		"claim_type",
		"operator",
		"value_expected",
		"period",
		"evidence_span",
		"verifiability",
		"needs_human_review",
	]
	with OUTPUT_CSV_FACTUAL_CLAIMS.open("w", encoding="utf-8", newline="") as handle:
		writer = csv.DictWriter(handle, fieldnames=headers)
		writer.writeheader()
		for claim in claims:
			row = dict(claim)
			row["value_expected"] = "|".join(claim.get("value_expected") or [])
			row["period"] = "|".join(claim.get("period") or [])
			writer.writerow({header: row.get(header) for header in headers})


def write_json_factual_claims(claims: List[Dict[str, Any]]) -> None:
	OUTPUT_JSON_FACTUAL_CLAIMS.parent.mkdir(parents=True, exist_ok=True)
	payload = {
		"providers": list(PROVIDERS.keys()),
		"claims": claims,
	}
	with OUTPUT_JSON_FACTUAL_CLAIMS.open("w", encoding="utf-8") as handle:
		json.dump(payload, handle, ensure_ascii=False, indent=2)


def write_csv_factual_check(verified_claims: List[Dict[str, Any]]) -> None:
	OUTPUT_CSV_FACTUAL_CHECK.parent.mkdir(parents=True, exist_ok=True)
	headers = [
		"claim_id",
		"provider",
		"case_id",
		"chart",
		"section",
		"claim_text",
		"claim_type",
		"operator",
		"value_expected",
		"period",
		"verifiability",
		"verification_status",
		"verification_reason",
		"gt_source",
		"gt_source_table",
		"gt_source_stats",
		"gt_values_count",
		"gt_category_count",
		"gt_series_count",
		"gt_min",
		"gt_max",
		"gt_mean",
		"gt_median",
		"gt_std_dev",
		"matched_values",
	]
	with OUTPUT_CSV_FACTUAL_CHECK.open("w", encoding="utf-8", newline="") as handle:
		writer = csv.DictWriter(handle, fieldnames=headers)
		writer.writeheader()
		for claim in verified_claims:
			row = dict(claim)
			row["value_expected"] = "|".join(str(item) for item in (claim.get("value_expected") or []))
			row["period"] = "|".join(str(item) for item in (claim.get("period") or []))
			row["matched_values"] = "|".join(str(item) for item in (claim.get("matched_values") or []))
			writer.writerow({header: row.get(header) for header in headers})


def write_json_factual_check(verified_claims: List[Dict[str, Any]]) -> None:
	OUTPUT_JSON_FACTUAL_CHECK.parent.mkdir(parents=True, exist_ok=True)
	payload = {
		"groundtruth": {
			"canonical_table": str(GROUNDTRUTH_CANONICAL_TABLE),
		},
		"claims": verified_claims,
	}
	with OUTPUT_JSON_FACTUAL_CHECK.open("w", encoding="utf-8") as handle:
		json.dump(payload, handle, ensure_ascii=False, indent=2)

def main() -> None:
	# Load metrics for all providers
	all_metrics = load_metrics()
	
	# Generate unified case data
	unified_cases = build_unified_cases(all_metrics)
	factual_claims = build_factual_claims(all_metrics)
	verified_claims = verify_factual_claims(factual_claims)
	add_verification_summary_to_unified_cases(unified_cases, verified_claims)
	
	# Write unified outputs
	write_csv_unified(unified_cases)
	write_json_unified(unified_cases)
	write_csv_factual_claims(factual_claims)
	write_json_factual_claims(factual_claims)
	write_csv_factual_check(verified_claims)
	write_json_factual_check(verified_claims)
	
	print("\n" + "=" * 80)
	print("FITXERS GENERATS:")
	print("=" * 80)
	print(f"CSV unificat: {OUTPUT_CSV_UNIFIED}")
	print(f"JSON unificat: {OUTPUT_JSON_UNIFIED}")
	print(f"CSV afirmacions factuals: {OUTPUT_CSV_FACTUAL_CLAIMS}")
	print(f"JSON afirmacions factuals: {OUTPUT_JSON_FACTUAL_CLAIMS}")
	print(f"CSV verificacio afirmacions: {OUTPUT_CSV_FACTUAL_CHECK}")
	print(f"JSON verificacio afirmacions: {OUTPUT_JSON_FACTUAL_CHECK}")


if __name__ == "__main__":
	main()
