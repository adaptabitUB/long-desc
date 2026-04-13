from __future__ import annotations

import csv
import json
import math
import re
import statistics
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List

import pronouncing
import pyphen


BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent.parent
LONG_DESCRIPTIONS_DIR = PROJECT_ROOT / "longdescriptions_byprompt"
OUTPUT_DIR = PROJECT_ROOT / "output"
METRICS_DIR = OUTPUT_DIR / "metriques"

OPENAI_DIR = LONG_DESCRIPTIONS_DIR / "openai"
GROUNDTRUTH_JSON = OUTPUT_DIR / "charts.json"

OUTPUT_CSV_UNIFIED = METRICS_DIR / "metriques_openai_unificat.csv"
OUTPUT_JSON_UNIFIED = METRICS_DIR / "metriques_openai_unificat.json"
OUTPUT_CSV_FACTUAL_CLAIMS = METRICS_DIR / "afirmacions_factuals_openai.csv"
OUTPUT_JSON_FACTUAL_CLAIMS = METRICS_DIR / "afirmacions_factuals_openai.json"
OUTPUT_CSV_FACTUAL_CHECK = METRICS_DIR / "verificacio_afirmacions_openai.csv"
OUTPUT_JSON_FACTUAL_CHECK = METRICS_DIR / "verificacio_afirmacions_openai.json"

WORD_PATTERN = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
NUMERIC_PATTERN = re.compile(r"(?<![A-Za-z])[+-]?(?:\d+[\.,]?\d*|\d*[\.,]\d+)(?![A-Za-z])")
CHART_HEADER_PATTERN = re.compile(r"^##\s+Chart\s+(\d+)\s*:\s*(.+?)\s*$")
SENTENCE_SPLIT_PATTERN = re.compile(r"(?:\.(?!\d)|[!?])+")
YEAR_PATTERN = re.compile(r"\b(?:19|20)\d{2}\b")
HYPHENATOR = pyphen.Pyphen(lang="en_US")

SECTION_KEYS = {
    "overview and main message": "overview_and_main_message",
    "chart structure": "chart_structure",
    "relevant patterns, trends, and comparisons": "relevant_patterns_trends_and_comparisons",
    "essential key details": "essential_key_details",
}

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


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def tokenize_words(text: str) -> List[str]:
    return WORD_PATTERN.findall(text)


def count_numeric_tokens(text: str) -> int:
    return len(NUMERIC_PATTERN.findall(text))


def count_sentences(text: str) -> int:
    return len([part.strip() for part in SENTENCE_SPLIT_PATTERN.split(text) if part.strip()])


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


def count_polysyllabic_words(text: str) -> int:
    count = 0
    for token in tokenize_words(text):
        parts = [part for part in re.split(r"[-']", token) if part]
        syllables = sum(estimate_syllables_in_word(part) for part in parts)
        if syllables >= 3:
            count += 1
    return count


def gunning_fog_index(word_count: int, sentence_count: int, polysyllabic_count: int) -> float | None:
    if word_count == 0 or sentence_count == 0 or word_count < 100:
        return None
    words_per_sentence = word_count / sentence_count
    polysyllabic_ratio = (polysyllabic_count / word_count) * 100
    return round(0.4 * (words_per_sentence + polysyllabic_ratio), 6)


def flesch_kincaid_grade(word_count: int, sentence_count: int, syllable_count: int) -> float | None:
    if word_count == 0 or sentence_count == 0:
        return None
    return round(0.39 * (word_count / sentence_count) + 11.8 * (syllable_count / word_count) - 15.59, 6)


def coleman_liau_index(char_count: int, sentence_count: int, word_count: int) -> float | None:
    if word_count == 0:
        return None
    l_val = (char_count / word_count) * 100
    s_val = (sentence_count / word_count) * 100
    return round(0.0588 * l_val - 0.296 * s_val - 15.8, 6)


def smog_index(polysyllabic_word_count: int, sentence_count: int) -> float | None:
    if polysyllabic_word_count <= 0:
        return 0.0
    return round(1.0430 * math.sqrt(polysyllabic_word_count * (30 / (sentence_count + 3.1291))), 6)


def count_number_words(text: str) -> int:
    normalized = text.lower().replace("—", " ").replace("–", " ").replace("-", " ")
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
    if case_id.startswith("CASE_"):
        return case_id.removeprefix("CASE_")
    if case_id.isdigit():
        return case_id.zfill(5)
    return case_id


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
            chart_number = header_match.group(1).strip()
            chart_title = header_match.group(2).strip()
            current_case = {
                "id": str(chart_number).zfill(5),
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
            current_case["sections"][current_section_key] += line.strip() + " "

    if current_case is not None:
        for section_name in SECTION_KEYS.values():
            current_case["sections"][section_name] = normalize_spaces(current_case["sections"].get(section_name, ""))
        cases.append(current_case)

    return cases


def analyse_case(case_entry: Dict[str, Any]) -> Dict[str, Any]:
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

    return {
        "id": str(case_entry.get("id") or "").strip(),
        "chart": str(case_entry.get("chart") or "").strip(),
        "overview_and_main_message": sections.get("overview_and_main_message"),
        "chart_structure": sections.get("chart_structure"),
        "relevant_patterns_trends_and_comparisons": sections.get("relevant_patterns_trends_and_comparisons"),
        "essential_key_details": sections.get("essential_key_details"),
        "nd_score": round((total_number_mentions / word_count) * 100, 6) if word_count else None,
        "smog_index": smog_index(polysyllabic_word_count, sentence_count),
        "gunning_fog_index": gunning_fog_index(word_count, sentence_count, polysyllabic_word_count),
        "flesch_kincaid_grade": flesch_kincaid_grade(word_count, sentence_count, total_syllables),
        "coleman_liau_index": coleman_liau_index(total_chars, sentence_count, word_count),
        "dictionary_confidence": count_dictionary_confidence(description_text),
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


def load_openai_metrics() -> List[Dict[str, Any]]:
    markdown_files = sorted(OPENAI_DIR.glob("*.md"))
    if not markdown_files:
        raise FileNotFoundError(f"No s'ha trobat cap fitxer markdown d'entrada a {OPENAI_DIR}")

    cases: List[Dict[str, Any]] = []
    for markdown_path in markdown_files:
        content = markdown_path.read_text(encoding="utf-8")
        for case_entry in extract_case_blocks_from_markdown(content):
            cases.append(analyse_case(case_entry))

    return sorted(cases, key=lambda c: c["id"])


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
    if any(token in lower for token in ("highest", "maximum", "peak", "top")):
        return "max"
    if any(token in lower for token in ("lowest", "minimum", "bottom")):
        return "min"
    if any(token in lower for token in ("higher than", "more than", "greater than", "above", "exceeds")):
        return ">"
    if any(token in lower for token in ("lower than", "less than", "below")):
        return "<"
    return None


def is_factual_candidate(sentence: str, has_numeric_value: bool) -> bool:
    if has_numeric_value:
        return True
    lower = sentence.lower()
    return any(hint in lower for hint in FACTUAL_HINT_WORDS)


def assess_verifiability(claim_type: str, numeric_values: List[str], periods: List[str]) -> str:
    if numeric_values and (periods or claim_type in {"comparison", "extreme", "value", "proportion"}):
        return "high"
    if claim_type in {"comparison", "trend", "ranking", "extreme", "proportion"}:
        return "medium"
    return "low"


def build_factual_claims(cases: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    claims: List[Dict[str, Any]] = []
    for case in cases:
        case_id = str(case.get("id") or "")
        chart = str(case.get("chart") or "")
        for section_name in SECTION_KEYS.values():
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
                claims.append(
                    {
                        "claim_id": f"openai_{case_id}_{len(claims) + 1:04d}",
                        "provider": "openai",
                        "case_id": case_id,
                        "chart": chart,
                        "section": section_name,
                        "claim_text": sentence,
                        "claim_type": claim_type,
                        "operator": detect_operator(sentence),
                        "value_expected": numeric_values,
                        "period": periods,
                        "evidence_span": sentence,
                        "verifiability": assess_verifiability(claim_type, numeric_values, periods),
                        "needs_human_review": assess_verifiability(claim_type, numeric_values, periods) != "high",
                    }
                )
    return claims


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
        return {"count": 0, "min": None, "max": None, "mean": None, "median": None, "std_dev": None, "range": None}
    std_dev = statistics.pstdev(values) if len(values) >= 2 else None
    return {
        "count": float(len(values)),
        "min": min(values),
        "max": max(values),
        "mean": statistics.fmean(values),
        "median": statistics.median(values),
        "std_dev": std_dev,
        "range": max(values) - min(values),
    }


def extract_stats_from_chart_entry(entry: Dict[str, Any], fallback_values: List[float]) -> Dict[str, float | None]:
    numeric_summary = entry.get("numeric_summary") or {}
    global_stats = numeric_summary.get("global_stats") or {}
    if not isinstance(global_stats, dict) or not global_stats:
        return build_groundtruth_stats(fallback_values)

    return {
        "count": safe_float(global_stats.get("n")),
        "min": safe_float(global_stats.get("min")),
        "max": safe_float(global_stats.get("max")),
        "mean": safe_float(global_stats.get("mean")),
        "median": safe_float(global_stats.get("median")),
        "std_dev": safe_float(global_stats.get("standard_deviation")),
        "range": safe_float(global_stats.get("range")),
    }


def load_groundtruth_cases() -> Dict[str, Dict[str, Any]]:
    if not GROUNDTRUTH_JSON.exists():
        raise FileNotFoundError(f"No s'ha trobat el fitxer de ground truth: {GROUNDTRUTH_JSON}")

    entries = json.loads(GROUNDTRUTH_JSON.read_text(encoding="utf-8"))
    if not isinstance(entries, list):
        raise ValueError(f"Format de ground truth invàlid a {GROUNDTRUTH_JSON}: s'esperava una llista")

    by_case: Dict[str, Dict[str, Any]] = {}
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        raw_case_id = str(entry.get("id") or (entry.get("source_case") or {}).get("case_id") or "")
        case_id = normalize_case_numeric_id(raw_case_id)
        if not case_id:
            continue
        values = collect_chart_values(entry)
        stats = extract_stats_from_chart_entry(entry, values)
        by_case[case_id] = {
            "case_id": case_id,
            "source": {"table": GROUNDTRUTH_JSON.name, "stats": "numeric_summary.global_stats"},
            "values": values,
            "category_count": collect_unique_field_count(entry, ["categoria", "category", "x"]),
            "series_count": collect_unique_field_count(entry, ["serie", "series", "color"]),
            "stats": stats,
        }
    return by_case


def approx_equal(a: float, b: float, tolerance: float = 0.2) -> bool:
    return abs(a - b) <= tolerance


def verify_claim_against_groundtruth(claim: Dict[str, Any], gt_case: Dict[str, Any] | None) -> Dict[str, Any]:
    result = dict(claim)
    if gt_case is None:
        result.update({
            "verification_status": "insufficient_data",
            "verification_reason": "case_not_found_in_groundtruth",
            "gt_source": None,
            "gt_values_count": None,
            "gt_min": None,
            "gt_max": None,
            "gt_mean": None,
            "gt_median": None,
            "gt_std_dev": None,
            "matched_values": [],
        })
        return result

    values = [safe_float(item) for item in (claim.get("value_expected") or [])]
    values = [value for value in values if value is not None]
    stats = gt_case["stats"]

    candidate_numbers = list(gt_case["values"])
    for key in ("min", "max", "mean", "median", "std_dev", "range", "count"):
        numeric = safe_float(stats.get(key))
        if numeric is not None:
            candidate_numbers.append(numeric)

    category_count = int(gt_case.get("category_count") or 0)
    series_count = int(gt_case.get("series_count") or 0)
    if category_count:
        candidate_numbers.append(float(category_count))
    if series_count:
        candidate_numbers.append(float(series_count))

    matched = [expected for expected in values if any(approx_equal(expected, candidate) for candidate in candidate_numbers)]
    status = "insufficient_data"
    reason = "no_numeric_values_in_claim"
    operator = str(claim.get("operator") or "")

    if values:
        if operator == "max" and safe_float(stats.get("max")) is not None:
            gt_max = float(stats["max"])
            if any(approx_equal(v, gt_max) for v in values):
                status = "supported"
                reason = "checked_against_groundtruth_max"
            elif matched:
                status = "insufficient_data"
                reason = "composite_extreme_not_fully_checkable"
            else:
                status = "contradicted"
                reason = "checked_against_groundtruth_max"
        elif operator == "min" and safe_float(stats.get("min")) is not None:
            gt_min = float(stats["min"])
            if any(approx_equal(v, gt_min) for v in values):
                status = "supported"
                reason = "checked_against_groundtruth_min"
            elif matched:
                status = "insufficient_data"
                reason = "composite_extreme_not_fully_checkable"
            else:
                status = "contradicted"
                reason = "checked_against_groundtruth_min"
        else:
            status = "supported" if len(matched) == len(values) else "contradicted"
            reason = "all_numeric_values_matched" if status == "supported" else "some_numeric_values_not_found"

    result.update({
        "verification_status": status,
        "verification_reason": reason,
        "gt_source": gt_case.get("source"),
        "gt_values_count": int(stats.get("count") or 0),
        "gt_category_count": category_count,
        "gt_series_count": series_count,
        "gt_min": stats.get("min"),
        "gt_max": stats.get("max"),
        "gt_mean": stats.get("mean"),
        "gt_median": stats.get("median"),
        "gt_std_dev": stats.get("std_dev"),
        "matched_values": matched,
    })
    return result


def verify_factual_claims(claims: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    gt = load_groundtruth_cases()
    verified: List[Dict[str, Any]] = []
    for claim in claims:
        case_id = normalize_case_numeric_id(str(claim.get("case_id") or ""))
        verified.append(verify_claim_against_groundtruth(claim, gt.get(case_id)))
    return verified


def add_verification_summary_to_cases(cases: List[Dict[str, Any]], verified_claims: List[Dict[str, Any]]) -> None:
    by_case: Dict[str, Dict[str, int]] = {}
    for claim in verified_claims:
        case_id = str(claim.get("case_id") or "").strip()
        status = str(claim.get("verification_status") or "insufficient_data")
        bucket = by_case.setdefault(case_id, {"total": 0, "supported": 0, "contradicted": 0, "insufficient_data": 0})
        bucket["total"] += 1
        if status not in bucket:
            status = "insufficient_data"
        bucket[status] += 1

    for case in cases:
        stats = by_case.get(str(case.get("id") or ""), {"total": 0, "supported": 0, "contradicted": 0, "insufficient_data": 0})
        total = stats["total"]
        case["verification_total_claims"] = total
        case["verification_supported_claims"] = stats["supported"]
        case["verification_contradicted_claims"] = stats["contradicted"]
        case["verification_insufficient_claims"] = stats["insufficient_data"]
        case["verification_supported_rate"] = round((stats["supported"] / total) * 100, 6) if total > 0 else None


def build_cases_by_id(cases: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {str(case["id"]): case for case in cases}


def write_csv(path: Path, rows: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers: List[str] = []
    for row in rows:
        for key in row.keys():
            if key not in headers:
                headers.append(key)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            record = dict(row)
            if isinstance(record.get("value_expected"), list):
                record["value_expected"] = "|".join(str(v) for v in record["value_expected"])
            if isinstance(record.get("period"), list):
                record["period"] = "|".join(str(v) for v in record["period"])
            if isinstance(record.get("matched_values"), list):
                record["matched_values"] = "|".join(str(v) for v in record["matched_values"])
            writer.writerow({h: record.get(h) for h in headers})


def main(experiment_dir: Path | None = None) -> None:
    """
    Generate programmatic metrics for OpenAI descriptions.
    
    Args:
        experiment_dir: Optional experiment directory for versioned input/output
    """
    # Determine paths
    if experiment_dir is not None:
        artifacts_dir = Path(experiment_dir) / "artifacts"
        openai_dir = artifacts_dir / "descriptions"
        groundtruth_json = artifacts_dir / "charts.json"
        metrics_dir = artifacts_dir / "metrics"
        
        # Temporarily override global constants for this execution
        # This is a workaround to avoid refactoring all helper functions
        global OPENAI_DIR, GROUNDTRUTH_JSON, METRICS_DIR
        global OUTPUT_CSV_UNIFIED, OUTPUT_JSON_UNIFIED
        global OUTPUT_CSV_FACTUAL_CLAIMS, OUTPUT_JSON_FACTUAL_CLAIMS
        global OUTPUT_CSV_FACTUAL_CHECK, OUTPUT_JSON_FACTUAL_CHECK
        
        OPENAI_DIR = openai_dir
        GROUNDTRUTH_JSON = groundtruth_json
        METRICS_DIR = metrics_dir
        OUTPUT_CSV_UNIFIED = metrics_dir / "metriques_openai_unificat.csv"
        OUTPUT_JSON_UNIFIED = metrics_dir / "metriques_openai_unificat.json"
        OUTPUT_CSV_FACTUAL_CLAIMS = metrics_dir / "afirmacions_factuals_openai.csv"
        OUTPUT_JSON_FACTUAL_CLAIMS = metrics_dir / "afirmacions_factuals_openai.json"
        OUTPUT_CSV_FACTUAL_CHECK = metrics_dir / "verificacio_afirmacions_openai.csv"
        OUTPUT_JSON_FACTUAL_CHECK = metrics_dir / "verificacio_afirmacions_openai.json"
    
    cases = load_openai_metrics()
    factual_claims = build_factual_claims(cases)
    verified_claims = verify_factual_claims(factual_claims)
    add_verification_summary_to_cases(cases, verified_claims)

    write_csv(OUTPUT_CSV_UNIFIED, cases)
    OUTPUT_JSON_UNIFIED.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_JSON_UNIFIED.open("w", encoding="utf-8") as handle:
        json.dump({"provider": "openai", "cases": build_cases_by_id(cases)}, handle, ensure_ascii=False, indent=2)

    write_csv(OUTPUT_CSV_FACTUAL_CLAIMS, factual_claims)
    with OUTPUT_JSON_FACTUAL_CLAIMS.open("w", encoding="utf-8") as handle:
        json.dump({"provider": "openai", "claims": factual_claims}, handle, ensure_ascii=False, indent=2)

    write_csv(OUTPUT_CSV_FACTUAL_CHECK, verified_claims)
    with OUTPUT_JSON_FACTUAL_CHECK.open("w", encoding="utf-8") as handle:
        json.dump(
            {
                "groundtruth": str(GROUNDTRUTH_JSON),
                "provider": "openai",
                "claims": verified_claims,
            },
            handle,
            ensure_ascii=False,
            indent=2,
        )

    print("\n" + "=" * 80)
    print("FITXERS GENERATS (OPENAI):")
    print("=" * 80)
    print(f"CSV unificat: {OUTPUT_CSV_UNIFIED}")
    print(f"JSON unificat: {OUTPUT_JSON_UNIFIED}")
    print(f"CSV afirmacions factuals: {OUTPUT_CSV_FACTUAL_CLAIMS}")
    print(f"JSON afirmacions factuals: {OUTPUT_JSON_FACTUAL_CLAIMS}")
    print(f"CSV verificacio afirmacions: {OUTPUT_CSV_FACTUAL_CHECK}")
    print(f"JSON verificacio afirmacions: {OUTPUT_JSON_FACTUAL_CHECK}")


if __name__ == "__main__":
    main()
