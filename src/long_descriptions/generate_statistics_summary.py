from __future__ import annotations

import csv
import json
import math
import re
from collections import Counter
from collections import OrderedDict
from pathlib import Path
from typing import Any, Dict, List


BASE_DIR = Path.cwd()
OUTPUT_ROOT = BASE_DIR / "output"
INPUT_JSON = OUTPUT_ROOT / "charts.json"
OUTPUT_CSV_DIR = OUTPUT_ROOT / "statistics"


def unique_in_order(values: List[Any]) -> List[Any]:
	return list(OrderedDict((value, None) for value in values).keys())


def extract_categories(instance: Dict[str, Any]) -> List[Any]:
	records = instance.get("data", {}).get("source", {}).get("values", [])
	if not records:
		return []

	encoding = instance.get("encoding", {})
	x_field = encoding.get("x", {}).get("field")
	if x_field:
		return unique_in_order([row.get(x_field) for row in records if row.get(x_field) is not None])

	for candidate in ("category", "region", "date"):
		if candidate in records[0]:
			return unique_in_order([row.get(candidate) for row in records if row.get(candidate) is not None])

	return []


def extract_series(instance: Dict[str, Any]) -> List[Any]:
	records = instance.get("data", {}).get("source", {}).get("values", [])
	if not records:
		return []

	encoding = instance.get("encoding", {})
	color_field = encoding.get("color", {}).get("field")
	if color_field:
		return unique_in_order([row.get(color_field) for row in records if row.get(color_field) is not None])

	for candidate in ("series", "group"):
		if candidate in records[0]:
			return unique_in_order([row.get(candidate) for row in records if row.get(candidate) is not None])

	if instance.get("chart", {}).get("excel_family") == "Stock":
		return ["Open", "High", "Low", "Close", "Volume"]

	return []


def extract_axes(instance: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
	axes = instance.get("axes", {})
	result: Dict[str, Dict[str, Any]] = {}

	for axis_key in ("x", "y", "secondary_y", "z"):
		axis = axes.get(axis_key)
		if not axis:
			continue

		result[axis_key] = {
			"min": axis.get("min"),
			"max": axis.get("max"),
			"interval": axis.get("interval"),
			"unit": axis.get("unit"),
		}

	return result


def get_field_name(instance: Dict[str, Any], role: str) -> str | None:
	return instance.get("encoding", {}).get(role, {}).get("field")


def to_float(value: Any) -> float | None:
	try:
		return float(value)
	except (TypeError, ValueError):
		return None


def detect_value_field(instance: Dict[str, Any], records: List[Dict[str, Any]]) -> str | None:
	if not records:
		return None

	field_from_encoding = get_field_name(instance, "y")
	if field_from_encoding:
		return field_from_encoding

	for field in ("value", "y", "close"):
		if field in records[0]:
			return field

	for key, value in records[0].items():
		if to_float(value) is not None:
			return key

	return None


def round_or_none(value: float | None, digits: int = 6) -> float | None:
	if value is None:
		return None
	return round(value, digits)


def percentile(values: List[float], p: float) -> float | None:
	if not values:
		return None
	if len(values) == 1:
		return values[0]

	sorted_values = sorted(values)
	position = (len(sorted_values) - 1) * p
	low_index = int(math.floor(position))
	high_index = int(math.ceil(position))

	if low_index == high_index:
		return sorted_values[low_index]

	weight = position - low_index
	return sorted_values[low_index] * (1 - weight) + sorted_values[high_index] * weight


def modes(values: List[float]) -> List[float]:
	if not values:
		return []
	counts = Counter(values)
	max_count = max(counts.values())
	if max_count <= 1:
		return []
	return sorted([value for value, count in counts.items() if count == max_count])


def linear_slope(values: List[float]) -> float | None:
	n = len(values)
	if n < 2:
		return None
	x_mean = (n - 1) / 2
	y_mean = sum(values) / n
	numerator = 0.0
	denominator = 0.0
	for index, value in enumerate(values):
		dx = index - x_mean
		numerator += dx * (value - y_mean)
		denominator += dx * dx
	if denominator == 0:
		return None
	return numerator / denominator


def pearson_correlation(xs: List[float], ys: List[float]) -> float | None:
	if len(xs) != len(ys) or len(xs) < 2:
		return None
	x_mean = sum(xs) / len(xs)
	y_mean = sum(ys) / len(ys)
	numerator = sum((x - x_mean) * (y - y_mean) for x, y in zip(xs, ys))
	den_x = math.sqrt(sum((x - x_mean) ** 2 for x in xs))
	den_y = math.sqrt(sum((y - y_mean) ** 2 for y in ys))
	denominator = den_x * den_y
	if denominator == 0:
		return None
	return numerator / denominator


def descriptive_stats(values: List[float]) -> Dict[str, Any]:
	if not values:
		return {
			"n": 0,
			"sum": None,
			"min": None,
			"max": None,
			"range": None,
			"mean": None,
			"median": None,
			"mode": [],
			"standard_deviation": None,
			"coefficient_of_variation": None,
			"quartiles": {"q1": None, "q2": None, "q3": None},
			"iqr": None,
			"percentiles": {"p10": None, "p90": None},
			"iqr_outliers": {"count": 0, "values": []},
			"linear_trend": None,
		}

	n = len(values)
	v_min = min(values)
	v_max = max(values)
	v_sum = sum(values)
	v_mean = v_sum / n
	v_median = percentile(values, 0.5)
	v_q1 = percentile(values, 0.25)
	v_q3 = percentile(values, 0.75)
	v_iqr = (v_q3 - v_q1) if (v_q1 is not None and v_q3 is not None) else None
	v_p10 = percentile(values, 0.10)
	v_p90 = percentile(values, 0.90)

	if n > 1:
		variance = sum((value - v_mean) ** 2 for value in values) / (n - 1)
		v_std = math.sqrt(variance)
	else:
		v_std = 0.0

	v_cv = (v_std / v_mean) if v_mean not in (0, 0.0) else None

	outliers: List[float] = []
	if v_iqr is not None:
		low_fence = v_q1 - 1.5 * v_iqr
		high_fence = v_q3 + 1.5 * v_iqr
		outliers = [value for value in values if value < low_fence or value > high_fence]

	return {
		"n": n,
		"sum": round_or_none(v_sum),
		"min": round_or_none(v_min),
		"max": round_or_none(v_max),
		"range": round_or_none(v_max - v_min),
		"mean": round_or_none(v_mean),
		"median": round_or_none(v_median),
		"mode": [round_or_none(value) for value in modes(values)],
		"standard_deviation": round_or_none(v_std),
		"coefficient_of_variation": round_or_none(v_cv),
		"quartiles": {
			"q1": round_or_none(v_q1),
			"q2": round_or_none(v_median),
			"q3": round_or_none(v_q3),
		},
		"iqr": round_or_none(v_iqr),
		"percentiles": {
			"p10": round_or_none(v_p10),
			"p90": round_or_none(v_p90),
		},
		"iqr_outliers": {
			"count": len(outliers),
			"values": [round_or_none(value) for value in sorted(outliers)],
		},
		"linear_trend": round_or_none(linear_slope(values)),
	}


def slugify_family(value: str | None) -> str:
	if not value:
		return "no_family"
	slug = re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")
	return slug or "no_family"


def json_cell(value: Any) -> str:
	if value is None:
		return ""
	return json.dumps(value, ensure_ascii=False)


def flatten_for_csv(summary: Dict[str, Any]) -> Dict[str, Any]:
	numeric_summary = summary.get("numeric_summary", {})
	global_stats = numeric_summary.get("global_stats", {})
	axes = summary.get("axes", {})

	row: Dict[str, Any] = {
		"id": summary.get("id"),
		"title": summary.get("title"),
		"type": summary.get("type"),
		"subtype": summary.get("subtype"),
		"chart_min_value": numeric_summary.get("chart_min_value"),
		"chart_max_value": numeric_summary.get("chart_max_value"),
		"xy_correlation": numeric_summary.get("xy_correlation"),
		"global_n": global_stats.get("n"),
		"global_sum": global_stats.get("sum"),
		"global_min": global_stats.get("min"),
		"global_max": global_stats.get("max"),
		"global_range": global_stats.get("range"),
		"global_mean": global_stats.get("mean"),
		"global_median": global_stats.get("median"),
		"global_mode": json_cell(global_stats.get("mode")),
		"global_standard_deviation": global_stats.get("standard_deviation"),
		"global_coefficient_of_variation": global_stats.get("coefficient_of_variation"),
		"global_q1": (global_stats.get("quartiles") or {}).get("q1"),
		"global_q2": (global_stats.get("quartiles") or {}).get("q2"),
		"global_q3": (global_stats.get("quartiles") or {}).get("q3"),
		"global_iqr": global_stats.get("iqr"),
		"global_p10": (global_stats.get("percentiles") or {}).get("p10"),
		"global_p90": (global_stats.get("percentiles") or {}).get("p90"),
		"global_outlier_count": (global_stats.get("iqr_outliers") or {}).get("count"),
		"global_outlier_values": json_cell((global_stats.get("iqr_outliers") or {}).get("values")),
		"global_linear_trend": global_stats.get("linear_trend"),
		"category_count": (numeric_summary.get("categories") or {}).get("count"),
		"category_names": "; ".join(str(value) for value in summary.get("categories", [])),
		"category_stats_json": json_cell((numeric_summary.get("categories") or {}).get("stats_by_category")),
		"series_count": (numeric_summary.get("series") or {}).get("count"),
		"series_names": "; ".join(str(value) for value in summary.get("series", [])),
		"series_stats_json": json_cell((numeric_summary.get("series") or {}).get("stats_by_series")),
		"numeric_summary_json": json_cell(numeric_summary),
	}

	for axis in ("x", "y", "secondary_y", "z"):
		axis_data = axes.get(axis, {})
		prefix = "y2" if axis == "secondary_y" else axis
		row[f"{prefix}_min"] = axis_data.get("min")
		row[f"{prefix}_max"] = axis_data.get("max")
		row[f"{prefix}_interval"] = axis_data.get("interval")
		row[f"{prefix}_unit"] = axis_data.get("unit")

	return row


def write_family_csvs(summaries: List[Dict[str, Any]]) -> List[Path]:
	OUTPUT_CSV_DIR.mkdir(parents=True, exist_ok=True)
	family_groups: Dict[str, List[Dict[str, Any]]] = OrderedDict()

	for summary in summaries:
		family = str(summary.get("type") or "NoFamily")
		family_groups.setdefault(family, []).append(summary)

	created_files: List[Path] = []
	for family, items in family_groups.items():
		rows = [flatten_for_csv(item) for item in items]
		all_headers: List[str] = []
		for row in rows:
			for key in row.keys():
				if key not in all_headers:
					all_headers.append(key)

		filename = f"statistics_summary_{slugify_family(family)}.csv"
		out_file = OUTPUT_CSV_DIR / filename
		with out_file.open("w", encoding="utf-8", newline="") as csv_file:
			writer = csv.DictWriter(csv_file, fieldnames=all_headers)
			writer.writeheader()
			for row in rows:
				writer.writerow(row)

		created_files.append(out_file)

	return created_files


def build_numeric_summary(instance: Dict[str, Any]) -> Dict[str, Any]:
	records = instance.get("data", {}).get("source", {}).get("values", [])
	if not records:
		return {
			"categories": {"count": 0, "stats_by_category": []},
			"series": {"count": 0, "stats_by_series": []},
			"global_stats": descriptive_stats([]),
			"chart_max_value": None,
			"chart_min_value": None,
			"xy_correlation": None,
		}

	category_field = get_field_name(instance, "x")
	series_field = get_field_name(instance, "color")
	value_field = detect_value_field(instance, records)

	if not category_field:
		for candidate in ("category", "region", "date"):
			if candidate in records[0]:
				category_field = candidate
				break

	if not series_field:
		for candidate in ("series", "group"):
			if candidate in records[0]:
				series_field = candidate
				break

	global_values: List[float] = []
	per_category: Dict[Any, List[float]] = OrderedDict()
	per_series: Dict[Any, List[float]] = OrderedDict()
	x_values: List[float] = []
	y_values: List[float] = []
	x_field = get_field_name(instance, "x")
	y_field = get_field_name(instance, "y") or value_field

	for row in records:
		if value_field is None:
			continue
		value = to_float(row.get(value_field))
		if value is None:
			continue

		global_values.append(value)

		category = row.get(category_field) if category_field else None
		if category is not None:
			per_category.setdefault(category, []).append(value)

		series_value = row.get(series_field) if series_field else "Value"
		if series_value is None:
			series_value = "Value"
		per_series.setdefault(series_value, []).append(value)

		x_numeric = to_float(row.get(x_field)) if x_field else None
		y_numeric = to_float(row.get(y_field)) if y_field else None
		if x_numeric is not None and y_numeric is not None:
			x_values.append(x_numeric)
			y_values.append(y_numeric)

	category_summary = [
		{"category": category, **descriptive_stats(values)}
		for category, values in per_category.items()
		if values
	]

	series_summary = [
		{"series": series_name, **descriptive_stats(values)}
		for series_name, values in per_series.items()
		if values
	]

	global_stats = descriptive_stats(global_values)

	return {
		"categories": {
			"count": len(per_category),
			"stats_by_category": category_summary,
		},
		"series": {
			"count": len(per_series),
			"stats_by_series": series_summary,
		},
		"global_stats": global_stats,
		"chart_max_value": global_stats.get("max"),
		"chart_min_value": global_stats.get("min"),
		"xy_correlation": round_or_none(pearson_correlation(x_values, y_values)),
	}


def build_summary(instances: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
	summaries: List[Dict[str, Any]] = []

	for instance in instances:
		chart = instance.get("chart", {})
		numeric_summary = build_numeric_summary(instance)
		summaries.append(
			{
				"id": instance.get("id"),
				"title": instance.get("title"),
				"type": chart.get("excel_family") or chart.get("family"),
				"subtype": chart.get("excel_subtype"),
				"axes": extract_axes(instance),
				"categories": extract_categories(instance),
				"series": extract_series(instance),
				"numeric_summary": numeric_summary,
			}
		)

	return summaries


def enrich_instances_with_summary(
	instances: List[Dict[str, Any]], summaries: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
	summary_by_id = {
		str(summary.get("id")): summary
		for summary in summaries
		if summary.get("id") is not None
	}

	enriched_instances: List[Dict[str, Any]] = []
	for instance in instances:
		instance_id = str(instance.get("id") or "")
		summary = summary_by_id.get(instance_id)
		enriched_instance = dict(instance)
		if summary is not None:
			enriched_instance["numeric_summary"] = summary.get("numeric_summary")
		enriched_instances.append(enriched_instance)

	return enriched_instances


def main(experiment_dir: Path | None = None) -> None:
	"""
	Generate statistical summaries for chart families.
	
	Args:
		experiment_dir: Optional experiment directory for versioned output
	"""
	# Determine paths
	if experiment_dir is not None:
		artifacts_dir = Path(experiment_dir) / "artifacts"
		input_json = artifacts_dir / "charts.json"
		csv_dir = artifacts_dir / "statistics"
	else:
		input_json = INPUT_JSON
		csv_dir = OUTPUT_CSV_DIR
	
	if not input_json.exists():
		raise FileNotFoundError(f"Input file not found: {input_json}")

	with input_json.open("r", encoding="utf-8") as file:
		instances = json.load(file)

	summary = build_summary(instances)
	enriched_instances = enrich_instances_with_summary(instances, summary)

	with input_json.open("w", encoding="utf-8") as file:
		json.dump(enriched_instances, file, ensure_ascii=False, indent=2)

	# Group summaries by family for CSV export
	summary_by_family = {}
	for item in summary:
		family = item.get("type", "unknown")
		if family not in summary_by_family:
			summary_by_family[family] = []
		summary_by_family[family].append(item)

	# Write CSVs with dynamic directory
	csv_dir.mkdir(parents=True, exist_ok=True)
	csv_files = []
	for family_name, family_data in summary_by_family.items():
		csv_file = csv_dir / f"statistics_summary_{family_name.lower().replace(' ', '_')}.csv"
		if family_data:
			first_instance = family_data[0]
			fieldnames = list(first_instance.keys())
			with csv_file.open("w", newline="", encoding="utf-8-sig") as f:
				writer = csv.DictWriter(f, fieldnames=fieldnames)
				writer.writeheader()
				writer.writerows(family_data)
			csv_files.append(csv_file)

	print(f"charts.json enriched with numeric_summary: {input_json}")
	print(f"Instances processed: {len(summary)}")
	print("CSV files generated per family:")
	for csv_file in csv_files:
		print(f" - {csv_file}")


if __name__ == "__main__":
	main()
