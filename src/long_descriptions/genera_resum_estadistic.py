from __future__ import annotations

import csv
import json
import math
import re
from collections import OrderedDict
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List


BASE_DIR = Path.cwd()
INPUT_JSON = BASE_DIR / "sortida_instancies_completa" / "instancies_canoniques.json"
OUTPUT_JSON = BASE_DIR / "sortida_instancies_completa" / "resum_estadistic_instancies.json"
OUTPUT_CSV_DIR = BASE_DIR / "sortida_instancies_completa"


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

	for candidate in ("categoria", "regio", "data"):
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

	for candidate in ("serie", "grup"):
		if candidate in records[0]:
			return unique_in_order([row.get(candidate) for row in records if row.get(candidate) is not None])

	if instance.get("chart", {}).get("excel_family") == "Stock":
		return ["Obertura", "Màxim", "Mínim", "Tancament", "Volum"]

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
			"unitat": axis.get("unit"),
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

	preferred = ("valor", "y", "tancament")
	for field in preferred:
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
	return sorted([v for v, c in counts.items() if c == max_count])


def linear_slope(values: List[float]) -> float | None:
	n = len(values)
	if n < 2:
		return None
	x_mean = (n - 1) / 2
	y_mean = sum(values) / n
	num = 0.0
	den = 0.0
	for i, y in enumerate(values):
		dx = i - x_mean
		num += dx * (y - y_mean)
		den += dx * dx
	if den == 0:
		return None
	return num / den


def pearson_correlation(xs: List[float], ys: List[float]) -> float | None:
	if len(xs) != len(ys) or len(xs) < 2:
		return None
	x_mean = sum(xs) / len(xs)
	y_mean = sum(ys) / len(ys)
	num = sum((x - x_mean) * (y - y_mean) for x, y in zip(xs, ys))
	den_x = math.sqrt(sum((x - x_mean) ** 2 for x in xs))
	den_y = math.sqrt(sum((y - y_mean) ** 2 for y in ys))
	den = den_x * den_y
	if den == 0:
		return None
	return num / den


def descriptive_stats(values: List[float]) -> Dict[str, Any]:
	if not values:
		return {
			"n": 0,
			"suma": None,
			"min": None,
			"max": None,
			"rang": None,
			"mitjana": None,
			"mediana": None,
			"moda": [],
			"desviacio_estandard": None,
			"coeficient_variacio": None,
			"quartils": {"q1": None, "q2": None, "q3": None},
			"iqr": None,
			"percentils": {"p10": None, "p90": None},
			"outliers_iqr": {"quantitat": 0, "valors": []},
			"tendencia_lineal": None,
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
		variance = sum((v - v_mean) ** 2 for v in values) / (n - 1)
		v_std = math.sqrt(variance)
	else:
		v_std = 0.0

	v_cv = (v_std / v_mean) if v_mean not in (0, 0.0) else None

	outliers: List[float] = []
	if v_iqr is not None:
		low_fence = v_q1 - 1.5 * v_iqr
		high_fence = v_q3 + 1.5 * v_iqr
		outliers = [v for v in values if v < low_fence or v > high_fence]

	return {
		"n": n,
		"suma": round_or_none(v_sum),
		"min": round_or_none(v_min),
		"max": round_or_none(v_max),
		"rang": round_or_none(v_max - v_min),
		"mitjana": round_or_none(v_mean),
		"mediana": round_or_none(v_median),
		"moda": [round_or_none(v) for v in modes(values)],
		"desviacio_estandard": round_or_none(v_std),
		"coeficient_variacio": round_or_none(v_cv),
		"quartils": {
			"q1": round_or_none(v_q1),
			"q2": round_or_none(v_median),
			"q3": round_or_none(v_q3),
		},
		"iqr": round_or_none(v_iqr),
		"percentils": {
			"p10": round_or_none(v_p10),
			"p90": round_or_none(v_p90),
		},
		"outliers_iqr": {
			"quantitat": len(outliers),
			"valors": [round_or_none(v) for v in sorted(outliers)],
		},
		"tendencia_lineal": round_or_none(linear_slope(values)),
	}


def slugify_family(value: str | None) -> str:
	if not value:
		return "sense_familia"
	slug = re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")
	return slug or "sense_familia"


def json_cell(value: Any) -> str:
	if value is None:
		return ""
	return json.dumps(value, ensure_ascii=False)


def flatten_for_csv(summary: Dict[str, Any]) -> Dict[str, Any]:
	resum_numeric = summary.get("resum_numeric", {})
	global_stats = resum_numeric.get("estadistica_global", {})
	eixos = summary.get("eixos", {})

	row: Dict[str, Any] = {
		"id": summary.get("id"),
		"titol": summary.get("titol"),
		"tipus": summary.get("tipus"),
		"subtipus": summary.get("subtipus"),
		"valor_minim_grafic": resum_numeric.get("valor_minim_grafic"),
		"valor_maxim_grafic": resum_numeric.get("valor_maxim_grafic"),
		"correlacio_xy": resum_numeric.get("correlacio_xy"),
		"global_n": global_stats.get("n"),
		"global_suma": global_stats.get("suma"),
		"global_min": global_stats.get("min"),
		"global_max": global_stats.get("max"),
		"global_rang": global_stats.get("rang"),
		"global_mitjana": global_stats.get("mitjana"),
		"global_mediana": global_stats.get("mediana"),
		"global_moda": json_cell(global_stats.get("moda")),
		"global_desviacio_estandard": global_stats.get("desviacio_estandard"),
		"global_coeficient_variacio": global_stats.get("coeficient_variacio"),
		"global_q1": (global_stats.get("quartils") or {}).get("q1"),
		"global_q2": (global_stats.get("quartils") or {}).get("q2"),
		"global_q3": (global_stats.get("quartils") or {}).get("q3"),
		"global_iqr": global_stats.get("iqr"),
		"global_p10": (global_stats.get("percentils") or {}).get("p10"),
		"global_p90": (global_stats.get("percentils") or {}).get("p90"),
		"global_outliers_quantitat": (global_stats.get("outliers_iqr") or {}).get("quantitat"),
		"global_outliers_valors": json_cell((global_stats.get("outliers_iqr") or {}).get("valors")),
		"global_tendencia_lineal": global_stats.get("tendencia_lineal"),
		"categories_quantitat": (resum_numeric.get("categories") or {}).get("quantitat"),
		"categories_nom_llista": "; ".join(str(v) for v in summary.get("categories", [])),
		"categories_estadistiques_json": json_cell((resum_numeric.get("categories") or {}).get("estadistiques_per_categoria")),
		"series_quantitat": (resum_numeric.get("series") or {}).get("quantitat"),
		"series_nom_llista": "; ".join(str(v) for v in summary.get("series", [])),
		"series_estadistiques_json": json_cell((resum_numeric.get("series") or {}).get("estadistiques_per_serie")),
		"resum_numeric_json": json_cell(resum_numeric),
	}

	for axis in ("x", "y", "secondary_y", "z"):
		axis_data = eixos.get(axis, {})
		prefix = "y2" if axis == "secondary_y" else axis
		row[f"{prefix}_min"] = axis_data.get("min")
		row[f"{prefix}_max"] = axis_data.get("max")
		row[f"{prefix}_interval"] = axis_data.get("interval")
		row[f"{prefix}_unitat"] = axis_data.get("unitat")

	return row


def write_family_csvs(summaries: List[Dict[str, Any]]) -> List[Path]:
	OUTPUT_CSV_DIR.mkdir(parents=True, exist_ok=True)
	family_groups: Dict[str, List[Dict[str, Any]]] = OrderedDict()

	for summary in summaries:
		family = str(summary.get("tipus") or "SenseFamilia")
		family_groups.setdefault(family, []).append(summary)

	created_files: List[Path] = []
	for family, items in family_groups.items():
		rows = [flatten_for_csv(item) for item in items]
		all_headers: List[str] = []
		for row in rows:
			for key in row.keys():
				if key not in all_headers:
					all_headers.append(key)

		filename = f"resum_estadistic_{slugify_family(family)}.csv"
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
			"categories": {"quantitat": 0, "min_max_per_categoria": []},
			"series": {"quantitat": 0, "min_max_per_serie": []},
			"estadistica_global": descriptive_stats([]),
			"valor_maxim_grafic": None,
			"valor_minim_grafic": None,
			"correlacio_xy": None,
		}

	category_field = get_field_name(instance, "x")
	series_field = get_field_name(instance, "color")
	value_field = detect_value_field(instance, records)

	if not category_field:
		for candidate in ("categoria", "regio", "data"):
			if candidate in records[0]:
				category_field = candidate
				break

	if not series_field:
		for candidate in ("serie", "grup"):
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

		series = row.get(series_field) if series_field else "Valor"
		if series is None:
			series = "Valor"
		per_series.setdefault(series, []).append(value)

		x_numeric = to_float(row.get(x_field)) if x_field else None
		y_numeric = to_float(row.get(y_field)) if y_field else None
		if x_numeric is not None and y_numeric is not None:
			x_values.append(x_numeric)
			y_values.append(y_numeric)

	category_summary = [
		{"categoria": category, **descriptive_stats(values)}
		for category, values in per_category.items()
		if values
	]

	series_summary = [
		{"serie": serie, **descriptive_stats(values)}
		for serie, values in per_series.items()
		if values
	]

	global_stats = descriptive_stats(global_values)

	return {
		"categories": {
			"quantitat": len(per_category),
			"estadistiques_per_categoria": category_summary,
		},
		"series": {
			"quantitat": len(per_series),
			"estadistiques_per_serie": series_summary,
		},
		"estadistica_global": global_stats,
		"valor_maxim_grafic": global_stats.get("max"),
		"valor_minim_grafic": global_stats.get("min"),
		"correlacio_xy": round_or_none(pearson_correlation(x_values, y_values)),
	}


def build_summary(instances: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
	summaries: List[Dict[str, Any]] = []

	for instance in instances:
		chart = instance.get("chart", {})
		numeric_summary = build_numeric_summary(instance)
		summaries.append(
			{
				"id": instance.get("id"),
				"titol": instance.get("title"),
				"tipus": chart.get("excel_family") or chart.get("family"),
				"subtipus": chart.get("excel_subtype"),
				"eixos": extract_axes(instance),
				"categories": extract_categories(instance),
				"series": extract_series(instance),
				"resum_numeric": numeric_summary,
			}
		)

	return summaries


def main() -> None:
	if not INPUT_JSON.exists():
		raise FileNotFoundError(f"No s'ha trobat l'entrada: {INPUT_JSON}")

	with INPUT_JSON.open("r", encoding="utf-8") as file:
		instances = json.load(file)

	summary = build_summary(instances)

	OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
	with OUTPUT_JSON.open("w", encoding="utf-8") as file:
		json.dump(summary, file, ensure_ascii=False, indent=2)

	csv_files = write_family_csvs(summary)

	print(f"Resum estadístic generat: {OUTPUT_JSON}")
	print(f"Instàncies processades: {len(summary)}")
	print("CSV per família generats:")
	for csv_file in csv_files:
		print(f" - {csv_file}")


if __name__ == "__main__":
	main()
