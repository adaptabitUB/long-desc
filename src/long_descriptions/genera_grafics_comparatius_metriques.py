from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
    import matplotlib.pyplot as plt
except ImportError as exc:
    raise SystemExit(
        "No s'ha trobat matplotlib. Instal·la'l amb: pip install matplotlib"
    ) from exc


BASE_DIR = Path.cwd()
METRICS_JSON = BASE_DIR / "long-descriptions" / "metriques" / "metriques_agregat.json"
OUTPUT_DIR = (
    BASE_DIR
    / "long-descriptions"
    / "metriques"
    / "grafics_comparatius"
    / f"claude_gemini_chatgpt_{date.today().isoformat()}"
)


def load_aggregated(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"No s'ha trobat el fitxer agregat: {path}")
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def get_nested(data: Dict[str, Any], keys: Tuple[str, ...]) -> float | None:
    current: Any = data
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            return None
        current = current[key]
    if isinstance(current, (int, float)):
        return float(current)
    return None


def prepare_series(
    aggregated: Dict[str, Any],
    metric_map: List[Tuple[str, Tuple[str, ...]]],
) -> Tuple[List[str], List[str], List[List[float | None]]]:
    providers = list(aggregated.keys())
    labels = [label for label, _ in metric_map]

    provider_values: List[List[float | None]] = []
    for provider in providers:
        values = []
        for _, path in metric_map:
            values.append(get_nested(aggregated[provider], path))
        provider_values.append(values)

    return providers, labels, provider_values


def plot_grouped_bars(
    title: str,
    ylabel: str,
    providers: List[str],
    labels: List[str],
    provider_values: List[List[float | None]],
    out_file: Path,
) -> None:
    out_file.parent.mkdir(parents=True, exist_ok=True)

    x = list(range(len(labels)))
    width = 0.8 / max(len(providers), 1)
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd"]

    fig, ax = plt.subplots(figsize=(12, 6))

    for idx, provider in enumerate(providers):
        vals = provider_values[idx]
        numeric_vals = [v if v is not None else 0.0 for v in vals]
        offset = (idx - (len(providers) - 1) / 2) * width
        positions = [xi + offset for xi in x]
        bars = ax.bar(
            positions,
            numeric_vals,
            width=width,
            label=provider.upper(),
            color=colors[idx % len(colors)],
            alpha=0.9,
        )

        for bar, original in zip(bars, vals):
            if original is None:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height(),
                    "N/A",
                    ha="center",
                    va="bottom",
                    fontsize=8,
                    rotation=90,
                )
            else:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height(),
                    f"{original:.3f}",
                    ha="center",
                    va="bottom",
                    fontsize=8,
                    rotation=90,
                )

    ax.set_title(title, fontsize=13, fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=20, ha="right")
    ax.grid(axis="y", alpha=0.25)
    ax.legend()

    fig.tight_layout()
    fig.savefig(out_file, dpi=150)
    plt.close(fig)



def main() -> None:
    aggregated = load_aggregated(METRICS_JSON)

    quality_metrics = [
        ("QCA_fair", ("qca_fair", "mean")),
        ("DVR", ("dvr", "mean")),
        ("HDR", ("hdr", "mean")),
        ("ICS", ("content_structure", "ics_score", "mean")),
        ("ND", ("number_mentions", "mean")),
    ]

    readability_metrics = [
        ("SMOG", ("smog", "mean")),
        ("Gunning Fog", ("gunning_fog", "mean")),
        ("Flesch-Kincaid", ("flesch_kincaid_grade", "mean")),
        ("Coleman-Liau", ("coleman_liau", "mean")),
    ]

    structure_metrics = [
        ("Has chart type (%)", ("content_structure", "has_chart_type_pct")),
        ("Has title (%)", ("content_structure", "has_title_pct")),
        ("Has axis labels (%)", ("content_structure", "has_axis_labels_pct")),
        ("Has categories (%)", ("content_structure", "has_categories_pct")),
        ("Has values (%)", ("content_structure", "has_values_pct")),
        ("Has scale info (%)", ("content_structure", "has_scale_info_pct")),
    ]

    length_metrics = [
        ("Word count", ("word_count", "mean")),
        ("Sentence count", ("smog", "mean_sentence_count")),
        ("Number mentions", ("number_mentions", "mean")),
    ]

    for filename, title, ylabel, metric_map in [
        (
            "comparativa_qualitat_numerica.png",
            "Comparativa Qualitat Numerica",
            "Valor mitja",
            quality_metrics,
        ),
        (
            "comparativa_llegibilitat.png",
            "Comparativa Llegibilitat",
            "Index mitja",
            readability_metrics,
        ),
        (
            "comparativa_estructura.png",
            "Comparativa Estructura del Contingut",
            "Percentatge",
            structure_metrics,
        ),
        (
            "comparativa_llargada_i_densitat.png",
            "Comparativa Llargada i Densitat",
            "Valor mitja",
            length_metrics,
        ),
    ]:
        providers, labels, provider_values = prepare_series(aggregated, metric_map)
        plot_grouped_bars(
            title=title,
            ylabel=ylabel,
            providers=providers,
            labels=labels,
            provider_values=provider_values,
            out_file=OUTPUT_DIR / filename,
        )

    print("Grafics generats a:")
    print(OUTPUT_DIR)


if __name__ == "__main__":
    main()
