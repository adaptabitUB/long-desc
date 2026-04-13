# Long_descriptions

## Instal·lació

Aquest projecte utilitza [UV](https://docs.astral.sh/uv/) com a gestor de paquets i entorns Python.

### Instal·lar UV

```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# O amb Homebrew
brew install uv
```

### Configurar l'entorn

```bash
# Crear entorn virtual i instal·lar dependències
uv sync

# Activar l'entorn virtual
source .venv/bin/activate  # macOS/Linux
```

## Ús

### Pipeline de generació completa

Executar tota la pipeline de generació en ordre:

```bash
# Executar tots els scripts en seqüència
uv run python run_pipeline.py

# O directament amb el mòdul
uv run python -m src.long_descriptions.main
```

L'ordre d'execució és:
1. **generate_coverage_matrix** - Generates coverage matrix with chart types
2. **generate_charts** - Generates canonical instances  
3. **generate_statistics_summary** - Generates statistical summary
4. **generate_macro_vba_statistics_summary** - Generates VBA macro for statistical summary
5. **generate_alt_text_openai** - Generates long alt texts from `charts.json` using an API key from the environment

### Scripts individuals

Executar scripts individualment:

```bash
# Matriu de cobertura (customizable sizes)
uv run python src/long_descriptions/generate_coverage_matrix.py

# Instàncies canòniques
uv run python src/long_descriptions/generate_charts.py

# Resum estadístic
uv run python src/long_descriptions/generate_statistics_summary.py

# Macro VBA
uv run python src/long_descriptions/generate_macro_vba_statistics_summary.py

# Textos alternatius llargs via OpenAI
uv run python src/long_descriptions/generate_alt_text_openai.py --start-case 1 --end-case 50

```

### Configurar l'API OpenAI de forma segura

No posis mai la clau al codi. Fes servir una variable d'entorn local o un fitxer `.env` no versionat.

Exemple de `.env` local:

```bash
OPENAI_API_KEY=la_teva_clau
OPENAI_MODEL=gpt-5.4
```

També tens un exemple sense secrets a `.env.example`.

El fitxer de prompt pot fer servir aquests placeholders:

```text
{case_id}
{chart_title}
{chart_json}
```

Exemple mínim de prompt:

```text
You are writing a long alt text for chart {case_id}.

Title: {chart_title}

Use the chart data and numeric_summary below to produce exactly these sections:
### Overview and main message
### Chart structure
### Relevant patterns, trends, and comparisons
### Essential key details

Chart JSON:
{chart_json}
```

Exemple d'execució per lots:

```bash
uv run python src/long_descriptions/generate_alt_text_openai.py --start-case 1 --end-case 500 --batch-size 50
```

### Neteja de fitxers generats

```bash
# Eliminar tots els fitxers de sortida
uv run python clean_outputs.py
```

## Sistema de Versionat d'Experiments

Aquest projecte implementa un sistema de versionat d'experiments basat en les best practices de l'article ["Reproducible AI: Versioning Models, Prompts, and Data"](https://medium.com/the-modern-scientist/reproducible-ai-versioning-models-prompts-and-data-96dd0337af65).

### Conceptes clau

El sistema trackeja 6 capes principals:
1. **Codi i entorn** - Git commit, Python version, dependencies
2. **Configuració del model** - Provider, model ID, paràmetres d'inferència
3. **Prompts** - Template versionat, hash del contingut
4. **Snapshots de dades** - Dades immutables amb random seed
5. **Suite d'avaluació** - Dataset de test, mètriques, thresholds
6. **Tracking d'experiments** - Lineage complet: codi → model → prompt → data → resultats

### Executar pipeline amb versionat

```bash
# Execució amb versionat activat (per defecte)
uv run python run_pipeline.py

# Això crea automàticament:
# - Un directori d'experiment: experiments/runs/exp_2026-04-13_14-30_gpt54_v1/
# - Un manifest.json amb tota la metadata
# - Artifacts organitzats en subdirectoris
```

### Estructura d'experiments

```
experiments/
├── runs/
│   └── exp_2026-04-13_14-30_gpt54_v1/    # Experiment individual
│       ├── manifest.json                  # Metadata completa
│       ├── artifacts/
│       │   ├── charts.json
│       │   ├── charts.xlsx
│       │   ├── statistics/
│       │   │   └── statistics_summary_*.csv
│       │   ├── descriptions/
│       │   │   └── alt_text_descriptions_*.md
│       │   └── metrics/
│       │       ├── metriques_openai_unificat.csv
│       │       └── ...
│       ├── configs/
│       │   ├── environment.yml
│       │   ├── chart_generation_config.json
│       │   └── ...
│       └── logs/
│           ├── run.log
│           └── errors.log
│
├── prompts/
│   ├── versions/
│   │   ├── v1_prompt.txt
│   │   └── v2_prompt.txt
│   └── active_prompt.txt               # Symlink al prompt actiu
│
├── data_snapshots/
│   ├── charts_seed20260311_v1.json
│   └── matrix_5000_500_seed20260311.csv
│
└── evaluations/
    └── eval_suite_v1/
        ├── eval_config.json
        ├── eval_dataset.json
        └── results/
            └── exp_2026-04-13_14-30_gpt54_v1_eval_results.json
```

### Manifest.json

Cada experiment té un `manifest.json` que conté:

```json
{
  "experiment_id": "exp_2026-04-13_14-30_gpt54_v1",
  "timestamp": "2026-04-13T14:30:00Z",
  "git_commit": "a13f92c7d8e...",
  "git_branch": "main",
  "git_dirty": false,
  
  "environment": {
    "python_version": "3.11.4",
    "packages": {
      "openai": "1.25.0",
      "pandas": "2.1.0"
    }
  },
  
  "model_config": {
    "provider": "openai",
    "model_name": "gpt-5.4",
    "inference_params": { ... }
  },
  
  "prompt": {
    "version": "v1",
    "hash": "sha256:8a3f9b2...",
    "commit": "a13f92c"
  },
  
  "data": {
    "random_seed": 20260311,
    "num_cases": 500
  },
  
  "results": { ... }
}
```

### Utilitats per gestionar experiments

#### Llistar experiments

```bash
# Veure tots els experiments
uv run python src/long_descriptions/list_experiments.py

# Limitar a 5 més recents
uv run python src/long_descriptions/list_experiments.py --limit 5

# Veure detalls d'un experiment específic
uv run python src/long_descriptions/list_experiments.py --details exp_2026-04-13_14-30_gpt54_v1
```

#### Comparar experiments

```bash
# Comparar 2 o més experiments
uv run python src/long_descriptions/compare_experiments.py \
  exp_2026-04-13_14-30_gpt54_v1 \
  exp_2026-04-14_10-15_gpt54_v2

# Mostra diferències en:
# - Git commit
# - Versions de packages
# - Configuració del model
# - Versió de prompt
# - Random seed
# - Resultats (temps, mètriques)
```

#### Avaluar experiments

```bash
# Executar avaluació amb suite v1
uv run python src/long_descriptions/run_evaluation.py \
  --experiment-id exp_2026-04-13_14-30_gpt54_v1 \
  --eval-suite v1

# Mostra:
# - Total cases avaluats
# - Mean score, min/max scores
# - Nombre de failures
# - Pass rate
# - PASSED/FAILED segons thresholds
```

### Versionar prompts

Els prompts es gestionen com a fitxers versionats:

```bash
# Prompts es guarden a experiments/prompts/versions/
# - v1_prompt.txt (versió inicial)
# - v2_prompt.txt (millores de llegibilitat)
# - v3_prompt.txt (més detall en patrons)

# El pipeline automàticament usa active_prompt.txt (symlink)
# Per canviar de versió:
cd experiments/prompts
rm active_prompt.txt
ln -s versions/v2_prompt.txt active_prompt.txt
```

### Best Practices

1. **Sempre comitejar codi abans d'executar experiments importants**
   - Això permet reproduir exactament l'experiment des del git commit

2. **No modificar artifacts d'experiments**
   - Els directoris a `experiments/runs/` són immutables
   - Per re-executar, crea un nou experiment

3. **Usar tags Git per marcar experiments importants**
   ```bash
   git tag -a exp_baseline_v1 -m "Baseline experiment with prompt v1"
   ```

4. **Documentar canvis de prompt a cada versió**
   - Afegir comentaris al principi del fitxer de prompt explicant canvis

5. **Revisar manifest.json abans de publicar resultats**
   - Verificar que git_dirty = false
   - Confirmar versions de packages

### Mode Legacy (sense versionat)

Si necessites executar amb el comportament legacy (output/ directament):

```python
# A src/long_descriptions/main.py
# Canviar enable_versioning=False

# O executar scripts individuals sense el pipeline:
uv run python src/long_descriptions/generate_charts.py
```

### Troubleshooting

**Error: "Experiment manifest not found"**
- Verifica que l'experiment_id és correcte
- Executa `list_experiments.py` per veure experiments disponibles

**Warning: "Git dirty = true"**
- Hi ha canvis no comitejats
- Commiteja els canvis o crea un stash abans d'executar experiments oficials

**Error: "No descriptions directory found"**
- L'experiment no ha generat descripcions (fase 5 del pipeline)
- Executa `generate_alt_text_openai.py` per aquest experiment

### Eines de desenvolupament

```bash
# Afegir una dependència
uv add <paquet>

# Afegir una dependència de desenvolupament
uv add --dev <paquet>

# Executar tests
uv run pytest

# Executar linter
uv run ruff check .
uv run ruff format .
```

## Estructura del projecte

```
.
├── src/
│   └── long_descriptions/
│       ├── __init__.py
│       ├── main.py                              # Orquestrador del pipeline
│       ├── generate_coverage_matrix.py          # Pas 1: Genera matrius CSV/Excel
│       ├── generate_charts.py                   # Pas 2: Genera instàncies canòniques
│       ├── generate_statistics_summary.py       # Pas 3: Genera resums estadístics
│       ├── generate_macro_vba_statistics_summary.py # Pas 4: Genera macro VBA
│       ├── generate_alt_text_openai.py          # Genera textos alternatius llargs via OpenAI
│       └── ...
├── tests/
│   ├── __init__.py
│   └── test_main.py
├── run_pipeline.py                # Executa el pipeline complet
├── clean_outputs.py               # Neteja fitxers generats
├── test_custom_sizes.py           # Test amb mides personalitzades
├── pyproject.toml
├── .env.example
├── .python-version
├── README.md
└── output/                        # Tota la sortida generada
    ├── coverage_matrix_5000_500.xlsx
    ├── matrix_5000.csv
    ├── matrix_500.csv
    ├── charts.json
    ├── charts.xlsx
    ├── manifest.json
    ├── StatisticsSummaryMacro.bas
    ├── statistics/
    │   └── statistics_summary_*.csv
    └── metriques/
```
