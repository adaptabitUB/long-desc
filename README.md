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
    ├── statistics_summary_*.csv
    ├── StatisticsSummaryMacro.bas
    └── metriques/
```
