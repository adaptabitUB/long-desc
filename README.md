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
1. **genera_matriu_cobertura** - Generates coverage matrix with chart types
2. **genera_instancies_canoniques** - Generates canonical instances  
3. **genera_resum_estadistic** - Generates statistical summary
4. **genera_macro_vba_resum_estadistic** - Generates VBA macro for statistical summary
5. **genera_metriques** - Generates metrics

### Scripts individuals

Executar scripts individualment:

```bash
# Matriu de cobertura (customizable sizes)
uv run python src/long_descriptions/genera_matriu_cobertura.py

# Instàncies canòniques
uv run python src/long_descriptions/genera_instancies_canoniques.py

# Resum estadístic
uv run python src/long_descriptions/genera_resum_estadistic.py

# Macro VBA
uv run python src/long_descriptions/genera_macro_vba_resum_estadistic.py

# Mètriques
uv run python src/long_descriptions/genera_metriques.py
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
│       ├── genera_matriu_cobertura.py           # Pas 1: Genera matrius CSV/Excel
│       ├── genera_instancies_canoniques.py      # Pas 2: Genera instàncies canòniques
│       ├── genera_resum_estadistic.py           # Pas 3: Genera resums estadístics
│       ├── genera_macro_vba_resum_estadistic.py # Pas 4: Genera macro VBA
│       ├── genera_metriques.py                  # Pas 5: Genera mètriques (desactivat)
│       └── genera_grafics_comparatius_metriques.py
├── tests/
│   ├── __init__.py
│   └── test_main.py
├── run_pipeline.py                # Executa el pipeline complet
├── clean_outputs.py               # Neteja fitxers generats
├── test_custom_sizes.py           # Test amb mides personalitzades
├── pyproject.toml
├── .python-version
├── README.md
├── matriu_500.csv                 # Fitxers generats (exemples)
├── matriu_5000.csv
└── sortida_instancies_completa/   # Directori de sortida principal
    ├── instancies_canoniques.json
    ├── instancies_canoniques.xlsx
    ├── manifest.json
    ├── resum_estadistic_instancies.json
    ├── resum_estadistic_*.csv     # Un per cada família de gràfics
    └── ResumEstadisticMacro.bas
```
