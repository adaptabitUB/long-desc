# Data Snapshots

Aquest directori conté snapshots immutables de dades utilitzades en experiments.

## Convenció de noms

- `charts_seed{SEED}_v{VERSION}.json` - Snapshots de charts.json generats amb un seed específic
- `matrix_{SIZE1}_{SIZE2}_{DATE}.csv` - Snapshots de matrius de cobertura

## Snapshots actuals

### charts_seed20260311_v1.json
- **Data creació**: 2026-04-13
- **Random seed**: 20260311
- **Número de casos**: 500
- **Font**: output/charts.json (snapshot inicial)
- **Descripció**: Primer snapshot dels 500 gràfics sintètics generats

## Ús

Cada experiment ha de referenciar el snapshot de data utilitzat en el seu manifest.json.
Els snapshots són immutables - mai es modifiquen després de crear-se.
