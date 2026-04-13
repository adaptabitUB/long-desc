# Resum de la Implementació del Sistema de Versionat d'Experiments

## ✅ Tasques Completades (17/17)

### Fase 1: Preparació del sistema (3 tasques)
- ✅ Crear estructura de directoris `experiments/`
- ✅ Migrar prompt actual a `experiments/prompts/`
- ✅ Crear snapshot inicial de data

### Fase 2: Implementar tracking de metadades (2 tasques)
- ✅ Crear mòdul `experiment_tracker.py`
- ✅ Modificar `run_pipeline.py` per versionat

### Fase 3: Modificar scripts per directoris versionats (6 tasques)
- ✅ Modificar `generate_coverage_matrix.py`
- ✅ Modificar `generate_charts.py`
- ✅ Modificar `generate_statistics_summary.py`
- ✅ Modificar `generate_macro_vba_statistics_summary.py`
- ✅ Modificar `generate_alt_text_openai.py`
- ✅ Modificar `metriques_programatiques_openai.py`

### Fase 4: Actualitzar manifest (1 tasca)
- ✅ Actualitzar manifest al final del pipeline

### Fase 5: Implementar sistema d'avaluació (2 tasques)
- ✅ Crear `eval_suite_v1` amb config i dataset
- ✅ Crear script `run_evaluation.py`

### Fase 6: Utilitats i documentació (3 tasques)
- ✅ Crear script `compare_experiments.py`
- ✅ Crear script `list_experiments.py`
- ✅ Actualitzar `README.md` amb documentació

## 📁 Estructura de Fitxers Creats

```
experiments/
├── runs/                          # Experiments versionats (buit inicialment)
├── prompts/
│   ├── versions/
│   │   └── v1_prompt.txt         # Prompt migrat
│   └── active_prompt.txt@        # Symlink → v1_prompt.txt
├── data_snapshots/
│   ├── charts_seed20260311_v1.json  # 7.6 MB snapshot
│   └── README.md
└── evaluations/
    └── eval_suite_v1/
        ├── eval_config.json      # Mètriques i thresholds
        ├── eval_dataset.json     # 5 casos de test
        └── README.md

src/long_descriptions/
├── experiment_tracker.py         # 9.3 KB - Classe per tracking
├── run_evaluation.py             # 11 KB - Script d'avaluació
├── compare_experiments.py        # 6 KB - Comparar experiments
└── list_experiments.py           # 6.5 KB - Llistar experiments

longdescriptions_byprompt/
└── prompt.txt                    # Còpia per compatibilitat legacy
```

## 🔧 Scripts Modificats

1. **src/long_descriptions/main.py**
   - Afegit paràmetres: `enable_versioning`, `model_name`, `prompt_version`, `random_seed`
   - Integració amb `ExperimentTracker`
   - Generació automàtica de `manifest.json`
   - Actualització de resultats després de cada etapa

2. **src/long_descriptions/generate_coverage_matrix.py**
   - Nou paràmetre: `experiment_dir`
   - Escriptura a `{experiment_dir}/artifacts/`
   - Creació automàtica de snapshot a `data_snapshots/`

3. **src/long_descriptions/generate_charts.py**
   - Nou paràmetre: `experiment_dir`
   - Creació de `chart_generation_config.json` amb paletes i seed
   - Snapshot automàtic de charts.json

4. **src/long_descriptions/generate_statistics_summary.py**
   - Nou paràmetre: `experiment_dir`
   - Paths dinàmics per input/output

5. **src/long_descriptions/generate_macro_vba_statistics_summary.py**
   - Nou paràmetre: `experiment_dir`
   - Output path dinàmic

6. **src/long_descriptions/metriques_programatiques_openai.py**
   - Nou paràmetre: `experiment_dir`
   - Override temporal de constants globals per compatibilitat

7. **README.md**
   - Nova secció: "Sistema de Versionat d'Experiments"
   - Exemples d'ús de les noves utilitats
   - Best practices per versionat

## 📊 Sistema de Lineage Tracking

Cada experiment ara trackeja:

### 1. Codi i Entorn
- Git commit SHA
- Git branch
- Git dirty status
- Python version
- Package versions (openai, pandas, numpy, openpyxl)
- Plataforma (macOS/Linux/Windows)

### 2. Model Configuration
- Provider (openai)
- Model name (gpt-5.4)
- Inference parameters (temperature, max_tokens, etc.)

### 3. Prompts
- Version (v1, v2, etc.)
- File path relatiu
- SHA256 hash del contingut
- Git commit associat

### 4. Data Snapshots
- Matrix file reference
- Charts snapshot reference
- Random seed (20260311)
- Number of cases (500)

### 5. Pipeline Parameters
- Coverage matrix params (5000, 500)
- Chart generation config (seed, paletes)
- Alt-text generation params (batch_size, ranges, model, sleep)

### 6. Results
- Timing per cada etapa
- Charts generated count
- Descriptions generated count
- Total pipeline time
- Completion timestamp

## 🎯 Funcionalitats Implementades

### Experiment Tracking
```bash
# Executar pipeline (auto-versioned)
python run_pipeline.py

# Genera: experiments/runs/exp_2026-04-13_HH-MM_gpt54_v1/
```

### Llistar Experiments
```bash
# Veure tots
python src/long_descriptions/list_experiments.py

# Últims 5
python src/long_descriptions/list_experiments.py --limit 5

# Detalls d'un
python src/long_descriptions/list_experiments.py --details exp_ID
```

### Comparar Experiments
```bash
python src/long_descriptions/compare_experiments.py exp_A exp_B

# Mostra diff de:
# - Git commits
# - Package versions
# - Model params
# - Prompt versions
# - Results
```

### Avaluar Quality
```bash
python src/long_descriptions/run_evaluation.py \
  --experiment-id exp_ID \
  --eval-suite v1

# Calcula:
# - Structure completeness (25%)
# - Readability score (20%)
# - Factual accuracy (35%)
# - Length appropriateness (10%)
# - No visual references (10%)
```

## 🔄 Compatibilitat Legacy

Els scripts mantenen compatibilitat amb el mode legacy:
- Si `experiment_dir=None`, usen `output/` directament
- Els scripts individuals continuen funcionant sense canvis
- Els directoris `output/` i `longdescriptions_byprompt/` es mantenen temporalment

## 📈 Mètriques d'Avaluació (eval_suite_v1)

### Mètriques Implementades
1. **structure_completeness** (25%)
   - Verifica 4 seccions obligatòries
   
2. **readability_score** (20%)
   - Flesch Reading Ease: 40-70
   - Gunning Fog: 10-15

3. **factual_accuracy** (35%)
   - Taxa de claims suportats >= 85%
   - Integra amb metriques_programatiques_openai.py

4. **length_appropriateness** (10%)
   - Word count: 100-500

5. **no_visual_references** (10%)
   - Detecta frases prohibides

### Thresholds
- Mean score: >= 75.0
- Min family score: >= 70.0
- Max failures: <= 5

### Dataset de Test
5 casos representatius:
- Column chart (case 22)
- Bar chart (case 100)
- Line chart (case 200)
- Pie chart (case 300)
- Scatter chart (case 400)

## 🎁 Beneficis Implementats

### Reproducibilitat
- Cada experiment és completament reproducible des del Git commit
- Snapshots immutables de dades
- Tracking de totes les dependencies

### Comparabilitat
- Fàcil comparació entre experiments
- Identificació ràpida de diferències en configuració
- Historial complet de canvis

### Qualitat
- Avaluació automàtica de resultats
- Mètriques objectives i transparents
- Thresholds configurables

### Organització
- Directoris estructurats i predictibles
- Metadata centralitzada en manifest.json
- Separació clara entre experiments

### Audit Trail
- Saber exactament què, quan i com es va generar
- Facilita debugging i optimització
- Permet roll-back a versions anteriors

## 🚀 Pròxims Passos Suggerits

### Opcionals (No Implementats)
1. **Integració MLflow/Weights & Biases** - Per visualització web
2. **DVC per data versioning** - Alternative a snapshots manuals
3. **CI/CD pipeline** - Executar experiments automàticament
4. **Model fine-tuning tracking** - Si es fan fine-tunings propis
5. **A/B testing framework** - Comparar múltiples prompts automàticament

### Millores Incrementals
1. Afegir més casos al eval_dataset.json
2. Crear eval_suite_v2 amb mètriques addicionals
3. Implementar cleanup automàtic d'experiments antics
4. Afegir alertes per experiments amb git_dirty=true
5. Dashboard simple per visualitzar tendències

## ⚠️ Consideracions Importants

### Estratègia de Cleanup
S'ha decidit mantenir tots els experiments indefinidament. Si l'espai és limitat:
- Opció A: Afegir script d'arxivat (comprimir experiments > 30 dies)
- Opció B: Afegir tag "production" per protegir experiments importants
- Opció C: Manual cleanup sense regles automàtiques

### Format d'Experiment ID
S'utilitza: `exp_{timestamp}_{model}_{prompt_version}`
- Fàcil d'ordenar cronològicament
- Identificable visualment
- Compatible amb filesystems

### Nivell de Logging
Actualment: INFO level (major events, API calls, errors)
- Es pot canviar a DEBUG per més detall
- Es pot canviar a WARNING per menys verbosity

## 📝 Verificació

Per verificar que tot funciona:

```bash
# 1. Executar pipeline complet
python run_pipeline.py

# 2. Verificar que es crea experiment
ls experiments/runs/

# 3. Revisar manifest
cat experiments/runs/exp_*/manifest.json

# 4. Llistar experiments
python src/long_descriptions/list_experiments.py

# 5. (Opcional) Executar avaluació quan hi hagi descripcions
python src/long_descriptions/run_evaluation.py \
  --experiment-id $(ls -t experiments/runs/ | head -1) \
  --eval-suite v1
```

## 📚 Documentació

Tota la documentació s'ha afegit a:
- ✅ README.md - Secció completa sobre versionat
- ✅ experiments/data_snapshots/README.md
- ✅ experiments/evaluations/eval_suite_v1/README.md
- ✅ Docstrings en tots els mòduls nous

---

**Data d'implementació**: 13 d'abril de 2026
**Basat en**: [Reproducible AI: Versioning Models, Prompts, and Data](https://medium.com/the-modern-scientist/reproducible-ai-versioning-models-prompts-and-data-96dd0337af65)
**Status**: ✅ Completat - 17/17 tasques
