# Evaluation Suite v1

Aquest directori conté la primera suite d'avaluació per a les descripcions alt-text de gràfics.

## Fitxers

### eval_config.json
Configuració de la suite d'avaluació incloent:
- Mètriques a calcular (estructura, llegibilitat, precisió factual, longitud, referències visuals)
- Pesos per cada mètrica
- Thresholds d'èxit
- Mètode d'agregació

### eval_dataset.json
Dataset d'avaluació manual amb casos de test representatius:
- 5 casos cobreixen diferents tipus de gràfics (Column, Bar, Line, Pie, Scatter)
- Defineix qualitats esperades per cada cas
- Serveix com a ground truth per validar outputs

## Ús

Executa l'avaluació amb:
```bash
python src/long_descriptions/run_evaluation.py \
  --experiment-id exp_2026-04-13_14-30_gpt54_v1 \
  --eval-suite v1
```

Els resultats es guarden a `experiments/evaluations/eval_suite_v1/results/{experiment_id}_eval_results.json`

## Mètriques

**structure_completeness** (25%)
- Verifica que totes les seccions obligatòries existeixin

**readability_score** (20%)
- Flesch Reading Ease: 40-70
- Gunning Fog: 10-15

**factual_accuracy** (35%)
- Taxa de claims suportats >= 85%

**length_appropriateness** (10%)
- Recompte de paraules: 100-500

**no_visual_references** (10%)
- Detecta frases prohibides ("as seen", "on the left", etc.)

## Thresholds d'èxit

- **Mean score**: >= 75.0
- **Min family score**: >= 70.0
- **Max failures**: <= 5 casos
