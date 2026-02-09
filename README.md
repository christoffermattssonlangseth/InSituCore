![InSituCore](assets/logo.png)

Local desktop app for Xenium-style spatial transcriptomics analysis.

`InSituCore` lets you run the pipeline and inspect results in one place:
- QC plots
- Spatial maps (KaroSpace)
- UMAP
- MANA-based compartment maps

## Quick start

Create and activate an environment (recommended):

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
```

Install core dependencies:

```bash
pip install -r requirements.txt
```

Optional dependencies:

```bash
pip install -r requirements-optional.txt
```

- `squidpy` is used when MANA needs to build spatial neighbors from coordinates.
- `louvain` is required only when using Louvain clustering.
- Embedded KaroSpace in-app requires `PySide6.QtWebEngineWidgets`.
  Depending on Python and platform, this may come from `PySide6` directly or require:

```bash
pip install PySide6-QtWebEngine
```

If `PySide6-QtWebEngine` is unavailable for your Python version, the app still works and opens KaroSpace in your browser.

Check your environment:

```bash
python3 check_env.py
```

Strict mode (fails if optional packages are missing):

```bash
python3 check_env.py --require-optional
```

Run the app:

```bash
python3 -m app.main
```

or:

```bash
python3 run_app.py
```

Generate/refresh macOS icon assets:

```bash
python3 scripts/generate_macos_icon.py
```

## Open as app (macOS)

Build a clickable app launcher (no Terminal window):

```bash
bash scripts/build_macos_app.sh
```

This creates:

- `dist/InSituCore.app`

Double-click `dist/InSituCore.app` in Finder to launch the app.
Logs are written to:

- `~/Library/Logs/InSituCore.log`

## Portable bundle (macOS)

Create a portable folder you can zip/share and run without opening Terminal:

```bash
bash scripts/build_portable_macos_bundle.sh
```

Outputs:

- `dist/InSituCore-portable/InSituCore.app`
- `dist/InSituCore-portable/InSituCore/` (project + local runtime)
- `dist/InSituCore-portable/InSituCore.log`

Create a zip in one step:

```bash
bash scripts/build_portable_macos_bundle.sh dist --zip
```

This will produce `dist/InSituCore-portable.zip`.

## App workflow

1. Set `Data dir` and `Output dir`.
2. Choose run discovery:
- `Run depth = Direct folders only` for `/dataset/output-*`.
- `Run depth = One level below samples` for `/dataset/sample_x/output-*`.
3. Choose sample ID source (`Auto`, `From run label`, `From parent folder`).
4. Optional: enable `MANA weighted aggregation`.
5. Optional: enable `Export KaroSpace HTML`.
6. Optional: open `Analysis` tab and click `Show advanced parameters` to configure:
- neighbors / PCs / UMAP min_dist
- clustering method (`Leiden`, `Louvain`, or `KMeans`)
- method-specific sweep values (resolutions or K list)
7. Click `Run`.
8. Use top-bar actions: `Load Outputs`, `Generate UMAP`, `Generate Compartments`.
9. For compartments, choose a compartment key in the `Compartment Map` tab:
- `Auto (primary)` uses the pipeline-selected default.
- Or choose a specific key like `compartment_gmm_k6` / `compartment_leiden_1.0`.

## Supported input layouts

Direct runs:

```text
/dataset/
  output-.../
    cell_feature_matrix.h5
    cells.csv.gz
```

Nested sample folders:

```text
/dataset/
  sample_A/
    output-.../
      cell_feature_matrix.h5
      cells.csv.gz
  sample_B/
    output-.../
      cell_feature_matrix.h5
      cells.csv.gz
```

## Outputs

Written under `--out-dir`:

- `data/raw.h5ad`
- `data/clustered.h5ad`
- `data/cluster_info.json`
- `data/markers_by_cluster.csv`
- `data/compartment_models.csv` (model summary for MANA compartments, including AIC/BIC for GMM)
- `xenium_qc/summary_by_run.csv`
- `xenium_qc/gene_detection_overall.csv`
- `xenium_qc/*.png`
- `plots/umap.png` (generated from app action)
- `plots/compartments.png` (generated from app action)
- `karospace.html` (if export enabled)

## CLI fallback

You can still run everything without the app:

```bash
python3 run_xenium_analysis.py \
  --data-dir /absolute/path/to/dataset \
  --out-dir /absolute/path/to/output
```

Useful options:

- `--run-prefix`
- `--run-search-depth` (`1` or `2`)
- `--sample-id-source` (`auto`, `run`, `parent`)
- `--cluster-method` (`leiden`, `louvain`, `kmeans`)
- `--n-neighbors`, `--n-pcs`, `--umap-min-dist`
- `--leiden-resolutions`, `--louvain-resolutions`, `--kmeans-clusters`
- `--mana-aggregate`
- `--mana-compartment-method` (`gmm`, `leiden`, `both`; default `gmm`)
- `--mana-gmm-components` (e.g. `6,10,14`)
- `--mana-compartment-resolutions` (for Leiden compartments)
- `--karospace-html /absolute/path/to/karospace.html`

## Visual themes

- App name: `InSituCore`
- Themes: `app/theme_light.qss` and `app/theme_dark.qss`
- Startup behavior: detects system appearance (dark/light) automatically
- Toggle in top bar: `Dark` / `Light` (manual override)

## Project notes

- Recent runs are stored at `~/.insitucore/recent.json`.
- Utility modules are vendored under `utils/` (`MANA` + `KaroSpace` helpers).
