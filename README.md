# Spatial Analysis for Dummies

Path-driven Xenium analysis workflow as:
- a reusable CLI script: `run_xenium_analysis.py`
- a step-by-step notebook: `01-path-driven-xenium-analysis.ipynb`

## What the pipeline does

1. Discovers run folders (default prefix: `output-`) under `--data-dir`.
2. Loads each run's:
   - `cell_feature_matrix.h5`
   - `cells.csv.gz`
3. Concatenates runs into one AnnData object and computes QC metrics.
4. Writes QC tables and plots.
5. Runs clustering workflow (normalize, log1p, PCA, neighbors, UMAP, Leiden).
6. Exports marker genes for the last Leiden resolution.

## Requirements

Python 3.10+ and these packages:
- `scanpy`
- `pandas`
- `numpy`
- `scipy`
- `matplotlib`
- `seaborn`

Install example:

```bash
pip install scanpy pandas numpy scipy matplotlib seaborn
```

## Expected input layout

`--data-dir` should look like:

```text
/path/to/dataset/
  output-.../
    cell_feature_matrix.h5
    cells.csv.gz
  output-.../
    cell_feature_matrix.h5
    cells.csv.gz
```

Runs missing either required file are skipped.

## Run

From this repo:

```bash
python3 run_xenium_analysis.py \
  --data-dir /absolute/path/to/new/dataset \
  --out-dir /absolute/path/to/output
```

### Useful flags

- `--run-prefix` (default: `output-`)
- `--run-search-depth` (default: `1`; set `2` for `sample_folder/output-*` layouts)
- `--min-counts` (default: `50`)
- `--min-genes` (default: `15`)
- `--target-sum` (default: `100`)
- `--n-neighbors` (default: `15`)
- `--n-pcs` (default: `30`)
- `--umap-min-dist` (default: `0.1`)
- `--leiden-resolutions` (default: `0.1,0.5,1,1.5,2`)
- `--sample-id-split` (default: `__`)
- `--sample-id-index` (default: `2`)
- `--sample-id-source` (default: `auto`; one of `auto`, `run`, `parent`)

For nested layouts like:

```text
/dataset/
  sample_A/
    output-.../
  sample_B/
    output-.../
```

use:

```bash
python3 run_xenium_analysis.py \
  --data-dir /dataset \
  --run-search-depth 2 \
  --sample-id-source parent
```

### Optional steps (CLI only)

**MANA weighted aggregation (optional):**

```bash
python3 run_xenium_analysis.py \
  --data-dir /absolute/path/to/new/dataset \
  --out-dir /absolute/path/to/output \
  --mana-aggregate
```

If `spatial_connectivities` is missing, the pipeline will try to build it with `squidpy`
(Delaunay graph). You can control MANA behavior via:

- `--mana-n-layers`
- `--mana-hop-decay`
- `--mana-distance-kernel`
- `--mana-distance-scale`
- `--mana-use-rep`
- `--mana-sample-key`
- `--mana-out-key`

If `ad.obsm['spatial']` is missing, the pipeline will try to infer coordinates
from common obs columns such as `x_centroid/y_centroid`, `x/y`, or `centroid_x/centroid_y`.

MANA runs downstream of the standard clustering/UMAP and produces separate
compartment labels (`compartment_leiden_*`) using the weighted representation.

**KaroSpace HTML export (optional):**

```bash
python3 run_xenium_analysis.py \
  --data-dir /absolute/path/to/new/dataset \
  --out-dir /absolute/path/to/output \
  --karospace-html /absolute/path/to/karospace.html
```

Use `--karospace-color`, `--karospace-groupby`, `--karospace-theme`, `--karospace-title`,
`--karospace-min-panel-size`, `--karospace-spot-size`, and `--karospace-downsample` to tune the export.

## Outputs

Written under `--out-dir`:

- `data/raw.h5ad`
- `data/clustered.h5ad`
- `data/markers_by_cluster.csv`
- `xenium_qc/summary_by_run.csv`
- `xenium_qc/gene_detection_overall.csv`
- `xenium_qc/*.png` (QC and composition plots)

## Notebook version

Use `01-path-driven-xenium-analysis.ipynb` if you want the guided workflow.

1. Set `DATA_DIR` and `OUT_DIR` in Step 1.
2. Run cells top-to-bottom.

## Included utils (MANA + KaroSpace)

This repo now vendors utility modules from:
- `/Users/christoffer/work/karolinska/development/MANA/utils`
- `/Users/christoffer/work/karolinska/development/spatial-viewer/karospace`

They live under `utils/` and can be imported directly from this repo:

```python
from utils import (
    aggregate_neighbors_weighted,
    aggregate_neighbors_weighted_simple,
    plot_spatial_compact_fast,
    load_spatial_data,
    export_to_html,
)
```

To generate a standalone KaroSpace HTML viewer from a `.h5ad` file:

```bash
python -m utils.karospace.cli /absolute/path/to/clustered.h5ad \
  --output karospace.html \
  --color leiden_1 \
  --groupby sample_id
```

## Desktop app (local)

The `app/` folder contains a local PySide6 desktop app to run the pipeline and view:
- QC plots
- KaroSpace spatial maps (embedded HTML)
- UMAP plots
- Compartment maps (MANA-based if enabled)

Install dependencies:

```bash
pip install PySide6
```

Optional (for embedded KaroSpace viewer):

```bash
pip install PySide6-QtWebEngine
```

Run the app:

```bash
python3 -m app.main
```

Or:

```bash
python3 run_app.py
```

The app keeps a small list of recent runs at `~/.spatial-analysis-for-dummies/recent.json`.
App name: `InSituViewer`.
Themes live in `app/theme_light.qss` and `app/theme_dark.qss`, with a top-bar dark mode toggle.
Primary actions are in the top bar:
`Run`, `Load Outputs`, `Generate UMAP`, `Generate Compartments`.
