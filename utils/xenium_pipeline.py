"""Xenium analysis pipeline helpers."""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Optional, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scanpy as sc
import seaborn as sns
from scipy import sparse
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.mixture import GaussianMixture

from .mana import aggregate_neighbors_weighted
from .karospace import export_to_html, load_spatial_data

warnings.filterwarnings("ignore")


def discover_runs(base_dir: Path, run_prefix: str, search_depth: int = 1) -> list[Path]:
    if search_depth < 1:
        raise ValueError("search_depth must be >= 1")

    runs: list[Path] = []
    for entry in sorted(base_dir.iterdir()):
        if not entry.is_dir():
            continue

        if entry.name.startswith(run_prefix):
            runs.append(entry)

        if search_depth >= 2:
            for child in sorted(entry.iterdir()):
                if child.is_dir() and child.name.startswith(run_prefix):
                    runs.append(child)

    return runs


def _split_csv_tokens(raw: str) -> list[str]:
    tokens = [token.strip() for token in raw.split(",") if token.strip()]
    if not tokens:
        raise ValueError("Expected at least one comma-separated token.")
    return tokens


def _coerce_bool_mask(series: pd.Series) -> pd.Series:
    if series.dtype == bool:
        return series.fillna(False)
    if np.issubdtype(series.dtype, np.number):
        return pd.to_numeric(series, errors="coerce").fillna(0).astype(int).eq(1)
    return (
        series.astype(str)
        .str.strip()
        .str.lower()
        .isin(["1", "true", "t", "yes", "y"])
    )


def _list_parquet_files(path: Path) -> list[Path]:
    if path.is_dir():
        files = sorted(path.glob("*.parquet"))
        if not files:
            raise FileNotFoundError(f"Directory contains no parquet shards: {path}")
        return files
    if path.is_file():
        return [path]
    raise FileNotFoundError(f"Not a file or directory: {path}")


def _get_available_parquet_columns(files: Sequence[Path]) -> set[str]:
    try:
        import pyarrow.parquet as pq

        schema = pq.read_schema(str(files[0]))
        return set(schema.names)
    except Exception:
        pass

    try:
        return set(pd.read_parquet(files[0]).columns)
    except Exception as exc:
        raise RuntimeError(
            "Could not inspect parquet columns. Install pyarrow or fastparquet."
        ) from exc


def _infer_distance_key_from_available(available: set[str]) -> Optional[str]:
    candidates = [
        "nucleus_distance",
        "distance_to_nucleus",
        "nucleus_distance_um",
        "distance_to_nucleus_um",
        "distance_to_nucleus_microns",
    ]
    for candidate in candidates:
        if candidate in available:
            return candidate
    return None


def _load_parquet_columns(files: Sequence[Path], columns: Sequence[str]) -> pd.DataFrame:
    try:
        import pyarrow as pa
        import pyarrow.parquet as pq

        tables = [pq.read_table(str(path), columns=list(columns)) for path in files]
        return pa.concat_tables(tables, promote=True).to_pandas()
    except Exception:
        frames = [pd.read_parquet(path, columns=list(columns)) for path in files]
        return pd.concat(frames, axis=0, ignore_index=True)


def _read_cells_metadata(run: Path) -> pd.DataFrame:
    cells_parquet_path = run / "cells.parquet"
    cells_csv_path = run / "cells.csv.gz"

    if cells_parquet_path.exists():
        cells_df = pd.read_parquet(cells_parquet_path)
    elif cells_csv_path.exists():
        cells_df = pd.read_csv(cells_csv_path, index_col=0)
    else:
        raise FileNotFoundError(
            f"Missing cells metadata for {run.name}: expected cells.parquet or cells.csv.gz."
        )

    if "cell_id" in cells_df.columns:
        cells_df["cell_id"] = cells_df["cell_id"].astype(str)
        cells_df = cells_df.drop_duplicates("cell_id").set_index("cell_id")
    else:
        cells_df.index = cells_df.index.astype(str)
        cells_df.index.name = "cell_id"
    return cells_df


def _build_adata_from_transcripts(
    run: Path,
    *,
    max_distance_to_nucleus_um: float,
    nucleus_distance_key: Optional[str],
    allowed_categories: Sequence[str],
) -> sc.AnnData:
    transcripts_path = run / "transcripts.parquet"
    transcript_files = _list_parquet_files(transcripts_path)
    available_columns = _get_available_parquet_columns(transcript_files)

    resolved_distance_key = nucleus_distance_key
    if resolved_distance_key is None or resolved_distance_key not in available_columns:
        inferred_key = _infer_distance_key_from_available(available_columns)
        if inferred_key is None:
            raise KeyError(
                "Could not find nucleus distance column in transcripts parquet. "
                "Provide --tx-nucleus-distance-key explicitly."
            )
        resolved_distance_key = inferred_key

    required_columns = [
        "cell_id",
        "feature_name",
        "codeword_category",
        "overlaps_nucleus",
        resolved_distance_key,
    ]
    missing = [column for column in required_columns if column not in available_columns]
    if missing:
        raise KeyError(
            f"Missing required transcript columns in {run.name}: {missing}. "
            f"Available: {sorted(available_columns)}"
        )

    tx_df = _load_parquet_columns(transcript_files, required_columns)
    allowed_set = set(allowed_categories)

    mask = tx_df["codeword_category"].astype(str).isin(allowed_set)
    mask &= ~tx_df["codeword_category"].astype(str).str.upper().eq("UNASSIGNED")
    mask &= ~tx_df["cell_id"].astype(str).str.upper().eq("UNASSIGNED")
    mask &= ~tx_df["feature_name"].astype(str).str.upper().eq("UNASSIGNED")

    overlap = _coerce_bool_mask(tx_df["overlaps_nucleus"])
    distance_um = pd.to_numeric(tx_df[resolved_distance_key], errors="coerce")
    near_nucleus = distance_um <= float(max_distance_to_nucleus_um)
    mask &= overlap | near_nucleus

    tx_df = tx_df.loc[mask, ["cell_id", "feature_name"]].dropna()
    if tx_df.empty:
        raise ValueError(
            f"No transcripts passed nucleus/distance filter in {run.name}. "
            f"distance_key={resolved_distance_key}, max_distance_um={max_distance_to_nucleus_um}"
        )

    pair_counts = (
        tx_df.value_counts(["cell_id", "feature_name"])
        .rename("count")
        .reset_index()
    )
    pair_counts["cell_id"] = pair_counts["cell_id"].astype(str)
    pair_counts["feature_name"] = pair_counts["feature_name"].astype(str)

    cell_labels = pd.Index(sorted(pair_counts["cell_id"].unique()), name="cell_id")
    gene_labels = pd.Index(sorted(pair_counts["feature_name"].unique()), name="gene")
    row_codes = pd.Categorical(pair_counts["cell_id"], categories=cell_labels).codes
    col_codes = pd.Categorical(pair_counts["feature_name"], categories=gene_labels).codes

    X = sparse.coo_matrix(
        (
            pair_counts["count"].to_numpy(dtype=np.int32),
            (row_codes, col_codes),
        ),
        shape=(len(cell_labels), len(gene_labels)),
        dtype=np.int32,
    ).tocsr()

    ad_int = sc.AnnData(X=X)
    ad_int.obs_names = cell_labels.astype(str)
    ad_int.var_names = gene_labels.astype(str)
    ad_int.obs.index.name = "cell_id"
    ad_int.obs["cell_id"] = ad_int.obs_names
    ad_int.var["gene"] = ad_int.var_names

    cells_df = _read_cells_metadata(run)
    obs_index = pd.Index(ad_int.obs_names)
    for column in cells_df.columns:
        ad_int.obs[column] = obs_index.map(cells_df[column]).to_numpy()

    for x_key, y_key in [("x_centroid", "y_centroid"), ("x", "y"), ("x_location", "y_location")]:
        if x_key in ad_int.obs.columns and y_key in ad_int.obs.columns:
            coords_df = ad_int.obs[[x_key, y_key]].apply(pd.to_numeric, errors="coerce")
            ad_int.obsm["spatial"] = coords_df.to_numpy()
            break

    return ad_int


def load_and_concat_runs(
    base_dir: Path,
    run_prefix: str,
    search_depth: int = 1,
    *,
    count_matrix_mode: str = "cell_feature_matrix",
    tx_max_distance_to_nucleus_um: float = 5.0,
    tx_nucleus_distance_key: Optional[str] = "nucleus_distance",
    tx_allowed_categories: str = "predesigned_gene,custom_gene",
) -> sc.AnnData:
    print("STEP: Discovering run folders")
    runs = discover_runs(base_dir, run_prefix, search_depth=search_depth)
    if not runs:
        raise FileNotFoundError(
            f"No run directories starting with '{run_prefix}' found in {base_dir}"
        )

    if count_matrix_mode not in {"cell_feature_matrix", "nucleus_or_distance"}:
        raise ValueError(
            f"Unsupported count_matrix_mode='{count_matrix_mode}'. "
            "Use 'cell_feature_matrix' or 'nucleus_or_distance'."
        )

    allowed_categories: list[str] = []
    if count_matrix_mode == "nucleus_or_distance":
        allowed_categories = _split_csv_tokens(tx_allowed_categories)
        print("STEP: Count matrix mode: nucleus OR distance-filtered transcripts")
        print(
            "Transcript filter: overlaps_nucleus OR "
            f"{tx_nucleus_distance_key} <= {tx_max_distance_to_nucleus_um} um"
        )

    ad_list: list[sc.AnnData] = []
    for run in runs:
        if count_matrix_mode == "cell_feature_matrix":
            h5_path = run / "cell_feature_matrix.h5"
            cell_info_path = run / "cells.csv.gz"
            if not h5_path.exists() or not cell_info_path.exists():
                print(f"Skipping {run} (missing cell_feature_matrix.h5 or cells.csv.gz)")
                continue

            print(f"STEP: Reading run data ({run.name})")
            print(f"Loading run: {run.name}")
            ad_int = sc.read_10x_h5(str(h5_path))
            print(f"STEP: Reading metadata ({run.name})")
            cell_info = pd.read_csv(cell_info_path, index_col=0)

            if len(cell_info) != ad_int.n_obs:
                raise ValueError(
                    f"Row mismatch in {run.name}: cell_info={len(cell_info)} vs matrix={ad_int.n_obs}"
                )

            ad_int.obs = cell_info
        else:
            transcripts_path = run / "transcripts.parquet"
            if not transcripts_path.exists():
                print(f"Skipping {run} (missing transcripts.parquet)")
                continue
            print(f"STEP: Building filtered transcript matrix ({run.name})")
            ad_int = _build_adata_from_transcripts(
                run,
                max_distance_to_nucleus_um=tx_max_distance_to_nucleus_um,
                nucleus_distance_key=tx_nucleus_distance_key,
                allowed_categories=allowed_categories,
            )

        rel_run = run.relative_to(base_dir)
        run_label = "__".join(rel_run.parts)
        ad_int.obs["run"] = run_label
        ad_int.obs["run_leaf"] = run.name
        if len(rel_run.parts) > 1:
            ad_int.obs["run_parent"] = rel_run.parts[-2]
        else:
            ad_int.obs["run_parent"] = ""
        ad_list.append(ad_int)

    if not ad_list:
        if count_matrix_mode == "cell_feature_matrix":
            raise RuntimeError(
                "No valid runs loaded. Ensure each run contains cell_feature_matrix.h5 and cells.csv.gz."
            )
        raise RuntimeError(
            "No valid runs loaded for nucleus/distance mode. Ensure each run contains "
            "transcripts.parquet and cells.parquet or cells.csv.gz."
        )

    print("STEP: Concatenating data")
    ad = sc.concat(ad_list)
    ad.layers["counts"] = ad.X.copy()
    return ad


def infer_sample_id(run_name: str, split_token: str, split_index: int) -> str:
    parts = run_name.split(split_token)
    if 0 <= split_index < len(parts):
        return parts[split_index]
    return run_name


def build_qc_outputs(ad: sc.AnnData, qc_dir: Path) -> None:
    run_col = "run"
    sample_col = "sample_id"
    cell_type_col = "cell_types"
    counts_col = "total_counts"
    ngenes_col = "n_genes_by_counts"

    def pctl(series: pd.Series, p: int) -> float:
        return float(np.nanpercentile(series, p))

    agg_dict: dict[str, tuple[str, object]] = {"n_cells": (run_col, "count")}
    if counts_col in ad.obs.columns:
        agg_dict |= {
            "counts_mean": (counts_col, "mean"),
            "counts_median": (counts_col, "median"),
            "counts_p10": (counts_col, lambda x: pctl(x, 10)),
            "counts_p90": (counts_col, lambda x: pctl(x, 90)),
        }
    if ngenes_col in ad.obs.columns:
        agg_dict |= {
            "genes_mean": (ngenes_col, "mean"),
            "genes_median": (ngenes_col, "median"),
            "genes_p10": (ngenes_col, lambda x: pctl(x, 10)),
            "genes_p90": (ngenes_col, lambda x: pctl(x, 90)),
        }

    summary = (
        ad.obs.groupby(sample_col)
        .agg(**agg_dict)
        .sort_values("n_cells", ascending=False)
    )
    summary.to_csv(qc_dir / "summary_by_run.csv")

    plt.figure(figsize=(9, 4.5))
    sns.barplot(y=summary.index, x=summary["n_cells"], palette="Set3")
    plt.title("Cells per Xenium run")
    plt.xlabel("# cells")
    plt.ylabel("Run")
    plt.tight_layout()
    plt.savefig(qc_dir / "cells_per_run_bar.png", dpi=200)
    plt.close()

    if ngenes_col in ad.obs.columns:
        plt.figure(figsize=(10, 4.5))
        sns.violinplot(
            data=ad.obs, x=sample_col, y=ngenes_col, inner="quartile", palette="rocket"
        )
        plt.title("n_genes_by_counts per run")
        plt.xlabel("Run")
        plt.ylabel("n_genes_by_counts")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        plt.savefig(qc_dir / "ngenes_violin.png", dpi=200)
        plt.close()

    if counts_col in ad.obs.columns:
        plt.figure(figsize=(10, 4.5))
        sns.violinplot(
            data=ad.obs, x=sample_col, y=counts_col, inner="quartile", palette="mako"
        )
        plt.title("total_counts per run")
        plt.xlabel("Run")
        plt.ylabel("total_counts")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        plt.savefig(qc_dir / "counts_violin.png", dpi=200)
        plt.close()

    if {counts_col, ngenes_col}.issubset(ad.obs.columns):
        plt.figure(figsize=(6, 5))
        plt.hexbin(ad.obs[counts_col], ad.obs[ngenes_col], gridsize=50, mincnt=1)
        plt.xlabel("total_counts")
        plt.ylabel("n_genes_by_counts")
        plt.title("Counts vs genes (all runs)")
        plt.tight_layout()
        plt.savefig(qc_dir / "counts_vs_genes_hex.png", dpi=200)
        plt.close()

    if cell_type_col in ad.obs.columns:
        ct_counts = ad.obs[cell_type_col].value_counts()
        plt.figure(figsize=(9, 5))
        sns.barplot(y=ct_counts.index[:20], x=ct_counts.values[:20], palette="Spectral")
        plt.title("Top cell types (all runs)")
        plt.xlabel("# cells")
        plt.ylabel("cell type")
        plt.tight_layout()
        plt.savefig(qc_dir / "celltypes_top20.png", dpi=200)
        plt.close()

        top_k = 14
        ct_top = ct_counts.index[:top_k].tolist()
        comp = (
            ad.obs.assign(
                ct_plot=lambda d: d[cell_type_col].where(d[cell_type_col].isin(ct_top), other="Other")
            )
            .groupby([sample_col, "ct_plot"])
            .size()
            .groupby(level=0)
            .apply(lambda x: x / x.sum())
            .unstack(fill_value=0)
        )
        comp.plot(kind="bar", stacked=True, figsize=(10, 5), colormap="tab20")
        plt.ylabel("fraction of cells")
        plt.title("Cell-type composition per run")
        plt.legend(bbox_to_anchor=(1.02, 1), loc="upper left", ncol=1)
        plt.tight_layout()
        plt.savefig(qc_dir / "celltypes_per_run_stacked.png", dpi=200)
        plt.close()

    x = ad.X
    is_sparse = sparse.issparse(x)
    if is_sparse:
        detected = (x > 0).astype(np.int8)
        det_overall = np.array(detected.sum(axis=0)).ravel() / ad.n_obs
    else:
        det_overall = np.asarray((x > 0).sum(axis=0)).ravel() / ad.n_obs

    det_overall_series = pd.Series(det_overall, index=ad.var_names, name="fraction_cells")
    det_overall_series.sort_values(ascending=False).to_csv(qc_dir / "gene_detection_overall.csv")

    top30 = det_overall_series.sort_values(ascending=False).head(30)
    plt.figure(figsize=(8, 5))
    sns.barplot(y=top30.index, x=top30.values, palette="coolwarm")
    plt.xlabel("fraction of cells detected")
    plt.ylabel("gene")
    plt.title("Panel coverage: top 30 genes")
    plt.tight_layout()
    plt.savefig(qc_dir / "gene_detection_top30.png", dpi=200)
    plt.close()

    runs = ad.obs[sample_col].astype(str).values
    run_idx = {run_name: np.where(runs == run_name)[0] for run_name in np.unique(runs)}
    det_run: dict[str, np.ndarray] = {}
    for run_name, idx in run_idx.items():
        if len(idx) == 0:
            continue
        if is_sparse:
            sub = x[idx, :]
            frac = np.array((sub > 0).sum(axis=0)).ravel() / len(idx)
        else:
            sub = x[idx, :]
            frac = np.asarray((sub > 0).sum(axis=0)).ravel() / len(idx)
        det_run[run_name] = frac

    if det_run:
        det_df = pd.DataFrame(det_run, index=ad.var_names).T
        var_rank = det_df.var(axis=0).sort_values(ascending=False)
        selected_genes = var_rank.head(40).index
        plt.figure(figsize=(min(12, len(selected_genes) * 0.3 + 4), max(4, len(det_df) * 0.35 + 2)))
        sns.heatmap(
            det_df[selected_genes],
            cmap="rocket",
            vmin=0,
            vmax=1,
            cbar_kws={"label": "fraction detected"},
        )
        plt.title("Gene detection per run (top variable genes)")
        plt.xlabel("gene")
        plt.ylabel("run")
        plt.tight_layout()
        plt.savefig(qc_dir / "gene_detection_heatmap_runs.png", dpi=200)
        plt.close()

    for cand in ["cell_area_um2", "cell_area", "area", "nucleus_area_um2"]:
        if cand in ad.obs.columns:
            plt.figure(figsize=(7, 4))
            sns.violinplot(
                data=ad.obs,
                x=sample_col,
                y=cand,
                inner="quartile",
                palette="PuBuGn",
            )
            plt.title(f"{cand} per run")
            plt.xlabel("Run")
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            plt.savefig(qc_dir / f"{cand}_violin.png", dpi=200)
            plt.close()
            break


def run_clustering(
    ad: sc.AnnData,
    args,
    output_data_dir: Path,
) -> tuple[sc.AnnData, str, list[str], str]:
    print("STEP: Running PCA")
    sc.tl.pca(ad)

    requested_graph_mode = getattr(args, "cluster_graph_mode", "auto")
    spatial_key = getattr(args, "cluster_spatial_key", "spatial")
    connectivity_key = getattr(args, "cluster_connectivity_key", "spatial_connectivities")
    sample_key = getattr(args, "cluster_sample_key", "sample_id")

    if requested_graph_mode not in {"auto", "expression", "spatial"}:
        raise ValueError(
            f"Unsupported --cluster-graph-mode '{requested_graph_mode}'. "
            "Use one of: auto, expression, spatial."
        )

    expr_neighbors_key = "cluster_expr"
    print("STEP: Computing expression neighbors")
    sc.pp.neighbors(
        ad,
        n_neighbors=args.n_neighbors,
        n_pcs=args.n_pcs,
        key_added=expr_neighbors_key,
    )

    spatial_neighbors_key = _neighbors_key_from_connectivity_key(connectivity_key)
    spatial_graph_ready = False
    try:
        print("STEP: Building spatial neighbors graph for clustering")
        ensure_spatial_connectivities(
            ad,
            spatial_key=spatial_key,
            connectivity_key=connectivity_key,
            sample_key=sample_key,
        )
        spatial_graph_ready = spatial_neighbors_key in ad.uns
        if not spatial_graph_ready:
            print(
                "STEP: Spatial graph connectivities exist but metadata is missing in ad.uns; "
                "spatial graph will not be used for clustering."
            )
    except (ImportError, ValueError, KeyError) as exc:
        print(f"STEP: Spatial graph unavailable for clustering ({exc})")
        if requested_graph_mode == "spatial":
            raise

    neighbors_key = expr_neighbors_key
    resolved_graph_mode = "expression"
    if requested_graph_mode == "spatial":
        if not spatial_graph_ready:
            raise ValueError(
                "Spatial clustering requested, but spatial graph metadata is unavailable. "
                "Ensure squidpy spatial_neighbors completed successfully."
            )
        neighbors_key = spatial_neighbors_key
        resolved_graph_mode = "spatial"
    elif requested_graph_mode == "auto" and spatial_graph_ready:
        neighbors_key = spatial_neighbors_key
        resolved_graph_mode = "spatial"

    print(f"STEP: Graph selected for UMAP/clustering: {resolved_graph_mode} ({neighbors_key})")

    print("STEP: Computing UMAP")
    sc.tl.umap(ad, min_dist=args.umap_min_dist, neighbors_key=neighbors_key)

    all_keys: list[str] = []
    method = getattr(args, "cluster_method", "leiden")

    if method == "leiden":
        resolution_tokens = [
            token.strip() for token in args.leiden_resolutions.split(",") if token.strip()
        ]
        if not resolution_tokens:
            raise ValueError("No valid Leiden resolutions were provided.")

        for resolution_token in resolution_tokens:
            resolution = float(resolution_token)
            key = f"leiden_{resolution_token}"
            print(f"STEP: Running Leiden clustering ({key})")
            sc.tl.leiden(ad, resolution=resolution, key_added=key, neighbors_key=neighbors_key)
            all_keys.append(key)
    elif method == "louvain":
        resolution_tokens = [
            token.strip() for token in args.louvain_resolutions.split(",") if token.strip()
        ]
        if not resolution_tokens:
            raise ValueError("No valid Louvain resolutions were provided.")

        for resolution_token in resolution_tokens:
            resolution = float(resolution_token)
            key = f"louvain_{resolution_token}"
            print(f"STEP: Running Louvain clustering ({key})")
            try:
                sc.tl.louvain(
                    ad,
                    resolution=resolution,
                    key_added=key,
                    neighbors_key=neighbors_key,
                )
            except ImportError as exc:
                raise ImportError(
                    "Louvain clustering requested, but required dependency is missing. "
                    "Install python-louvain (package: louvain) and retry."
                ) from exc
            all_keys.append(key)
    elif method == "kmeans":
        k_tokens = [token.strip() for token in args.kmeans_clusters.split(",") if token.strip()]
        if not k_tokens:
            raise ValueError("No valid KMeans K values were provided.")
        if "X_pca" not in ad.obsm:
            raise ValueError("KMeans clustering requires PCA coordinates in adata.obsm['X_pca'].")

        x_pca = np.asarray(ad.obsm["X_pca"])
        for token in k_tokens:
            k = int(token)
            key = f"kmeans_k{k}"
            print(f"STEP: Running KMeans clustering ({key})")
            km = KMeans(
                n_clusters=k,
                random_state=args.kmeans_random_state,
                n_init=args.kmeans_n_init,
            )
            labels = km.fit_predict(x_pca)
            ad.obs[key] = pd.Categorical(labels.astype(str))
            all_keys.append(key)
    else:
        raise ValueError(f"Unsupported cluster method: {method}")

    if not all_keys:
        raise ValueError("No clustering outputs were produced.")
    last_key = all_keys[-1]

    marker_path = output_data_dir / "markers_by_cluster.csv"
    print(f"STEP: Ranking marker genes ({last_key})")
    sc.tl.rank_genes_groups(ad, groupby=last_key, method="t-test")
    markers = sc.get.rank_genes_groups_df(ad, group=None)
    markers.to_csv(marker_path, index=False)

    return ad, last_key, all_keys, resolved_graph_mode


def run_compartment_clustering(
    ad: sc.AnnData,
    args,
    output_data_dir: Path,
) -> dict[str, object]:
    if args.mana_out_key not in ad.obsm:
        raise KeyError(
            f"Expected MANA output '{args.mana_out_key}' in adata.obsm. "
            "Run with --mana-aggregate first."
        )

    all_keys: list[str] = []
    model_rows: list[dict[str, object]] = []

    if args.mana_compartment_method in {"leiden", "both"}:
        print("STEP: Computing compartment neighbors")
        sc.pp.neighbors(
            ad,
            n_neighbors=args.mana_compartment_neighbors,
            use_rep=args.mana_out_key,
            key_added="mana_compartments",
        )

        resolution_tokens = [
            token.strip()
            for token in args.mana_compartment_resolutions.split(",")
            if token.strip()
        ]
        if not resolution_tokens:
            raise ValueError("No valid MANA compartment resolutions were provided.")

        for resolution_token in resolution_tokens:
            resolution = float(resolution_token)
            key = f"compartment_leiden_{resolution_token}"
            print(f"STEP: Running compartment Leiden ({key})")
            sc.tl.leiden(
                ad,
                resolution=resolution,
                key_added=key,
                neighbors_key="mana_compartments",
            )
            all_keys.append(key)
            n_clusters = int(ad.obs[key].astype(str).nunique())
            model_rows.append(
                {
                    "key": key,
                    "method": "leiden",
                    "parameter": resolution,
                    "n_clusters": n_clusters,
                    "aic": np.nan,
                    "bic": np.nan,
                }
            )

    if args.mana_compartment_method in {"gmm", "both"}:
        component_tokens = [
            token.strip()
            for token in args.mana_gmm_components.split(",")
            if token.strip()
        ]
        if not component_tokens:
            raise ValueError("No valid MANA GMM components were provided.")

        rep = np.asarray(ad.obsm[args.mana_out_key], dtype=np.float32)
        if rep.ndim != 2:
            raise ValueError(
                f"Expected 2D MANA representation in adata.obsm['{args.mana_out_key}'], got shape {rep.shape}."
            )
        n_obs, n_dim = rep.shape
        print(f"STEP: Preparing GMM input ({n_obs} cells x {n_dim} dims)")
        max_dims = int(getattr(args, "mana_gmm_max_dims", 30))
        if max_dims > 0 and n_dim > max_dims:
            n_components = min(max_dims, max(2, n_obs - 1))
            print(f"STEP: Reducing MANA representation for GMM with PCA ({n_dim} -> {n_components} dims)")
            pca = PCA(n_components=n_components, random_state=args.mana_gmm_random_state)
            rep = pca.fit_transform(rep).astype(np.float32, copy=False)

        for token in component_tokens:
            n_components = int(token)
            key = f"compartment_gmm_k{n_components}"
            print(f"STEP: Running compartment GMM ({key})")
            gmm = GaussianMixture(
                n_components=n_components,
                covariance_type=args.mana_gmm_covariance_type,
                random_state=args.mana_gmm_random_state,
                n_init=args.mana_gmm_n_init,
            )
            labels = gmm.fit_predict(rep)
            ad.obs[key] = pd.Categorical(labels.astype(str))
            all_keys.append(key)
            model_rows.append(
                {
                    "key": key,
                    "method": "gmm",
                    "parameter": n_components,
                    "n_clusters": int(ad.obs[key].astype(str).nunique()),
                    "aic": float(gmm.aic(rep)),
                    "bic": float(gmm.bic(rep)),
                }
            )

    if not all_keys:
        raise ValueError("No compartment clustering outputs were produced.")

    model_df = pd.DataFrame(model_rows)
    model_df.to_csv(output_data_dir / "compartment_models.csv", index=False)

    if args.mana_compartment_method == "gmm":
        primary_key = next((k for k in all_keys if k.startswith("compartment_gmm_")), all_keys[0])
    elif args.mana_compartment_method == "leiden":
        primary_key = all_keys[-1]
    else:
        primary_key = next((k for k in all_keys if k.startswith("compartment_gmm_")), all_keys[-1])

    return {
        "primary_key": primary_key,
        "all_keys": all_keys,
    }


def preprocess_for_clustering(ad: sc.AnnData, args) -> None:
    filter_for_clustering(ad, args)
    normalize_for_clustering(ad, args)


def filter_for_clustering(ad: sc.AnnData, args) -> None:
    print("STEP: Filtering cells by min counts")
    sc.pp.filter_cells(ad, min_counts=args.min_counts)
    print("STEP: Filtering cells by min genes")
    sc.pp.filter_cells(ad, min_genes=args.min_genes)


def normalize_for_clustering(ad: sc.AnnData, args) -> None:
    print("STEP: Normalizing counts")
    sc.pp.normalize_total(ad, inplace=True, target_sum=args.target_sum)
    print("STEP: Log1p transform")
    sc.pp.log1p(ad)


def prepare_scvi_representation(ad: sc.AnnData, args) -> str:
    latent_key = getattr(args, "scvi_latent_key", "X_scVI")
    if latent_key in ad.obsm:
        return latent_key

    try:
        import scvi
    except ImportError as exc:
        raise ImportError(
            "scVI representation requested for MANA, but scvi-tools is not installed. "
            "Install scvi-tools or switch MANA representation mode."
        ) from exc

    batch_key = getattr(args, "scvi_batch_key", None)
    if batch_key and batch_key not in ad.obs.columns:
        print(f"STEP: scVI batch key '{batch_key}' not found; continuing without batch_key")
        batch_key = None

    hvg_top_genes = int(getattr(args, "scvi_hvg_top_genes", 500))
    hvg_flavor = str(getattr(args, "scvi_hvg_flavor", "seurat_v3"))
    print(f"STEP: Selecting HVGs for scVI (n_top_genes={hvg_top_genes}, flavor={hvg_flavor})")
    sc.pp.highly_variable_genes(ad, n_top_genes=hvg_top_genes, flavor=hvg_flavor)
    if "highly_variable" not in ad.var.columns:
        raise ValueError("Failed to compute highly variable genes for scVI preparation.")
    ad_scvi = ad[:, ad.var["highly_variable"]].copy()
    if ad_scvi.n_vars == 0:
        raise ValueError("No highly variable genes selected for scVI.")
    if "counts" not in ad_scvi.layers:
        ad_scvi.layers["counts"] = ad_scvi.X.copy()

    print("STEP: Preparing scVI model")
    scvi.model.SCVI.setup_anndata(ad_scvi, layer="counts", batch_key=batch_key)
    model = scvi.model.SCVI(
        ad_scvi,
        n_latent=int(getattr(args, "scvi_n_latent", 30)),
    )
    print("STEP: Training scVI model")
    model.train(
        max_epochs=int(getattr(args, "scvi_max_epochs", 30)),
        early_stopping=True,
        enable_progress_bar=True,
    )
    print("STEP: Extracting scVI latent representation")
    ad.obsm[latent_key] = model.get_latent_representation(ad_scvi).astype(np.float32, copy=False)
    return latent_key


def _infer_spatial_from_obs(ad: sc.AnnData) -> Optional[np.ndarray]:
    candidate_pairs = [
        ("x_centroid", "y_centroid"),
        ("x", "y"),
        ("x_coord", "y_coord"),
        ("x_um", "y_um"),
        ("x_umap", "y_umap"),
        ("centroid_x", "centroid_y"),
        ("center_x", "center_y"),
        ("spatial_x", "spatial_y"),
        ("x_pos", "y_pos"),
        ("x_position", "y_position"),
    ]
    for x_col, y_col in candidate_pairs:
        if x_col in ad.obs.columns and y_col in ad.obs.columns:
            coords = ad.obs[[x_col, y_col]].to_numpy(dtype=float)
            return coords
    return None


def ensure_spatial_coordinates(
    ad: sc.AnnData,
    *,
    target_key: str = "spatial",
    preferred_source_key: Optional[str] = None,
) -> bool:
    if target_key in ad.obsm:
        return True

    if preferred_source_key and preferred_source_key in ad.obsm:
        ad.obsm[target_key] = np.asarray(ad.obsm[preferred_source_key])
        print(f"Copied ad.obsm['{preferred_source_key}'] -> ad.obsm['{target_key}']")
        return True

    for source_key in ("spatial", "X_spatial"):
        if source_key in ad.obsm:
            ad.obsm[target_key] = np.asarray(ad.obsm[source_key])
            print(f"Copied ad.obsm['{source_key}'] -> ad.obsm['{target_key}']")
            return True

    inferred = _infer_spatial_from_obs(ad)
    if inferred is not None:
        ad.obsm[target_key] = inferred
        print(f"Inferred spatial coordinates from obs columns -> ad.obsm['{target_key}']")
        return True

    return False


def _neighbors_key_from_connectivity_key(connectivity_key: str) -> str:
    if connectivity_key.endswith("_connectivities"):
        return connectivity_key[: -len("_connectivities")]
    return connectivity_key


def ensure_spatial_connectivities(
    ad: sc.AnnData,
    *,
    spatial_key: str,
    connectivity_key: str,
    sample_key: Optional[str] = None,
) -> str:
    if not ensure_spatial_coordinates(
        ad,
        target_key=spatial_key,
        preferred_source_key="spatial" if spatial_key != "spatial" else None,
    ):
        raise ValueError(
            f"Spatial coordinates missing for graph construction: ad.obsm['{spatial_key}']."
        )

    if connectivity_key in ad.obsp:
        return connectivity_key

    try:
        import squidpy as sq
    except ImportError as exc:
        raise ImportError(
            "Spatial graph requested, but squidpy is not installed. "
            "Install squidpy or precompute spatial connectivities."
        ) from exc

    print("STEP: Building spatial neighbors graph")
    print("Building spatial neighbors graph with squidpy...")
    library_key = None
    if sample_key and sample_key in ad.obs.columns:
        library_key = sample_key
    elif "sample_id" in ad.obs.columns:
        library_key = "sample_id"

    key_added = _neighbors_key_from_connectivity_key(connectivity_key)
    print(
        f"Spatial neighbors config: key_added='{key_added}', library_key='{library_key}', delaunay=True"
    )
    sq.gr.spatial_neighbors(
        ad,
        coord_type="generic",
        delaunay=True,
        key_added=key_added,
        library_key=library_key,
    )

    generated_connectivity_key = f"{key_added}_connectivities"
    generated_distance_key = f"{key_added}_distances"
    if connectivity_key != generated_connectivity_key and generated_connectivity_key in ad.obsp:
        ad.obsp[connectivity_key] = ad.obsp[generated_connectivity_key]
        custom_distance_key = connectivity_key.replace("_connectivities", "_distances")
        if generated_distance_key in ad.obsp and custom_distance_key not in ad.obsp:
            ad.obsp[custom_distance_key] = ad.obsp[generated_distance_key]
        print(
            f"Aliased connectivity key '{generated_connectivity_key}' -> '{connectivity_key}'"
        )

    try:
        import cellcharter as cc

        print("STEP: Removing long spatial links (cellcharter)")
        cc.gr.remove_long_links(ad)
    except ImportError:
        pass

    if connectivity_key not in ad.obsp:
        raise KeyError(
            f"Expected connectivity key '{connectivity_key}' after squidpy graph build, but not found."
        )

    return connectivity_key


def maybe_run_mana(
    ad: sc.AnnData,
    *,
    enabled: bool,
    spatial_key: str,
    connectivity_key: str,
    n_layers: int,
    hop_decay: float,
    distance_kernel: str,
    distance_scale: Optional[float],
    use_rep: Optional[str],
    sample_key: Optional[str],
    out_key: str,
    normalize_weights: bool,
    include_self: bool,
) -> None:
    if not enabled:
        return

    ensure_spatial_connectivities(
        ad,
        spatial_key=spatial_key,
        connectivity_key=connectivity_key,
        sample_key=sample_key,
    )

    resolved_use_rep = use_rep
    if resolved_use_rep is None and "X_pca" in ad.obsm:
        resolved_use_rep = "X_pca"
        print("STEP: Using X_pca as input representation for MANA aggregation")

    print("STEP: Computing weighted neighborhood representation")
    print("Running MANA weighted aggregation...")
    aggregate_neighbors_weighted(
        ad,
        n_layers=n_layers,
        aggregations="mean",
        connectivity_key=connectivity_key,
        use_rep=resolved_use_rep,
        sample_key=sample_key,
        out_key=out_key,
        hop_decay=hop_decay,
        distance_kernel=distance_kernel,
        distance_scale=distance_scale,
        spatial_key=spatial_key,
        normalize_weights=normalize_weights,
        include_self=include_self,
    )


def export_karospace_html(
    *,
    h5ad_path: Path,
    output_path: Path,
    color: str,
    groupby: str,
    title: str,
    theme: str,
    min_panel_size: int,
    spot_size: float,
    downsample: Optional[int],
) -> Path:
    print("Loading data for KaroSpace export...")
    dataset = load_spatial_data(str(h5ad_path), groupby=groupby)

    print("Exporting KaroSpace HTML...")
    output = export_to_html(
        dataset,
        output_path=str(output_path),
        color=color,
        title=title,
        min_panel_size=min_panel_size,
        spot_size=spot_size,
        downsample=downsample,
        theme=theme,
    )
    return Path(output)
