"""Xenium analysis pipeline helpers."""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Optional

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scanpy as sc
import seaborn as sns
from scipy import sparse

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


def load_and_concat_runs(base_dir: Path, run_prefix: str, search_depth: int = 1) -> sc.AnnData:
    print("STEP: Discovering run folders")
    runs = discover_runs(base_dir, run_prefix, search_depth=search_depth)
    if not runs:
        raise FileNotFoundError(
            f"No run directories starting with '{run_prefix}' found in {base_dir}"
        )

    ad_list: list[sc.AnnData] = []
    for run in runs:
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
        raise RuntimeError(
            "No valid runs loaded. Ensure each run contains cell_feature_matrix.h5 and cells.csv.gz."
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
) -> tuple[sc.AnnData, str]:
    print("STEP: Running PCA")
    sc.tl.pca(ad)
    print("STEP: Computing neighbors")
    sc.pp.neighbors(ad, n_neighbors=args.n_neighbors, n_pcs=args.n_pcs)
    print("STEP: Computing UMAP")
    sc.tl.umap(ad, min_dist=args.umap_min_dist)

    resolution_tokens = [token.strip() for token in args.leiden_resolutions.split(",") if token.strip()]
    if not resolution_tokens:
        raise ValueError("No valid leiden resolutions were provided.")

    last_key = ""
    for resolution_token in resolution_tokens:
        resolution = float(resolution_token)
        key = f"leiden_{resolution_token}"
        print(f"STEP: Running Leiden clustering ({key})")
        sc.tl.leiden(ad, resolution=resolution, key_added=key)
        last_key = key

    marker_path = output_data_dir / "markers_by_cluster.csv"
    print(f"STEP: Ranking marker genes ({last_key})")
    sc.tl.rank_genes_groups(ad, groupby=last_key, method="t-test")
    markers = sc.get.rank_genes_groups_df(ad, group=None)
    markers.to_csv(marker_path, index=False)

    return ad, last_key


def run_compartment_clustering(
    ad: sc.AnnData,
    args,
    output_data_dir: Path,
) -> str:
    if args.mana_out_key not in ad.obsm:
        raise KeyError(
            f"Expected MANA output '{args.mana_out_key}' in adata.obsm. "
            "Run with --mana-aggregate first."
        )

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

    last_key = ""
    for resolution_token in resolution_tokens:
        resolution = float(resolution_token)
        key = f"compartment_leiden_{resolution_token}"
        print(f"STEP: Running compartment Leiden ({key})")
        sc.tl.leiden(ad, resolution=resolution, key_added=key, neighbors_key="mana_compartments")
        last_key = key

    return last_key


def preprocess_for_clustering(ad: sc.AnnData, args) -> None:
    print("STEP: Filtering cells by min counts")
    sc.pp.filter_cells(ad, min_counts=args.min_counts)
    print("STEP: Filtering cells by min genes")
    sc.pp.filter_cells(ad, min_genes=args.min_genes)

    print("STEP: Normalizing counts")
    sc.pp.normalize_total(ad, inplace=True, target_sum=args.target_sum)
    print("STEP: Log1p transform")
    sc.pp.log1p(ad)


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

    if spatial_key not in ad.obsm:
        print("STEP: Resolving spatial coordinates")
        inferred = _infer_spatial_from_obs(ad)
        if inferred is not None:
            ad.obsm[spatial_key] = inferred
            print(f"Inferred spatial coordinates from obs columns -> ad.obsm['{spatial_key}']")
        else:
            raise ValueError(
                f"MANA requested, but ad.obsm['{spatial_key}'] is missing. "
                "Provide spatial coordinates or choose another --mana-spatial-key."
            )

    if connectivity_key not in ad.obsp:
        try:
            import squidpy as sq
        except ImportError as exc:
            raise ImportError(
                "MANA requested, but spatial connectivity is missing and squidpy is not installed. "
                "Install squidpy or precompute ad.obsp['spatial_connectivities']."
            ) from exc

        print("STEP: Building spatial neighbors graph")
        print("Building spatial neighbors graph with squidpy...")
        sq.gr.spatial_neighbors(ad, coord_type="generic", delaunay=True)

        if connectivity_key not in ad.obsp:
            raise KeyError(
                f"Expected connectivity key '{connectivity_key}' after squidpy graph build, but not found."
            )

    print("STEP: Computing weighted neighborhood representation")
    print("Running MANA weighted aggregation...")
    aggregate_neighbors_weighted(
        ad,
        n_layers=n_layers,
        aggregations="mean",
        connectivity_key=connectivity_key,
        use_rep=use_rep,
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
