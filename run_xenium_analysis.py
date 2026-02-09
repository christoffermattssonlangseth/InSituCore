#!/usr/bin/env python3
"""
Replicate the core Xenium workflow from intership1 as a reusable CLI pipeline.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import scanpy as sc
from utils.xenium_pipeline import (
    build_qc_outputs,
    export_karospace_html,
    ensure_spatial_coordinates,
    filter_for_clustering,
    infer_sample_id,
    load_and_concat_runs,
    maybe_run_mana,
    normalize_for_clustering,
    prepare_scvi_representation,
    run_compartment_clustering,
    run_clustering,
)


def _sanitize_obs_for_h5ad(ad: sc.AnnData) -> None:
    idx_name = ad.obs.index.name
    if not idx_name:
        return
    if idx_name not in ad.obs.columns:
        return

    idx_vals = ad.obs.index.astype(str).to_numpy()
    col_vals = ad.obs[idx_name].astype(str).to_numpy()
    if len(idx_vals) != len(col_vals):
        return
    if (idx_vals == col_vals).all():
        return

    new_name = f"{idx_name}_column"
    suffix = 1
    while new_name in ad.obs.columns:
        suffix += 1
        new_name = f"{idx_name}_column_{suffix}"

    ad.obs.rename(columns={idx_name: new_name}, inplace=True)
    print(
        f"STEP: Renaming conflicting obs column '{idx_name}' to '{new_name}' for h5ad compatibility"
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run Xenium analysis from a dataset path.")
    parser.add_argument(
        "--data-dir",
        required=True,
        help="Base directory containing output-* Xenium run folders.",
    )
    parser.add_argument(
        "--out-dir",
        default="analysis_output",
        help="Directory where outputs are written.",
    )
    parser.add_argument(
        "--run-prefix",
        default="output-",
        help="Prefix used to discover run directories.",
    )
    parser.add_argument(
        "--run-search-depth",
        type=int,
        default=1,
        help="Run directory search depth: 1=direct only, 2=one level below (default: 1).",
    )
    parser.add_argument(
        "--count-matrix-mode",
        choices=["cell_feature_matrix", "nucleus_or_distance"],
        default="cell_feature_matrix",
        help=(
            "Count matrix source: standard Xenium cell_feature_matrix.h5 "
            "or transcript-level nucleus/distance filtering."
        ),
    )
    parser.add_argument(
        "--tx-max-distance-um",
        type=float,
        default=5.0,
        help=(
            "Maximum transcript distance to nucleus in microns for nucleus_or_distance mode "
            "(default: 5.0)."
        ),
    )
    parser.add_argument(
        "--tx-nucleus-distance-key",
        default="nucleus_distance",
        help=(
            "Transcript parquet column containing nucleus distance "
            "(default: nucleus_distance)."
        ),
    )
    parser.add_argument(
        "--tx-allowed-categories",
        default="predesigned_gene,custom_gene",
        help=(
            "Comma-separated transcript codeword_category values kept when building counts "
            "(default: predesigned_gene,custom_gene)."
        ),
    )
    parser.add_argument("--min-counts", type=int, default=50, help="Filter threshold for min counts.")
    parser.add_argument("--min-genes", type=int, default=15, help="Filter threshold for min genes.")
    parser.add_argument(
        "--target-sum",
        type=float,
        default=100.0,
        help="Target sum used in normalize_total.",
    )
    parser.add_argument(
        "--n-neighbors",
        type=int,
        default=15,
        help="n_neighbors used by scanpy neighbors graph.",
    )
    parser.add_argument(
        "--n-pcs",
        type=int,
        default=30,
        help="Number of PCs used for neighbors graph.",
    )
    parser.add_argument(
        "--umap-min-dist",
        type=float,
        default=0.1,
        help="UMAP min_dist parameter.",
    )
    parser.add_argument(
        "--cluster-graph-mode",
        choices=["auto", "expression", "spatial"],
        default="auto",
        help=(
            "Graph source for UMAP/Leiden/Louvain: expression neighbors, spatial neighbors, "
            "or auto (prefer spatial, fallback to expression)."
        ),
    )
    parser.add_argument(
        "--cluster-spatial-key",
        default="spatial",
        help="Spatial coordinates key in adata.obsm for clustering graph (default: spatial).",
    )
    parser.add_argument(
        "--cluster-connectivity-key",
        default="spatial_connectivities",
        help=(
            "Connectivity key in adata.obsp for clustering graph "
            "(default: spatial_connectivities)."
        ),
    )
    parser.add_argument(
        "--cluster-sample-key",
        default="sample_id",
        help=(
            "obs column used to avoid cross-sample spatial edges in clustering graph "
            "(default: sample_id)."
        ),
    )
    parser.add_argument(
        "--leiden-resolutions",
        default="0.1,0.5,1,1.5,2",
        help="Comma-separated Leiden resolutions.",
    )
    parser.add_argument(
        "--cluster-method",
        choices=["leiden", "louvain", "kmeans"],
        default="leiden",
        help="Primary clustering method after graph/UMAP (default: leiden).",
    )
    parser.add_argument(
        "--louvain-resolutions",
        default="0.5,1.0",
        help="Comma-separated Louvain resolutions.",
    )
    parser.add_argument(
        "--kmeans-clusters",
        default="8,12",
        help="Comma-separated K values for KMeans clustering.",
    )
    parser.add_argument(
        "--kmeans-random-state",
        type=int,
        default=0,
        help="Random seed for KMeans clustering (default: 0).",
    )
    parser.add_argument(
        "--kmeans-n-init",
        type=int,
        default=10,
        help="Number of KMeans initializations per K (default: 10).",
    )
    parser.add_argument(
        "--sample-id-split",
        default="__",
        help="Delimiter used to split run name for sample ID extraction.",
    )
    parser.add_argument(
        "--sample-id-index",
        type=int,
        default=2,
        help="Index in split run name used for sample ID extraction.",
    )
    parser.add_argument(
        "--sample-id-source",
        choices=["auto", "run", "parent"],
        default="auto",
        help="How to derive sample_id: from run label, parent folder, or auto (default: auto).",
    )
    parser.add_argument(
        "--mana-aggregate",
        action="store_true",
        help="Run optional MANA weighted aggregation (requires spatial coordinates).",
    )
    parser.add_argument(
        "--mana-spatial-key",
        default="spatial",
        help="Spatial coordinates key in adata.obsm (default: spatial).",
    )
    parser.add_argument(
        "--mana-connectivity-key",
        default="spatial_connectivities",
        help="Connectivity key in adata.obsp (default: spatial_connectivities).",
    )
    parser.add_argument(
        "--mana-n-layers",
        type=int,
        default=3,
        help="Number of neighbor hops for MANA aggregation (default: 3).",
    )
    parser.add_argument(
        "--mana-hop-decay",
        type=float,
        default=0.2,
        help="Hop decay for MANA aggregation (default: 0.2).",
    )
    parser.add_argument(
        "--mana-distance-kernel",
        choices=["exponential", "inverse", "gaussian", "none"],
        default="gaussian",
        help="Distance kernel for MANA aggregation (default: gaussian).",
    )
    parser.add_argument(
        "--mana-distance-scale",
        type=float,
        default=None,
        help="Distance scale for MANA kernel (default: auto).",
    )
    parser.add_argument(
        "--mana-representation-mode",
        choices=["scvi", "pca", "custom", "auto"],
        default="scvi",
        help="Feature representation for MANA aggregation (default: scvi).",
    )
    parser.add_argument(
        "--mana-use-rep",
        default=None,
        help="Representation key in adata.obsm when --mana-representation-mode=custom.",
    )
    parser.add_argument(
        "--scvi-latent-key",
        default="X_scVI",
        help="Output key in adata.obsm for scVI latent representation (default: X_scVI).",
    )
    parser.add_argument(
        "--scvi-n-latent",
        type=int,
        default=30,
        help="Latent dimensions for scVI (default: 30).",
    )
    parser.add_argument(
        "--scvi-max-epochs",
        type=int,
        default=30,
        help="Max training epochs for scVI (default: 30).",
    )
    parser.add_argument(
        "--scvi-hvg-top-genes",
        type=int,
        default=500,
        help="Number of highly variable genes used to train scVI (default: 500).",
    )
    parser.add_argument(
        "--scvi-hvg-flavor",
        default="seurat_v3",
        help="Flavor for highly_variable_genes before scVI (default: seurat_v3).",
    )
    parser.add_argument(
        "--scvi-batch-key",
        default="sample_id",
        help="Optional obs column used as batch key for scVI (default: sample_id).",
    )
    parser.add_argument(
        "--mana-sample-key",
        default="sample_id",
        help="obs column to aggregate per-sample (default: sample_id).",
    )
    parser.add_argument(
        "--mana-out-key",
        default="X_mana_gauss",
        help="Output key in adata.obsm for MANA aggregation (default: X_mana_gauss).",
    )
    parser.add_argument(
        "--mana-compartment-resolutions",
        default="1.0",
        help="Comma-separated Leiden resolutions for MANA compartments (default: 1.0).",
    )
    parser.add_argument(
        "--mana-compartment-neighbors",
        type=int,
        default=15,
        help="n_neighbors for MANA compartment clustering (default: 15).",
    )
    parser.add_argument(
        "--mana-compartment-method",
        choices=["gmm", "leiden", "both"],
        default="gmm",
        help="Compartment clustering method for MANA output (default: gmm).",
    )
    parser.add_argument(
        "--mana-gmm-components",
        default="5,8,10,12,15,20",
        help="Comma-separated number of components for MANA GMM compartments (default: 5,8,10,12,15,20).",
    )
    parser.add_argument(
        "--mana-gmm-covariance-type",
        choices=["full", "tied", "diag", "spherical"],
        default="diag",
        help="Covariance type for MANA GMM compartments (default: diag).",
    )
    parser.add_argument(
        "--mana-gmm-random-state",
        type=int,
        default=0,
        help="Random seed for MANA GMM compartments (default: 0).",
    )
    parser.add_argument(
        "--mana-gmm-n-init",
        type=int,
        default=1,
        help="Number of GMM initializations per component count (default: 1).",
    )
    parser.add_argument(
        "--mana-gmm-max-dims",
        type=int,
        default=30,
        help="Max dimensions used for MANA GMM input; if exceeded, PCA reduction is applied (default: 30).",
    )
    parser.add_argument(
        "--mana-no-normalize",
        dest="mana_normalize_weights",
        action="store_false",
        help="Disable MANA weight normalization (default: normalize).",
    )
    parser.add_argument(
        "--mana-no-include-self",
        dest="mana_include_self",
        action="store_false",
        help="Exclude self features from MANA aggregation (default: include).",
    )
    parser.set_defaults(mana_normalize_weights=True, mana_include_self=True)

    parser.add_argument(
        "--karospace-html",
        default=None,
        help="Optional path to export a KaroSpace HTML viewer.",
    )
    parser.add_argument(
        "--karospace-color",
        default=None,
        help="Initial color for KaroSpace (default: last Leiden key).",
    )
    parser.add_argument(
        "--karospace-groupby",
        default="sample_id",
        help="Groupby column for KaroSpace sections (default: sample_id).",
    )
    parser.add_argument(
        "--karospace-theme",
        choices=["light", "dark"],
        default="light",
        help="KaroSpace theme (default: light).",
    )
    parser.add_argument(
        "--karospace-title",
        default="KaroSpace",
        help="Title for KaroSpace HTML (default: KaroSpace).",
    )
    parser.add_argument(
        "--karospace-min-panel-size",
        type=int,
        default=150,
        help="Minimum panel width in pixels (default: 150).",
    )
    parser.add_argument(
        "--karospace-spot-size",
        type=float,
        default=2.0,
        help="Spot size for KaroSpace plots (default: 2.0).",
    )
    parser.add_argument(
        "--karospace-downsample",
        type=int,
        default=None,
        help="Downsample per section for KaroSpace export (default: no downsample).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_dir = Path(args.data_dir).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    data_out_dir = out_dir / "data"
    qc_dir = out_dir / "xenium_qc"

    data_out_dir.mkdir(parents=True, exist_ok=True)
    qc_dir.mkdir(parents=True, exist_ok=True)

    if not base_dir.exists() or not base_dir.is_dir():
        raise NotADirectoryError(f"Invalid data directory: {base_dir}")

    print("STEP: Reading data")
    ad = load_and_concat_runs(
        base_dir=base_dir,
        run_prefix=args.run_prefix,
        search_depth=args.run_search_depth,
        count_matrix_mode=args.count_matrix_mode,
        tx_max_distance_to_nucleus_um=args.tx_max_distance_um,
        tx_nucleus_distance_key=args.tx_nucleus_distance_key,
        tx_allowed_categories=args.tx_allowed_categories,
    )
    print("STEP: Calculating QC metrics")
    sc.pp.calculate_qc_metrics(ad, percent_top=None, log1p=False, inplace=True)

    use_parent_sample = False
    has_parent = "run_parent" in ad.obs.columns and ad.obs["run_parent"].astype(str).str.len().gt(0).any()
    if args.sample_id_source == "parent":
        use_parent_sample = True
    elif args.sample_id_source == "auto":
        use_parent_sample = bool(args.run_search_depth > 1 and has_parent)

    if use_parent_sample and has_parent:
        print("STEP: Deriving sample IDs from parent folder")
        ad.obs["sample_id"] = ad.obs["run_parent"].astype(str)
    else:
        print("STEP: Deriving sample IDs from run label")
        ad.obs["sample_id"] = ad.obs["run"].astype(str).apply(
            lambda run: infer_sample_id(run, args.sample_id_split, args.sample_id_index)
        )

    print("STEP: Saving raw AnnData")
    raw_path = data_out_dir / "raw.h5ad"
    _sanitize_obs_for_h5ad(ad)
    ad.write(raw_path)
    print(f"Saved raw AnnData: {raw_path}")

    print("STEP: Building QC outputs")
    build_qc_outputs(ad, qc_dir)
    print(f"Saved QC outputs: {qc_dir}")

    ad_clustered = ad.copy()
    print("STEP: Preprocessing for clustering")
    filter_for_clustering(ad_clustered, args)

    mana_use_rep = args.mana_use_rep
    if args.mana_aggregate:
        mode = args.mana_representation_mode
        if mode == "custom":
            if not mana_use_rep:
                raise ValueError(
                    "MANA representation mode is 'custom' but no --mana-use-rep was provided."
                )
        elif mode == "pca":
            mana_use_rep = "X_pca"
        elif mode == "scvi":
            mana_use_rep = prepare_scvi_representation(ad_clustered, args)
        elif mode == "auto":
            if mana_use_rep:
                pass
            else:
                try:
                    mana_use_rep = prepare_scvi_representation(ad_clustered, args)
                except ImportError:
                    print("STEP: scVI unavailable; falling back to X_pca for MANA")
                    mana_use_rep = "X_pca"

    normalize_for_clustering(ad_clustered, args)
    print("STEP: Running clustering workflow")
    ad_clustered, cluster_key, cluster_keys, cluster_graph_mode = run_clustering(
        ad_clustered, args, data_out_dir
    )

    compartment_key = None
    compartment_keys: list[str] = []
    if args.mana_aggregate:
        print("STEP: Running MANA weighted representation")
        maybe_run_mana(
            ad_clustered,
            enabled=True,
            spatial_key=args.mana_spatial_key,
            connectivity_key=args.mana_connectivity_key,
            n_layers=args.mana_n_layers,
            hop_decay=args.mana_hop_decay,
            distance_kernel=args.mana_distance_kernel,
            distance_scale=args.mana_distance_scale,
            use_rep=mana_use_rep,
            sample_key=args.mana_sample_key,
            out_key=args.mana_out_key,
            normalize_weights=args.mana_normalize_weights,
            include_self=args.mana_include_self,
        )
        print("STEP: Running compartment clustering")
        compartment_result = run_compartment_clustering(ad_clustered, args, data_out_dir)
        compartment_key = str(compartment_result.get("primary_key"))
        compartment_keys = [str(k) for k in compartment_result.get("all_keys", [])]
    cluster_info = {
        "cluster_key": cluster_key,
        "cluster_keys": cluster_keys,
        "cluster_method": args.cluster_method,
        "cluster_graph_mode_requested": args.cluster_graph_mode,
        "cluster_graph_mode_resolved": cluster_graph_mode,
        "count_matrix_mode": args.count_matrix_mode,
        "tx_max_distance_um": args.tx_max_distance_um,
        "tx_nucleus_distance_key": args.tx_nucleus_distance_key,
        "tx_allowed_categories": args.tx_allowed_categories,
        "mana_representation_mode": args.mana_representation_mode,
        "mana_use_rep": mana_use_rep,
        "compartment_key": compartment_key,
        "compartment_keys": compartment_keys,
        "mana_compartment_method": args.mana_compartment_method,
        "mana_gmm_components": args.mana_gmm_components,
        "leiden_resolutions": args.leiden_resolutions,
        "mana_compartment_resolutions": args.mana_compartment_resolutions,
    }
    print("STEP: Ensuring spatial coordinates for outputs")
    has_spatial = ensure_spatial_coordinates(
        ad_clustered,
        target_key="spatial",
        preferred_source_key=args.mana_spatial_key if args.mana_spatial_key != "spatial" else None,
    )
    if not has_spatial and args.karospace_html:
        raise ValueError(
            "KaroSpace export requested, but spatial coordinates are unavailable in clustered data. "
            "Expected ad.obsm['spatial'] or inferable x/y columns in ad.obs."
        )

    print("STEP: Saving clustered outputs")
    (data_out_dir / "cluster_info.json").write_text(json.dumps(cluster_info, indent=2))
    clustered_path = data_out_dir / "clustered.h5ad"
    _sanitize_obs_for_h5ad(ad_clustered)
    ad_clustered.write(clustered_path)

    print(f"Saved clustered AnnData: {clustered_path}")
    print(f"Marker genes saved: {data_out_dir / 'markers_by_cluster.csv'}")
    print(f"Cluster key used for marker ranking: {cluster_key}")
    if cluster_keys:
        print(f"Cluster keys: {', '.join(cluster_keys)}")
    if compartment_keys:
        print(f"MANA compartment keys: {', '.join(compartment_keys)}")

    if args.karospace_html:
        print("STEP: Exporting KaroSpace HTML")
        karospace_color = args.karospace_color or cluster_key
        karospace_path = Path(args.karospace_html).expanduser().resolve()
        export_karospace_html(
            h5ad_path=clustered_path,
            output_path=karospace_path,
            color=karospace_color,
            groupby=args.karospace_groupby,
            title=args.karospace_title,
            theme=args.karospace_theme,
            min_panel_size=args.karospace_min_panel_size,
            spot_size=args.karospace_spot_size,
            downsample=args.karospace_downsample,
        )
        print(f"Saved KaroSpace HTML: {karospace_path}")


if __name__ == "__main__":
    main()
