"""CLI helpers to generate plots for the desktop app."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Optional

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import scanpy as sc

from .mana import plot_spatial_compact_fast


def _infer_default_color(
    adata: sc.AnnData,
    cluster_info_path: Optional[Path],
    fallback: str = "leiden",
    preferred_key: Optional[str] = None,
) -> str:
    def _to_candidates(value: object) -> list[str]:
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        if isinstance(value, str) and value.strip():
            return [value.strip()]
        return []

    cluster_key = None

    if cluster_info_path and cluster_info_path.is_file():
        try:
            payload = json.loads(cluster_info_path.read_text())
            candidates: list[str] = []
            if preferred_key:
                candidates.extend(_to_candidates(payload.get(preferred_key)))
                if preferred_key.endswith("s"):
                    candidates.extend(_to_candidates(payload.get(preferred_key[:-1])))
            if not candidates:
                candidates.extend(_to_candidates(payload.get("cluster_key")))
            for candidate in candidates:
                if candidate in adata.obs.columns:
                    cluster_key = candidate
                    break
        except json.JSONDecodeError:
            cluster_key = None

    if cluster_key and cluster_key in adata.obs.columns:
        return cluster_key

    leiden_cols = [c for c in adata.obs.columns if c.startswith("leiden_")]
    if leiden_cols:
        return sorted(leiden_cols)[-1]

    if fallback in adata.obs.columns:
        return fallback

    if len(adata.obs.columns) > 0:
        return str(adata.obs.columns[0])

    return fallback


def _load_cluster_info_path(h5ad_path: Path) -> Optional[Path]:
    cluster_info = h5ad_path.parent / "cluster_info.json"
    if cluster_info.exists():
        return cluster_info
    return None


def generate_umap_plot(h5ad_path: Path, output_path: Path, color: Optional[str]) -> None:
    adata = sc.read_h5ad(h5ad_path)
    cluster_info = _load_cluster_info_path(h5ad_path)

    if "X_umap" not in adata.obsm:
        raise ValueError("UMAP coordinates not found in adata.obsm['X_umap']")

    color_key = color or _infer_default_color(adata, cluster_info)

    sc.pl.umap(adata, color=color_key, show=False)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()


def generate_compartment_map(h5ad_path: Path, output_path: Path, color: Optional[str]) -> None:
    adata = sc.read_h5ad(h5ad_path)
    cluster_info = _load_cluster_info_path(h5ad_path)

    color_key = color or _infer_default_color(
        adata,
        cluster_info,
        preferred_key="compartment_keys",
    )

    plot_spatial_compact_fast(
        adata,
        color=color_key,
        groupby="sample_id",
        cols=3,
        height=8,
        shared_scale=False,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()


def generate_spatial_map(h5ad_path: Path, output_path: Path, color: Optional[str]) -> None:
    adata = sc.read_h5ad(h5ad_path)
    cluster_info = _load_cluster_info_path(h5ad_path)

    color_key = color or _infer_default_color(
        adata,
        cluster_info,
        preferred_key="cluster_keys",
    )
    groupby_key = "sample_id" if "sample_id" in adata.obs.columns else "run"
    if groupby_key not in adata.obs.columns and len(adata.obs.columns) > 0:
        groupby_key = str(adata.obs.columns[0])

    plot_spatial_compact_fast(
        adata,
        color=color_key,
        groupby=groupby_key,
        cols=3,
        height=8,
        shared_scale=False,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()


def generate_gene_expression_dotplot(
    h5ad_path: Path,
    output_path: Path,
    groupby: Optional[str],
    top_n: int,
) -> None:
    adata = sc.read_h5ad(h5ad_path)
    cluster_info = _load_cluster_info_path(h5ad_path)

    groupby_key = groupby or _infer_default_color(
        adata,
        cluster_info,
        preferred_key="cluster_keys",
    )
    if groupby_key not in adata.obs.columns:
        raise ValueError(f"Groupby key '{groupby_key}' not found in adata.obs.")
    if not str(adata.obs[groupby_key].dtype).startswith("category"):
        adata.obs[groupby_key] = adata.obs[groupby_key].astype("category")

    top_n = max(1, int(top_n))
    rank_key = f"rank_genes_groups__{groupby_key}"
    sc.tl.rank_genes_groups(adata, groupby=groupby_key, method="t-test", key_added=rank_key)
    sc.pl.rank_genes_groups_dotplot(
        adata,
        key=rank_key,
        n_genes=top_n,
        show=False,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close("all")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate plots for the desktop app.")
    sub = parser.add_subparsers(dest="command", required=True)

    umap = sub.add_parser("umap", help="Generate UMAP plot")
    umap.add_argument("--h5ad", required=True, help="Path to clustered.h5ad")
    umap.add_argument("--output", required=True, help="Output PNG path")
    umap.add_argument("--color", default=None, help="Color key in adata.obs")

    comp = sub.add_parser("compartments", help="Generate compartment map")
    comp.add_argument("--h5ad", required=True, help="Path to clustered.h5ad")
    comp.add_argument("--output", required=True, help="Output PNG path")
    comp.add_argument("--color", default=None, help="Color key in adata.obs")

    spatial = sub.add_parser("spatial", help="Generate static spatial map")
    spatial.add_argument("--h5ad", required=True, help="Path to clustered.h5ad")
    spatial.add_argument("--output", required=True, help="Output PNG path")
    spatial.add_argument("--color", default=None, help="Color key in adata.obs")

    dotplot = sub.add_parser("dotplot", help="Generate gene-expression dotplot")
    dotplot.add_argument("--h5ad", required=True, help="Path to clustered.h5ad")
    dotplot.add_argument("--output", required=True, help="Output PNG path")
    dotplot.add_argument("--groupby", default=None, help="Groupby key in adata.obs")
    dotplot.add_argument("--top-n", type=int, default=10, help="Top genes per group (default: 10)")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    h5ad_path = Path(args.h5ad).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if args.command == "umap":
        generate_umap_plot(h5ad_path, output_path, args.color)
    elif args.command == "compartments":
        generate_compartment_map(h5ad_path, output_path, args.color)
    elif args.command == "spatial":
        generate_spatial_map(h5ad_path, output_path, args.color)
    elif args.command == "dotplot":
        generate_gene_expression_dotplot(h5ad_path, output_path, args.groupby, args.top_n)


if __name__ == "__main__":
    main()
