#!/usr/bin/env python3
"""Check local Python environment for InSituCore dependencies."""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import sys
from importlib import metadata


CORE_MODULES = [
    ("anndata", "anndata"),
    ("matplotlib", "matplotlib"),
    ("numpy", "numpy"),
    ("pandas", "pandas"),
    ("PySide6", "PySide6"),
    ("scanpy", "scanpy"),
    ("sklearn", "scikit-learn"),
    ("scipy", "scipy"),
    ("seaborn", "seaborn"),
]

OPTIONAL_MODULES = [
    ("squidpy", "squidpy"),
    ("louvain", "louvain"),
    ("PySide6.QtWebEngineWidgets", "PySide6-QtWebEngine"),
]


def _is_available(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def _pkg_version(package_name: str) -> str:
    try:
        return metadata.version(package_name)
    except metadata.PackageNotFoundError:
        return "-"


def _check_group(group_name: str, items: list[tuple[str, str]]) -> list[tuple[str, str, bool, str]]:
    rows: list[tuple[str, str, bool, str]] = []
    print(f"\n{group_name}:")
    for module_name, package_name in items:
        ok = _is_available(module_name)
        version = _pkg_version(package_name) if ok else "-"
        status = "OK" if ok else "MISSING"
        print(f"  {status:8} {module_name:32} {version}")
        rows.append((module_name, package_name, ok, version))
    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description="Check dependencies for InSituCore.")
    parser.add_argument(
        "--require-optional",
        action="store_true",
        help="Return non-zero exit code when optional packages are missing.",
    )
    args = parser.parse_args()

    print("InSituCore environment check")
    print(f"Python: {sys.version.split()[0]}")

    core_rows = _check_group("Core dependencies", CORE_MODULES)
    optional_rows = _check_group("Optional dependencies", OPTIONAL_MODULES)

    missing_core = [row for row in core_rows if not row[2]]
    missing_optional = [row for row in optional_rows if not row[2]]

    print("\nSummary:")
    if missing_core:
        print("  Missing core packages:")
        for module_name, _, _, _ in missing_core:
            print(f"  - {module_name}")
    else:
        print("  All core packages are installed.")

    if missing_optional:
        print("  Missing optional packages:")
        for module_name, _, _, _ in missing_optional:
            print(f"  - {module_name}")
        print("  The app still works without optional packages.")
    else:
        print("  All optional packages are installed.")

    if missing_core:
        return 1
    if args.require_optional and missing_optional:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
