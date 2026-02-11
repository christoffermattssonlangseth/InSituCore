"""Local desktop app for running and visualizing Xenium analysis outputs."""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from PySide6 import QtCore, QtGui, QtWidgets

try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
    WEB_AVAILABLE = True
except Exception:
    QWebEngineView = None
    WEB_AVAILABLE = False


ROOT_DIR = Path(__file__).resolve().parents[1]
APP_ICON_PATH = ROOT_DIR / "assets" / "app_icon_1024.png"
if not APP_ICON_PATH.exists():
    APP_ICON_PATH = ROOT_DIR / "assets" / "logo.png"
RECENT_PATH = Path.home() / ".insitucore" / "recent.json"
LEGACY_RECENT_PATH = Path.home() / ".spatial-analysis-for-dummies" / "recent.json"
THEME_LIGHT_PATH = Path(__file__).with_name("theme_light.qss")
THEME_DARK_PATH = Path(__file__).with_name("theme_dark.qss")


@dataclass
class RecentProject:
    data_dir: str
    out_dir: str
    karospace_html: Optional[str]
    last_used: str

    def label(self) -> str:
        return f"{self.data_dir} -> {self.out_dir}"


def _load_recent() -> List[RecentProject]:
    path = RECENT_PATH if RECENT_PATH.exists() else LEGACY_RECENT_PATH
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text())
    except json.JSONDecodeError:
        return []
    projects: List[RecentProject] = []
    for item in payload.get("projects", []):
        projects.append(
            RecentProject(
                data_dir=item.get("data_dir", ""),
                out_dir=item.get("out_dir", ""),
                karospace_html=item.get("karospace_html"),
                last_used=item.get("last_used", ""),
            )
        )
    return projects


def _save_recent(projects: List[RecentProject]) -> None:
    RECENT_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "projects": [
            {
                "data_dir": p.data_dir,
                "out_dir": p.out_dir,
                "karospace_html": p.karospace_html,
                "last_used": p.last_used,
            }
            for p in projects
        ]
    }
    RECENT_PATH.write_text(json.dumps(payload, indent=2))


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("InSituCore")
        self.resize(1200, 800)
        self.setMinimumSize(780, 520)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QtGui.QIcon(str(APP_ICON_PATH)))

        self.process: Optional[QtCore.QProcess] = None
        self._plot_processes: List[QtCore.QProcess] = []
        self._busy_counter = 0
        self._busy_base_text = "Running"
        self._runner_frames = ["o-/", "o_/", "o-\\", "o_\\"]
        self._busy_tick = 0
        self._busy_has_error = False
        self.current_out_dir: Optional[Path] = None
        self.current_karospace_html: Optional[Path] = None
        self.recent_projects: List[RecentProject] = _load_recent()
        self._theme_mode = self._detect_system_theme()
        self._manual_theme_override = False

        self._build_ui()
        self.activity_stage.setVisible(self.width() >= 980)
        self._busy_timer = QtCore.QTimer(self)
        self._busy_timer.setInterval(280)
        self._busy_timer.timeout.connect(self._animate_busy_state)
        self._apply_theme(self._theme_mode)
        self._connect_system_theme_signal()
        self._populate_recent()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        root.setObjectName("Root")
        root_layout = QtWidgets.QVBoxLayout(root)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(12)

        root_layout.addWidget(self._build_top_bar())

        self.recent_list = QtWidgets.QListWidget()
        self.recent_list.setObjectName("RecentList")
        self.recent_list.itemSelectionChanged.connect(self._on_recent_selected)

        recent_box, recent_layout = self._create_card(
            "Recent Projects",
            "Load existing runs without recomputing.",
        )
        recent_box.setMinimumWidth(220)
        recent_box.setMaximumWidth(320)
        recent_layout.addWidget(self.recent_list)

        self.load_recent_btn = QtWidgets.QPushButton("Load Outputs")
        self.load_recent_btn.clicked.connect(self._load_selected_recent)
        recent_layout.addWidget(self.load_recent_btn)

        self.workspace_nav = QtWidgets.QListWidget()
        self.workspace_nav.setObjectName("WorkspaceNav")
        self.workspace_nav.setMinimumWidth(170)
        self.workspace_nav.setMaximumWidth(240)
        self.workspace_nav.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

        self.workspace_stack = QtWidgets.QStackedWidget()
        workspace_pages = [
            ("Run Pipeline", self._wrap_scroll(self._build_run_tab())),
            ("Analysis Controls", self._wrap_scroll(self._build_analysis_tab())),
            ("QC Gallery", self._build_qc_tab()),
            ("Spatial Static", self._build_spatial_static_tab()),
            ("Spatial Interactive", self._build_spatial_tab()),
            ("UMAP", self._build_umap_tab()),
            ("Compartment Map", self._build_compartment_tab()),
            ("Gene Expression", self._build_gene_expression_tab()),
        ]
        for label, page in workspace_pages:
            self.workspace_nav.addItem(label)
            self.workspace_stack.addWidget(page)
        self.workspace_nav.currentRowChanged.connect(self._on_workspace_nav_changed)
        self.workspace_nav.setCurrentRow(0)

        workspace_body = QtWidgets.QWidget()
        workspace_layout = QtWidgets.QHBoxLayout(workspace_body)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(10)
        workspace_layout.addWidget(self.workspace_nav)
        workspace_layout.addWidget(self.workspace_stack, stretch=1)
        tabs_card, tabs_layout = self._create_card("Workspace")
        tabs_layout.addWidget(workspace_body, stretch=1)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.setChildrenCollapsible(True)
        splitter.addWidget(recent_box)
        splitter.addWidget(tabs_card)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([280, 920])
        root_layout.addWidget(splitter, stretch=1)

        self.setCentralWidget(root)

    def _wrap_scroll(self, content: QtWidgets.QWidget) -> QtWidgets.QScrollArea:
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setWidget(content)
        return scroll

    def _on_workspace_nav_changed(self, index: int) -> None:
        if index < 0:
            return
        if index >= self.workspace_stack.count():
            return
        self.workspace_stack.setCurrentIndex(index)

    def _build_top_bar(self) -> QtWidgets.QWidget:
        top_bar = QtWidgets.QFrame()
        top_bar.setObjectName("TopBar")
        layout = QtWidgets.QHBoxLayout(top_bar)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(1)
        title = QtWidgets.QLabel("InSituCore")
        title.setObjectName("TopTitle")
        subtitle = QtWidgets.QLabel("Local Spatial Analysis")
        subtitle.setObjectName("TopSubtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        layout.addLayout(title_col)
        layout.addStretch(1)

        self.top_run_btn = QtWidgets.QPushButton("Run")
        self.top_run_btn.setProperty("variant", "primary")
        self.top_run_btn.clicked.connect(self._run_pipeline)
        layout.addWidget(self.top_run_btn)

        self.top_load_btn = QtWidgets.QPushButton("Load Outputs")
        self.top_load_btn.clicked.connect(self._load_outputs_only)
        layout.addWidget(self.top_load_btn)

        self.theme_toggle_btn = QtWidgets.QPushButton("Dark")
        self.theme_toggle_btn.setCheckable(True)
        self.theme_toggle_btn.toggled.connect(self._toggle_theme)
        layout.addWidget(self.theme_toggle_btn)

        self.activity_stage = QtWidgets.QLabel("Idle")
        self.activity_stage.setObjectName("ActivityStage")
        self.activity_stage.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        layout.addWidget(self.activity_stage)

        self.runner_glyph = QtWidgets.QLabel("o-/")
        self.runner_glyph.setObjectName("RunnerGlyph")
        self.runner_glyph.setText("   ")
        self.runner_glyph.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        layout.addWidget(self.runner_glyph)

        self.status_chip = QtWidgets.QLabel("Ready")
        self.status_chip.setObjectName("StatusChip")
        self.status_chip.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        layout.addWidget(self.status_chip)
        return top_bar

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        show_stage = self.width() >= 980
        self.activity_stage.setVisible(show_stage)

    def _create_card(
        self,
        title: str,
        subtitle: Optional[str] = None,
    ) -> tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        card = QtWidgets.QFrame()
        card.setObjectName("Card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        if title:
            title_label = QtWidgets.QLabel(title)
            title_label.setObjectName("CardTitle")
            layout.addWidget(title_label)
        if subtitle:
            subtitle_label = QtWidgets.QLabel(subtitle)
            subtitle_label.setObjectName("CardSubtitle")
            layout.addWidget(subtitle_label)
        return card, layout

    def _build_run_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(10)

        data_card, data_layout = self._create_card(
            "Dataset",
            "Configure source and output folders for a pipeline run.",
        )

        form = QtWidgets.QFormLayout()
        form.setSpacing(10)

        self.data_dir_edit = QtWidgets.QLineEdit()
        self.data_dir_edit.setPlaceholderText("/path/to/dataset")
        data_btn = QtWidgets.QPushButton("Browse")
        data_btn.clicked.connect(lambda: self._choose_dir(self.data_dir_edit))
        data_row_w = QtWidgets.QWidget()
        data_row = QtWidgets.QHBoxLayout(data_row_w)
        data_row.setContentsMargins(0, 0, 0, 0)
        data_row.addWidget(self.data_dir_edit)
        data_row.addWidget(data_btn)
        form.addRow("Data dir", data_row_w)

        self.out_dir_edit = QtWidgets.QLineEdit()
        self.out_dir_edit.setPlaceholderText("/path/to/output")
        out_btn = QtWidgets.QPushButton("Browse")
        out_btn.clicked.connect(lambda: self._choose_dir(self.out_dir_edit))
        out_row_w = QtWidgets.QWidget()
        out_row = QtWidgets.QHBoxLayout(out_row_w)
        out_row.setContentsMargins(0, 0, 0, 0)
        out_row.addWidget(self.out_dir_edit)
        out_row.addWidget(out_btn)
        form.addRow("Output dir", out_row_w)

        self.run_prefix_edit = QtWidgets.QLineEdit("output-")
        form.addRow("Run prefix", self.run_prefix_edit)

        self.run_search_depth_combo = QtWidgets.QComboBox()
        self.run_search_depth_combo.addItem("Direct folders only", 1)
        self.run_search_depth_combo.addItem("One level below samples", 2)
        form.addRow("Run depth", self.run_search_depth_combo)

        self.sample_id_source_combo = QtWidgets.QComboBox()
        self.sample_id_source_combo.addItem("Auto", "auto")
        self.sample_id_source_combo.addItem("From run label", "run")
        self.sample_id_source_combo.addItem("From parent folder", "parent")
        form.addRow("Sample ID source", self.sample_id_source_combo)

        self.count_matrix_mode_combo = QtWidgets.QComboBox()
        self.count_matrix_mode_combo.addItem(
            "Standard Xenium cell matrix",
            "cell_feature_matrix",
        )
        self.count_matrix_mode_combo.addItem(
            "Nucleus OR distance-filtered transcripts",
            "nucleus_or_distance",
        )
        form.addRow("Count matrix mode", self.count_matrix_mode_combo)

        self.tx_max_distance_spin = QtWidgets.QDoubleSpinBox()
        self.tx_max_distance_spin.setDecimals(2)
        self.tx_max_distance_spin.setRange(0.0, 200.0)
        self.tx_max_distance_spin.setSingleStep(0.5)
        self.tx_max_distance_spin.setValue(5.0)
        form.addRow("Tx max distance (um)", self.tx_max_distance_spin)

        self.tx_nucleus_distance_key_edit = QtWidgets.QLineEdit("nucleus_distance")
        self.tx_nucleus_distance_key_edit.setPlaceholderText("e.g. nucleus_distance")
        form.addRow("Tx distance key", self.tx_nucleus_distance_key_edit)

        self.tx_allowed_categories_edit = QtWidgets.QLineEdit("predesigned_gene,custom_gene")
        self.tx_allowed_categories_edit.setPlaceholderText("comma-separated categories")
        form.addRow("Tx categories", self.tx_allowed_categories_edit)
        data_layout.addLayout(form)
        self.count_matrix_mode_combo.currentIndexChanged.connect(self._sync_count_matrix_controls)
        self._sync_count_matrix_controls()
        layout.addWidget(data_card)

        options_box, options_layout = self._create_card(
            "Optional Steps",
            "Enable downstream modules for weighted compartments and viewer export.",
        )

        self.mana_check = QtWidgets.QCheckBox("Enable MANA weighted aggregation")
        self.mana_layers = QtWidgets.QSpinBox()
        self.mana_layers.setRange(1, 10)
        self.mana_layers.setValue(3)
        self.mana_hop_decay = QtWidgets.QDoubleSpinBox()
        self.mana_hop_decay.setRange(0.0, 1.0)
        self.mana_hop_decay.setSingleStep(0.05)
        self.mana_hop_decay.setValue(0.2)
        self.mana_kernel = QtWidgets.QComboBox()
        self.mana_kernel.addItems(["exponential", "inverse", "gaussian", "none"])
        self.mana_kernel.setCurrentText("gaussian")
        self.mana_rep_mode = QtWidgets.QComboBox()
        self.mana_rep_mode.addItem("scVI latent (recommended)", "scvi")
        self.mana_rep_mode.addItem("PCA embedding", "pca")
        self.mana_rep_mode.addItem("Auto (scVI -> PCA)", "auto")
        self.mana_rep_mode.addItem("Custom obsm key", "custom")
        self.mana_custom_rep_edit = QtWidgets.QLineEdit("X_scVI")
        self.mana_custom_rep_edit.setPlaceholderText("e.g. X_scVI")

        mana_form = QtWidgets.QFormLayout()
        mana_form.addRow(self.mana_check)
        mana_form.addRow("Layers", self.mana_layers)
        mana_form.addRow("Hop decay", self.mana_hop_decay)
        mana_form.addRow("Distance kernel", self.mana_kernel)
        mana_form.addRow("Representation", self.mana_rep_mode)
        mana_form.addRow("Custom rep key", self.mana_custom_rep_edit)
        options_layout.addLayout(mana_form)
        self.mana_rep_mode.currentIndexChanged.connect(self._sync_mana_rep_controls)
        self._sync_mana_rep_controls()

        self.karospace_check = QtWidgets.QCheckBox("Export KaroSpace HTML")
        self.karospace_path_edit = QtWidgets.QLineEdit()
        self.karospace_path_edit.setPlaceholderText("/path/to/karospace.html")
        karospace_btn = QtWidgets.QPushButton("Browse")
        karospace_btn.clicked.connect(self._choose_karospace_path)
        karospace_row = QtWidgets.QHBoxLayout()
        karospace_row.addWidget(self.karospace_path_edit)
        karospace_row.addWidget(karospace_btn)

        options_layout.addWidget(self.karospace_check)
        options_layout.addLayout(karospace_row)

        layout.addWidget(options_box)

        action_hint = QtWidgets.QLabel("Primary actions are in the top bar.")
        action_hint.setObjectName("CardSubtitle")
        layout.addWidget(action_hint)

        logs_card, logs_layout = self._create_card("Run Log")
        self.log_view = QtWidgets.QPlainTextEdit()
        self.log_view.setObjectName("LogView")
        self.log_view.setReadOnly(True)
        logs_layout.addWidget(self.log_view, stretch=1)
        layout.addWidget(logs_card, stretch=1)

        return widget

    def _build_qc_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card("QC Gallery", "Generated plots from xenium_qc.")
        self.qc_scroll = QtWidgets.QScrollArea()
        self.qc_scroll.setWidgetResizable(True)
        self.qc_container = QtWidgets.QWidget()
        self.qc_layout = QtWidgets.QVBoxLayout(self.qc_container)
        self.qc_layout.addStretch(1)
        self.qc_scroll.setWidget(self.qc_container)
        card_layout.addWidget(self.qc_scroll, stretch=1)
        layout.addWidget(card, stretch=1)
        return widget

    def _build_analysis_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(10)

        intro_card, intro_layout = self._create_card(
            "Analysis Controls",
            "Graph, UMAP, and clustering settings.",
        )

        self.analysis_toggle_btn = QtWidgets.QToolButton()
        self.analysis_toggle_btn.setText("Hide advanced parameters")
        self.analysis_toggle_btn.setCheckable(True)
        self.analysis_toggle_btn.setChecked(True)
        self.analysis_toggle_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.analysis_toggle_btn.setArrowType(QtCore.Qt.DownArrow)
        intro_layout.addWidget(self.analysis_toggle_btn)

        self.analysis_panel = QtWidgets.QWidget()
        self.analysis_panel.setVisible(True)
        panel_layout = QtWidgets.QVBoxLayout(self.analysis_panel)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(10)

        graph_card, graph_layout = self._create_card(
            "Graph + UMAP",
            "Tune neighborhood graph and UMAP embedding before clustering.",
        )
        graph_form = QtWidgets.QFormLayout()
        graph_form.setSpacing(10)

        self.n_neighbors_spin = QtWidgets.QSpinBox()
        self.n_neighbors_spin.setRange(2, 200)
        self.n_neighbors_spin.setValue(15)
        graph_form.addRow("Neighbors (n_neighbors)", self.n_neighbors_spin)

        self.n_pcs_spin = QtWidgets.QSpinBox()
        self.n_pcs_spin.setRange(2, 200)
        self.n_pcs_spin.setValue(30)
        graph_form.addRow("PCs (n_pcs)", self.n_pcs_spin)

        self.umap_min_dist_spin = QtWidgets.QDoubleSpinBox()
        self.umap_min_dist_spin.setDecimals(3)
        self.umap_min_dist_spin.setRange(0.0, 1.0)
        self.umap_min_dist_spin.setSingleStep(0.05)
        self.umap_min_dist_spin.setValue(0.1)
        graph_form.addRow("UMAP min_dist", self.umap_min_dist_spin)

        self.cluster_graph_mode_combo = QtWidgets.QComboBox()
        self.cluster_graph_mode_combo.addItem("Auto (prefer spatial)", "auto")
        self.cluster_graph_mode_combo.addItem("Expression graph", "expression")
        self.cluster_graph_mode_combo.addItem("Spatial graph", "spatial")
        graph_form.addRow("Graph source", self.cluster_graph_mode_combo)
        graph_layout.addLayout(graph_form)
        panel_layout.addWidget(graph_card)

        cluster_card, cluster_layout = self._create_card(
            "Clustering",
            "Choose Leiden, Louvain, or KMeans and define sweep values.",
        )
        cluster_form = QtWidgets.QFormLayout()
        cluster_form.setSpacing(10)

        self.cluster_method_combo = QtWidgets.QComboBox()
        self.cluster_method_combo.addItem("Leiden", "leiden")
        self.cluster_method_combo.addItem("Louvain", "louvain")
        self.cluster_method_combo.addItem("KMeans", "kmeans")
        cluster_form.addRow("Method", self.cluster_method_combo)

        self.leiden_res_edit = QtWidgets.QLineEdit("0.1,0.5,1,1.5,2")
        self.leiden_res_edit.setPlaceholderText("e.g. 0.5,1.0,1.5")
        cluster_form.addRow("Leiden resolutions", self.leiden_res_edit)

        self.louvain_res_edit = QtWidgets.QLineEdit("0.5,1.0")
        self.louvain_res_edit.setPlaceholderText("e.g. 0.5,1.0")
        cluster_form.addRow("Louvain resolutions", self.louvain_res_edit)

        self.kmeans_clusters_edit = QtWidgets.QLineEdit("8,12")
        self.kmeans_clusters_edit.setPlaceholderText("e.g. 8,12,16")
        cluster_form.addRow("KMeans k values", self.kmeans_clusters_edit)

        self.kmeans_random_state_spin = QtWidgets.QSpinBox()
        self.kmeans_random_state_spin.setRange(0, 999999)
        self.kmeans_random_state_spin.setValue(0)
        cluster_form.addRow("KMeans random_state", self.kmeans_random_state_spin)

        self.kmeans_n_init_spin = QtWidgets.QSpinBox()
        self.kmeans_n_init_spin.setRange(1, 50)
        self.kmeans_n_init_spin.setValue(10)
        cluster_form.addRow("KMeans n_init", self.kmeans_n_init_spin)

        cluster_layout.addLayout(cluster_form)
        panel_layout.addWidget(cluster_card)
        panel_layout.addStretch(1)

        intro_layout.addWidget(self.analysis_panel)
        layout.addWidget(intro_card, stretch=1)

        self.analysis_toggle_btn.toggled.connect(self._toggle_analysis_panel)
        self._toggle_analysis_panel(True)
        self.cluster_method_combo.currentIndexChanged.connect(self._sync_cluster_controls)
        self._sync_cluster_controls()
        return widget

    def _toggle_analysis_panel(self, checked: bool) -> None:
        self.analysis_panel.setVisible(checked)
        self.analysis_toggle_btn.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.analysis_toggle_btn.setText("Hide advanced parameters" if checked else "Show advanced parameters")

    def _sync_cluster_controls(self) -> None:
        method = str(self.cluster_method_combo.currentData())
        self.leiden_res_edit.setEnabled(method == "leiden")
        self.louvain_res_edit.setEnabled(method == "louvain")
        kmeans_enabled = method == "kmeans"
        self.kmeans_clusters_edit.setEnabled(kmeans_enabled)
        self.kmeans_random_state_spin.setEnabled(kmeans_enabled)
        self.kmeans_n_init_spin.setEnabled(kmeans_enabled)

    def _sync_count_matrix_controls(self) -> None:
        transcript_mode = str(self.count_matrix_mode_combo.currentData()) == "nucleus_or_distance"
        self.tx_max_distance_spin.setEnabled(transcript_mode)
        self.tx_nucleus_distance_key_edit.setEnabled(transcript_mode)
        self.tx_allowed_categories_edit.setEnabled(transcript_mode)

    def _sync_mana_rep_controls(self) -> None:
        mode = str(self.mana_rep_mode.currentData())
        self.mana_custom_rep_edit.setEnabled(mode == "custom")

    def _build_spatial_static_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card(
            "Spatial Map (Static)",
            "Fast map from clustered.h5ad using plot_spatial_compact_fast.",
        )

        controls_row = QtWidgets.QHBoxLayout()
        controls_row.setSpacing(10)
        key_label = QtWidgets.QLabel("Color key")
        key_label.setObjectName("CardSubtitle")
        self.spatial_key_combo = QtWidgets.QComboBox()
        self.spatial_key_combo.addItem("Auto (cluster key)", "")
        self.spatial_key_combo.setEnabled(False)
        self.generate_spatial_btn = QtWidgets.QPushButton("Generate Spatial Map")
        self.generate_spatial_btn.clicked.connect(self._generate_spatial_map)
        controls_row.addWidget(key_label)
        controls_row.addWidget(self.spatial_key_combo, stretch=1)
        controls_row.addWidget(self.generate_spatial_btn)
        card_layout.addLayout(controls_row)

        self.spatial_static_label = QtWidgets.QLabel("No spatial map found. Click Generate Spatial Map.")
        self.spatial_static_label.setObjectName("PreviewSurface")
        self.spatial_static_label.setAlignment(QtCore.Qt.AlignCenter)
        self.spatial_static_label.setMinimumHeight(260)
        card_layout.addWidget(self.spatial_static_label, stretch=1)
        layout.addWidget(card, stretch=1)
        return widget

    def _build_spatial_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card(
            "Spatial Map (Interactive)",
            "KaroSpace viewer output for section-level inspection.",
        )

        if WEB_AVAILABLE:
            self.spatial_view = QWebEngineView()
            card_layout.addWidget(self.spatial_view, stretch=1)
        else:
            self.spatial_view = None
            self.spatial_fallback_label = QtWidgets.QLabel(
                "Qt WebEngine not available. Spatial viewer will open in your browser."
            )
            self.spatial_fallback_label.setAlignment(QtCore.Qt.AlignCenter)
            card_layout.addWidget(self.spatial_fallback_label, stretch=1)
            self.spatial_open_btn = QtWidgets.QPushButton("Open KaroSpace in Browser")
            self.spatial_open_btn.clicked.connect(self._open_karospace_external)
            card_layout.addWidget(self.spatial_open_btn)

        layout.addWidget(card, stretch=1)
        return widget

    def _build_umap_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card("UMAP", "Top bar action: Generate UMAP.")
        key_row = QtWidgets.QHBoxLayout()
        key_label = QtWidgets.QLabel("Cluster key")
        key_label.setObjectName("CardSubtitle")
        self.umap_key_combo = QtWidgets.QComboBox()
        self.umap_key_combo.addItem("Auto (cluster key)", "")
        self.umap_key_combo.setEnabled(False)
        self.generate_umap_btn = QtWidgets.QPushButton("Generate UMAP")
        self.generate_umap_btn.clicked.connect(self._generate_umap_plot)
        key_row.addWidget(key_label)
        key_row.addWidget(self.umap_key_combo, stretch=1)
        key_row.addWidget(self.generate_umap_btn)
        card_layout.addLayout(key_row)

        self.umap_label = QtWidgets.QLabel("No UMAP image loaded.")
        self.umap_label.setObjectName("PreviewSurface")
        self.umap_label.setAlignment(QtCore.Qt.AlignCenter)
        self.umap_label.setMinimumHeight(260)
        card_layout.addWidget(self.umap_label, stretch=1)
        layout.addWidget(card, stretch=1)
        return widget

    def _build_compartment_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card(
            "Compartment Map",
            "Top bar action: Generate Compartments.",
        )
        key_row = QtWidgets.QHBoxLayout()
        key_label = QtWidgets.QLabel("Compartment key")
        key_label.setObjectName("CardSubtitle")
        self.compartment_key_combo = QtWidgets.QComboBox()
        self.compartment_key_combo.addItem("Auto (primary)", "")
        self.compartment_key_combo.setEnabled(False)
        self.generate_compartment_btn = QtWidgets.QPushButton("Generate Compartments")
        self.generate_compartment_btn.clicked.connect(self._generate_compartment_map)
        key_row.addWidget(key_label)
        key_row.addWidget(self.compartment_key_combo, stretch=1)
        key_row.addWidget(self.generate_compartment_btn)
        card_layout.addLayout(key_row)
        self.compartment_label = QtWidgets.QLabel("No compartment map loaded.")
        self.compartment_label.setObjectName("PreviewSurface")
        self.compartment_label.setAlignment(QtCore.Qt.AlignCenter)
        self.compartment_label.setMinimumHeight(260)
        card_layout.addWidget(self.compartment_label, stretch=1)
        layout.addWidget(card, stretch=1)
        return widget

    def _build_gene_expression_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card(
            "Gene Expression Dotplot",
            "Top markers per cluster/compartment key.",
        )
        controls_row = QtWidgets.QHBoxLayout()
        controls_row.setSpacing(10)

        key_label = QtWidgets.QLabel("Group by")
        key_label.setObjectName("CardSubtitle")
        self.gene_expr_key_combo = QtWidgets.QComboBox()
        self.gene_expr_key_combo.addItem("Auto (cluster key)", "")
        self.gene_expr_key_combo.setEnabled(False)

        top_n_label = QtWidgets.QLabel("Top genes")
        top_n_label.setObjectName("CardSubtitle")
        self.gene_expr_top_n_spin = QtWidgets.QSpinBox()
        self.gene_expr_top_n_spin.setRange(1, 50)
        self.gene_expr_top_n_spin.setValue(10)

        self.generate_gene_expr_btn = QtWidgets.QPushButton("Generate Dotplot")
        self.generate_gene_expr_btn.clicked.connect(self._generate_gene_expression_dotplot)

        controls_row.addWidget(key_label)
        controls_row.addWidget(self.gene_expr_key_combo, stretch=1)
        controls_row.addWidget(top_n_label)
        controls_row.addWidget(self.gene_expr_top_n_spin)
        controls_row.addWidget(self.generate_gene_expr_btn)
        card_layout.addLayout(controls_row)

        self.gene_expr_label = QtWidgets.QLabel("No dotplot loaded.")
        self.gene_expr_label.setObjectName("PreviewSurface")
        self.gene_expr_label.setAlignment(QtCore.Qt.AlignCenter)
        self.gene_expr_label.setMinimumHeight(260)
        card_layout.addWidget(self.gene_expr_label, stretch=1)

        layout.addWidget(card, stretch=1)
        return widget

    def _choose_dir(self, line_edit: QtWidgets.QLineEdit) -> None:
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select directory")
        if path:
            line_edit.setText(path)

    def _choose_karospace_path(self) -> None:
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Select KaroSpace HTML",
            "karospace.html",
            "HTML files (*.html)"
        )
        if path:
            self.karospace_path_edit.setText(path)

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_view.appendPlainText(f"[{timestamp}] {message}")

    def _collect_existing_pipeline_outputs(
        self,
        out_dir: Path,
        karospace_path: Optional[Path],
    ) -> List[Path]:
        candidates = [
            out_dir / "data" / "raw.h5ad",
            out_dir / "data" / "clustered.h5ad",
            out_dir / "data" / "cluster_info.json",
            out_dir / "data" / "markers_by_cluster.csv",
            out_dir / "xenium_qc" / "summary_by_run.csv",
            out_dir / "xenium_qc" / "gene_detection_overall.csv",
            out_dir / "plots" / "spatial.png",
            out_dir / "plots" / "umap.png",
            out_dir / "plots" / "compartments.png",
        ]
        if karospace_path is not None:
            candidates.append(karospace_path)
        return [p for p in candidates if p.exists()]

    def _confirm_overwrite(
        self,
        paths: List[Path],
        *,
        title: str,
        prompt: str,
    ) -> bool:
        if not paths:
            return True
        shown = [str(p) for p in paths[:10]]
        details = "\n".join(shown)
        if len(paths) > 10:
            details += f"\n... and {len(paths) - 10} more"

        msg = QtWidgets.QMessageBox(self)
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowTitle(title)
        msg.setText(prompt)
        msg.setInformativeText("Existing files will be replaced.")
        msg.setDetailedText(details)
        msg.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        msg.setDefaultButton(QtWidgets.QMessageBox.No)
        return msg.exec() == QtWidgets.QMessageBox.Yes

    def _enter_busy(self, stage_text: str) -> None:
        if self._busy_counter == 0:
            self._busy_has_error = False
            self._busy_tick = 0
            self._busy_timer.start()
        self._busy_counter += 1
        self._busy_base_text = stage_text
        self._set_activity_stage(stage_text)
        self.status_chip.setText("Running")

    def _leave_busy(self, failed: bool = False) -> None:
        if failed:
            self._busy_has_error = True
        if self._busy_counter > 0:
            self._busy_counter -= 1
        if self._busy_counter == 0:
            self._busy_timer.stop()
            self.runner_glyph.setText("   ")
            self._set_activity_stage("Idle")
            self.status_chip.setText("Error" if self._busy_has_error else "Ready")

    def _animate_busy_state(self) -> None:
        if self._busy_counter <= 0:
            self.runner_glyph.setText("   ")
            self.status_chip.setText("Ready")
            return
        frame = self._runner_frames[self._busy_tick % len(self._runner_frames)]
        self.runner_glyph.setText(frame)
        self.status_chip.setText("Running")
        self._busy_tick += 1

    def _set_activity_stage(self, text: str) -> None:
        # Keep top bar width stable while still showing useful per-step status.
        fm = self.activity_stage.fontMetrics()
        elided = fm.elidedText(text, QtCore.Qt.ElideRight, self.activity_stage.width() or 240)
        self.activity_stage.setText(elided)
        self.activity_stage.setToolTip(text)

    def _update_stage_from_log(self, line: str) -> None:
        if line.startswith("STEP: "):
            stage_text = line.replace("STEP: ", "", 1).strip()
            self._busy_base_text = stage_text
            self._set_activity_stage(stage_text)
            return

        stage_map = [
            ("Loading run:", "Loading runs"),
            ("Saved raw AnnData:", "Preparing QC"),
            ("Saved QC outputs:", "QC complete"),
            ("Running MANA weighted aggregation...", "MANA aggregation"),
            ("Saved clustered AnnData:", "Saving clustered data"),
            ("Exporting KaroSpace HTML...", "Exporting KaroSpace"),
        ]
        for token, stage_text in stage_map:
            if token in line:
                self._busy_base_text = stage_text
                self._set_activity_stage(stage_text)
                return

    def _theme_path(self) -> Path:
        return THEME_DARK_PATH if self._theme_mode == "dark" else THEME_LIGHT_PATH

    def _detect_system_theme(self) -> str:
        app = QtWidgets.QApplication.instance()
        if app is None:
            return "light"

        style_hints = app.styleHints()
        color_scheme = getattr(style_hints, "colorScheme", lambda: None)()
        dark_enum = getattr(getattr(QtCore.Qt, "ColorScheme", object), "Dark", None)
        light_enum = getattr(getattr(QtCore.Qt, "ColorScheme", object), "Light", None)
        if dark_enum is not None and color_scheme == dark_enum:
            return "dark"
        if light_enum is not None and color_scheme == light_enum:
            return "light"

        # Fallback for environments without colorScheme support.
        bg = app.palette().color(QtGui.QPalette.Window)
        return "dark" if bg.lightness() < 128 else "light"

    def _connect_system_theme_signal(self) -> None:
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        style_hints = app.styleHints()
        signal = getattr(style_hints, "colorSchemeChanged", None)
        if signal is not None:
            signal.connect(self._on_system_color_scheme_changed)

    def _toggle_theme(self, checked: bool) -> None:
        self._manual_theme_override = True
        self._theme_mode = "dark" if checked else "light"
        self._apply_theme(self._theme_mode)

    def _on_system_color_scheme_changed(self, _scheme: object) -> None:
        if self._manual_theme_override:
            return
        self._apply_theme(self._detect_system_theme())

    def _apply_theme(self, mode: Optional[str] = None) -> None:
        if mode in {"light", "dark"}:
            self._theme_mode = mode
        app = QtWidgets.QApplication.instance()
        theme_path = self._theme_path()
        if app is None or not theme_path.exists():
            return
        app.setStyleSheet(theme_path.read_text())
        self.theme_toggle_btn.blockSignals(True)
        self.theme_toggle_btn.setChecked(self._theme_mode == "dark")
        self.theme_toggle_btn.setText("Light" if self._theme_mode == "dark" else "Dark")
        self.theme_toggle_btn.blockSignals(False)
        self._log(f"Theme set: {self._theme_mode}")

    def _run_pipeline(self) -> None:
        data_dir = self.data_dir_edit.text().strip()
        out_dir = self.out_dir_edit.text().strip()
        if not data_dir or not out_dir:
            QtWidgets.QMessageBox.warning(self, "Missing paths", "Please set data and output directories.")
            return

        out_dir_path = Path(out_dir).expanduser().resolve()
        karospace_path_obj: Optional[Path] = None

        args = [
            sys.executable,
            "-u",
            str(ROOT_DIR / "run_xenium_analysis.py"),
            "--data-dir",
            data_dir,
            "--out-dir",
            out_dir,
            "--run-prefix",
            self.run_prefix_edit.text().strip() or "output-",
            "--run-search-depth",
            str(self.run_search_depth_combo.currentData()),
            "--sample-id-source",
            str(self.sample_id_source_combo.currentData()),
            "--count-matrix-mode",
            str(self.count_matrix_mode_combo.currentData()),
            "--n-neighbors",
            str(self.n_neighbors_spin.value()),
            "--n-pcs",
            str(self.n_pcs_spin.value()),
            "--umap-min-dist",
            str(self.umap_min_dist_spin.value()),
            "--cluster-graph-mode",
            str(self.cluster_graph_mode_combo.currentData()),
            "--cluster-method",
            str(self.cluster_method_combo.currentData()),
            "--leiden-resolutions",
            self.leiden_res_edit.text().strip() or "0.1,0.5,1,1.5,2",
            "--louvain-resolutions",
            self.louvain_res_edit.text().strip() or "0.5,1.0",
            "--kmeans-clusters",
            self.kmeans_clusters_edit.text().strip() or "8,12",
            "--kmeans-random-state",
            str(self.kmeans_random_state_spin.value()),
            "--kmeans-n-init",
            str(self.kmeans_n_init_spin.value()),
        ]

        if str(self.count_matrix_mode_combo.currentData()) == "nucleus_or_distance":
            args += [
                "--tx-max-distance-um",
                str(self.tx_max_distance_spin.value()),
                "--tx-nucleus-distance-key",
                self.tx_nucleus_distance_key_edit.text().strip() or "nucleus_distance",
                "--tx-allowed-categories",
                self.tx_allowed_categories_edit.text().strip() or "predesigned_gene,custom_gene",
            ]

        if self.mana_check.isChecked():
            rep_mode = str(self.mana_rep_mode.currentData())
            args += [
                "--mana-aggregate",
                "--mana-n-layers",
                str(self.mana_layers.value()),
                "--mana-hop-decay",
                str(self.mana_hop_decay.value()),
                "--mana-distance-kernel",
                self.mana_kernel.currentText(),
                "--mana-representation-mode",
                rep_mode,
            ]
            if rep_mode == "custom":
                custom_rep = self.mana_custom_rep_edit.text().strip()
                if not custom_rep:
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Missing custom representation key",
                        "Please provide a custom obsm key for MANA representation.",
                    )
                    return
                args += ["--mana-use-rep", custom_rep]

        if self.karospace_check.isChecked():
            karospace_path = self.karospace_path_edit.text().strip()
            if not karospace_path:
                karospace_path = str(out_dir_path / "karospace.html")
                self.karospace_path_edit.setText(karospace_path)
            karospace_path_obj = Path(karospace_path).expanduser().resolve()
            args += ["--karospace-html", karospace_path]

        existing_outputs = self._collect_existing_pipeline_outputs(
            out_dir=out_dir_path,
            karospace_path=karospace_path_obj,
        )
        if not self._confirm_overwrite(
            existing_outputs,
            title="Existing outputs found",
            prompt="Run pipeline and overwrite existing outputs?",
        ):
            self._log("Run cancelled by user (existing outputs).")
            return

        self._log("Starting pipeline...")
        self._enter_busy("Running pipeline")
        self.top_run_btn.setEnabled(False)
        self.process = QtCore.QProcess(self)
        self.process.setProgram(args[0])
        self.process.setArguments(args[1:])
        self.process.setWorkingDirectory(str(ROOT_DIR))
        self.process.setProcessChannelMode(QtCore.QProcess.MergedChannels)
        self.process.readyReadStandardOutput.connect(self._on_process_output)
        self.process.finished.connect(self._on_process_finished)
        self.process.start()

    def _on_process_output(self) -> None:
        if not self.process:
            return
        text = self.process.readAllStandardOutput().data().decode("utf-8", errors="ignore")
        for line in text.splitlines():
            if line.strip():
                self._log(line)
                self._update_stage_from_log(line)

    def _on_process_finished(self, exit_code: int, _status: QtCore.QProcess.ExitStatus) -> None:
        self.top_run_btn.setEnabled(True)
        self._log(f"Pipeline finished (exit code {exit_code}).")
        self._leave_busy(failed=exit_code != 0)
        out_dir = self.out_dir_edit.text().strip()
        if out_dir:
            self._load_outputs(Path(out_dir))
            self._update_recent()

    def _load_outputs_only(self) -> None:
        out_dir = self.out_dir_edit.text().strip()
        if not out_dir:
            out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select output directory")
            if not out_dir:
                return
            self.out_dir_edit.setText(out_dir)
        self._load_outputs(Path(out_dir))
        self._update_recent()

    def _load_outputs(self, out_dir: Path) -> None:
        self.current_out_dir = out_dir
        self._refresh_spatial_keys(out_dir)
        self._refresh_umap_keys(out_dir)
        self._refresh_compartment_keys(out_dir)
        self._refresh_gene_expression_keys(out_dir)
        self._load_qc_images(out_dir)
        self._load_karospace(out_dir)
        self._load_spatial_image(out_dir)
        self._load_umap_image(out_dir)
        self._load_compartment_image(out_dir)
        self._load_gene_expression_image(out_dir)

    def _refresh_spatial_keys(self, out_dir: Path) -> None:
        keys: List[str] = []
        cluster_info_path = out_dir / "data" / "cluster_info.json"
        if cluster_info_path.exists():
            try:
                payload = json.loads(cluster_info_path.read_text())
            except json.JSONDecodeError:
                payload = {}

            ordered_sources = [
                payload.get("cluster_key"),
                payload.get("compartment_key"),
                payload.get("cluster_keys"),
                payload.get("compartment_keys"),
            ]
            for source in ordered_sources:
                if isinstance(source, list):
                    candidates = source
                else:
                    candidates = [source]
                for key in candidates:
                    key_text = str(key or "").strip()
                    if key_text and key_text not in keys:
                        keys.append(key_text)

        self.spatial_key_combo.blockSignals(True)
        self.spatial_key_combo.clear()
        self.spatial_key_combo.addItem("Auto (cluster key)", "")
        for key in keys:
            self.spatial_key_combo.addItem(key, key)
        self.spatial_key_combo.setEnabled(bool(keys))
        self.spatial_key_combo.blockSignals(False)

    def _refresh_compartment_keys(self, out_dir: Path) -> None:
        keys: List[str] = []
        cluster_info_path = out_dir / "data" / "cluster_info.json"
        if cluster_info_path.exists():
            try:
                payload = json.loads(cluster_info_path.read_text())
            except json.JSONDecodeError:
                payload = {}

            raw_keys = payload.get("compartment_keys", [])
            if isinstance(raw_keys, list):
                for key in raw_keys:
                    key_text = str(key).strip()
                    if key_text and key_text not in keys:
                        keys.append(key_text)

            primary_key = str(payload.get("compartment_key") or "").strip()
            if primary_key and primary_key not in keys:
                keys.insert(0, primary_key)

        self.compartment_key_combo.blockSignals(True)
        self.compartment_key_combo.clear()
        self.compartment_key_combo.addItem("Auto (primary)", "")
        for key in keys:
            self.compartment_key_combo.addItem(key, key)
        self.compartment_key_combo.setEnabled(bool(keys))
        self.compartment_key_combo.blockSignals(False)

    def _refresh_umap_keys(self, out_dir: Path) -> None:
        keys: List[str] = []
        cluster_info_path = out_dir / "data" / "cluster_info.json"
        if cluster_info_path.exists():
            try:
                payload = json.loads(cluster_info_path.read_text())
            except json.JSONDecodeError:
                payload = {}

            ordered_sources = [
                payload.get("cluster_key"),
                payload.get("cluster_keys"),
                payload.get("compartment_key"),
                payload.get("compartment_keys"),
            ]
            for source in ordered_sources:
                if isinstance(source, list):
                    candidates = source
                else:
                    candidates = [source]
                for key in candidates:
                    key_text = str(key or "").strip()
                    if key_text and key_text not in keys:
                        keys.append(key_text)

        self.umap_key_combo.blockSignals(True)
        self.umap_key_combo.clear()
        self.umap_key_combo.addItem("Auto (cluster key)", "")
        for key in keys:
            self.umap_key_combo.addItem(key, key)
        self.umap_key_combo.setEnabled(bool(keys))
        self.umap_key_combo.blockSignals(False)

    def _refresh_gene_expression_keys(self, out_dir: Path) -> None:
        keys: List[str] = []
        cluster_info_path = out_dir / "data" / "cluster_info.json"
        if cluster_info_path.exists():
            try:
                payload = json.loads(cluster_info_path.read_text())
            except json.JSONDecodeError:
                payload = {}

            ordered_sources = [
                payload.get("cluster_key"),
                payload.get("cluster_keys"),
                payload.get("compartment_key"),
                payload.get("compartment_keys"),
            ]
            for source in ordered_sources:
                if isinstance(source, list):
                    candidates = source
                else:
                    candidates = [source]
                for key in candidates:
                    key_text = str(key or "").strip()
                    if key_text and key_text not in keys:
                        keys.append(key_text)

        self.gene_expr_key_combo.blockSignals(True)
        self.gene_expr_key_combo.clear()
        self.gene_expr_key_combo.addItem("Auto (cluster key)", "")
        for key in keys:
            self.gene_expr_key_combo.addItem(key, key)
        self.gene_expr_key_combo.setEnabled(bool(keys))
        self.gene_expr_key_combo.blockSignals(False)

    def _load_qc_images(self, out_dir: Path) -> None:
        for i in reversed(range(self.qc_layout.count())):
            item = self.qc_layout.takeAt(i)
            if item and item.widget():
                item.widget().deleteLater()
        qc_dir = out_dir / "xenium_qc"
        if not qc_dir.exists():
            self.qc_layout.addWidget(QtWidgets.QLabel("No QC outputs found."))
            self.qc_layout.addStretch(1)
            return

        images = sorted(qc_dir.glob("*.png"))
        if not images:
            self.qc_layout.addWidget(QtWidgets.QLabel("No QC images found."))
            self.qc_layout.addStretch(1)
            return

        for img_path in images:
            label = QtWidgets.QLabel()
            label.setAlignment(QtCore.Qt.AlignCenter)
            pixmap = QtGui.QPixmap(str(img_path))
            if not pixmap.isNull():
                label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
            else:
                label.setText(str(img_path))
            self.qc_layout.addWidget(label)

        self.qc_layout.addStretch(1)

    def _load_karospace(self, out_dir: Path) -> None:
        karospace_path = None
        candidate = out_dir / "karospace.html"
        if candidate.exists():
            karospace_path = candidate
        elif self.karospace_path_edit.text().strip():
            path = Path(self.karospace_path_edit.text().strip())
            if path.exists():
                karospace_path = path

        self.current_karospace_html = karospace_path
        if WEB_AVAILABLE and self.spatial_view is not None:
            if karospace_path:
                self.spatial_view.load(QtCore.QUrl.fromLocalFile(str(karospace_path)))
            else:
                self.spatial_view.setHtml("<p>No KaroSpace HTML found. Run with export enabled.</p>")
        elif not WEB_AVAILABLE:
            if karospace_path:
                self.spatial_fallback_label.setText(f"KaroSpace HTML: {karospace_path}")
            else:
                self.spatial_fallback_label.setText("No KaroSpace HTML found. Run with export enabled.")

    def _open_karospace_external(self) -> None:
        if not self.current_karospace_html or not self.current_karospace_html.exists():
            QtWidgets.QMessageBox.information(
                self,
                "No KaroSpace HTML",
                "No KaroSpace HTML found. Run with export enabled first.",
            )
            return
        QtGui.QDesktopServices.openUrl(
            QtCore.QUrl.fromLocalFile(str(self.current_karospace_html))
        )

    def _load_spatial_image(self, out_dir: Path) -> None:
        plot_path = out_dir / "plots" / "spatial.png"
        if plot_path.exists():
            pixmap = QtGui.QPixmap(str(plot_path))
            self.spatial_static_label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
        else:
            self.spatial_static_label.setText("No spatial map found. Click Generate Spatial Map.")

    def _load_umap_image(self, out_dir: Path) -> None:
        plot_path = out_dir / "plots" / "umap.png"
        if plot_path.exists():
            pixmap = QtGui.QPixmap(str(plot_path))
            self.umap_label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
        else:
            self.umap_label.setText("No UMAP image found. Click Generate UMAP Plot.")

    def _load_compartment_image(self, out_dir: Path) -> None:
        plot_path = out_dir / "plots" / "compartments.png"
        if plot_path.exists():
            pixmap = QtGui.QPixmap(str(plot_path))
            self.compartment_label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
        else:
            self.compartment_label.setText("No compartment map found. Click Generate Compartment Map.")

    def _load_gene_expression_image(self, out_dir: Path) -> None:
        plot_path = out_dir / "plots" / "gene_expression_dotplot.png"
        if plot_path.exists():
            pixmap = QtGui.QPixmap(str(plot_path))
            self.gene_expr_label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
        else:
            self.gene_expr_label.setText("No gene expression dotplot found. Click Generate Dotplot.")

    def _generate_spatial_map(self) -> None:
        if not self.current_out_dir:
            QtWidgets.QMessageBox.warning(self, "Missing output", "Load outputs first.")
            return

        h5ad_path = self.current_out_dir / "data" / "clustered.h5ad"
        if not h5ad_path.exists():
            QtWidgets.QMessageBox.warning(self, "Missing file", "clustered.h5ad not found.")
            return

        output_dir = self.current_out_dir / "plots"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "spatial.png"
        if not self._confirm_overwrite(
            [output_path] if output_path.exists() else [],
            title="Existing plot found",
            prompt="Regenerate spatial map and overwrite existing file?",
        ):
            self._log("Spatial map generation cancelled by user.")
            return

        args = [
            sys.executable,
            "-u",
            "-m",
            "utils.app_visuals",
            "spatial",
            "--h5ad",
            str(h5ad_path),
            "--output",
            str(output_path),
        ]
        selected_key = str(self.spatial_key_combo.currentData() or "").strip()
        if selected_key:
            args += ["--color", selected_key]

        self._run_visual_process(args, output_path, self.spatial_static_label)

    def _generate_umap_plot(self) -> None:
        if not self.current_out_dir:
            QtWidgets.QMessageBox.warning(self, "Missing output", "Load outputs first.")
            return

        h5ad_path = self.current_out_dir / "data" / "clustered.h5ad"
        if not h5ad_path.exists():
            QtWidgets.QMessageBox.warning(self, "Missing file", "clustered.h5ad not found.")
            return

        output_dir = self.current_out_dir / "plots"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "umap.png"
        if not self._confirm_overwrite(
            [output_path] if output_path.exists() else [],
            title="Existing plot found",
            prompt="Regenerate UMAP plot and overwrite existing file?",
        ):
            self._log("UMAP generation cancelled by user.")
            return

        args = [
            sys.executable,
            "-u",
            "-m",
            "utils.app_visuals",
            "umap",
            "--h5ad",
            str(h5ad_path),
            "--output",
            str(output_path),
        ]
        selected_key = str(self.umap_key_combo.currentData() or "").strip()
        if selected_key:
            args += ["--color", selected_key]

        self._run_visual_process(args, output_path, self.umap_label)

    def _generate_compartment_map(self) -> None:
        if not self.current_out_dir:
            QtWidgets.QMessageBox.warning(self, "Missing output", "Load outputs first.")
            return

        h5ad_path = self.current_out_dir / "data" / "clustered.h5ad"
        if not h5ad_path.exists():
            QtWidgets.QMessageBox.warning(self, "Missing file", "clustered.h5ad not found.")
            return

        output_dir = self.current_out_dir / "plots"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "compartments.png"
        if not self._confirm_overwrite(
            [output_path] if output_path.exists() else [],
            title="Existing plot found",
            prompt="Regenerate compartment plot and overwrite existing file?",
        ):
            self._log("Compartment generation cancelled by user.")
            return

        args = [
            sys.executable,
            "-u",
            "-m",
            "utils.app_visuals",
            "compartments",
            "--h5ad",
            str(h5ad_path),
            "--output",
            str(output_path),
        ]
        selected_key = str(self.compartment_key_combo.currentData() or "").strip()
        if selected_key:
            args += ["--color", selected_key]

        self._run_visual_process(args, output_path, self.compartment_label)

    def _generate_gene_expression_dotplot(self) -> None:
        if not self.current_out_dir:
            QtWidgets.QMessageBox.warning(self, "Missing output", "Load outputs first.")
            return

        h5ad_path = self.current_out_dir / "data" / "clustered.h5ad"
        if not h5ad_path.exists():
            QtWidgets.QMessageBox.warning(self, "Missing file", "clustered.h5ad not found.")
            return

        output_dir = self.current_out_dir / "plots"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "gene_expression_dotplot.png"
        if not self._confirm_overwrite(
            [output_path] if output_path.exists() else [],
            title="Existing plot found",
            prompt="Regenerate gene expression dotplot and overwrite existing file?",
        ):
            self._log("Gene expression dotplot generation cancelled by user.")
            return

        args = [
            sys.executable,
            "-u",
            "-m",
            "utils.app_visuals",
            "dotplot",
            "--h5ad",
            str(h5ad_path),
            "--output",
            str(output_path),
            "--top-n",
            str(self.gene_expr_top_n_spin.value()),
        ]
        selected_key = str(self.gene_expr_key_combo.currentData() or "").strip()
        if selected_key:
            args += ["--groupby", selected_key]

        self._run_visual_process(args, output_path, self.gene_expr_label)

    def _run_visual_process(
        self,
        args: List[str],
        output_path: Path,
        target_label: QtWidgets.QLabel,
    ) -> None:
        self._log(f"Generating plot: {output_path.name}")
        self._enter_busy(f"Generating {output_path.stem}")
        proc = QtCore.QProcess(self)
        self._plot_processes.append(proc)
        proc.setProgram(args[0])
        proc.setArguments(args[1:])
        proc.setWorkingDirectory(str(ROOT_DIR))
        proc.setProcessChannelMode(QtCore.QProcess.MergedChannels)

        def _on_plot_output() -> None:
            text = proc.readAllStandardOutput().data().decode("utf-8", errors="ignore")
            for line in text.splitlines():
                if line.strip():
                    self._log(line)

        def _on_finished(exit_code: int, _status: QtCore.QProcess.ExitStatus) -> None:
            if proc in self._plot_processes:
                self._plot_processes.remove(proc)
            if exit_code != 0:
                self._log(f"Plot generation failed: {output_path.name}")
                self._leave_busy(failed=True)
                return
            if output_path.exists():
                pixmap = QtGui.QPixmap(str(output_path))
                target_label.setPixmap(pixmap.scaledToWidth(900, QtCore.Qt.SmoothTransformation))
            self._log(f"Plot ready: {output_path.name}")
            self._leave_busy(failed=False)

        proc.readyReadStandardOutput.connect(_on_plot_output)
        proc.finished.connect(_on_finished)
        proc.start()

    def _populate_recent(self) -> None:
        self.recent_list.clear()
        for project in sorted(self.recent_projects, key=lambda p: p.last_used, reverse=True):
            item = QtWidgets.QListWidgetItem(project.label())
            item.setData(QtCore.Qt.UserRole, project)
            self.recent_list.addItem(item)

    def _on_recent_selected(self) -> None:
        items = self.recent_list.selectedItems()
        if not items:
            return
        project: RecentProject = items[0].data(QtCore.Qt.UserRole)
        self.data_dir_edit.setText(project.data_dir)
        self.out_dir_edit.setText(project.out_dir)
        if project.karospace_html:
            self.karospace_path_edit.setText(project.karospace_html)

    def _load_selected_recent(self) -> None:
        items = self.recent_list.selectedItems()
        if not items:
            return
        project: RecentProject = items[0].data(QtCore.Qt.UserRole)
        self.data_dir_edit.setText(project.data_dir)
        self.out_dir_edit.setText(project.out_dir)
        if project.karospace_html:
            self.karospace_path_edit.setText(project.karospace_html)
        self._load_outputs(Path(project.out_dir))

    def _update_recent(self) -> None:
        data_dir = self.data_dir_edit.text().strip()
        out_dir = self.out_dir_edit.text().strip()
        if not data_dir or not out_dir:
            return
        karospace_html = self.karospace_path_edit.text().strip() or None
        now = datetime.now().isoformat(timespec="seconds")

        filtered = [
            p for p in self.recent_projects
            if not (p.data_dir == data_dir and p.out_dir == out_dir)
        ]
        filtered.append(RecentProject(data_dir, out_dir, karospace_html, now))
        self.recent_projects = filtered
        _save_recent(self.recent_projects)
        self._populate_recent()


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("InSituCore")
    app.setApplicationDisplayName("InSituCore")
    app.setOrganizationName("InSituCore")
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QtGui.QIcon(str(APP_ICON_PATH)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
