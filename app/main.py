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
RECENT_PATH = Path.home() / ".spatial-analysis-for-dummies" / "recent.json"
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
    if not RECENT_PATH.exists():
        return []
    try:
        payload = json.loads(RECENT_PATH.read_text())
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
        self.setWindowTitle("InSituViewer")
        self.resize(1200, 800)

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
        self._theme_mode = "light"

        self._build_ui()
        self._busy_timer = QtCore.QTimer(self)
        self._busy_timer.setInterval(280)
        self._busy_timer.timeout.connect(self._animate_busy_state)
        self._apply_theme(self._theme_mode)
        self._populate_recent()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        root.setObjectName("Root")
        root_layout = QtWidgets.QVBoxLayout(root)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(12)

        root_layout.addWidget(self._build_top_bar())

        layout = QtWidgets.QHBoxLayout()
        layout.setSpacing(12)
        root_layout.addLayout(layout, stretch=1)

        self.recent_list = QtWidgets.QListWidget()
        self.recent_list.setObjectName("RecentList")
        self.recent_list.setMaximumWidth(320)
        self.recent_list.itemSelectionChanged.connect(self._on_recent_selected)

        recent_box, recent_layout = self._create_card(
            "Recent Projects",
            "Load existing runs without recomputing.",
        )
        recent_box.setMaximumWidth(340)
        recent_layout.addWidget(self.recent_list)

        self.load_recent_btn = QtWidgets.QPushButton("Load Outputs")
        self.load_recent_btn.clicked.connect(self._load_selected_recent)
        recent_layout.addWidget(self.load_recent_btn)

        layout.addWidget(recent_box)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setObjectName("MainTabs")
        self.tabs.addTab(self._build_run_tab(), "Run")
        self.tabs.addTab(self._build_qc_tab(), "QC")
        self.tabs.addTab(self._build_spatial_tab(), "Spatial")
        self.tabs.addTab(self._build_umap_tab(), "UMAP")
        self.tabs.addTab(self._build_compartment_tab(), "Compartments")
        tabs_card, tabs_layout = self._create_card("Workspace")
        tabs_layout.addWidget(self.tabs, stretch=1)
        layout.addWidget(tabs_card, stretch=1)

        self.setCentralWidget(root)

    def _build_top_bar(self) -> QtWidgets.QWidget:
        top_bar = QtWidgets.QFrame()
        top_bar.setObjectName("TopBar")
        layout = QtWidgets.QHBoxLayout(top_bar)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(10)

        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(1)
        title = QtWidgets.QLabel("InSituViewer")
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

        self.top_umap_btn = QtWidgets.QPushButton("Generate UMAP")
        self.top_umap_btn.clicked.connect(self._generate_umap_plot)
        layout.addWidget(self.top_umap_btn)

        self.top_comp_btn = QtWidgets.QPushButton("Generate Compartments")
        self.top_comp_btn.clicked.connect(self._generate_compartment_map)
        layout.addWidget(self.top_comp_btn)

        self.theme_toggle_btn = QtWidgets.QPushButton("Dark")
        self.theme_toggle_btn.setCheckable(True)
        self.theme_toggle_btn.toggled.connect(self._toggle_theme)
        layout.addWidget(self.theme_toggle_btn)

        self.activity_stage = QtWidgets.QLabel("Idle")
        self.activity_stage.setObjectName("ActivityStage")
        layout.addWidget(self.activity_stage)

        self.runner_glyph = QtWidgets.QLabel("o-/")
        self.runner_glyph.setObjectName("RunnerGlyph")
        self.runner_glyph.setText("   ")
        layout.addWidget(self.runner_glyph)

        self.status_chip = QtWidgets.QLabel("Ready")
        self.status_chip.setObjectName("StatusChip")
        layout.addWidget(self.status_chip)
        return top_bar

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
        data_layout.addLayout(form)
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
        self.mana_hop_decay.setValue(0.5)
        self.mana_kernel = QtWidgets.QComboBox()
        self.mana_kernel.addItems(["exponential", "inverse", "gaussian", "none"])

        mana_form = QtWidgets.QFormLayout()
        mana_form.addRow(self.mana_check)
        mana_form.addRow("Layers", self.mana_layers)
        mana_form.addRow("Hop decay", self.mana_hop_decay)
        mana_form.addRow("Distance kernel", self.mana_kernel)
        options_layout.addLayout(mana_form)

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

    def _build_spatial_tab(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)

        card, card_layout = self._create_card(
            "Spatial Map",
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
        self.umap_label = QtWidgets.QLabel("No UMAP image loaded.")
        self.umap_label.setObjectName("PreviewSurface")
        self.umap_label.setAlignment(QtCore.Qt.AlignCenter)
        self.umap_label.setMinimumHeight(400)
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
        self.compartment_label = QtWidgets.QLabel("No compartment map loaded.")
        self.compartment_label.setObjectName("PreviewSurface")
        self.compartment_label.setAlignment(QtCore.Qt.AlignCenter)
        self.compartment_label.setMinimumHeight(400)
        card_layout.addWidget(self.compartment_label, stretch=1)
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

    def _toggle_theme(self, checked: bool) -> None:
        self._theme_mode = "dark" if checked else "light"
        self._apply_theme(self._theme_mode)

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
        ]

        if self.mana_check.isChecked():
            args += [
                "--mana-aggregate",
                "--mana-n-layers",
                str(self.mana_layers.value()),
                "--mana-hop-decay",
                str(self.mana_hop_decay.value()),
                "--mana-distance-kernel",
                self.mana_kernel.currentText(),
            ]

        if self.karospace_check.isChecked():
            karospace_path = self.karospace_path_edit.text().strip()
            if not karospace_path:
                karospace_path = str(Path(out_dir) / "karospace.html")
                self.karospace_path_edit.setText(karospace_path)
            args += ["--karospace-html", karospace_path]

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
        self._load_qc_images(out_dir)
        self._load_karospace(out_dir)
        self._load_umap_image(out_dir)
        self._load_compartment_image(out_dir)

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

        self._run_visual_process(args, output_path, self.compartment_label)

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
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
