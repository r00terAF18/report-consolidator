import sys
from pathlib import Path

from PySide6.QtCore import Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from my_qt_app.consolidator import Consolidator
from my_qt_app.themes import THEMES


class NullWriter:
    def write(self, text): pass
    def flush(self): pass

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()


class DropZone(QLabel):
    def __init__(self, on_folder_selected):
        super().__init__()
        self.on_folder_selected = on_folder_selected
        self.is_busy = False
        self.setObjectName("dropZone")
        self.setAlignment(Qt.AlignCenter)
        self.setText("Drop a daily report folder or month folder here\nor click 'Browse Folder'")
        self.setAcceptDrops(True)

    def set_busy(self, busy: bool):
        self.is_busy = busy
        if busy:
            self.setObjectName("dropZoneDisabled")
            self.setText("Processing in background...\nPlease wait.")
        else:
            self.setObjectName("dropZone")
            self.setText("Drop a daily report folder or month folder here\nor click 'Browse Folder'")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if self.is_busy:
            event.ignore()
            return

        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for u in urls:
                local_file = u.toLocalFile()
                if local_file and Path(local_file).is_dir():
                    self.setObjectName("dropZoneActive")
                    self.style().unpolish(self)
                    self.style().polish(self)
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.reset_style()

    def dropEvent(self, event):
        self.reset_style()
        if self.is_busy:
            event.ignore()
            return

        target_folder: Path | None = None
        for url in event.mimeData().urls():
            local_file = url.toLocalFile()
            if local_file:
                p = Path(local_file)
                if p.is_dir():
                    target_folder = p
                    break

        if target_folder:
            event.acceptProposedAction()
            QTimer.singleShot(0, lambda: self.on_folder_selected(target_folder))
        else:
            event.ignore()

    def reset_style(self):
        if self.is_busy:
            self.setObjectName("dropZoneDisabled")
        else:
            self.setObjectName("dropZone")
        self.style().unpolish(self)
        self.style().polish(self)


class ConsolidationWorker(QThread):
    log_signal = Signal(str)
    status_signal = Signal(str)
    finished_signal = Signal(bool)
    error_signal = Signal(str)
    output_signal = Signal(str)

    def __init__(self, folder_path: Path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        try:
            consolidator = Consolidator(
                folder_path=self.folder_path,
                log_callback=self.log_signal.emit,
                status_callback=self.status_signal.emit,
            )
            success, _ = consolidator.run()
            for output in consolidator.outputs:
                self.output_signal.emit(str(output.resolve()))
            self.finished_signal.emit(success)
        except Exception as e:
            self.error_signal.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DDR Report Consolidator")
        self.resize(700, 540)

        self.worker: ConsolidationWorker | None = None
        self.close_when_finished = False
        self.last_directory: str = str(Path.home())
        self.output_path: str | None = None

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        header_layout = QHBoxLayout()
        title = QLabel("DDR REPORT CONSOLIDATOR")
        title.setObjectName("titleLabel")
        header_layout.addWidget(title)
        header_layout.addStretch()

        theme_label = QLabel("Theme:")
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(list(THEMES.keys()))
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        header_layout.addWidget(theme_label)
        header_layout.addWidget(self.theme_combo)
        layout.addLayout(header_layout)

        self.drop_zone = DropZone(self.process_folder)
        layout.addWidget(self.drop_zone)

        instructions = QLabel(
            "Select a folder containing the Tool Pusher DDR and source reports.\n"
            "A month folder processes each daily subfolder. Originals are preserved;\n"
            "new files end in _CONSOLIDATED.xlsx; full source details are saved separately.\n"
            "Review warnings before issuing."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        btn_layout = QHBoxLayout()
        self.browse_btn = QPushButton("Browse Folder...")
        self.browse_btn.clicked.connect(self.browse_folder)
        btn_layout.addWidget(self.browse_btn)

        self.clear_btn = QPushButton("Clear Logs")
        self.clear_btn.clicked.connect(self.clear_logs)
        btn_layout.addWidget(self.clear_btn)
        self.open_btn = QPushButton('Open Saved DDR')
        self.open_btn.setEnabled(False)
        self.open_btn.clicked.connect(self.open_output)
        btn_layout.addWidget(self.open_btn)
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setAcceptRichText(False)
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("System ready. Drag a folder or browse to start consolidation...")
        layout.addWidget(self.log_output)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        layout.addWidget(self.status_label)

        self.change_theme("Synthwave")

    def change_theme(self, theme_name: str):
        stylesheet = THEMES.get(theme_name, THEMES["Synthwave"])
        self.setStyleSheet(stylesheet)
        self.drop_zone.reset_style()

    def clear_logs(self):
        if self.worker and self.worker.isRunning():
            return
        self.log_output.clear()
        self.status_label.setText("Ready")

    def set_ui_busy(self, busy: bool):
        self.browse_btn.setEnabled(not busy)
        self.clear_btn.setEnabled(not busy)
        self.theme_combo.setEnabled(not busy)
        self.drop_zone.set_busy(busy)
        self.progress_bar.setVisible(busy)

    def browse_folder(self):
        if self.worker and self.worker.isRunning():
            return

        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Report Directory",
            self.last_directory,
            QFileDialog.ShowDirsOnly,
        )
        if folder:
            self.last_directory = folder
            target = Path(folder)
            QTimer.singleShot(0, lambda: self.process_folder(target))

    def process_folder(self, folder: Path):
        if self.worker and self.worker.isRunning():
            return

        self.set_ui_busy(True)
        self.output_path = None
        self.open_btn.setEnabled(False)
        self.log_output.clear()
        self.status_label.setText(f"Processing: {folder.name}...")
        self.log_output.append(f"Starting consolidation for: {folder.resolve()}")

        self.worker = ConsolidationWorker(folder)
        self.worker.log_signal.connect(self.log_output.append)
        self.worker.status_signal.connect(self.status_label.setText)
        self.worker.finished_signal.connect(self.on_consolidation_finished)
        self.worker.error_signal.connect(self.on_consolidation_error)
        self.worker.output_signal.connect(self.on_output_saved)
        self.worker.finished.connect(self.on_worker_stopped)
        self.worker.start()

    def on_output_saved(self, path: str):
        self.output_path = path
        self.open_btn.setToolTip(path)
        self.open_btn.setEnabled(True)
        self.log_output.append(f'Open Saved DDR: {path}')

    def open_output(self):
        if self.output_path:
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_path)):
                self.log_output.append(f'Could not open the saved DDR: {self.output_path}')

    def on_consolidation_finished(self, success: bool):
        self.status_label.setText("Saved — review warnings in the log." if success else "Some reports could not be saved — see log.")

    def on_consolidation_error(self, error_message: str):
        self.log_output.append(f"Fatal error during processing: {error_message}")
        self.status_label.setText("Consolidation halted with errors.")

    def on_worker_stopped(self):
        self.worker.deleteLater()
        self.worker = None
        self.set_ui_busy(False)
        if self.close_when_finished:
            self.close()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.close_when_finished = True
            self.status_label.setText("Finishing report writes before closing...")
            event.ignore()
            return
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
