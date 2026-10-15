import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from my_qt_app.consolidator import Consolidator
from my_qt_app.themes import THEMES


class DropZone(QLabel):
    def __init__(self, on_folder_selected):
        super().__init__()
        self.on_folder_selected = on_folder_selected
        self.setObjectName("dropZone")
        self.setAlignment(Qt.AlignCenter)
        self.setText("Drag and drop daily report folder here\nor click 'Browse Folder'")
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(Path(u.toLocalFile()).is_dir() for u in urls):
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
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.is_dir():
                self.on_folder_selected(path)
                event.acceptProposedAction()
                return

    def reset_style(self):
        self.setObjectName("dropZone")
        self.style().unpolish(self)
        self.style().polish(self)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DDR Report Consolidator")
        self.resize(680, 520)

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        # Header with Theme Selector
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

        # Drop Zone
        self.drop_zone = DropZone(self.process_folder)
        layout.addWidget(self.drop_zone)

        # Action Buttons
        btn_layout = QHBoxLayout()
        self.browse_btn = QPushButton("Browse Folder...")
        self.browse_btn.clicked.connect(self.browse_folder)
        btn_layout.addWidget(self.browse_btn)

        self.clear_btn = QPushButton("Clear Logs")
        self.clear_btn.clicked.connect(self.clear_logs)
        btn_layout.addWidget(self.clear_btn)
        layout.addLayout(btn_layout)

        # Activity Log Output
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("System ready. Drag a folder or browse to start consolidation...")
        layout.addWidget(self.log_output)

        # Status Label
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        layout.addWidget(self.status_label)

        # Apply default theme
        self.change_theme("Synthwave")

    def change_theme(self, theme_name: str):
        stylesheet = THEMES.get(theme_name, THEMES["Synthwave"])
        self.setStyleSheet(stylesheet)
        self.drop_zone.reset_style()

    def clear_logs(self):
        self.log_output.clear()
        self.status_label.setText("Ready")

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Report Directory")
        if folder:
            self.process_folder(Path(folder))

    def process_folder(self, folder: Path):
        self.log_output.clear()
        self.status_label.setText(f"Processing: {folder.name}...")
        self.log_output.append(f"Processing directory: {folder.resolve()}")
        consolidator = Consolidator(folder)
        success, logs = consolidator.run()
        for line in logs:
            self.log_output.append(line)
        self.status_label.setText("Finished successfully!" if success else "Consolidation halted with errors.")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
