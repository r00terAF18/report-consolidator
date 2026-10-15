"""Comprehensive QSS theme definitions for the application."""

SYNTHWAVE = """
QMainWindow, QWidget#centralWidget {
    background-color: #1a102f;
    color: #f0e6ff;
    font-family: 'Segoe UI', 'Ubuntu', sans-serif;
}
QMenuBar {
    background-color: #120b22;
    color: #ff71ce;
    border-bottom: 1px solid #01cdfe;
}
QMenuBar::item:selected {
    background-color: #2b194a;
    color: #05ffa1;
}
QLabel {
    color: #f0e6ff;
    font-size: 13px;
}
QLabel#titleLabel {
    color: #ff71ce;
    font-size: 18px;
    font-weight: bold;
    letter-spacing: 1px;
}
QLabel#statusLabel {
    color: #05ffa1;
    font-size: 12px;
}
QComboBox {
    background-color: #2b194a;
    border: 1px solid #ff71ce;
    border-radius: 6px;
    padding: 6px 12px;
    color: #01cdfe;
    font-weight: bold;
}
QComboBox:hover {
    border: 1px solid #05ffa1;
}
QComboBox:disabled {
    background-color: #1f1238;
    color: #6b5585;
    border: 1px solid #4a3363;
}
QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox QAbstractItemView {
    background-color: #1f1238;
    border: 1px solid #ff71ce;
    selection-background-color: #ff71ce;
    selection-color: #120b22;
    color: #01cdfe;
    padding: 4px;
}
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff71ce, stop:1 #b967ff);\n    color: #120b22;
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-weight: bold;
    font-size: 13px;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #01cdfe, stop:1 #05ffa1);\n    color: #0d0221;
}
QPushButton:pressed {
    background-color: #b967ff;
}
QPushButton:disabled {
    background: #2b194a;
    color: #6b5585;
}
QProgressBar {
    border: 1px solid #ff71ce;
    border-radius: 3px;
    background-color: #120b22;
    max-height: 6px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff71ce, stop:1 #01cdfe);
    border-radius: 2px;
}
QTextEdit {
    background-color: #120b22;
    border: 1px solid #2b194a;
    border-radius: 8px;
    padding: 10px;
    color: #01cdfe;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}
QScrollBar:vertical {
    border: none;
    background: #120b22;
    width: 10px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #ff71ce;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background: #05ffa1;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QLabel#dropZone {
    border: 2px dashed #ff71ce;
    border-radius: 12px;
    background-color: rgba(43, 25, 74, 0.45);
    color: #f0e6ff;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneActive {
    border: 2px solid #05ffa1;
    border-radius: 12px;
    background-color: rgba(5, 255, 161, 0.15);
    color: #05ffa1;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneDisabled {
    border: 2px dashed #4a3363;
    border-radius: 12px;
    background-color: rgba(30, 18, 50, 0.35);
    color: #6b5585;
    font-size: 14px;
    padding: 24px;
}
"""

LIQUID_GLASS = """
QMainWindow, QWidget#centralWidget {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1a2233, stop:0.5 #202b40, stop:1 #141b29);
    color: #e4edf8;
    font-family: 'SF Pro Display', 'Segoe UI', 'Ubuntu', sans-serif;
}
QMenuBar {
    background-color: rgba(255, 255, 255, 0.05);
    border-bottom: 1px solid rgba(255, 255, 255, 0.12);
    color: #d1e2f7;
}
QMenuBar::item:selected {
    background-color: rgba(255, 255, 255, 0.15);
    border-radius: 4px;
}
QLabel {
    color: #d8e5f5;
    font-size: 13px;
}
QLabel#titleLabel {
    color: #ffffff;
    font-size: 18px;
    font-weight: 600;
}
QLabel#statusLabel {
    color: #72b2ff;
    font-size: 12px;
}
QComboBox {
    background-color: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.2);
    border-radius: 8px;
    padding: 6px 12px;
    color: #e6f0fa;
}
QComboBox:hover {
    background-color: rgba(255, 255, 255, 0.14);
    border: 1px solid rgba(130, 185, 255, 0.5);
}
QComboBox:disabled {
    background-color: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.08);
    color: rgba(255, 255, 255, 0.3);
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #1a2334;
    border: 1px solid rgba(255, 255, 255, 0.2);
    selection-background-color: rgba(114, 178, 255, 0.35);
    selection-color: #ffffff;
    color: #e6f0fa;
    padding: 6px;
    border-radius: 6px;
}
QPushButton {
    background-color: rgba(255, 255, 255, 0.12);
    border: 1px solid rgba(255, 255, 255, 0.25);
    border-radius: 8px;
    color: #ffffff;
    padding: 8px 18px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton:hover {
    background-color: rgba(255, 255, 255, 0.22);
    border: 1px solid rgba(255, 255, 255, 0.45);
}
QPushButton:pressed {
    background-color: rgba(255, 255, 255, 0.08);
}
QPushButton:disabled {
    background-color: rgba(255, 255, 255, 0.04);
    color: rgba(255, 255, 255, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.08);
}
QProgressBar {
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 3px;
    background-color: rgba(15, 22, 34, 0.5);
    max-height: 6px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4388e6, stop:1 #72b2ff);
    border-radius: 2px;
}
QTextEdit {
    background-color: rgba(15, 22, 34, 0.65);
    border: 1px solid rgba(255, 255, 255, 0.14);
    border-radius: 10px;
    padding: 10px;
    color: #b8d2ee;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}
QScrollBar:vertical {
    border: none;
    background: transparent;
    width: 8px;
    margin: 4px;
}
QScrollBar::handle:vertical {
    background: rgba(255, 255, 255, 0.25);
    min-height: 24px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: rgba(255, 255, 255, 0.45);
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QLabel#dropZone {
    border: 2px dashed rgba(255, 255, 255, 0.3);
    border-radius: 12px;
    background-color: rgba(255, 255, 255, 0.05);
    color: #d1e2f7;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneActive {
    border: 2px solid #72b2ff;
    border-radius: 12px;
    background-color: rgba(114, 178, 255, 0.15);
    color: #ffffff;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneDisabled {
    border: 2px dashed rgba(255, 255, 255, 0.1);
    border-radius: 12px;
    background-color: rgba(255, 255, 255, 0.02);
    color: rgba(209, 226, 247, 0.4);
    font-size: 14px;
    padding: 24px;
}
"""

MATERIAL = """
QMainWindow, QWidget#centralWidget {
    background-color: #121212;
    color: #e0e0e0;
    font-family: 'Roboto', 'Segoe UI', sans-serif;
}
QMenuBar {
    background-color: #1f1b24;
    color: #ffffff;
    border-bottom: 1px solid #2d2d2d;
}
QMenuBar::item:selected {
    background-color: #332940;
    color: #bb86fc;
}
QLabel {
    color: #e0e0e0;
    font-size: 13px;
}
QLabel#titleLabel {
    color: #bb86fc;
    font-size: 18px;
    font-weight: 500;
    letter-spacing: 0.5px;
}
QLabel#statusLabel {
    color: #03dac6;
    font-size: 12px;
    font-weight: 500;
}
QComboBox {
    background-color: #1e1e1e;
    border: 1px solid #383838;
    border-radius: 4px;
    padding: 6px 12px;
    color: #ffffff;
}
QComboBox:hover {
    border: 1px solid #bb86fc;
}
QComboBox:disabled {
    background-color: #181818;
    border: 1px solid #282828;
    color: #555555;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #1e1e1e;
    border: 1px solid #383838;
    selection-background-color: #3700b3;
    selection-color: #ffffff;
    color: #e0e0e0;
    padding: 4px;
}
QPushButton {
    background-color: #bb86fc;
    color: #000000;
    border: none;
    border-radius: 4px;
    padding: 8px 18px;
    font-weight: 500;
    font-size: 13px;
    letter-spacing: 0.5px;
}
QPushButton:hover {
    background-color: #d7b7fd;
}
QPushButton:pressed {
    background-color: #9955e8;
}
QPushButton:disabled {
    background-color: #2c2533;
    color: #665c70;
}
QProgressBar {
    border: 1px solid #383838;
    border-radius: 3px;
    background-color: #1e1e1e;
    max-height: 6px;
}
QProgressBar::chunk {
    background-color: #03dac6;
    border-radius: 2px;
}
QTextEdit {
    background-color: #1e1e1e;
    border: 1px solid #2c2c2c;
    border-radius: 4px;
    padding: 10px;
    color: #03dac6;
    font-family: 'Roboto Mono', 'Consolas', monospace;
    font-size: 12px;
}
QScrollBar:vertical {
    border: none;
    background: #121212;
    width: 8px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #383838;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #bb86fc;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QLabel#dropZone {
    border: 2px dashed #383838;
    border-radius: 8px;
    background-color: #1e1e1e;
    color: #a0a0a0;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneActive {
    border: 2px solid #03dac6;
    border-radius: 8px;
    background-color: rgba(3, 218, 198, 0.08);
    color: #03dac6;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneDisabled {
    border: 2px dashed #282828;
    border-radius: 8px;
    background-color: #161616;
    color: #555555;
    font-size: 14px;
    padding: 24px;
}
"""

FLUENT = """
QMainWindow, QWidget#centralWidget {
    background-color: #202020;
    color: #ffffff;
    font-family: 'Segoe UI Variable', 'Segoe UI', sans-serif;
}
QMenuBar {
    background-color: #2c2c2c;
    color: #ffffff;
    border-bottom: 1px solid #383838;
}
QMenuBar::item:selected {
    background-color: #383838;
    border-radius: 4px;
}
QLabel {
    color: #cccccc;
    font-size: 13px;
}
QLabel#titleLabel {
    color: #ffffff;
    font-size: 18px;
    font-weight: 600;
}
QLabel#statusLabel {
    color: #60cdff;
    font-size: 12px;
}
QComboBox {
    background-color: #2d2d2d;
    border: 1px solid #3c3c3c;
    border-radius: 5px;
    padding: 5px 12px;
    color: #ffffff;
}
QComboBox:hover {
    background-color: #323232;
    border: 1px solid #4a4a4a;
}
QComboBox:disabled {
    background-color: #1a1a1a;
    border: 1px solid #282828;
    color: #555555;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #2c2c2c;
    border: 1px solid #383838;
    selection-background-color: #383838;
    selection-color: #60cdff;
    color: #ffffff;
    padding: 4px;
    border-radius: 5px;
}
QPushButton {
    background-color: #2d2d2d;
    border: 1px solid #383838;
    border-bottom: 1px solid #242424;
    border-radius: 5px;
    color: #ffffff;
    padding: 7px 18px;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #383838;
    border: 1px solid #484848;
}
QPushButton:pressed {
    background-color: #252525;
    color: #8a8a8a;
}
QPushButton:disabled {
    background-color: #1c1c1c;
    color: #555555;
    border: 1px solid #282828;
}
QProgressBar {
    border: 1px solid #383838;
    border-radius: 3px;
    background-color: #1e1e1e;
    max-height: 6px;
}
QProgressBar::chunk {
    background-color: #60cdff;
    border-radius: 2px;
}
QTextEdit {
    background-color: #1e1e1e;
    border: 1px solid #333333;
    border-radius: 6px;
    padding: 10px;
    color: #e0e0e0;
    font-family: 'Consolas', monospace;
    font-size: 12px;
}
QScrollBar:vertical {
    border: none;
    background: transparent;
    width: 6px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #484848;
    min-height: 24px;
    border-radius: 3px;
}
QScrollBar::handle:vertical:hover {
    background: #60cdff;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QLabel#dropZone {
    border: 2px dashed #404040;
    border-radius: 8px;
    background-color: #282828;
    color: #999999;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneActive {
    border: 2px solid #60cdff;
    border-radius: 8px;
    background-color: #2b333a;
    color: #60cdff;
    font-size: 14px;
    padding: 24px;
}
QLabel#dropZoneDisabled {
    border: 2px dashed #303030;
    border-radius: 8px;
    background-color: #1c1c1c;
    color: #505050;
    font-size: 14px;
    padding: 24px;
}
"""

THEMES = {
    "Synthwave": SYNTHWAVE,
    "Liquid Glass": LIQUID_GLASS,
    "Material": MATERIAL,
    "Fluent": FLUENT,
}
