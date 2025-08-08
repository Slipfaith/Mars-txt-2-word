# styles.py
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QPalette, QColor

QSS = """
/* base */
* {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Ubuntu, "Noto Sans", "Helvetica Neue", Arial, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol";
    font-size: 12.5pt;
}

/* main window bg */
QMainWindow, QWidget {
    background: #0f1115;
    color: #e6e6e6;
}

/* Tabs */
QTabWidget::pane {
    border: 1px solid #1f2430;
    border-radius: 12px;
    padding: 8px;
    background: #12151b;
}
QTabBar::tab {
    background: transparent;
    color: #aab3c0;
    padding: 8px 14px;
    margin: 2px;
    border-radius: 10px;
}
QTabBar::tab:selected {
    color: #ffffff;
    background: #1a1f29;
}
QTabBar::tab:hover {
    color: #ffffff;
    background: #161a22;
}

/* Buttons */
QPushButton {
    background: #1e2532;
    color: #e6e6e6;
    border: 1px solid #232a39;
    border-radius: 10px;
    padding: 8px 14px;
}
QPushButton:hover {
    background: #222a39;
}
QPushButton:pressed {
    background: #1a2130;
}
QPushButton:disabled {
    background: #1a1f29;
    color: #7f8793;
    border-color: #1f2430;
}

/* Line edit */
QLineEdit {
    background: #0f131a;
    border: 1px solid #1f2430;
    border-radius: 10px;
    padding: 8px 10px;
    color: #e6e6e6;
    selection-background-color: #2e7ef5;
}
QLineEdit:focus {
    border: 1px solid #2e7ef5;
    background: #0c1016;
}

/* Labels */
QLabel {
    color: #b8c1ce;
}

/* Progress dialog */
QProgressDialog, QDialog {
    background: #10141b;
}
QProgressBar {
    background: #0f131a;
    border: 1px solid #1f2430;
    border-radius: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #2e7ef5;
    border-radius: 8px;
}

/* File dialogs / menus */
QMenu {
    background: #0f131a;
    color: #e6e6e6;
    border: 1px solid #1f2430;
}
QMenu::item:selected {
    background: #1a1f29;
}
QToolTip {
    background: #1a1f29;
    color: #e6e6e6;
    border: 1px solid #232a39;
    padding: 6px 8px;
    border-radius: 8px;
}

/* Scrollbars */
QScrollBar:vertical {
    background: transparent;
    width: 12px;
    margin: 4px;
}
QScrollBar::handle:vertical {
    background: #2a2f3d;
    min-height: 24px;
    border-radius: 6px;
}
QScrollBar::handle:vertical:hover {
    background: #343b4d;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0; background: transparent;
}
QScrollBar:horizontal {
    background: transparent;
    height: 12px; margin: 4px;
}
QScrollBar::handle:horizontal {
    background: #2a2f3d;
    min-width: 24px;
    border-radius: 6px;
}
QScrollBar::handle:horizontal:hover {
    background: #343b4d;
}
"""

def apply_modern_style(app: QApplication) -> None:
    # Базовая палитра (dark)
    pal = app.palette()
    pal.setColor(QPalette.Window, QColor("#0f1115"))
    pal.setColor(QPalette.WindowText, QColor("#e6e6e6"))
    pal.setColor(QPalette.Base, QColor("#0f131a"))
    pal.setColor(QPalette.AlternateBase, QColor("#12151b"))
    pal.setColor(QPalette.Text, QColor("#e6e6e6"))
    pal.setColor(QPalette.Button, QColor("#1e2532"))
    pal.setColor(QPalette.ButtonText, QColor("#e6e6e6"))
    pal.setColor(QPalette.Highlight, QColor("#2e7ef5"))
    pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(pal)

    app.setStyleSheet(QSS)
