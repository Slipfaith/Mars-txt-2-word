from PySide6.QtWidgets import QLineEdit, QFileDialog
from PySide6.QtCore import Qt, Signal
import os

class DragDropField(QLineEdit):
    """Read-only line edit supporting drag and drop and double click selection."""

    pathChanged = Signal(str)

    def __init__(self, mode: str = "file", parent=None):
        super().__init__(parent)
        self.mode = mode  # "file" or "folder"
        self.setReadOnly(True)
        self.setAcceptDrops(True)
        placeholder = "Перетащите {} или двойной клик".format(
            "файл" if self.mode == "file" else "папку"
        )
        self.setPlaceholderText(placeholder)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if self._valid_path(path):
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if self._valid_path(path):
                self.setText(path)
                self.pathChanged.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()

    def mouseDoubleClickEvent(self, event):
        if self.mode == "folder":
            folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
            if folder:
                self.setText(folder)
                self.pathChanged.emit(folder)
        else:
            file, _ = QFileDialog.getOpenFileName(self, "Выберите файл")
            if file:
                self.setText(file)
                self.pathChanged.emit(file)
        super().mouseDoubleClickEvent(event)

    def _valid_path(self, path: str) -> bool:
        return (
            os.path.isdir(path)
            if self.mode == "folder"
            else os.path.isfile(path)
        )
