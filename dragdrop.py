from PySide6.QtWidgets import QLineEdit, QFileDialog
from PySide6.QtCore import Signal
import os


class DragDropField(QLineEdit):
    """Read-only line edit supporting drag&drop and double click selection.

    mode: "file", "folder" или "files_or_folder"
    """

    pathChanged = Signal(str)
    pathsChanged = Signal(list)

    def __init__(self, mode: str = "file", parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setAcceptDrops(True)
        self.mode = mode
        self._paths = []

    @property
    def paths(self):
        return self._paths

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if self._valid_path(url.toLocalFile()):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            paths = [u.toLocalFile() for u in urls if self._valid_path(u.toLocalFile())]
            if not paths:
                event.ignore()
                return
            self._paths = paths
            display = "; ".join(paths)
            self.setText(display)
            if len(paths) == 1:
                self.pathChanged.emit(paths[0])
            self.pathsChanged.emit(paths)
            event.acceptProposedAction()
            return
        event.ignore()

    def mouseDoubleClickEvent(self, event):
        if self.mode == "folder":
            folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
            if folder:
                self._paths = [folder]
                self.setText(folder)
                self.pathChanged.emit(folder)
                self.pathsChanged.emit([folder])
        elif self.mode == "files_or_folder":
            files, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы")
            if files:
                self._paths = files
                self.setText("; ".join(files))
                if len(files) == 1:
                    self.pathChanged.emit(files[0])
                self.pathsChanged.emit(files)
            else:
                folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
                if folder:
                    self._paths = [folder]
                    self.setText(folder)
                    self.pathChanged.emit(folder)
                    self.pathsChanged.emit([folder])
        else:  # file
            file, _ = QFileDialog.getOpenFileName(self, "Выберите файл")
            if file:
                self._paths = [file]
                self.setText(file)
                self.pathChanged.emit(file)
                self.pathsChanged.emit([file])
        super().mouseDoubleClickEvent(event)

    def _valid_path(self, path: str) -> bool:
        if self.mode == "folder":
            return os.path.isdir(path)
        if self.mode == "file":
            return os.path.isfile(path)
        if self.mode == "files_or_folder":
            return os.path.isdir(path) or os.path.isfile(path)
        return False
