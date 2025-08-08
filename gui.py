# gui.py
import sys
import os
import logging

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QProgressDialog,
    QStyle,
    QFrame,
)

from dragdrop import DragDropField
from converter import (
    export_to_word,
    export_paths_to_word,
    import_from_word,
)
from styles import apply_modern_style


class ExportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.last_result_path = None
        self._temp_log_handler = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        card = QFrame()
        card.setObjectName("card")
        card.setFrameShape(QFrame.NoFrame)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(10)

        # ENG
        row_eng = QHBoxLayout()
        row_eng.setSpacing(10)
        self.eng_field = DragDropField(mode="files_or_folder")
        row_eng.addWidget(QLabel("ENG:"))
        row_eng.addWidget(self.eng_field)
        card_layout.addLayout(row_eng)

        # RUS
        row_rus = QHBoxLayout()
        row_rus.setSpacing(10)
        self.rus_field = DragDropField(mode="files_or_folder")
        row_rus.addWidget(QLabel("RUS:"))
        row_rus.addWidget(self.rus_field)
        card_layout.addLayout(row_rus)

        # Кодировка (RU force)
        row_enc = QHBoxLayout()
        row_enc.setSpacing(10)
        self.ru_enc_edit = QLineEdit()
        self.ru_enc_edit.setPlaceholderText("Кодировка RU (опционально, напр. cp1251)")
        row_enc.addWidget(QLabel("RU encoding:"))
        row_enc.addWidget(self.ru_enc_edit)
        card_layout.addLayout(row_enc)

        # Кнопки
        row_btns = QHBoxLayout()
        row_btns.setSpacing(10)
        self.btn_export = QPushButton("Создать Word")
        self.btn_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_export.clicked.connect(self._on_export)
        row_btns.addWidget(self.btn_export)

        self.btn_open_result = QPushButton("Открыть результат")
        self.btn_open_result.setEnabled(False)
        self.btn_open_result.clicked.connect(self._open_result_folder)
        row_btns.addWidget(self.btn_open_result)

        card_layout.addLayout(row_btns)

        layout.addWidget(card)
        self.setLayout(layout)

    @staticmethod
    def _all_exist(paths: list[str]) -> bool:
        if not paths:
            return False
        return all(os.path.exists(p) for p in paths)

    def _attach_file_logger(self, log_path: str):
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s")
        fh.setFormatter(fmt)
        logging.getLogger().addHandler(fh)
        self._temp_log_handler = fh

    def _detach_file_logger(self):
        if self._temp_log_handler:
            logging.getLogger().removeHandler(self._temp_log_handler)
            self._temp_log_handler.close()
            self._temp_log_handler = None

    def _on_export(self):
        eng_sel = self.eng_field.paths
        rus_sel = self.rus_field.paths
        if not eng_sel or not rus_sel:
            QMessageBox.warning(self, "Ошибка", "Выбери ENG и RUS: файлы или папки.")
            return
        if not self._all_exist(eng_sel) or not self._all_exist(rus_sel):
            QMessageBox.warning(self, "Ошибка", "Некоторые пути не существуют.")
            return

        out_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Word", "", "Word файлы (*.docx)"
        )
        if not out_path:
            return

        log_path = os.path.splitext(out_path)[0] + "_log.txt"
        self._attach_file_logger(log_path)

        try:
            ru_force = self.ru_enc_edit.text().strip() or None

            progress = QProgressDialog("Экспорт...", "Отмена", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoClose(True)
            progress.show()

            def progress_cb(done: int, total: int):
                if total > 0:
                    progress.setValue(int(done * 100 / total))
                    QApplication.processEvents()

            def is_single_folder(lst: list[str]) -> bool:
                return len(lst) == 1 and os.path.isdir(lst[0])

            if is_single_folder(eng_sel) and is_single_folder(rus_sel):
                export_to_word(
                    eng_folder=eng_sel[0],
                    rus_folder=rus_sel[0],
                    output_path=out_path,
                    file_extension=None,
                    rus_force_encoding=ru_force,
                    progress=progress_cb,
                )
            else:
                eng_files = [p for p in eng_sel if os.path.isfile(p)]
                rus_files = [p for p in rus_sel if os.path.isfile(p)]
                if not eng_files or not rus_files:
                    QMessageBox.warning(self, "Ошибка", "Укажи корректно файлы/папки.")
                    return
                export_paths_to_word(
                    eng_files, rus_files, out_path, rus_force_encoding=ru_force, progress=progress_cb
                )

            progress.setValue(100)
            self.last_result_path = out_path
            self.btn_open_result.setEnabled(True)
            QMessageBox.information(self, "Успех", f"Word сохранён:\n{out_path}\n\nЛог: {log_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))
            raise
        finally:
            self._detach_file_logger()

    def _open_result_folder(self):
        if self.last_result_path and os.path.exists(self.last_result_path):
            folder = os.path.dirname(self.last_result_path) or os.getcwd()
            try:
                os.startfile(folder)
            except Exception:
                QMessageBox.information(self, "Открытие папки", folder)


class ImportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._last_eng = None
        self._last_rus = None
        self._temp_log_handler = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        card = QFrame()
        card.setObjectName("card")
        card.setFrameShape(QFrame.NoFrame)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(10)

        # Word
        row_word = QHBoxLayout()
        row_word.setSpacing(10)
        self.word_field = DragDropField(mode="file")
        row_word.addWidget(QLabel("Word:"))
        row_word.addWidget(self.word_field)
        card_layout.addLayout(row_word)

        # ENG out
        row_eng = QHBoxLayout()
        row_eng.setSpacing(10)
        self.eng_out_field = DragDropField(mode="folder")
        row_eng.addWidget(QLabel("ENG out:"))
        row_eng.addWidget(self.eng_out_field)
        card_layout.addLayout(row_eng)

        # RUS out
        row_rus = QHBoxLayout()
        row_rus.setSpacing(10)
        self.rus_out_field = DragDropField(mode="folder")
        row_rus.addWidget(QLabel("RUS out:"))
        row_rus.addWidget(self.rus_out_field)
        card_layout.addLayout(row_rus)

        # buttons
        row_btns = QHBoxLayout()
        row_btns.setSpacing(10)
        self.btn_import = QPushButton("Разбить Word")
        self.btn_import.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.btn_import.clicked.connect(self._on_import)
        row_btns.addWidget(self.btn_import)

        self.btn_open_folders = QPushButton("Открыть папки")
        self.btn_open_folders.setEnabled(False)
        self.btn_open_folders.clicked.connect(self._open_result_folders)
        row_btns.addWidget(self.btn_open_folders)

        card_layout.addLayout(row_btns)

        layout.addWidget(card)
        self.setLayout(layout)

    def _attach_file_logger(self, word_file: str) -> str:
        log_path = os.path.splitext(word_file)[0] + "_import_log.txt"
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s")
        fh.setFormatter(fmt)
        logging.getLogger().addHandler(fh)
        self._temp_log_handler = fh
        return log_path

    def _detach_file_logger(self):
        if self._temp_log_handler:
            logging.getLogger().removeHandler(self._temp_log_handler)
            self._temp_log_handler.close()
            self._temp_log_handler = None

    def _on_import(self):
        word_file = self.word_field.text().strip()
        eng_out = self.eng_out_field.text().strip()
        rus_out = self.rus_out_field.text().strip()

        if not word_file or not eng_out or not rus_out:
            QMessageBox.warning(self, "Ошибка", "Выбери Word и обе папки для вывода.")
            return
        if not os.path.isfile(word_file):
            QMessageBox.warning(self, "Ошибка", "Файл Word не найден.")
            return
        os.makedirs(eng_out, exist_ok=True)
        os.makedirs(rus_out, exist_ok=True)

        log_path = self._attach_file_logger(word_file)

        try:
            progress = QProgressDialog("Импорт...", "Отмена", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoClose(True)
            progress.show()

            def progress_cb(done: int, total: int):
                if total > 0:
                    progress.setValue(int(done * 100 / total))
                    QApplication.processEvents()

            import_from_word(
                word_path=word_file,
                eng_output_folder=eng_out,
                rus_output_folder=rus_out,
                overwrite=False,
                progress=progress_cb,
            )
            progress.setValue(100)

            self._last_eng = eng_out
            self._last_rus = rus_out
            self.btn_open_folders.setEnabled(True)
            QMessageBox.information(self, "Успех", f"Файлы сохранены.\nЛог: {log_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))
            raise
        finally:
            self._detach_file_logger()

    def _open_result_folders(self):
        try:
            if self._last_eng and os.path.isdir(self._last_eng):
                os.startfile(self._last_eng)
            if self._last_rus and os.path.isdir(self._last_rus):
                os.startfile(self._last_rus)
        except Exception:
            # На всякий — fallback показать путь
            paths = "\n".join(p for p in (self._last_eng, self._last_rus) if p)
            if paths:
                QMessageBox.information(self, "Папки результатов", paths)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TXT/SRT ⇄ DOCX")
        tabs = QTabWidget()
        tabs.addTab(ExportTab(), "Экспорт в Word")
        tabs.addTab(ImportTab(), "Импорт из Word")
        self.setCentralWidget(tabs)
        self.resize(980, 520)


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
    )
    app = QApplication(sys.argv)
    apply_modern_style(app)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
