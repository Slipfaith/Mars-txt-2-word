import sys
import logging
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QTabWidget,
    QLineEdit,
    QMessageBox,
    QLabel,
)

from dragdrop import DragDropField

from converter import export_to_word, export_paths_to_word, import_from_word

logger = logging.getLogger(__name__)


class ExportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()

        hbox_eng = QHBoxLayout()
        self.eng_folder_edit = DragDropField(mode="files_or_folder")
        hbox_eng.addWidget(QLabel("Английские файлы или папка:"))
        hbox_eng.addWidget(self.eng_folder_edit)
        layout.addLayout(hbox_eng)

        hbox_rus = QHBoxLayout()
        self.rus_folder_edit = DragDropField(mode="files_or_folder")
        hbox_rus.addWidget(QLabel("Русские файлы или папка:"))
        hbox_rus.addWidget(self.rus_folder_edit)
        layout.addLayout(hbox_rus)

        hbox_encoding = QHBoxLayout()
        encoding_label = QLabel("Кодировка русских файлов (пусто = авто):")
        self.encoding_edit = QLineEdit()
        hbox_encoding.addWidget(encoding_label)
        hbox_encoding.addWidget(self.encoding_edit)
        layout.addLayout(hbox_encoding)

        btn_export = QPushButton("Создать Word документ")
        btn_export.clicked.connect(self.do_export)
        layout.addWidget(btn_export)

        self.setLayout(layout)

    def do_export(self):
        eng_paths = self.eng_folder_edit.paths
        rus_paths = self.rus_folder_edit.paths
        if not eng_paths or not rus_paths:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Сначала выберите файлы или папки для обоих языков",
            )
            return
        rus_enc = self.encoding_edit.text().strip() or None
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Word документ", "", "Word файлы (*.docx)"
        )
        if not output_path:
            return
        try:
            if (
                len(eng_paths) == 1
                and os.path.isdir(eng_paths[0])
                and len(rus_paths) == 1
                and os.path.isdir(rus_paths[0])
            ):
                export_to_word(
                    eng_paths[0],
                    rus_paths[0],
                    output_path,
                    file_extension=None,
                    rus_force_encoding=rus_enc,
                )
            else:
                def expand(paths):
                    out = []
                    for p in paths:
                        if os.path.isdir(p):
                            for name in os.listdir(p):
                                if name.lower().endswith((".txt", ".srt")):
                                    out.append(os.path.join(p, name))
                        else:
                            out.append(p)
                    return out

                eng_files = expand(eng_paths)
                rus_files = expand(rus_paths)
                export_paths_to_word(
                    sorted(eng_files),
                    sorted(rus_files),
                    output_path,
                    rus_force_encoding=rus_enc,
                )
            QMessageBox.information(self, "Успех", "Word документ успешно создан!")
        except Exception as exc:
            logger.exception("Export failed")
            QMessageBox.critical(self, "Ошибка", str(exc))


class ImportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()

        hbox_word = QHBoxLayout()
        self.word_file_edit = DragDropField(mode="file")
        hbox_word.addWidget(QLabel("Word документ:"))
        hbox_word.addWidget(self.word_file_edit)
        layout.addLayout(hbox_word)

        hbox_eng = QHBoxLayout()
        self.eng_output_edit = DragDropField(mode="folder")
        hbox_eng.addWidget(QLabel("Папка для английских файлов:"))
        hbox_eng.addWidget(self.eng_output_edit)
        layout.addLayout(hbox_eng)

        hbox_rus = QHBoxLayout()
        self.rus_output_edit = DragDropField(mode="folder")
        hbox_rus.addWidget(QLabel("Папка для русских файлов:"))
        hbox_rus.addWidget(self.rus_output_edit)
        layout.addLayout(hbox_rus)

        btn_import = QPushButton("Разбить Word документ")
        btn_import.clicked.connect(self.do_import)
        layout.addWidget(btn_import)

        self.setLayout(layout)

    def do_import(self):
        word_file = self.word_file_edit.text().strip()
        eng_output = self.eng_output_edit.text().strip()
        rus_output = self.rus_output_edit.text().strip()
        if not word_file or not eng_output or not rus_output:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Сначала выберите Word документ и папки для сохранения файлов",
            )
            return
        try:
            import_from_word(word_file, eng_output, rus_output)
            QMessageBox.information(self, "Успех", "Файлы успешно сохранены!")
        except Exception as exc:
            logger.exception("Import failed")
            QMessageBox.critical(self, "Ошибка", str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Конвертер субтитров: TXT/SRT <-> Word")
        self.resize(600, 300)
        tabs = QTabWidget()
        tabs.addTab(ExportTab(), "Создать Word документ")
        tabs.addTab(ImportTab(), "Разбить Word документ")
        self.setCentralWidget(tabs)


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
    )
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
