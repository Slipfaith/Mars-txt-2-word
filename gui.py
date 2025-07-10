import sys
import logging
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QTabWidget, QLineEdit, QMessageBox, QLabel
)

from converter import export_to_word, import_from_word

logger = logging.getLogger(__name__)


class ExportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.eng_folder = ""
        self.rus_folder = ""
        layout = QVBoxLayout()

        hbox_eng = QHBoxLayout()
        self.eng_folder_edit = QLineEdit()
        btn_select_eng = QPushButton("Выбрать папку с английскими файлами")
        btn_select_eng.clicked.connect(self.select_eng_folder)
        hbox_eng.addWidget(self.eng_folder_edit)
        hbox_eng.addWidget(btn_select_eng)
        layout.addLayout(hbox_eng)

        hbox_rus = QHBoxLayout()
        self.rus_folder_edit = QLineEdit()
        btn_select_rus = QPushButton("Выбрать папку с русскими файлами")
        btn_select_rus.clicked.connect(self.select_rus_folder)
        hbox_rus.addWidget(self.rus_folder_edit)
        hbox_rus.addWidget(btn_select_rus)
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

    def select_eng_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с английскими файлами")
        if folder:
            self.eng_folder = folder
            self.eng_folder_edit.setText(folder)

    def select_rus_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с русскими файлами")
        if folder:
            self.rus_folder = folder
            self.rus_folder_edit.setText(folder)

    def do_export(self):
        if not self.eng_folder or not self.rus_folder:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите обе папки с файлами")
            return
        rus_enc = self.encoding_edit.text().strip() or None
        output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить Word документ", "", "Word файлы (*.docx)")
        if not output_path:
            return
        try:
            export_to_word(self.eng_folder, self.rus_folder, output_path, rus_force_encoding=rus_enc)
            QMessageBox.information(self, "Успех", "Word документ успешно создан!")
        except Exception as exc:
            logger.exception("Export failed")
            QMessageBox.critical(self, "Ошибка", str(exc))


class ImportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.word_file = ""
        self.eng_output_folder = ""
        self.rus_output_folder = ""
        layout = QVBoxLayout()

        hbox_word = QHBoxLayout()
        self.word_file_edit = QLineEdit()
        btn_select_word = QPushButton("Выбрать Word документ")
        btn_select_word.clicked.connect(self.select_word_file)
        hbox_word.addWidget(self.word_file_edit)
        hbox_word.addWidget(btn_select_word)
        layout.addLayout(hbox_word)

        hbox_eng = QHBoxLayout()
        self.eng_output_edit = QLineEdit()
        btn_select_eng_out = QPushButton("Папка для английских файлов")
        btn_select_eng_out.clicked.connect(self.select_eng_output)
        hbox_eng.addWidget(self.eng_output_edit)
        hbox_eng.addWidget(btn_select_eng_out)
        layout.addLayout(hbox_eng)

        hbox_rus = QHBoxLayout()
        self.rus_output_edit = QLineEdit()
        btn_select_rus_out = QPushButton("Папка для русских файлов")
        btn_select_rus_out.clicked.connect(self.select_rus_output)
        hbox_rus.addWidget(self.rus_output_edit)
        hbox_rus.addWidget(btn_select_rus_out)
        layout.addLayout(hbox_rus)

        btn_import = QPushButton("Разбить Word документ")
        btn_import.clicked.connect(self.do_import)
        layout.addWidget(btn_import)

        self.setLayout(layout)

    def select_word_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Выберите Word документ", "", "Word файлы (*.docx)")
        if file:
            self.word_file = file
            self.word_file_edit.setText(file)

    def select_eng_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для английских файлов")
        if folder:
            self.eng_output_folder = folder
            self.eng_output_edit.setText(folder)

    def select_rus_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для русских файлов")
        if folder:
            self.rus_output_folder = folder
            self.rus_output_edit.setText(folder)

    def do_import(self):
        if not self.word_file or not self.eng_output_folder or not self.rus_output_folder:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите Word документ и папки для сохранения файлов")
            return
        try:
            import_from_word(self.word_file, self.eng_output_folder, self.rus_output_folder)
            QMessageBox.information(self, "Успех", "Файлы успешно сохранены!")
        except Exception as exc:
            logger.exception("Import failed")
            QMessageBox.critical(self, "Ошибка", str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Конвертер субтитров: TXT <-> Word")
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
